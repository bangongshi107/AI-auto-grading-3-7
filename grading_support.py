"""阅卷流程的通用支撑：停止原因、异常体系、分数处理、错误分类与重试。"""
import re
import time
import random
from decimal import Decimal, ROUND_HALF_UP
from enum import Enum
from functools import wraps
from typing import Optional, Callable, Tuple


# ==================== 停止原因分类枚举 ====================

class StopReason(Enum):
    """阅卷停止原因分类
    
    用于统一管理所有导致阅卷停止的原因，便于：
    1. UI层根据不同原因显示不同的提示和建议
    2. 日志系统分类统计停止原因
    3. 决定是否可以自动恢复/重试
    """
    # 正常完成
    COMPLETED = "completed"                    # 正常完成所有阅卷
    
    # 用户主动操作
    USER_STOPPED = "user_stopped"              # 用户手动点击停止按钮
    
    # 需要人工介入（AI判断）
    MANUAL_INTERVENTION = "manual_intervention"  # AI判断需要人工介入（如无法识别答案）
    ZERO_SCORE_STREAK = "zero_score_streak"    # 连续多份卷子全部题目均为0分，疑似异常试卷/AI持续误判
    STUCK_PAGE = "stuck_page"                  # 卡页：连续多轮截图与上一轮高度相似，疑似页面未刷新
    SCREENSHOT_MISMATCH = "screenshot_mismatch"  # 写入分数前二次核验发现页面内容已变化，为避免错评已停止
    THRESHOLD_EXCEEDED = "threshold_exceeded"  # 双评分差超过阈值
    
    # 技术错误（可能可重试）
    NETWORK_ERROR = "network_error"            # 网络错误（超时、连接失败等）
    API_ERROR = "api_error"                    # API错误（两个API都失败）
    
    # 配置/资源错误（需要修改配置）
    CONFIG_ERROR = "config_error"              # 配置错误（缺少必要配置）
    RESOURCE_ERROR = "resource_error"          # 资源错误（文件读写、截图失败等）
    
    # 业务逻辑错误
    SCORE_PARSE_ERROR = "score_parse_error"    # 分数解析错误
    
    # 未知错误
    UNKNOWN_ERROR = "unknown_error"            # 未知错误
    
    @property
    def is_recoverable(self) -> bool:
        """判断该停止原因是否可能通过重试恢复"""
        return self in (
            StopReason.NETWORK_ERROR,
            StopReason.API_ERROR,
        )
    
    @property
    def needs_config_fix(self) -> bool:
        """判断是否需要用户修改配置才能继续"""
        return self in (
            StopReason.CONFIG_ERROR,
            StopReason.RESOURCE_ERROR,
        )
    
    @property
    def needs_manual_review(self) -> bool:
        """判断是否需要人工审核当前试卷"""
        return self in (
            StopReason.MANUAL_INTERVENTION,
            StopReason.ZERO_SCORE_STREAK,
            StopReason.STUCK_PAGE,
            StopReason.SCREENSHOT_MISMATCH,
            StopReason.THRESHOLD_EXCEEDED,
        )
    
    @property
    def user_friendly_name(self) -> str:
        """返回用户友好的停止原因名称"""
        names = {
            StopReason.COMPLETED: "阅卷完成",
            StopReason.USER_STOPPED: "用户停止",
            StopReason.MANUAL_INTERVENTION: "需人工介入",
            StopReason.ZERO_SCORE_STREAK: "连续多份0分",
            StopReason.STUCK_PAGE: "卡页未刷新",
            StopReason.SCREENSHOT_MISMATCH: "评分对象与页面不匹配",
            StopReason.THRESHOLD_EXCEEDED: "双评分差过大",
            StopReason.NETWORK_ERROR: "网络错误",
            StopReason.API_ERROR: "AI接口错误",
            StopReason.CONFIG_ERROR: "配置错误",
            StopReason.RESOURCE_ERROR: "资源错误",
            StopReason.SCORE_PARSE_ERROR: "分数解析错误",
            StopReason.UNKNOWN_ERROR: "未知错误",
        }
        return names.get(self, "未知")


# ==================== 面向老师的错误提示分类 ====================
# main.py / ui_components/main_window.py / api_service.py 三处都需要把底层错误
# 翻译成老师能看懂的中文提示，曾经各自维护一份关键词表，容易互相漂移。
# 现在统一由 classify_teacher_facing_error() 判定类别，各处只按类别挑选自己的措辞。

class TeacherErrorCategory(Enum):
    """老师可见错误提示的分类依据"""
    EMPTY = "empty"
    TIMEOUT = "timeout"
    AUTH_401 = "auth_401"
    QUOTA_403 = "quota_403"
    RATE_LIMIT_429 = "rate_limit_429"
    SERVICE_5XX = "service_5xx"
    FILE_PERMISSION = "file_permission"
    OTHER = "other"


def classify_teacher_facing_error(text: Optional[str]) -> TeacherErrorCategory:
    """将原始错误文本分类，供UI层据此挑选对应的中文提示与建议。

    关键词取三处历史实现的并集，只会让分类更宽松，不会丢失原有匹配。
    """
    s = (text or "").strip()
    if not s:
        return TeacherErrorCategory.EMPTY

    low = s.lower()
    if any(k in low for k in ["timed out", "timeout"]):
        return TeacherErrorCategory.TIMEOUT
    if any(k in low for k in ["401", "unauthorized", "invalid api key", "api key"]):
        return TeacherErrorCategory.AUTH_401
    if any(k in low for k in ["403", "forbidden", "quota", "余额", "payment", "insufficient"]):
        return TeacherErrorCategory.QUOTA_403
    if any(k in low for k in ["429", "rate limit", "too many", "请求太频繁"]):
        return TeacherErrorCategory.RATE_LIMIT_429
    if any(k in low for k in ["502", "503", "504", "service unavailable", "bad gateway"]):
        return TeacherErrorCategory.SERVICE_5XX
    if any(k in low for k in ["permission", "permissionerror", "access is denied", "被占用", "正在使用"]):
        return TeacherErrorCategory.FILE_PERMISSION
    return TeacherErrorCategory.OTHER


# ==================== 自定义异常层次结构 ====================

class GradingError(Exception):
    """阅卷系统基础异常类
    
    所有自定义异常的基类，提供统一的错误信息格式和恢复建议。
    """
    
    def __init__(self, message: str, recoverable: bool = False, 
                 recovery_action: str = "", original_error: Optional[Exception] = None):
        """
        Args:
            message: 错误描述信息
            recoverable: 是否可自动恢复
            recovery_action: 建议的恢复操作
            original_error: 原始异常（用于异常链）
        """
        super().__init__(message)
        self.message = message
        self.recoverable = recoverable
        self.recovery_action = recovery_action
        self.original_error = original_error
    
    def __str__(self):
        base = self.message
        if self.recovery_action:
            base += f" [建议操作: {self.recovery_action}]"
        return base


class ConfigError(GradingError):
    """配置相关错误
    
    包括：配置文件缺失/格式错误、必需参数未设置、参数值无效等。
    通常需要用户修改配置后重试。
    """
    
    def __init__(self, message: str, config_key: str = "", 
                 expected_type: str = "", original_error: Optional[Exception] = None):
        recovery = "请检查配置文件或在设置界面修正配置"
        if config_key:
            recovery = f"请检查配置项 '{config_key}'"
            if expected_type:
                recovery += f"，期望类型: {expected_type}"
        super().__init__(message, recoverable=False, 
                        recovery_action=recovery, original_error=original_error)
        self.config_key = config_key
        self.expected_type = expected_type


class NetworkError(GradingError):
    """网络相关错误
    
    包括：连接超时、网络不可达、API服务不可用、限流等。
    通常可以通过重试恢复。
    """
    
    # 网络错误子类型
    TYPE_TIMEOUT = "timeout"           # 连接/读取超时
    TYPE_CONNECTION = "connection"     # 连接失败
    TYPE_RATE_LIMIT = "rate_limit"     # API限流（429）
    TYPE_SERVICE_DOWN = "service_down" # 服务不可用（503）
    TYPE_SERVER_ERROR = "server_error" # 服务器内部错误（5xx）
    
    def __init__(self, message: str, error_type: str = "", 
                 retry_after: int = 0, original_error: Optional[Exception] = None):
        # 根据错误类型设置恢复建议
        recovery_map = {
            self.TYPE_TIMEOUT: "请检查网络连接，稍后重试",
            self.TYPE_CONNECTION: "请检查网络连接和API地址配置",
            self.TYPE_RATE_LIMIT: f"API请求过于频繁，请等待{retry_after}秒后重试" if retry_after else "API请求过于频繁，请稍后重试",
            self.TYPE_SERVICE_DOWN: "API服务暂时不可用，请稍后重试",
            self.TYPE_SERVER_ERROR: "API服务器错误，请稍后重试",
        }
        recovery = recovery_map.get(error_type, "请检查网络连接后重试")
        
        # 网络错误通常可重试
        super().__init__(message, recoverable=True, 
                        recovery_action=recovery, original_error=original_error)
        self.error_type = error_type
        self.retry_after = retry_after


class BusinessError(GradingError):
    """业务逻辑错误
    
    包括：评分解析失败、分数超出范围、答案区域无效等。
    根据具体情况可能需要人工介入或可以自动恢复。
    """
    
    # 业务错误子类型
    TYPE_SCORE_PARSE = "score_parse"       # 分数解析失败
    TYPE_SCORE_RANGE = "score_range"       # 分数超出范围
    TYPE_AREA_INVALID = "area_invalid"     # 答案区域无效
    TYPE_API_RESPONSE = "api_response"     # API响应格式错误
    TYPE_DUAL_EVAL = "dual_eval"           # 双评分差超阈值
    
    def __init__(self, message: str, error_type: str = "", 
                 question_index: int = 0, recoverable: bool = False,
                 original_error: Optional[Exception] = None):
        # 根据错误类型设置恢复建议
        recovery_map = {
            self.TYPE_SCORE_PARSE: "AI返回的分数格式无效，请检查评分细则或手动评分",
            self.TYPE_SCORE_RANGE: "分数已自动修正到有效范围",
            self.TYPE_AREA_INVALID: "请重新配置答案区域",
            self.TYPE_API_RESPONSE: "API响应格式异常，可能需要更换模型",
            self.TYPE_DUAL_EVAL: "双评分差超过阈值，需要人工复核",
        }
        recovery = recovery_map.get(error_type, "请检查相关配置或手动处理")
        
        super().__init__(message, recoverable=recoverable, 
                        recovery_action=recovery, original_error=original_error)
        self.error_type = error_type
        self.question_index = question_index


class ResourceError(GradingError):
    """资源相关错误
    
    包括：文件读写失败、内存不足、截图失败等系统资源问题。
    """
    
    TYPE_FILE_IO = "file_io"           # 文件读写错误
    TYPE_SCREENSHOT = "screenshot"     # 截图失败
    TYPE_MEMORY = "memory"             # 内存不足
    
    def __init__(self, message: str, error_type: str = "",
                 resource_path: str = "", original_error: Optional[Exception] = None):
        recovery_map = {
            self.TYPE_FILE_IO: f"文件操作失败: {resource_path}" if resource_path else "文件操作失败，请检查权限",
            self.TYPE_SCREENSHOT: "截图失败，请检查屏幕访问权限",
            self.TYPE_MEMORY: "内存不足，请关闭其他程序后重试",
        }
        recovery = recovery_map.get(error_type, "请检查系统资源")
        
        super().__init__(message, recoverable=False,
                        recovery_action=recovery, original_error=original_error)
        self.error_type = error_type
        self.resource_path = resource_path


# ==================== 异常恢复策略管理器 ====================

class ErrorRecoveryManager:
    """异常恢复策略管理器
    
    根据不同类型的异常提供相应的恢复策略和建议。
    """
    
    @staticmethod
    def classify_exception(error: Exception) -> GradingError:
        """将标准异常转换为自定义异常类型
        
        Args:
            error: 原始异常
            
        Returns:
            对应的GradingError子类实例
        """
        # 网络类错误复用统一的关键词分类，避免两套规则并存
        network_types = {
            'timeout': NetworkError.TYPE_TIMEOUT,
            'network': NetworkError.TYPE_CONNECTION,
            'rate_limit': NetworkError.TYPE_RATE_LIMIT,
            'service_unavailable': NetworkError.TYPE_SERVICE_DOWN,
            'server_error': NetworkError.TYPE_SERVER_ERROR,
        }
        error_type, _ = extract_error_type_and_classify(error)
        if error_type in network_types:
            return NetworkError(str(error), network_types[error_type], original_error=error)

        error_str = str(error).lower()
        # 检测配置相关错误
        if isinstance(error, KeyError):
            return ConfigError(f"配置字段缺失: {error}", config_key=str(error), original_error=error)
        
        if isinstance(error, ValueError):
            # 尝试区分配置错误和业务错误
            if any(kw in error_str for kw in ['config', '配置', 'parameter', '参数']):
                return ConfigError(str(error), original_error=error)
            else:
                return BusinessError(str(error), BusinessError.TYPE_SCORE_PARSE, original_error=error)
        
        # 检测资源相关错误
        if isinstance(error, (IOError, OSError, FileNotFoundError, PermissionError)):
            return ResourceError(str(error), ResourceError.TYPE_FILE_IO, original_error=error)
        
        if isinstance(error, MemoryError):
            return ResourceError(str(error), ResourceError.TYPE_MEMORY, original_error=error)
        
        # 默认作为业务错误
        return BusinessError(str(error), original_error=error)
    
    @staticmethod
    def get_recovery_strategy(error: GradingError) -> dict:
        """获取错误恢复策略
        
        Args:
            error: GradingError实例
            
        Returns:
            恢复策略字典，包含:
            - should_retry: 是否应该重试
            - retry_delay: 重试延迟（秒）
            - max_retries: 最大重试次数
            - should_stop: 是否应该停止整个流程
            - notify_user: 是否需要通知用户
            - log_level: 日志级别
        """
        strategy = {
            'should_retry': False,
            'retry_delay': 1.0,
            'max_retries': 3,
            'should_stop': True,
            'notify_user': True,
            'log_level': 'ERROR'
        }
        
        if isinstance(error, NetworkError):
            # 网络错误：通常可重试
            strategy['should_retry'] = True
            strategy['should_stop'] = False
            strategy['log_level'] = 'WARNING'
            
            if error.error_type == NetworkError.TYPE_RATE_LIMIT:
                strategy['retry_delay'] = max(error.retry_after, 5.0)
                strategy['max_retries'] = 5
            elif error.error_type == NetworkError.TYPE_TIMEOUT:
                strategy['retry_delay'] = 2.0
                strategy['max_retries'] = 3
            elif error.error_type == NetworkError.TYPE_SERVER_ERROR:
                strategy['retry_delay'] = 3.0
                strategy['max_retries'] = 2
        
        elif isinstance(error, ConfigError):
            # 配置错误：需要停止并通知用户
            strategy['should_retry'] = False
            strategy['should_stop'] = True
            strategy['notify_user'] = True
            strategy['log_level'] = 'ERROR'
        
        elif isinstance(error, BusinessError):
            # 业务错误：根据子类型决定
            if error.error_type == BusinessError.TYPE_SCORE_RANGE:
                # 分数范围错误：已自动修正，可继续
                strategy['should_retry'] = False
                strategy['should_stop'] = False
                strategy['notify_user'] = False
                strategy['log_level'] = 'WARNING'
            elif error.error_type == BusinessError.TYPE_DUAL_EVAL:
                # 双评差异：需要人工介入
                strategy['should_stop'] = True
                strategy['notify_user'] = True
            else:
                # 其他业务错误：停止当前题目
                strategy['should_stop'] = True
                strategy['notify_user'] = True
        
        elif isinstance(error, ResourceError):
            # 资源错误：通常需要停止
            strategy['should_retry'] = False
            strategy['should_stop'] = True
            strategy['notify_user'] = True
            strategy['log_level'] = 'ERROR'
        
        return strategy
    
    @staticmethod
    def format_error_message(error: GradingError, include_recovery: bool = True) -> str:
        """格式化错误消息
        
        Args:
            error: GradingError实例
            include_recovery: 是否包含恢复建议
            
        Returns:
            格式化的错误消息
        """
        # 确定错误类型前缀
        type_prefix = {
            ConfigError: "[配置错误]",
            NetworkError: "[网络错误]",
            BusinessError: "[业务错误]",
            ResourceError: "[资源错误]",
            GradingError: "[系统错误]"
        }
        
        prefix = "[错误]"
        for err_type, pref in type_prefix.items():
            if isinstance(error, err_type):
                prefix = pref
                break
        
        message = f"{prefix} {error.message}"
        
        if include_recovery and error.recovery_action:
            message += f"\n  → 建议: {error.recovery_action}"
        
        return message


# ==================== 分数处理管道类 ====================

class ScoreProcessor:
    """
    统一的分数处理管道类，负责分数的清洗→校验→四舍五入→范围限制。
    确保所有分数处理逻辑集中在一个地方，避免边界情况漏处理。
    """
    
    @staticmethod
    def sanitize(val) -> float:
        """
        清洗和标准化分数输入，确保返回有效的浮点数。
        如果无法提取有效数字，抛出 ValueError 以确保评分准确性。
        
        Args:
            val: 待清洗的分数值（可以是数字、字符串等）
            
        Returns:
            清洗后的浮点数
            
        Raises:
            ValueError: 无法转换为有效分数时
        """
        if isinstance(val, (int, float)):
            return float(val)
        
        # 尝试从字符串中提取数字
        try:
            # 提取浮点数（包括负数）
            match = re.search(r'-?\d+\.?\d*', str(val))
            if match:
                return float(match.group())
        except Exception:
            pass
        
        raise ValueError(f"无法将 {val} 转换为有效的分数")
    
    @staticmethod
    def round_to_step(value: float, step: float) -> float:
        """
        将数值四舍五入到指定步长的倍数。
        
        Args:
            value: 要四舍五入的数值
            step: 步长（如0.5或1）
        
        Returns:
            四舍五入后的值
            
        Examples:
            round_to_step(7.3, 0.5) -> 7.5
            round_to_step(7.3, 1.0) -> 7.0
            round_to_step(7.8, 0.5) -> 8.0
        """
        if step <= 0:
            return value
        try:
            value_dec = Decimal(str(value))
            step_dec = Decimal(str(step))
            if step_dec == 0:
                return value
            scaled = value_dec / step_dec
            rounded = scaled.quantize(Decimal('0'), rounding=ROUND_HALF_UP)
            return float(rounded * step_dec)
        except Exception:
            return round(value / step) * step
    
    @staticmethod
    def validate_range(score: float, min_score: float, max_score: float, 
                      logger: Optional[Callable] = None) -> float:
        """
        验证分数是否在有效范围内，超出则修正并记录日志。
        
        Args:
            score: 待验证的分数
            min_score: 最低分
            max_score: 最高分
            logger: 可选的日志记录函数，签名为 logger(message, is_error, level)
            
        Returns:
            修正后的分数
        """
        if score < min_score:
            if logger:
                logger(f"分数 {score} 低于最低分 {min_score}，修正为 {min_score}。", True, "ERROR")
            return min_score
        elif score > max_score:
            if logger:
                logger(f"分数 {score} 超出最高分 {max_score}，修正为 {max_score}。", True, "ERROR")
            return max_score
        return score
    
    @classmethod
    def process_pipeline(cls, raw_score, min_score: float, max_score: float, 
                        rounding_step: float = 0.5,
                        logger: Optional[Callable] = None) -> Tuple[float, str]:
        """
        完整的分数处理管道：清洗→四舍五入→范围校验。
        
        Args:
            raw_score: 原始分数（任意类型）
            min_score: 最低分
            max_score: 最高分
            rounding_step: 四舍五入步长（默认0.5）
            logger: 可选的日志记录函数
            
        Returns:
            (处理后的最终分数, 处理过程描述)
            
        Raises:
            ValueError: 无法清洗分数时
        """
        steps_log = []
        
        # 步骤1: 清洗分数
        try:
            sanitized = cls.sanitize(raw_score)
            steps_log.append(f"清洗: {raw_score} → {sanitized}")
        except ValueError as e:
            raise ValueError(f"分数清洗失败: {e}")
        
        # 步骤2: 四舍五入到步长
        rounded = cls.round_to_step(sanitized, rounding_step)
        if rounded != sanitized:
            steps_log.append(f"四舍五入(步长{rounding_step}): {sanitized} → {rounded}")
        
        # 步骤3: 范围校验和修正
        validated = cls.validate_range(rounded, min_score, max_score, logger)
        if validated != rounded:
            steps_log.append(f"范围修正: {rounded} → {validated}")
        
        process_desc = " | ".join(steps_log) if steps_log else f"无需处理: {validated}"
        return validated, process_desc
    
    @classmethod
    def process_itemized_scores(cls, itemized_scores_list, 
                                min_score: float, max_score: float,
                                rounding_step: float = 0.5,
                                logger: Optional[Callable] = None) -> Tuple[list, float]:
        """
        处理分项得分列表，返回清洗后的分数列表和总分。
        
        Args:
            itemized_scores_list: 分项得分列表（可能包含字符串等）
            min_score: 单项最低分
            max_score: 单项最高分（用于单项校验，总分可能超出）
            rounding_step: 四舍五入步长
            logger: 可选的日志记录函数
            
        Returns:
            (清洗后的分数列表, 计算的总分)
            
        Raises:
            ValueError: 任何分项无法清洗时
        """
        cleaned_scores = []
        for idx, score in enumerate(itemized_scores_list):
            try:
                cleaned = cls.sanitize(score)
                cleaned_scores.append(cleaned)
            except ValueError as e:
                raise ValueError(f"分项得分[{idx}] 清洗失败: {e}")
        
        total = sum(cleaned_scores)
        return cleaned_scores, total


# ==================== 统一重试机制 ====================

class ErrorRetryability(Enum):
    """错误的可重试性分级（优先级从高到低）"""
    DEFINITELY_RETRYABLE = 1    # 明确可重试：网络超时、429限流、服务暂时不可用
    POSSIBLY_RETRYABLE = 2      # 可能可重试：Token过期、偶发5xx错误
    NOT_WORTH_RETRYING = 3      # 不值得重试：JSON格式错误、业务逻辑错误
    MANUAL_INTERVENTION = 4     # 需要人工介入：权限问题、功能缺陷


def extract_error_type_and_classify(error: Exception) -> Tuple[str, ErrorRetryability]:
    """提取错误类型并分类其可重试性
    
    Returns:
        (错误类型名称, 可重试性级别)
    """
    s = str(error).lower()

    def has_code(*codes: str) -> bool:
        # 状态码必须是独立数字，避免把“第500字”之类误判为 HTTP 500
        return re.search(r'(?<!\w)(?:' + '|'.join(codes) + r')(?!\w)', s) is not None
    
    # 1. 明确可重试的错误
    if 'timeout' in s or '超时' in s or 'timed out' in s:
        return ('timeout', ErrorRetryability.DEFINITELY_RETRYABLE)
    
    if has_code('429') or 'rate limit' in s or '限流' in s or 'too many requests' in s:
        return ('rate_limit', ErrorRetryability.DEFINITELY_RETRYABLE)
    
    if 'connection' in s or '连接' in s or 'network' in s or '网络' in s:
        return ('network', ErrorRetryability.DEFINITELY_RETRYABLE)
    
    if has_code('503') or 'service unavailable' in s or '服务不可用' in s:
        return ('service_unavailable', ErrorRetryability.DEFINITELY_RETRYABLE)
    
    # 2. 可能可重试的错误
    if 'token' in s or 'access_token' in s:
        # Token问题可能是过期，可以尝试刷新
        return ('token', ErrorRetryability.POSSIBLY_RETRYABLE)
    
    if has_code('500', '502', '504') or 'internal server error' in s:
        # 偶发的服务器错误可能恢复
        return ('server_error', ErrorRetryability.POSSIBLY_RETRYABLE)
    
    # 3. 不值得重试的错误
    if 'json' in s or '格式' in s or 'parse' in s or '解析' in s:
        return ('json_parse', ErrorRetryability.NOT_WORTH_RETRYING)
    
    if has_code('400') or 'bad request' in s or '请求错误' in s:
        return ('bad_request', ErrorRetryability.NOT_WORTH_RETRYING)
    
    if has_code('404') or 'not found' in s:
        return ('not_found', ErrorRetryability.NOT_WORTH_RETRYING)
    
    if 'invalid' in s or '无效' in s or '非法' in s:
        return ('invalid_input', ErrorRetryability.NOT_WORTH_RETRYING)
    
    # 4. 需要人工介入的错误
    if has_code('401', '403') or 'unauthorized' in s or 'forbidden' in s or '权限' in s or '认证失败' in s:
        # 权限问题通常需要修改配置
        return ('permission', ErrorRetryability.MANUAL_INTERVENTION)
    
    if 'not implemented' in s or '未实现' in s or 'unsupported' in s:
        return ('not_implemented', ErrorRetryability.MANUAL_INTERVENTION)
    
    # 默认：未知错误，可能可重试
    return ('unknown', ErrorRetryability.POSSIBLY_RETRYABLE)


def calculate_smart_retry_delay(attempt: int, error_type: str, base_delay: float = 1.0) -> float:
    """根据错误类型和重试次数智能计算延迟时间（指数退避+错误感知）
    
    Args:
        attempt: 第几次重试（从1开始）
        error_type: 错误类型名称
        base_delay: 基础延迟时间（秒）
    
    Returns:
        延迟时间（秒）
    """
    # 不同错误类型的基础延迟倍数
    error_base_multipliers = {
        'rate_limit': 3.0,          # 限流：延迟长一些
        'timeout': 1.5,             # 超时：中等延迟
        'network': 1.0,             # 网络：正常延迟
        'token': 2.0,               # Token：稍长延迟（给时间刷新）
        'server_error': 2.0,        # 服务器错误：稍长延迟
        'service_unavailable': 2.5, # 服务不可用：较长延迟
    }
    
    multiplier = error_base_multipliers.get(error_type, 1.0)
    
    # 指数退避：第1次重试 = 基础延迟，第2次 = 2倍，第3次 = 4倍...
    exponential_factor = 2 ** (attempt - 1)
    
    # 添加随机抖动（±20%），避免多个请求同时重试
    jitter = random.uniform(0.8, 1.2)
    
    delay = base_delay * multiplier * exponential_factor * jitter
    
    # 设置最大延迟上限（避免等待太久）
    max_delay = 10.0
    return min(delay, max_delay)
def unified_retry(
    max_retries: int = 1,
    transient_error_checker: Optional[Callable[[Exception], bool]] = None,
    retry_delay: float = 1.0,
    log_callback: Optional[Callable[[str, bool, str], None]] = None,
    operation_name: str = "操作"
):
    """
    统一重试装饰器：对短暂性错误（网络超时、限流等）最多重试 max_retries 次，
    对业务/配置错误立即失败。延迟采用指数退避+错误感知策略。
    """
    def decorator(func: Callable) -> Callable:
        @wraps(func)
        def wrapper(*args, **kwargs):
            last_exception = None
            last_error_type = 'unknown'
            last_retryability = ErrorRetryability.POSSIBLY_RETRYABLE
            
            for attempt in range(max_retries + 1):  # +1 因为包含首次尝试
                try:
                    if attempt > 0:
                        # 计算智能延迟（指数退避+错误感知）
                        smart_delay = calculate_smart_retry_delay(
                            attempt=attempt,
                            error_type=last_error_type,
                            base_delay=retry_delay
                        )
                        
                        if log_callback:
                            # 显示更详细的重试信息
                            log_callback(
                                f"{operation_name}第{attempt}次重试（错误类型:{last_error_type}, 延迟{smart_delay:.1f}秒）...",
                                False, "DETAIL"
                            )
                        
                        time.sleep(smart_delay)
                    
                    # 执行实际操作
                    return func(*args, **kwargs)
                    
                except Exception as e:
                    last_exception = e
                    
                    # 提取错误类型并分类
                    error_type, retryability = extract_error_type_and_classify(e)
                    last_error_type = error_type
                    last_retryability = retryability
                    
                    # 判断是否应该重试（使用精细分类）
                    should_retry = False
                    
                    if retryability == ErrorRetryability.DEFINITELY_RETRYABLE:
                        # 明确可重试
                        should_retry = True
                    elif retryability == ErrorRetryability.POSSIBLY_RETRYABLE:
                        # 可能可重试，使用旧的检查器兼容
                        if transient_error_checker:
                            try:
                                should_retry = transient_error_checker(e)
                            except Exception:
                                should_retry = True  # 默认重试一次
                        else:
                            should_retry = True
                    elif retryability == ErrorRetryability.NOT_WORTH_RETRYING:
                        # 不值得重试（如JSON格式错误）
                        should_retry = False
                        if log_callback:
                            log_callback(
                                f"{operation_name}失败（{error_type}错误不值得重试）: {str(e)}",
                                True, "ERROR"
                            )
                    else:  # MANUAL_INTERVENTION
                        # 需要人工介入（如权限问题）
                        should_retry = False
                        if log_callback:
                            log_callback(
                                f"{operation_name}失败（{error_type}错误需要人工介入）: {str(e)}",
                                True, "ERROR"
                            )
                    
                    # 根据判断决定是否重试
                    if not should_retry:
                        raise
                    
                    # 短暂性错误的处理
                    if attempt < max_retries:
                        # 还有重试机会
                        if log_callback:
                            log_callback(
                                f"{operation_name}尝试{attempt+1}/{max_retries+1}失败（{error_type}错误）: {str(e)[:100]}，将智能重试",
                                True, "WARNING"
                            )
                    else:
                        # 最后一次尝试也失败了
                        if log_callback:
                            log_callback(
                                f"{operation_name}失败（已重试{max_retries}次，{error_type}错误）: {str(e)}",
                                True, "ERROR"
                            )
                        raise
            
            # 理论上不会到这里，但为了安全
            if last_exception:
                raise last_exception
            
        return wrapper
    return decorator
