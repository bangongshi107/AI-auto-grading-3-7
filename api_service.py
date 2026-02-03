# api_service.py - AI评分API服务模块
# 支持多平台：火山引擎、阿里通义、百度千帆、腾讯混元、智谱、月之暗面、OpenRouter、OpenAI、Google Gemini

import requests
import logging
import traceback
from typing import Tuple, Optional, Dict, Any
import hashlib
import hmac
import time
import json
import sys
from datetime import datetime, timezone
from threading import Lock, local

# ==============================================================================
#  UI文本到提供商ID的映射字典 (UI Text to Provider ID Mapping)
#  这是连接UI显示文本和后台代码的桥梁。
#  UI上的"火山引擎 (豆包)" 对应到代码里的 "volcengine"。
#  现在基于 PROVIDER_CONFIGS 动态生成，避免数据冗余。
#  注意：只包含AI评分模型提供商
# ==============================================================================
def generate_ui_text_to_provider_id():
    """基于 PROVIDER_CONFIGS 动态生成 UI_TEXT_TO_PROVIDER_ID 映射"""
    return {
        config["name"]: provider_id 
        for provider_id, config in PROVIDER_CONFIGS.items()
    }

# ==============================================================================
#  权威供应商配置字典 (Authoritative Provider Configuration)
#  这是整个系统的"单一事实来源 (Single Source of Truth)"。
#
#  腾讯混元更新历史 (Tencent Hunyuan Update History):
#  - 2025-09-13: 重大更新 - 统一使用 ChatCompletions 接口
#    * 替换 ImageQuestion 为 ChatCompletions action (无频率限制)
#    * 实现腾讯云签名方法 v3 完整认证
#    * 支持所有视觉模型自动适配 (hunyuan-vision, hunyuan-turbos-vision 等)
#    * 智能检测视觉模型并自动选择正确的 payload 格式
#    * API Key 格式: SecretId:SecretKey
# ==============================================================================
PROVIDER_CONFIGS = {
    # 这里的 key ('volcengine', 'moonshot'等) 是程序内部使用的【内部标识】
    "volcengine": {
        "name": "火山引擎 (推荐)",
        "url": "https://ark.cn-beijing.volces.com/api/v3/chat/completions",
        "auth_method": "bearer",
        "payload_builder": "_build_volcengine_payload",
    },
    "moonshot": {
        "name": "月之暗面",
        "url": "https://api.moonshot.cn/v1/chat/completions",
        "auth_method": "bearer",
        "payload_builder": "_build_openai_compatible_payload",
    },
    "zhipu": {
        "name": "智谱清言",
        "url": "https://open.bigmodel.cn/api/paas/v4/chat/completions",
        "auth_method": "bearer", # 智谱的Key虽然是JWT，但用法和Bearer完全一样
        "payload_builder": "_build_openai_compatible_payload",
    },
    # "deepseek": {
    #     "name": "deepseek",
    #     "url": "https://api.deepseek.com/chat/completions",
    #     "auth_method": "bearer",
    #     "payload_builder": "_build_openai_compatible_payload",
    # },
    "aliyun": {
        "name": "阿里通义千问",
        "url": "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions",
        "auth_method": "bearer",
        "payload_builder": "_build_openai_compatible_payload",
    },
    "baidu": {
        "name": "百度文心千帆",
        "url": "https://qianfan.baidubce.com/v2/chat/completions",
        "auth_method": "bearer",
        "payload_builder": "_build_openai_compatible_payload",
    },
    "tencent": {
        "name": "腾讯混元",
        "url": "https://hunyuan.tencentcloudapi.com/",
        "auth_method": "tencent_signature_v3", # 使用腾讯云签名方法 v3
        "payload_builder": "_build_tencent_payload",
        "service_info": {  # 新增服务信息配置，避免硬编码
            "service": "hunyuan",
            "region": "ap-guangzhou",
            "version": "2023-09-01",
            "host": "hunyuan.tencentcloudapi.com",
            "action": "ChatCompletions"
        }
    },
    "openrouter": {
        "name": "OpenRouter",
        "url": "https://openrouter.ai/api/v1/chat/completions",
        "auth_method": "bearer",
        "payload_builder": "_build_openai_compatible_payload",
    },
    "openai": { # 新增
        "name": "OpenAI",
        "url": "https://api.openai.com/v1/chat/completions",
        "auth_method": "bearer",
        "payload_builder": "_build_openai_compatible_payload",
    },
    "gemini": { # 新增
        "name": "Google Gemini",
        "url": "https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent",  # {model} 将被动态替换
        "auth_method": "google_api_key_in_url",
        "payload_builder": "_build_gemini_payload",
        "dynamic_url": True,  # 标记需要动态URL替换
    }
}

# ==============================================================================
#  生成UI文本到提供商ID的映射常量
# ==============================================================================
UI_TEXT_TO_PROVIDER_ID = generate_ui_text_to_provider_id()

# ==============================================================================
#  辅助函数，用于UI和内部ID之间的转换
# ==============================================================================
def get_provider_id_from_ui_text(ui_text: str) -> Optional[str]:
    mapping = generate_ui_text_to_provider_id()
    return mapping.get(ui_text.strip())

def get_ui_text_from_provider_id(provider_id: str) -> Optional[str]:
    config = PROVIDER_CONFIGS.get(provider_id)
    return config["name"] if config else None

class ApiService:
    def __init__(self, config_manager):
        self.config_manager = config_manager
        # requests.Session 不是严格线程安全；双评并发时会同时发起两次请求。
        # 使用 thread-local 的 Session，既保留连接复用，又避免跨线程共享 Session。
        self._thread_local = local()
        self.logger = logging.getLogger(__name__)
        # 初始化当前题目索引，虽然主要逻辑在AutoThread中，但这里有个默认值更安全
        self.current_question_index = 1

    def _should_print_raw_ai_response(self) -> bool:
        """是否在控制台打印AI接口的原始响应。

        这是给开发者调试用的：打印的是 HTTP 原始响应体（response.text），
        不会包含请求头里的 Authorization，因此不会主动泄露 API Key。
        """
        try:
            return bool(getattr(self.config_manager, "debug_print_raw_ai_response", False))
        except Exception:
            return False

    def _print_raw_ai_response(self, provider_name: str, url: str, status_code: int, raw_text: str) -> None:
        """把原始 AI 响应完整输出到控制台（stdout）。"""
        if not self._should_print_raw_ai_response():
            return
        try:
            # 使用明确的 begin/end 标记，便于在控制台/日志里搜索定位。
            ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            sys.stdout.write(
                f"\n========== AI RAW RESPONSE BEGIN [{provider_name}] {ts} status={status_code} =========="
                f"\nURL: {url}\n"
            )
            sys.stdout.write(raw_text or "")
            if raw_text and not raw_text.endswith("\n"):
                sys.stdout.write("\n")
            sys.stdout.write(f"========== AI RAW RESPONSE END [{provider_name}] ==========" + "\n")
            sys.stdout.flush()
        except Exception:
            # 绝不让调试输出影响业务流程
            pass

    def _get_session(self) -> requests.Session:
        """获取当前线程专属的 requests.Session。"""
        sess = getattr(self._thread_local, "session", None)
        if sess is None:
            sess = requests.Session()
            self._thread_local.session = sess
        return sess

    def reset(self):
        """重置 API 服务状态
        
        执行以下操作：
        1. 关闭当前线程的 requests.Session（释放连接）
        2. 用新的 threading.local() 替换 _thread_local（清空线程本地存储）
        3. 重置 current_question_index = 1
        
        用于在批量评卷过程中定期清理 HTTP 会话，避免长连接状态污染
        """
        sess = getattr(self._thread_local, "session", None)
        if sess is not None:
            try:
                sess.close()  # 关闭连接池
            except Exception:
                pass
        
        self._thread_local = local()  # 创建全新的线程本地存储
        self.current_question_index = 1

    # ==========================================================================
    #  腾讯云签名方法 v3 实现 (Tencent Cloud Signature Method v3)
    #
    #  更新历史 (Update History):
    #  - 2025-09-13: 首次实现完整的 TC3-HMAC-SHA256 签名流程
    #    * 实现规范请求字符串构建
    #    * 实现 HMAC-SHA256 多层签名计算
    #    * 支持动态时间戳和凭证范围
    #    * 自动生成 Authorization header
    #
    #  技术要点 (Technical Notes):
    #  - 使用 UTC 时间戳确保时区一致性
    #  - 签名顺序: SecretKey -> Date -> Service -> "tc3_request"
    #  - 支持的 Service: "hunyuan"
    #  - 支持的 Region: "ap-guangzhou" (默认)
    # ==========================================================================
    def _build_tencent_signature_v3(self, secret_id: str, secret_key: str, service: str, region: str,
                                   action: str, version: str, payload: str, host: str) -> Tuple[str, str]:
        """构建腾讯云 API 签名方法 v3

        Args:
            secret_id: 腾讯云 SecretId
            secret_key: 腾讯云 SecretKey
            service: 服务名称 (hunyuan)
            region: 地域 (ap-guangzhou)
            action: API 动作 (ChatCompletions)
            version: API 版本 (2023-09-01)
            payload: 请求 payload 的 JSON 字符串

        Returns:
            Tuple[str, str]: (authorization_header, timestamp)
        """

        # 1. 创建规范请求字符串
        algorithm = "TC3-HMAC-SHA256"
        timestamp = int(time.time())
        date = datetime.fromtimestamp(timestamp, tz=timezone.utc).strftime('%Y-%m-%d')  # 腾讯云签名要求 YYYY-MM-DD 格式

        # 规范请求
        canonical_request = self._build_canonical_request(action, payload, host)

        # 2. 创建待签字符串
        credential_scope = f"{date}/{service}/tc3_request"
        string_to_sign = f"{algorithm}\n{timestamp}\n{credential_scope}\n{hashlib.sha256(canonical_request.encode('utf-8')).hexdigest()}"

        # 3. 计算签名
        secret_date = hmac.new(f"TC3{secret_key}".encode('utf-8'), date.encode('utf-8'), hashlib.sha256).digest()
        secret_service = hmac.new(secret_date, service.encode('utf-8'), hashlib.sha256).digest()
        secret_signing = hmac.new(secret_service, "tc3_request".encode('utf-8'), hashlib.sha256).digest()
        signature = hmac.new(secret_signing, string_to_sign.encode('utf-8'), hashlib.sha256).hexdigest()

        # 4. 构建 Authorization
        authorization = f"{algorithm} Credential={secret_id}/{credential_scope}, SignedHeaders=content-type;host, Signature={signature}"

        return authorization, str(timestamp)

    def _build_canonical_request(self, action: str, payload: str, host: str) -> str:
        """构建规范请求字符串"""
        # HTTP 请求方法
        http_request_method = "POST"
        # 规范 URI
        canonical_uri = "/"
        # 规范查询字符串
        canonical_querystring = ""
        # 规范头部
        canonical_headers = f"content-type:application/json\nhost:{host}\n"
        # 签名的头部列表
        signed_headers = "content-type;host"
        # 请求载荷的哈希值
        hashed_request_payload = hashlib.sha256(payload.encode('utf-8')).hexdigest()

        canonical_request = f"{http_request_method}\n{canonical_uri}\n{canonical_querystring}\n{canonical_headers}\n{signed_headers}\n{hashed_request_payload}"

        return canonical_request

    # 新增: 设置当前题目索引的方法
    def set_current_question(self, index: int):
        self.current_question_index = index

    def call_first_api(self, img_str: str, prompt: Any) -> Tuple[Optional[str], Optional[str]]:
        return self._call_api_by_group("first", img_str, prompt)

    def call_second_api(self, img_str: str, prompt: Any) -> Tuple[Optional[str], Optional[str]]:
        return self._call_api_by_group("second", img_str, prompt)

    def _call_api_by_group(self, api_group: str, img_str: str, prompt: Any) -> Tuple[Optional[str], Optional[str]]:
        """根据API组别调用对应的预设供应商API"""
        try:
            if api_group == "first":
                provider = self.config_manager.first_api_provider
                api_key = self.config_manager.first_api_key
                model_id = self.config_manager.first_modelID
            elif api_group == "second":
                provider = self.config_manager.second_api_provider
                api_key = self.config_manager.second_api_key
                model_id = self.config_manager.second_modelID
            else:
                return None, "无效的API组别"

            # 兼容：provider 可能是 UI 文本（如“火山引擎 (推荐)”）而不是内部ID（如 volcengine）
            if provider and provider not in PROVIDER_CONFIGS:
                mapped_provider = get_provider_id_from_ui_text(str(provider))
                if mapped_provider:
                    provider = mapped_provider
                    # 写回内存，确保后续保存会落盘为内部ID
                    try:
                        if api_group == "first":
                            self.config_manager.first_api_provider = provider
                        elif api_group == "second":
                            self.config_manager.second_api_provider = provider
                    except Exception:
                        pass

            if not all([provider, api_key, model_id]):
                return None, f"第{api_group}组API配置不完整 (供应商、Key或模型ID为空)"

            self.logger.debug(f"[API] 准备调用 {api_group} API, 供应商: {provider}")
            return self._execute_api_call(provider, api_key, model_id, img_str, prompt)
        except Exception as e:
            error_detail = traceback.format_exc()
            self.logger.exception(f"[API] 调用 {api_group} API 时发生严重错误: {str(e)}\n{error_detail}")
            return None, f"API调用失败: {str(e)}"

    def test_api_connection(self, api_group: str) -> Tuple[bool, str]:
        """测试指定API组的连接"""
        try:
            if api_group == "first":
                provider, api_key, model_id, group_name = (
                    self.config_manager.first_api_provider, self.config_manager.first_api_key,
                    self.config_manager.first_modelID, "第一个"
                )
            elif api_group == "second":
                provider, api_key, model_id, group_name = (
                    self.config_manager.second_api_provider, self.config_manager.second_api_key,
                    self.config_manager.second_modelID, "第二个"
                )
            else:
                return False, "无效的API组别"
            
            if not all([provider, api_key.strip(), model_id.strip()]):
                return False, f"{group_name}组信息没填完整（平台/密钥/模型ID）"

            # 兼容：provider 可能是 UI 文本
            if provider and provider not in PROVIDER_CONFIGS:
                mapped_provider = get_provider_id_from_ui_text(str(provider))
                if mapped_provider:
                    provider = mapped_provider
                    try:
                        if api_group == "first":
                            self.config_manager.first_api_provider = provider
                        elif api_group == "second":
                            self.config_manager.second_api_provider = provider
                    except Exception:
                        pass

            # 测试AI评分API
            self.logger.info(f"[API Test] 测试{group_name}API, 供应商: {provider}")
            result, error = self._execute_api_call(provider, api_key, model_id, img_str="", prompt="你好")

            provider_name = PROVIDER_CONFIGS.get(provider, {}).get("name", provider)
            
            def _friendly_reason(err: Optional[str]) -> str:
                s = (err or "").strip()
                low = s.lower()
                if any(k in low for k in ["timed out", "timeout"]):
                    return "网络可能不稳定（连接超时）"
                if any(k in low for k in ["401", "unauthorized", "invalid api key"]):
                    return "密钥可能不正确或已失效"
                if any(k in low for k in ["403", "forbidden", "quota", "余额", "payment", "insufficient"]):
                    return "账号可能没有权限或余额/额度不足"
                if any(k in low for k in ["429", "rate limit", "too many"]):
                    return "请求太频繁，平台临时限制"
                if any(k in low for k in ["502", "503", "504", "service unavailable", "bad gateway"]):
                    return "平台服务繁忙或临时不可用"
                if not s:
                    return "原因不明"
                return s

            if not (result and not error):
                reason = _friendly_reason(error)
                msg = (
                    f"{provider_name}：连接失败。\n"
                    f"可能原因：{reason}。\n"
                    "建议：检查密钥/模型ID是否填写正确；确认账号余额/额度；网络正常后再试。"
                )
                return False, msg
            
            # AI评分API连接成功，构建结果信息
            result_info = f"{provider_name}：连接成功"
            
            return True, result_info
        except Exception as e:
            error_detail = traceback.format_exc()
            self.logger.exception(f"[API Test] API测试过程中发生异常: {str(e)}\n{error_detail}")
            return False, f"测试时出错：{str(e)}"

    def _preprocess_api_key(self, api_key: str, auth_method: str) -> Tuple[str, Optional[str]]:
        """
        预处理API Key，增强格式验证和兼容性

        Args:
            api_key: 原始API Key
            auth_method: 鉴权方法

        Returns:
            tuple: (processed_key, error_message)
        """
        if not api_key or not api_key.strip():
            return "", "API Key不能为空"

        api_key = api_key.strip()

        if auth_method == "bearer":
            # 处理Bearer token的重复前缀问题
            if api_key.lower().startswith("bearer "):
                api_key = api_key[7:].strip()  # 移除"Bearer "前缀
            return api_key, None

        elif auth_method == "tencent_signature_v3":
            # 处理腾讯API Key格式
            # 支持中文冒号自动转换
            api_key = api_key.replace("：", ":")  # 中文冒号转英文冒号

            # 检查冒号数量
            colon_count = api_key.count(":")
            if colon_count == 0:
                return "", "腾讯API Key格式错误：缺少冒号分隔符，应为 'SecretId:SecretKey' 格式"
            elif colon_count > 1:
                return "", "腾讯API Key格式错误：冒号数量过多，应为 'SecretId:SecretKey' 格式"

            # 分离SecretId和SecretKey
            parts = api_key.split(":", 1)
            secret_id, secret_key = parts[0].strip(), parts[1].strip()

            # 验证格式合理性
            if not secret_id:
                return "", "腾讯API Key格式错误：SecretId不能为空"
            if not secret_key:
                return "", "腾讯API Key格式错误：SecretKey不能为空"
            if len(secret_id) < 10:
                return "", "腾讯API Key格式错误：SecretId长度过短"
            if len(secret_key) < 10:
                return "", "腾讯API Key格式错误：SecretKey长度过短"

            return f"{secret_id}:{secret_key}", None

        elif auth_method == "google_api_key_in_url":
            # Google Gemini API Key - 直接使用，无特殊格式要求
            # API Key会被添加到URL参数中，不需要特殊处理
            if len(api_key) < 20:  # 基本长度检查
                return "", "Google API Key格式错误：Key长度过短"
            return api_key, None

        # 其他鉴权方法直接返回
        return api_key, None

    def _execute_api_call(self, provider: str, api_key: str, model_id: str, img_str: str, prompt) -> Tuple[Optional[str], Optional[str]]:
        # 在函数开始就获取provider_name，避免异常处理时未定义
        provider_name = PROVIDER_CONFIGS.get(provider, {}).get("name", provider)
        
        if provider not in PROVIDER_CONFIGS:
            return None, f"未知的供应商标识: {provider}"

        config = PROVIDER_CONFIGS[provider]
        url = config["url"]
        
        # 支持动态URL（例如Gemini需要在URL中包含模型名称）
        if config.get("dynamic_url", False):
            url = url.replace("{model}", model_id)
        
        headers = {}
        auth_method = config.get("auth_method", "bearer")

        # 预处理API Key
        processed_key, key_error = self._preprocess_api_key(api_key, auth_method)
        if key_error:
            return None, key_error

        # 先构建 payload，因为腾讯签名需要用到它
        try:
            builder_func = getattr(self, config["payload_builder"])
            payload = builder_func(model_id, img_str, prompt)
        except Exception as e:
            return None, f"构建请求体失败: {e}"

        # 鉴权处理
        if auth_method == "bearer":
            headers["Authorization"] = f"Bearer {processed_key}"
        elif auth_method == "google_api_key_in_url": # For Gemini
             url += f"?key={processed_key}"
        elif auth_method == "tencent_signature_v3":
            # 腾讯云签名方法 v3 - 使用预处理后的Key
            secret_id, secret_key = processed_key.split(":", 1)
            payload_str = json.dumps(payload, separators=(',', ':'))

            # 从配置中读取服务信息，避免硬编码
            service_info = config.get("service_info", {})
            service = service_info.get("service", "hunyuan")
            region = service_info.get("region", "ap-guangzhou")
            version = service_info.get("version", "2023-09-01")
            action = service_info.get("action", "ChatCompletions")

            host = service_info.get("host", "hunyuan.tencentcloudapi.com")
            authorization, timestamp = self._build_tencent_signature_v3(
                secret_id, secret_key, service, region, action, version, payload_str, host
            )
            headers["Authorization"] = authorization
            headers["X-TC-Timestamp"] = timestamp
            headers["X-TC-Version"] = version
            headers["X-TC-Action"] = action
            headers["X-TC-Region"] = region

        # 通用请求发送逻辑（所有认证方式共享）
        try:
            self.logger.debug(f"[{provider_name}] 发送API请求到: {url}")
            
            headers["Content-Type"] = "application/json"
            response = self._get_session().post(url, headers=headers, json=payload, timeout=60)

            self.logger.debug(f"[{provider_name}] 收到响应: 状态码 {response.status_code}")

            # 开发者调试：把原始响应完整输出到控制台（不走UI日志，不截断）。
            self._print_raw_ai_response(provider_name, url, response.status_code, response.text)
            
            if response.status_code == 200:
                try:
                    data = response.json()
                except Exception as e:
                    self.logger.warning(f"[{provider_name}] 响应JSON解析失败: {e}")
                    return None, f"API响应JSON解析失败：{e}"

                content = self._extract_response_content(data, provider)
                if content:
                    self.logger.debug(f"[{provider_name}] 成功提取响应内容")
                    return content, None
                else:
                    self.logger.warning(f"[{provider_name}] 响应内容为空或无法解析")
                    return None, f"API响应内容为空或无法解析。原始响应: {str(data)[:200]}"
            else:
                error_text = response.text[:200]
                self.logger.warning(f"[{provider_name}] API请求失败: {response.status_code}")
                friendly_error = self._create_api_error_message(provider, response.status_code, error_text)
                return None, friendly_error
        except requests.exceptions.Timeout:
            self.logger.warning(f"[{provider_name}] 请求超时")
            return None, f"[{provider_name}] 请求超时，请检查网络连接或稍后重试"
        except requests.exceptions.ConnectionError as e:
            self.logger.warning(f"[{provider_name}] 连接失败: {str(e)[:100]}")
            return None, f"[{provider_name}] 无法连接到服务器，请检查网络设置"
        except requests.exceptions.RequestException as e:
            self.logger.exception(f"[{provider_name}] 网络请求异常")
            friendly_error = self._create_network_error_message(e)
            return None, friendly_error

    def _extract_response_content(self, data: Dict[str, Any], provider: str) -> Optional[str]:
        """从API响应中提取内容
        
        支持的提供商响应格式：
        - OpenAI兼容格式: openai, moonshot, openrouter, zhipu, volcengine, aliyun, baidu
        - 腾讯混元格式: tencent
        - Google Gemini格式: gemini
        """
        try:
            # OpenAI兼容格式 - 标准的 choices[0].message.content
            if provider in ["openai", "moonshot", "openrouter", "zhipu", "volcengine", "aliyun", "baidu"]:
                return data["choices"][0]["message"]["content"]
            
            # 腾讯混元 - 使用相同的OpenAI兼容格式
            if provider == "tencent":
                return data["choices"][0]["message"]["content"]
            
            # Google Gemini - 特殊格式
            if provider == "gemini":
                return data["candidates"][0]["content"]["parts"][0]["text"]
        except (KeyError, IndexError, TypeError) as e:
            self.logger.warning(f"解析{provider}响应失败: {e}")
            return None # 解析失败
        return str(data) # Fallback

    def _get_pure_base64(self, img_str: str) -> str:
        if not img_str: return ""
        marker = "base64,"
        pos = img_str.find(marker)
        return img_str[pos + len(marker):] if pos != -1 else img_str

    # ==========================================================================
    #  各厂商专属的Payload构建函数
    # ==========================================================================
    def _build_openai_compatible_payload(self, model_id, img_str, prompt):
        """
        适用于大多数与OpenAI兼容的厂商 (Moonshot, 智谱, Baidu V2, Aliyun-Compatible等)
        核心原则: 图片在前，文本在后，以保证最大兼容性。
        """
        # 支持 prompt 为字符串或 {system,user} 结构
        system_text = ""
        user_text = ""
        thinking_cfg = {"type": "disabled"}
        if isinstance(prompt, dict):
            system_text = str(prompt.get("system", "") or "")
            user_text = str(prompt.get("user", "") or "")
            raw_thinking = prompt.get("thinking")
            if raw_thinking:
                if isinstance(raw_thinking, dict):
                    thinking_type = raw_thinking.get("type") or raw_thinking.get("mode") or raw_thinking.get("value")
                    thinking_cfg = {"type": str(thinking_type)} if thinking_type else raw_thinking
                elif isinstance(raw_thinking, str):
                    thinking_cfg = {"type": raw_thinking}
        else:
            user_text = str(prompt)

        messages = []
        if system_text.strip():
            messages.append({"role": "system", "content": system_text})

        if not img_str:
            messages.append({"role": "user", "content": user_text})
            return {"model": model_id, "messages": messages, "max_tokens": 4096}

        pure_base64 = self._get_pure_base64(img_str)
        # 视觉模式：system 作为单独消息，user 带 image+text
        messages.append({
            "role": "user",
            "content": [
                {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{pure_base64}"}},
                {"type": "text", "text": user_text}
            ]
        })
        return {"model": model_id, "messages": messages, "max_tokens": 4096}



    def _build_volcengine_payload(self, model_id, img_str, prompt):
        """
        专为火山引擎定制 - 符合官方API文档格式

        AI自动改卷程序专用优化 (2025-09-13 更新):
        ============================================
        当前优化: 默认使用高细节模式提升手写文字识别精度
        适用场景: AI批改学生答案图片，需准确识别手写内容

        优化详情:
        - detail: "high" - 高细节模式，适用于复杂手写识别
        - 优势: 更好的文字识别精度，适合教育场景
        - 权衡: 可能增加响应时间和token消耗

        后续优化计划:
        ============================================
        1. 图片质量自适应: 根据图片复杂度自动选择detail等级
        2. 模型验证: 确保用户选择的模型支持视觉输入
        3. 性能监控: 添加图片大小和处理时间统计
        4. 配置选项: 允许用户自定义detail参数
        5. 批量优化: 支持多图片同时处理
        """
        system_text = ""
        user_text = ""
        thinking_cfg = {"type": "disabled"}
        if isinstance(prompt, dict):
            system_text = str(prompt.get("system", "") or "")
            user_text = str(prompt.get("user", "") or "")
            prompt_thinking = prompt.get("thinking")
            if isinstance(prompt_thinking, dict) and prompt_thinking.get("type"):
                thinking_cfg = prompt_thinking
        else:
            user_text = str(prompt)

        messages = []
        if system_text.strip():
            messages.append({"role": "system", "content": system_text})

        if not img_str:
            # 纯文本模式
            messages.append({"role": "user", "content": user_text})
            return {
                "model": model_id,
                "messages": messages,
                "max_tokens": 4096,
                "thinking": thinking_cfg
            }

        # 视觉模式 - AI改卷专用配置
        # 按照火山引擎官方文档：image在前，text在后
        pure_base64 = self._get_pure_base64(img_str)
        messages.append({
            "role": "user",
            "content": [
                {
                    "type": "image_url",
                    "image_url": {
                        "url": f"data:image/jpeg;base64,{pure_base64}",
                        "detail": "high"
                    }
                },
                {"type": "text", "text": user_text}
            ]
        })
        return {
            "model": model_id,
            "messages": messages,
            "max_tokens": 4096,
            "thinking": thinking_cfg
        }





    def _build_tencent_payload(self, model_id, img_str, prompt):
        """专为腾讯混元定制 - 支持所有视觉模型

        更新历史 (Update History):
        - 2025-09-13: 重构 payload 构建逻辑
          * 统一使用 ChatCompletions 接口格式
          * 实现智能视觉模型检测
          * 支持动态模型名称输入
          * 自动选择 Contents vs Content 格式

        支持的视觉模型包括：
        - hunyuan-vision (基础多模态)
        - hunyuan-turbos-vision (旗舰模型)
        - hunyuan-turbos-vision-20250619 (最新旗舰)
        - hunyuan-t1-vision (深度思考)
        - hunyuan-t1-vision-20250619 (最新深度思考)
        - hunyuan-large-vision (多语言支持)

        未来维护注意事项 (Future Maintenance Notes):
        - 如果新模型名称不含 "vision"，需要更新检测逻辑
        - 如果腾讯改变 payload 格式，需要相应调整
        - 支持的图像格式：JPEG (base64编码)
        - 图像URL格式：data:image/jpeg;base64,{base64_data}

        Args:
            model_id: 模型名称，由用户界面输入
            img_str: 图像base64字符串（可选）
            prompt: 文本提示

        Returns:
            dict: 符合腾讯API格式的请求payload
        """
        system_text = ""
        user_text = ""
        if isinstance(prompt, dict):
            system_text = str(prompt.get("system", "") or "")
            user_text = str(prompt.get("user", "") or "")
        else:
            user_text = str(prompt)

        # 腾讯所有视觉模型都支持图像输入，通过模型名中的 "vision" 标识
        is_vision_model = "vision" in model_id.lower()

        if not img_str or not is_vision_model:
            # 纯文本模式或非视觉模型
            messages = []
            if system_text.strip():
                messages.append({"Role": "system", "Content": system_text})
            messages.append({"Role": "user", "Content": user_text})
            return {"Model": model_id, "Messages": messages, "Stream": False}

        # 视觉模型支持图像输入
        pure_base64 = self._get_pure_base64(img_str)
        messages = []
        if system_text.strip():
            messages.append({"Role": "system", "Content": system_text})
        messages.append({
            "Role": "user",
            "Contents": [
                {"Type": "text", "Text": user_text},
                {"Type": "image_url", "ImageUrl": {"Url": f"data:image/jpeg;base64,{pure_base64}"}}
            ]
        })
        return {"Model": model_id, "Messages": messages, "Stream": False}



    def _build_gemini_payload(self, model_id, img_str, prompt):
        """专为 Google Gemini 定制"""
        system_text = ""
        user_text = ""
        if isinstance(prompt, dict):
            system_text = str(prompt.get("system", "") or "")
            user_text = str(prompt.get("user", "") or "")
        else:
            user_text = str(prompt)

        payload = {}
        if system_text.strip():
            payload["system_instruction"] = {"parts": [{"text": system_text}]}

        if not img_str:
            payload["contents"] = [{"parts": [{"text": user_text}]}]
            return payload

        pure_base64 = self._get_pure_base64(img_str)
        payload["contents"] = [{
            "parts": [
                {"text": user_text},
                {"inline_data": {"mime_type": "image/jpeg", "data": pure_base64}}
            ]
        }]
        return payload


    def _create_api_error_message(self, provider: str, status_code: int, response_text: str) -> str:
        """根据API返回的错误，生成对用户更友好的错误信息。"""
        provider_name = PROVIDER_CONFIGS.get(provider, {}).get("name", provider)

        if status_code == 401 or status_code == 403:
            return (f"【认证失败】{provider_name} 的 API Key 无效或已过期。\n"
                    f"解决方案：请前往 {provider_name} 官网，检查并重新复制粘贴您的 API Key。")

        if status_code == 400:
            if "zhipu" in provider and "1210" in response_text:
                return (f"【参数错误】发送给 {provider_name} 的模型ID可能有误。\n"
                        f"解决方案：请检查您为 {provider_name} 设置的模型ID是否正确、可用，且您的账户有权访问。")
            else:
                return (f"【请求错误】发送给 {provider_name} 的请求参数有误。\n"
                        f"常见原因：模型ID填写错误或不兼容。请核对后重试。")

        if status_code == 429:
            return (f"【请求超限】您对 {provider_name} 的API请求过于频繁，已触发限流。\n"
                    f"解决方案：请稍等片刻再试，或在程序中增大'等待时间'。")

        # 返回一个通用的、但更清晰的错误
        return (f"【服务异常】{provider_name} 服务器返回了未处理的错误 (状态码: {status_code})。\n"
                f"服务器响应(部分): {response_text[:100]}")

    def _create_network_error_message(self, error: requests.exceptions.RequestException) -> str:
        """根据网络异常类型，生成用户友好的信息"""
        error_str = str(error)
        if "Invalid leading whitespace" in error_str:
            return ("【格式错误】您的 API Key 中可能包含了非法字符（如换行或多余的文字）。\n"
                    "解决方案：请彻底清空API Key输入框，然后从官网【精确地】只复制Key本身，再粘贴回来。")

        if "timed out" in error_str.lower():
            return ("【网络超时】连接API服务器超时。\n"
                    "解决方案：请检查您的网络连接是否通畅，或稍后再试。")

        # 通用网络错误
        return f"【网络连接失败】无法连接到API服务器。\n请检查您的网络设置和防火墙。错误详情: {error_str[:150]}"

    def update_config_from_manager(self):
        """
        这个方法在我们的新架构中不再需要。
        因为 `call_api` 等方法每次都会直接从 `config_manager` 读取最新的配置。
        保留此空方法以防止旧代码调用时出错。
        """
        pass

