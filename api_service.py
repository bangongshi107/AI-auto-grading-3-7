# api_service.py - AI评分API服务模块
# 支持多平台：火山引擎、阿里通义、百度千帆、DeepSeek、智谱、月之暗面、OpenRouter、OpenAI、Google Gemini

import requests
import re
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
import logging
import traceback
from typing import Tuple, Optional, Dict, Any
import sys
from datetime import datetime
from threading import local

from grading_support import classify_teacher_facing_error, TeacherErrorCategory

# 思考模式会占用输出额度并拖慢响应，固定取较宽松的值（仍不超过多数视觉模型的上限）
MAX_OUTPUT_TOKENS = 8192
# DeepSeek / OpenAI 的思考过程计入输出上限，8192 容易被思考吃满导致正文被截断
REASONING_MAX_OUTPUT_TOKENS = 32768
REQUEST_TIMEOUT = 180  # 秒
# 思考模式 + 高清手写图响应更慢；先按 300 秒放宽，待实测耗时后再调整
REASONING_REQUEST_TIMEOUT = 300  # 秒

# ==============================================================================
#  权威供应商配置字典 (Authoritative Provider Configuration)
#  这里定义每个供应商的接口行为（地址、鉴权、请求体格式）。
#  下拉框里“提供哪些选项、显示什么文字”由 setting/七题.ui 决定（见 ui_schema.py），
#  "name" 必须与 .ui 里的选项文字一致；"aliases" 仅用于兼容旧配置里写过的名称。
# ==============================================================================
PROVIDER_CONFIGS = {
    # 这里的 key ('volcengine', 'moonshot'等) 是程序内部使用的【内部标识】
    "volcengine": {
        "name": "火山引擎",
        "aliases": ("火山引擎 (推荐)",),
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
    "deepseek": {
        "name": "DeepSeek",
        "url": "https://api.deepseek.com/chat/completions",
        "auth_method": "bearer",
        "payload_builder": "_build_deepseek_payload",
        "max_output_tokens": REASONING_MAX_OUTPUT_TOKENS,
        "timeout": REASONING_REQUEST_TIMEOUT,
    },
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
    "openrouter": {
        "name": "OpenRouter",
        "url": "https://openrouter.ai/api/v1/chat/completions",
        "auth_method": "bearer",
        "payload_builder": "_build_openai_compatible_payload",
    },
    "openai": {
        "name": "ChatGPT (OpenAI)",
        "aliases": ("ChatGPT", "OpenAI"),
        "url": "https://api.openai.com/v1/chat/completions",
        "auth_method": "bearer",
        "payload_builder": "_build_openai_payload",
        "max_output_tokens": REASONING_MAX_OUTPUT_TOKENS,
        "timeout": REASONING_REQUEST_TIMEOUT,
    },
    "gemini": {
        "name": "Google Gemini",
        "url": "https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent",  # {model} 将被动态替换
        "auth_method": "google_api_key_in_url",
        "payload_builder": "_build_gemini_payload",
        "dynamic_url": True,  # 标记需要动态URL替换
    }
}

# ==============================================================================
#  辅助函数，用于UI显示文本和内部ID之间的转换
# ==============================================================================
def get_provider_id_from_ui_text(ui_text: str) -> Optional[str]:
    text = (ui_text or "").strip()
    for provider_id, config in PROVIDER_CONFIGS.items():
        if text == config["name"] or text in config.get("aliases", ()):
            return provider_id
    return None

def build_test_image_data_uri() -> str:
    """生成一张小的测试图片（JPEG data URI），用于连接测试时验证模型能否接收图片。
    生成失败时返回空串，调用方退化为纯文本测试。"""
    try:
        import base64
        import io
        from PIL import Image, ImageDraw

        img = Image.new("RGB", (256, 256), "white")
        draw = ImageDraw.Draw(img)
        draw.rectangle((40, 40, 216, 216), outline="black", width=6)
        draw.line((40, 40, 216, 216), fill="black", width=6)
        draw.line((40, 216, 216, 40), fill="black", width=6)
        buf = io.BytesIO()
        img.save(buf, format="JPEG", quality=85)
        return "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode("ascii")
    except Exception:
        return ""


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
        """获取当前线程专属的 requests.Session。
        
        配置智能重试和连接池管理：
        - 自动重试连接错误（Max retries exceeded问题）
        - 禁用keep-alive连接复用超时导致的断连问题
        - 在底层连接失败时自动重建连接
        """
        sess = getattr(self._thread_local, "session", None)
        if sess is None:
            sess = requests.Session()
            
            # 重试策略：只重试“连接阶段”失败（请求还没发出，安全且很快）。
            # 读超时、5xx 都不在这里重试——它们由 AutoThread 的双API交叉重试统一处理，
            # 否则 300 秒的超时会被底层重试放大成十几分钟的卡死。
            retry_strategy = Retry(
                total=2,
                connect=2,                  # 连接失败最多重试2次
                read=0,                     # 读超时不重试
                status=0,                   # 5xx 不重试
                other=0,
                backoff_factor=0.5,         # 退避：0.5s, 1s
                raise_on_status=False,      # 不抛出状态码异常，由业务层处理
            )
            
            # 创建适配器，配置连接池参数
            adapter = HTTPAdapter(
                max_retries=retry_strategy,
                pool_connections=5,         # 连接池大小
                pool_maxsize=5,             # 每个主机的最大连接数
                pool_block=False,           # 不阻塞等待连接
            )
            
            # 为HTTP和HTTPS都挂载适配器
            sess.mount("https://", adapter)
            sess.mount("http://", adapter)

            # 禁用 HTTP keep-alive：防止服务端空闲超时关闭连接后客户端仍复用
            # 导致 "Max retries exceeded" / 僵尸连接问题
            sess.headers.update({"Connection": "close"})

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

    # 新增: 设置当前题目索引的方法
    def set_current_question(self, index: int):
        self.current_question_index = index

    def call_first_api(self, img_str: str, prompt: Any) -> Tuple[Optional[str], Optional[str]]:
        return self._call_api_by_group("first", img_str, prompt)

    def call_second_api(self, img_str: str, prompt: Any) -> Tuple[Optional[str], Optional[str]]:
        return self._call_api_by_group("second", img_str, prompt)

    def _get_group_config(self, api_group: str) -> Tuple[str, str, str]:
        """读取指定API组的 (provider内部ID, api_key, model_id)。api_group 必须是 first/second。"""
        cm = self.config_manager
        provider = getattr(cm, f"{api_group}_api_provider")
        api_key = getattr(cm, f"{api_group}_api_key")
        model_id = getattr(cm, f"{api_group}_modelID")

        # provider 可能是 UI 文本（如“火山引擎”）而不是内部ID（如 volcengine）
        if provider and provider not in PROVIDER_CONFIGS:
            mapped_provider = get_provider_id_from_ui_text(str(provider))
            if mapped_provider:
                provider = mapped_provider
                setattr(cm, f"{api_group}_api_provider", provider)  # 写回内存，保存时落盘为内部ID
        return provider, api_key, model_id

    def _call_api_by_group(self, api_group: str, img_str: str, prompt: Any) -> Tuple[Optional[str], Optional[str]]:
        """根据API组别调用对应的预设供应商API"""
        try:
            if api_group not in ("first", "second"):
                return None, "无效的API组别"
            provider, api_key, model_id = self._get_group_config(api_group)

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
            if api_group not in ("first", "second"):
                return False, "无效的API组别"
            group_name = "第一个" if api_group == "first" else "第二个"
            provider, api_key, model_id = self._get_group_config(api_group)

            if not all([provider, api_key.strip(), model_id.strip()]):
                return False, f"{group_name}组信息没填完整（平台/密钥/模型ID）"

            # 测试请求带一张小图：评分和OCR都要看图，只测文字通不通无法发现“模型不支持图片输入”
            self.logger.info(f"[API Test] 测试{group_name}API, 供应商: {provider}")
            test_image = build_test_image_data_uri()
            result, error = self._execute_api_call(
                provider, api_key, model_id, img_str=test_image, prompt="这是连通性测试，请只回复：OK"
            )

            provider_name = PROVIDER_CONFIGS.get(provider, {}).get("name", provider)
            
            def _friendly_reason(err: Optional[str]) -> str:
                messages = {
                    TeacherErrorCategory.TIMEOUT: "网络可能不稳定（连接超时）",
                    TeacherErrorCategory.AUTH_401: "密钥可能不正确或已失效",
                    TeacherErrorCategory.QUOTA_403: "账号可能没有权限或余额/额度不足",
                    TeacherErrorCategory.RATE_LIMIT_429: "请求太频繁，平台临时限制",
                    TeacherErrorCategory.SERVICE_5XX: "平台服务繁忙或临时不可用",
                }
                category = classify_teacher_facing_error(err)
                if category in messages:
                    return messages[category]
                s = (err or "").strip()
                return s if s else "原因不明"

            if not (result and not error):
                reason = _friendly_reason(error)
                msg = (
                    f"{provider_name}：连接失败。\n"
                    f"可能原因：{reason}。\n"
                    "建议：检查密钥/模型ID是否填写正确；确认所选模型支持图片输入（测试会带一张小图）；"
                    "确认账号余额/额度；网络正常后再试。"
                )
                return False, msg
            
            # AI评分API连接成功，构建结果信息
            result_info = f"{provider_name}：连接成功" + ("（已验证可接收图片）" if test_image else "")
            
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

        # 先构建 payload
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

        # 通用请求发送逻辑（所有认证方式共享）
        try:
            self.logger.debug(f"[{provider_name}] 发送API请求到: {url}")
            
            headers["Content-Type"] = "application/json"
            response = self._get_session().post(
                url, headers=headers, json=payload, timeout=config.get("timeout", REQUEST_TIMEOUT)
            )

            self.logger.debug(f"[{provider_name}] 收到响应: 状态码 {response.status_code}")

            # 开发者调试：把原始响应完整输出到控制台（不走UI日志，不截断）。
            self._print_raw_ai_response(provider_name, url, response.status_code, response.text)
            
            if response.status_code == 200:
                try:
                    data = response.json()
                except Exception as e:
                    self.logger.warning(f"[{provider_name}] 响应JSON解析失败: {e}")
                    return None, f"API响应JSON解析失败：{e}"

                if self._is_truncated(data, provider):
                    self.logger.warning(f"[{provider_name}] 输出被截断")
                    token_limit = config.get("max_output_tokens", MAX_OUTPUT_TOKENS)
                    return None, f"[{provider_name}] AI输出被截断（超过{token_limit} tokens上限，思考过程也计入），请精简评分细则或更换模型"

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
            # 超时可能是连接池问题，重置Session确保下次使用新连接
            self.reset()
            return None, f"[{provider_name}] 请求超时"
        except requests.exceptions.ConnectionError as e:
            self.logger.warning(f"[{provider_name}] 连接失败: {str(e)[:100]}")
            # 连接出错时重置Session，清理可能损坏的连接池
            self.reset()
            return None, f"[{provider_name}] 连接失败: {str(e)[:100]}"
        except requests.exceptions.RequestException as e:
            self.logger.exception(f"[{provider_name}] 网络请求异常")
            # 其他网络异常也重置Session
            self.reset()
            friendly_error = self._create_network_error_message(e)
            return None, friendly_error

    @staticmethod
    def _is_truncated(data: Dict[str, Any], provider: str) -> bool:
        """响应是否因达到输出上限而被截断。"""
        try:
            if provider == "gemini":
                return data["candidates"][0].get("finishReason") == "MAX_TOKENS"
            return data["choices"][0].get("finish_reason") == "length"
        except (KeyError, IndexError, TypeError, AttributeError):
            return False

    @staticmethod
    def _strip_reasoning(text: Optional[str]) -> str:
        """去掉混在正文里的思考过程，只保留最终回答。

        DeepSeek / OpenAI 官方接口把思考放在独立字段（reasoning_content 等），正文本身是干净的；
        这里额外兜底一些网关/开源部署把 <think>...</think> 直接拼进 content 的情况。
        """
        if not text:
            return ""
        text = re.sub(r"<think(?:ing)?>.*?</think(?:ing)?>", "", text, flags=re.DOTALL | re.IGNORECASE)
        # 只有结束标签、没有开始标签：结束标签之前都是思考
        match = re.search(r"</think(?:ing)?>", text, flags=re.IGNORECASE)
        if match:
            text = text[match.end():]
        return text.strip()

    def _extract_response_content(self, data: Dict[str, Any], provider: str) -> Optional[str]:
        """从API响应中提取最终回答（不含思考过程）。

        Gemini 用 candidates 结构，其余均为 OpenAI 兼容的 choices 结构；
        只读取 message.content，不读取 reasoning_content。
        """
        try:
            if provider == "gemini":
                parts = data["candidates"][0]["content"]["parts"]
                # Gemini 的思考摘要带 thought=True 标记，需排除
                text = "".join(p.get("text", "") for p in parts if not p.get("thought"))
                return text or None
            content = data["choices"][0]["message"]["content"]
            if isinstance(content, list):
                content = "".join(
                    str(block.get("text", "")) for block in content
                    if isinstance(block, dict) and block.get("type", "text") == "text"
                )
            return self._strip_reasoning(content) or None
        except (KeyError, IndexError, TypeError, AttributeError) as e:
            self.logger.warning(f"解析{provider}响应失败: {e}")
            return None

    def _get_pure_base64(self, img_str: str) -> str:
        if not img_str: return ""
        marker = "base64,"
        pos = img_str.find(marker)
        return img_str[pos + len(marker):] if pos != -1 else img_str

    # ==========================================================================
    #  各厂商专属的Payload构建函数
    # ==========================================================================
    def _build_chat_messages(self, img_str, prompt, detail=None):
        """构造 OpenAI 风格的 messages。prompt 可为字符串或 {system,user} 结构。
        核心原则: 文本在前、图片在后。图片每份卷子都不同，放在最后才能让 system + 固定文本
        （题型提示/评分细则等）成为相同前缀，命中 OpenAI/DeepSeek 等的自动前缀缓存。
        项目约定: 程序使用的所有模型均支持图片输入，无需再做兼容性假设。"""
        system_text = ""
        user_text = ""
        if isinstance(prompt, dict):
            system_text = str(prompt.get("system", "") or "")
            user_text = str(prompt.get("user", "") or "")
        else:
            user_text = str(prompt)

        messages = []
        if system_text.strip():
            messages.append({"role": "system", "content": system_text})

        if not img_str:
            messages.append({"role": "user", "content": user_text})
            return messages

        image_url = {"url": f"data:image/jpeg;base64,{self._get_pure_base64(img_str)}"}
        if detail:
            image_url["detail"] = detail
        messages.append({
            "role": "user",
            "content": [
                {"type": "text", "text": user_text},
                {"type": "image_url", "image_url": image_url},
            ]
        })
        return messages

    def _build_openai_compatible_payload(self, model_id, img_str, prompt):
        """适用于大多数与OpenAI兼容的厂商 (Moonshot, 智谱, Baidu V2, Aliyun-Compatible等)"""
        return {
            "model": model_id,
            "messages": self._build_chat_messages(img_str, prompt),
            "max_tokens": MAX_OUTPUT_TOKENS,
        }

    def _build_deepseek_payload(self, model_id, img_str, prompt):
        """DeepSeek：思考恒开（关闭会明显降低评分正确率），思考内容由响应里的 reasoning_content 单独返回，
        解析时只取 content。不显式传 reasoning_effort：官方默认即 high，足够手写评分；
        再调高到 max 只会更慢更贵，调低则可能影响正确率。思考过程计入 max_tokens，因此上限放宽。
        不传 detail：DeepSeek 默认（auto）即保留原图，并统一缩放到约 1300×1300 以内，传 high 没有区别。"""
        return {
            "model": model_id,
            "messages": self._build_chat_messages(img_str, prompt),
            "max_tokens": REASONING_MAX_OUTPUT_TOKENS,
            "thinking": {"type": "enabled"},
        }

    def _build_openai_payload(self, model_id, img_str, prompt):
        """OpenAI：新一代推理模型不再接受 max_tokens，必须使用 max_completion_tokens（旧模型同样支持）。
        不传 reasoning_effort，沿用模型默认的思考强度，避免不支持该参数的模型报错。
        detail 固定为 high：auto 可能对中等尺寸图片选 low（缩到 512×512），会糊掉手写笔画；
        high 只多出几百个输入 token，相对思考用掉的输出 token 可以忽略。截图本身不放大，也不再额外提高分辨率。"""
        return {
            "model": model_id,
            "messages": self._build_chat_messages(img_str, prompt, detail="high"),
            "max_completion_tokens": REASONING_MAX_OUTPUT_TOKENS,
        }



    def _build_volcengine_payload(self, model_id, img_str, prompt):
        """火山引擎：图片在前文本在后，detail=high 保留手写笔画细节，思考模式恒开。"""
        return {
            "model": model_id,
            "messages": self._build_chat_messages(img_str, prompt, detail="high"),
            "max_tokens": MAX_OUTPUT_TOKENS,
            "thinking": {"type": "enabled"},
        }

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

