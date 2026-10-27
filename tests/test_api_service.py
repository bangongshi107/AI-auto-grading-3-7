import base64
import io

import pytest
from PIL import Image

from api_service import (
    ApiService,
    PROVIDER_CONFIGS,
    REASONING_MAX_OUTPUT_TOKENS,
    REASONING_REQUEST_TIMEOUT,
    REQUEST_TIMEOUT,
    build_test_image_data_uri,
    get_provider_id_from_ui_text,
)

PROMPT = {"system": "S", "user": "U"}
IMG = "data:image/jpeg;base64,AAAA"


@pytest.fixture
def api():
    return ApiService(object())


class _FakeResponse:
    status_code = 200
    text = "{}"

    def json(self):
        return {"choices": [{"finish_reason": "stop", "message": {"content": "OK"}}]}


class _FakeSession:
    def __init__(self):
        self.timeout = None

    def post(self, url, headers=None, json=None, timeout=None):
        self.timeout = timeout
        return _FakeResponse()


def test_provider_aliases():
    assert get_provider_id_from_ui_text("ChatGPT (OpenAI)") == "openai"
    assert get_provider_id_from_ui_text("ChatGPT") == "openai"
    assert get_provider_id_from_ui_text("OpenAI") == "openai"
    assert get_provider_id_from_ui_text("火山引擎") == "volcengine"
    assert get_provider_id_from_ui_text("火山引擎 (推荐)") == "volcengine"
    assert get_provider_id_from_ui_text("腾讯混元") is None


def test_every_provider_has_a_payload_builder(api):
    for config in PROVIDER_CONFIGS.values():
        assert callable(getattr(api, config["payload_builder"]))


@pytest.mark.parametrize(
    "provider, expected",
    [
        ("deepseek", REASONING_REQUEST_TIMEOUT),
        ("openai", REASONING_REQUEST_TIMEOUT),
        ("volcengine", REQUEST_TIMEOUT),
        ("moonshot", REQUEST_TIMEOUT),
    ],
)
def test_request_timeout_per_provider(api, provider, expected):
    session = _FakeSession()
    api._get_session = lambda: session
    content, error = api._execute_api_call(provider, "key", "model", IMG, PROMPT)
    assert (content, error) == ("OK", None)
    assert session.timeout == expected


def test_reasoning_providers_get_longer_timeout():
    assert REASONING_REQUEST_TIMEOUT == 300
    assert REASONING_REQUEST_TIMEOUT > REQUEST_TIMEOUT


def test_deepseek_payload_keeps_thinking_on_with_default_effort(api):
    payload = api._build_deepseek_payload("deepseek-flash", IMG, PROMPT)
    assert payload["thinking"] == {"type": "enabled"}
    assert "reasoning_effort" not in payload
    assert payload["max_tokens"] == REASONING_MAX_OUTPUT_TOKENS
    content = payload["messages"][-1]["content"]
    assert content[0]["type"] == "text" and content[1]["type"] == "image_url"
    assert payload["messages"][0] == {"role": "system", "content": "S"}


def test_openai_payload_uses_max_completion_tokens(api):
    payload = api._build_openai_payload("gpt-6-luna", IMG, PROMPT)
    assert "max_tokens" not in payload
    assert payload["max_completion_tokens"] == REASONING_MAX_OUTPUT_TOKENS
    assert "reasoning_effort" not in payload
    assert payload["messages"][-1]["content"][-1]["image_url"]["detail"] == "high"


def test_deepseek_payload_does_not_send_detail(api):
    payload = api._build_deepseek_payload("deepseek-flash", IMG, PROMPT)
    assert "detail" not in payload["messages"][-1]["content"][-1]["image_url"]


def test_volcengine_payload_unchanged_semantics(api):
    payload = api._build_volcengine_payload("doubao", IMG, PROMPT)
    assert payload["thinking"] == {"type": "enabled"}
    assert payload["max_tokens"] == 8192
    assert payload["messages"][-1]["content"][-1]["image_url"]["detail"] == "high"
    text_only = api._build_volcengine_payload("doubao", "", "hi")
    assert text_only["messages"] == [{"role": "user", "content": "hi"}]


def test_http_retry_only_retries_connection_errors(api):
    session = api._get_session()
    retry = session.get_adapter("https://api.deepseek.com").max_retries
    assert retry.connect == 2
    assert retry.read == 0
    assert retry.status == 0
    assert not retry.status_forcelist


def test_text_only_payload_has_no_image(api):
    payload = api._build_openai_compatible_payload("m", "", "hello")
    assert payload["messages"] == [{"role": "user", "content": "hello"}]


@pytest.mark.parametrize(
    "raw, expected",
    [
        ('{"score": 5}', '{"score": 5}'),
        ('<think>内部推理</think>\n{"score": 5}', '{"score": 5}'),
        ("<THINK>x</THINK>答案", "答案"),
        ("推理过程……</think>最终答案", "最终答案"),
        ("<think>只有思考</think>", ""),
        ("", ""),
        (None, ""),
    ],
)
def test_strip_reasoning(raw, expected):
    assert ApiService._strip_reasoning(raw) == expected


def test_extract_ignores_reasoning_content_field(api):
    data = {"choices": [{"message": {"reasoning_content": "思考", "content": "最终"}}]}
    assert api._extract_response_content(data, "deepseek") == "最终"


def test_extract_returns_none_when_only_reasoning(api):
    data = {"choices": [{"message": {"reasoning_content": "思考", "content": ""}}]}
    assert api._extract_response_content(data, "deepseek") is None


def test_extract_handles_content_blocks(api):
    data = {"choices": [{"message": {"content": [{"type": "text", "text": "甲"}, {"type": "text", "text": "乙"}]}}]}
    assert api._extract_response_content(data, "openai") == "甲乙"


def test_extract_gemini_skips_thought_parts(api):
    data = {"candidates": [{"content": {"parts": [{"text": "思考摘要", "thought": True}, {"text": "答案"}]}}]}
    assert api._extract_response_content(data, "gemini") == "答案"


def test_truncation_detection():
    assert ApiService._is_truncated({"choices": [{"finish_reason": "length"}]}, "deepseek")
    assert not ApiService._is_truncated({"choices": [{"finish_reason": "stop"}]}, "deepseek")
    assert ApiService._is_truncated({"candidates": [{"finishReason": "MAX_TOKENS"}]}, "gemini")


def test_connection_test_image_is_a_valid_jpeg():
    uri = build_test_image_data_uri()
    assert uri.startswith("data:image/jpeg;base64,")
    image = Image.open(io.BytesIO(base64.b64decode(uri.split(",", 1)[1])))
    assert image.format == "JPEG" and image.size == (256, 256)
