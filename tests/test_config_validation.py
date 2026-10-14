import pytest

from config_manager import ConfigManager
from ui_schema import load_ui_schema


@pytest.fixture
def cm():
    """不读写真实 config.ini 的 ConfigManager。"""
    manager = ConfigManager.__new__(ConfigManager)
    manager.max_questions = load_ui_schema().max_questions
    manager._init_default_config()
    return manager


def _configure_ready(cm):
    cm.first_api_provider = "openai"
    cm.first_api_key = "k1"
    cm.first_modelID = "gpt-6-luna"
    cm.second_api_provider = "deepseek"
    cm.second_api_key = "k2"
    cm.second_modelID = "deepseek-flash"
    q1 = cm.question_configs["1"]
    q1.update(
        standard_answer="细则",
        answer_area={"x1": 1, "y1": 1, "x2": 9, "y2": 9},
        confirm_button_pos=(10, 10),
        score_input_pos=(20, 20),
    )


def test_defaults_report_every_missing_item(cm):
    errors = cm.validate_before_run()
    assert any("第一组：密钥" in e for e in errors)
    assert any("第二组：密钥" in e for e in errors)
    assert any("评分细则" in e for e in errors)
    assert any("答案区域" in e for e in errors)


def test_fully_configured_passes(cm):
    _configure_ready(cm)
    assert cm.validate_before_run() == []


def test_provider_ui_text_is_normalized_to_id(cm):
    _configure_ready(cm)
    cm.first_api_provider = "ChatGPT (OpenAI)"
    cm.second_api_provider = "火山引擎 (推荐)"  # 旧版界面文字
    assert cm.validate_before_run() == []
    assert cm.first_api_provider == "openai"
    assert cm.second_api_provider == "volcengine"


def test_unknown_provider_is_rejected(cm):
    _configure_ready(cm)
    cm.first_api_provider = "腾讯混元"
    assert any("第一组：请选择" in e for e in cm.validate_before_run())


def test_zero_coordinates_are_invalid(cm):
    _configure_ready(cm)
    cm.question_configs["1"]["score_input_pos"] = (0, 0)
    assert any("分数输入坐标" in e for e in cm.validate_before_run())


def test_three_step_requires_three_positions(cm):
    _configure_ready(cm)
    cm.question_configs["1"]["enable_three_step_scoring"] = True
    errors = cm.validate_before_run()
    assert sum("三步打分" in e for e in errors) == 3
    assert not any("分数输入坐标" in e for e in errors)


def test_dual_evaluation_blocker(cm):
    assert cm.dual_evaluation_blocker() is None

    cm.question_configs["2"]["enabled"] = True
    assert "多题" in cm.dual_evaluation_blocker()

    cm.question_configs["2"]["enabled"] = False
    cm.question_configs["1"]["work_mode"] = "ocr_then_grade"
    assert "识评分离" in cm.dual_evaluation_blocker()


def test_max_questions_comes_from_ui(cm):
    assert cm.max_questions == load_ui_schema().max_questions
    assert len(cm.question_configs) == cm.max_questions
