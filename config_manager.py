# config_manager.py - 配置管理模块

import configparser
import os
import sys
import appdirs

class ConfigManager:
    """配置管理器,负责保存和加载配置"""
    _instance = None
    _initialized = False

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super(ConfigManager, cls).__new__(cls)
        return cls._instance

    def __init__(self):
        if ConfigManager._initialized:
            return
        self.parser = configparser.ConfigParser(allow_no_value=True, interpolation=None)

        app_name = "AutoGraderApp"
        app_author = "Mr.Why"

        if getattr(sys, 'frozen', False):
            self.config_dir = appdirs.user_config_dir(app_name, app_author)
        else:
            self.config_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "setting")

        self.config_file_path = os.path.join(self.config_dir, "config.ini")
        os.makedirs(self.config_dir, exist_ok=True)

        self.max_questions = 7
        self._init_default_config()
        self.load_config()
        ConfigManager._initialized = True

    def _init_default_config(self):
        """初始化默认配置值
        
        注意：
        - first_api_provider 和 second_api_provider 是AI评分模型提供商
        """
        # --- AI评分模型提供商配置 ---
        self.first_api_provider = "volcengine" # 默认使用火山引擎
        self.first_api_key = ""
        self.first_modelID = ""
        self.second_api_provider = "moonshot" # 默认使用 Moonshot
        self.second_api_key = ""
        self.second_modelID = ""
        
        self.dual_evaluation_enabled = False
        self.score_diff_threshold = 5
        
        # 无人模式配置
        self.unattended_mode_enabled = False  # 无人模式开关
        self.unattended_retry_delay = 120  # 重试延迟（秒）
        self.unattended_max_retry_rounds = 10  # 最大重试轮数
        
        self.subject = ""
        self.cycle_number = 1
        self.wait_time = 1.5  # 等待时间（支持小数，默认1.5秒）
        self.api_reset_interval = 30  # API重置间隔（默认30秒）
        self.score_rounding_step = 0.5  # 分数步长（默认0.5）
        
        self.question_configs = {}
        for i in range(1, self.max_questions + 1):
            is_q1 = (i == 1)
            self.question_configs[str(i)] = {
                'enabled': is_q1, # 第一题默认启用，其他默认禁用
                'score_input_pos': None,
                'confirm_button_pos': None,
                'standard_answer': "",
                'answer_area': None,
                'min_score': 0,
                'max_score': 100,
                'enable_next_button': False,
                'next_button_pos': None,
                'enable_anomaly_button': False,  # 异常卷按钮开关
                'anomaly_button_pos': None,  # 异常卷按钮位置
                'question_type': 'Subjective_PointBased_QA',
                'score_rounding_step': 0.5,  # 每题独立步长，默认0.5
            }
            if is_q1:
                self.question_configs[str(i)].update({
                    'enable_three_step_scoring': False,
                    'score_input_pos_step1': None,
                    'score_input_pos_step2': None,
                    'score_input_pos_step3': None
                })

    def load_config(self):
        """加载配置文件，如果不存在则创建默认配置"""
        if not os.path.exists(self.config_file_path):
            print(f"配置文件不存在，创建默认配置: {self.config_file_path}")
            self._save_config_to_file()
            return
        try:
            self.parser.read(self.config_file_path, encoding='utf-8')
        except configparser.Error as e:
            print(f"配置文件格式错误，使用默认配置: {e}")
            return
        self._safe_load_config()

    def _safe_load_config(self):
        """安全地加载配置，缺失项使用默认值"""

        # 兼容旧/错误配置：允许 provider 字段写入 UI 文本（如“火山引擎 (推荐)”），自动映射为内部 provider_id（如 volcengine）。
        self.first_api_provider = self._normalize_ai_provider_value(
            self._get_config_safe('API', 'first_api_provider', "volcengine"),
            default_provider_id="volcengine",
            field_label="first_api_provider",
        )
        self.first_api_key = self._get_config_safe('API', 'first_api_key', "")
        self.first_modelID = self._get_config_safe('API', 'first_modelID', "")
        self.second_api_provider = self._normalize_ai_provider_value(
            self._get_config_safe('API', 'second_api_provider', "moonshot"),
            default_provider_id="moonshot",
            field_label="second_api_provider",
        )
        self.second_api_key = self._get_config_safe('API', 'second_api_key', "")
        self.second_modelID = self._get_config_safe('API', 'second_modelID', "")
        
        self.dual_evaluation_enabled = self._get_config_safe('DualEvaluation', 'enabled', False, bool)
        self.score_diff_threshold = self._get_config_safe('DualEvaluation', 'score_diff_threshold', 5, int)
        
        # 加载无人模式配置
        self.unattended_mode_enabled = self._get_config_safe('UnattendedMode', 'enabled', False, bool)
        self.unattended_retry_delay = self._get_config_safe('UnattendedMode', 'retry_delay', 120, int)
        self.unattended_max_retry_rounds = self._get_config_safe('UnattendedMode', 'max_retry_rounds', 10, int)
        
        self.subject = self._get_config_safe('UI', 'subject', "")
        self.cycle_number = self._get_config_safe('Auto', 'cycle_number', 1, int)
        self.wait_time = float(self._get_config_safe('Auto', 'wait_time', '1.5'))
        self.api_reset_interval = self._get_config_safe('Auto', 'api_reset_interval', 30, int)
        self.score_rounding_step = float(self._get_config_safe('Settings', 'score_rounding_step', '0.5'))
        
        # 不再从配置文件读取/写入 UI 字号与字体族（移除用户自行调整字号的设定）
        
        for i in range(1, self.max_questions + 1):
            section_name = f'Question{i}'
            q_idx_str = str(i)
            
            # 第一题的 enabled 状态在加载后会被强制设为 True
            default_enabled = (i == 1)
            
            current_q_config = {
                'enabled': self._get_config_safe(section_name, 'enabled', default_enabled, bool),
                'score_input_pos': self._parse_position(self._get_config_safe(section_name, 'score_input', None)),
                'confirm_button_pos': self._parse_position(self._get_config_safe(section_name, 'confirm_button', None)),
                'standard_answer': self._get_config_safe(section_name, 'standard_answer', ""),
                'answer_area': self._parse_area(self._get_config_safe(section_name, 'answer_area', None)),
                'min_score': self._get_config_safe(section_name, 'min_score', 0, int),
                'max_score': self._get_config_safe(section_name, 'max_score', 100, int),
                'enable_next_button': self._get_config_safe(section_name, 'enable_next_button', False, bool),
                'next_button_pos': self._parse_position(self._get_config_safe(section_name, 'next_button_pos', None)),
                'enable_anomaly_button': self._get_config_safe(section_name, 'enable_anomaly_button', False, bool),
                'anomaly_button_pos': self._parse_position(self._get_config_safe(section_name, 'anomaly_button_pos', None)),
                'question_type': self._get_config_safe(section_name, 'question_type', 'Subjective_PointBased_QA', str),
                'score_rounding_step': float(self._get_config_safe(section_name, 'score_rounding_step', '0.5')),  # 每题独立步长
            }
            if i == 1:
                current_q_config['enable_three_step_scoring'] = self._get_config_safe(section_name, 'enable_three_step_scoring', False, bool)
                current_q_config['score_input_pos_step1'] = self._parse_position(self._get_config_safe(section_name, 'score_input_pos_step1', None))
                current_q_config['score_input_pos_step2'] = self._parse_position(self._get_config_safe(section_name, 'score_input_pos_step2', None))
                current_q_config['score_input_pos_step3'] = self._parse_position(self._get_config_safe(section_name, 'score_input_pos_step3', None))
            self.question_configs[q_idx_str] = current_q_config
        
        # 强制确保第一题始终启用
        if '1' in self.question_configs:
            self.question_configs['1']['enabled'] = True

    def _normalize_ai_provider_value(self, raw_value, default_provider_id: str, field_label: str) -> str:
        """将配置中的供应商字段标准化为内部 provider_id。

        兼容输入：
        - 内部ID：volcengine / moonshot / ...
        - UI文本：火山引擎 (推荐) / 月之暗面 / ...
        """
        if raw_value is None:
            return default_provider_id

        value = str(raw_value).strip()
        if not value:
            return default_provider_id

        # 延迟导入，避免潜在循环依赖与启动开销
        try:
            from api_service import PROVIDER_CONFIGS, get_provider_id_from_ui_text
        except Exception:
            # 如果映射不可用，至少返回原始值（后续由ApiService再兜底）
            return value

        # 已经是内部ID
        if value in PROVIDER_CONFIGS:
            return value

        # 尝试从 UI 文本映射回内部ID
        provider_id = get_provider_id_from_ui_text(value)
        if provider_id:
            return provider_id

        # 未知值：保留原值，方便UI显示用户填写的内容，但后续需要UI校验阻止启动
        try:
            print(f"[ConfigManager] 未识别的AI供应商配置({field_label}): {value}")
        except Exception:
            pass
        return value

    def _get_config_safe(self, section, option, default_value, value_type: type = str):
        """安全地获取配置值"""
        try:
            if not self.parser.has_section(section) or not self.parser.has_option(section, option):
                return default_value
            raw_val = self.parser.get(section, option)
            if value_type == str: return raw_val
            elif value_type == int: return int(raw_val) if raw_val and raw_val.strip() else default_value
            elif value_type == bool: return self.parser.getboolean(section, option)
            return default_value
        except (ValueError, TypeError):
            return default_value

    def _parse_position(self, pos_str):
        try:
            if not pos_str or not pos_str.strip(): return None
            x, y = map(int, map(str.strip, pos_str.split(',')))
            return (x, y)
        except (ValueError, AttributeError, TypeError): return None

    def _parse_area(self, area_str):
        try:
            if not area_str or not area_str.strip(): return None
            coords = [int(c.strip()) for c in area_str.split(',')]
            if len(coords) != 4: return None
            return {'x1': coords[0], 'y1': coords[1], 'x2': coords[2], 'y2': coords[3]}
        except (ValueError, TypeError): return None

    def update_config_in_memory(self, field_name, value):
        """更新内存中的配置项。"""
        try:
            self._update_memory_config(field_name, value)
        except Exception as e:
            print(f"ConfigManager: Error updating memory for {field_name}: {e}")

    def _update_memory_config(self, field_name, value):
        """更新内存中的配置"""
        if field_name == 'first_api_provider': self.first_api_provider = str(value) if value else ""
        elif field_name == 'first_api_key': self.first_api_key = str(value) if value else ""
        elif field_name == 'first_modelID': self.first_modelID = str(value) if value else ""
        elif field_name == 'second_api_provider': self.second_api_provider = str(value) if value else ""
        elif field_name == 'second_api_key': self.second_api_key = str(value) if value else ""
        elif field_name == 'second_modelID': self.second_modelID = str(value) if value else ""
        elif field_name == 'subject': self.subject = str(value) if value else ""
        elif field_name == 'cycle_number': self.cycle_number = max(1, int(value)) if value else 1
        elif field_name == 'wait_time': 
            # 支持小数点，范围 0.1-9.9
            try:
                wait_val = float(value) if value else 1.5
                self.wait_time = max(0.1, min(9.9, wait_val))
            except (ValueError, TypeError):
                self.wait_time = 1.5
        elif field_name == 'api_reset_interval': self.api_reset_interval = max(0, int(value)) if value else 30
        elif field_name == 'dual_evaluation_enabled': self.dual_evaluation_enabled = bool(value)
        elif field_name == 'score_diff_threshold': self.score_diff_threshold = max(1, int(value)) if value else 5
        elif field_name == 'unattended_mode_enabled': self.unattended_mode_enabled = bool(value)
        elif field_name == 'unattended_retry_delay': self.unattended_retry_delay = max(10, int(value)) if value else 120
        elif field_name == 'unattended_max_retry_rounds': self.unattended_max_retry_rounds = max(1, int(value)) if value else 10
        elif field_name == 'score_rounding_step':
            try:
                self.score_rounding_step = float(value) if value is not None else 0.5
            except (ValueError, TypeError):
                self.score_rounding_step = 0.5
        elif field_name.startswith('question_'): self._update_question_config_from_field_name(field_name, value)
        else:
            # 忽略未知的配置字段，比如旧的 'first_api_url'
            pass

    def _update_question_config_from_field_name(self, field_name, value):
        """从字段名解析并更新题目配置"""
        parts = field_name.split('_')
        if len(parts) < 3: return
        
        q_index, field_type = parts[1], '_'.join(parts[2:])
        if q_index not in self.question_configs: return

        if field_type == 'enabled': self.question_configs[q_index]['enabled'] = bool(value)
        elif field_type == 'standard_answer': self.question_configs[q_index]['standard_answer'] = str(value) if value else ""
        elif field_type == 'score_input_pos': self.question_configs[q_index]['score_input_pos'] = value
        elif field_type == 'confirm_button_pos': self.question_configs[q_index]['confirm_button_pos'] = value
        elif field_type == 'answer_area': self.question_configs[q_index]['answer_area'] = value
        elif field_type == 'min_score': self.question_configs[q_index]['min_score'] = int(value) if value is not None else 0
        elif field_type == 'max_score': self.question_configs[q_index]['max_score'] = int(value) if value is not None else 100
        elif field_type == 'enable_next_button': self.question_configs[q_index]['enable_next_button'] = bool(value)
        elif field_type == 'next_button_pos': self.question_configs[q_index]['next_button_pos'] = value
        elif field_type == 'enable_anomaly_button': self.question_configs[q_index]['enable_anomaly_button'] = bool(value)
        elif field_type == 'anomaly_button_pos': self.question_configs[q_index]['anomaly_button_pos'] = value
        elif field_type == 'question_type': self.question_configs[q_index]['question_type'] = str(value) if value else 'Subjective_PointBased_QA'
        elif field_type == 'score_rounding_step':  # 每题独立步长
            try:
                self.question_configs[q_index]['score_rounding_step'] = float(value) if value is not None else 0.5
            except (ValueError, TypeError):
                self.question_configs[q_index]['score_rounding_step'] = 0.5
        elif q_index == '1': # 仅第一题
            if field_type == 'enable_three_step_scoring': self.question_configs[q_index]['enable_three_step_scoring'] = bool(value)
            elif field_type == 'score_input_pos_step1': self.question_configs[q_index]['score_input_pos_step1'] = value
            elif field_type == 'score_input_pos_step2': self.question_configs[q_index]['score_input_pos_step2'] = value
            elif field_type == 'score_input_pos_step3': self.question_configs[q_index]['score_input_pos_step3'] = value

    def update_question_config(self, question_index, field_type, value):
        field_name = f"question_{question_index}_{field_type}"
        self._update_memory_config(field_name, value)

    def save_all_configs_to_file(self):
        return self._save_config_to_file()

    def _save_config_to_file(self):
        """将内存中的配置保存到文件"""
        try:
            config = configparser.ConfigParser(interpolation=None)
            
            config['API'] = {
                'first_api_provider': str(self.first_api_provider),
                'first_api_key': str(self.first_api_key),
                'first_modelID': str(self.first_modelID),
                'second_api_provider': str(self.second_api_provider),
                'second_api_key': str(self.second_api_key),
                'second_modelID': str(self.second_modelID),
            }
            config['UI'] = {'subject': str(self.subject)}
            config['Auto'] = {
                'cycle_number': str(self.cycle_number), 
                'wait_time': f"{self.wait_time:.1f}",
                'api_reset_interval': str(self.api_reset_interval)
            }
            config['DualEvaluation'] = {'enabled': str(self.dual_evaluation_enabled), 'score_diff_threshold': str(self.score_diff_threshold)}
            config['UnattendedMode'] = {
                'enabled': str(self.unattended_mode_enabled),
                'retry_delay': str(self.unattended_retry_delay),
                'max_retry_rounds': str(self.unattended_max_retry_rounds)
            }
            config['Settings'] = {
                'score_rounding_step': str(self.score_rounding_step),
            }
            
            for i in range(1, self.max_questions + 1):
                section_name = f'Question{i}'
                q_idx_str = str(i)
                q_config = self.question_configs[q_idx_str]
                
                is_enabled_for_saving = q_config['enabled']
                if q_idx_str == '1': is_enabled_for_saving = True

                section_data = {
                    'enabled': str(is_enabled_for_saving),
                    'standard_answer': q_config['standard_answer'],
                    'min_score': str(q_config['min_score']),
                    'max_score': str(q_config['max_score']),
                    'enable_next_button': str(q_config['enable_next_button']),
                    'enable_anomaly_button': str(q_config.get('enable_anomaly_button', False)),
                    'question_type': q_config.get('question_type', 'Subjective_PointBased_QA'),
                    'score_rounding_step': str(q_config.get('score_rounding_step', 0.5)),
                    'score_input': f"{q_config['score_input_pos'][0]},{q_config['score_input_pos'][1]}" if q_config['score_input_pos'] else "",
                    'confirm_button': f"{q_config['confirm_button_pos'][0]},{q_config['confirm_button_pos'][1]}" if q_config['confirm_button_pos'] else "",
                    'next_button_pos': f"{q_config['next_button_pos'][0]},{q_config['next_button_pos'][1]}" if q_config['next_button_pos'] else "",
                    'anomaly_button_pos': f"{q_config['anomaly_button_pos'][0]},{q_config['anomaly_button_pos'][1]}" if q_config.get('anomaly_button_pos') else "",
                    'answer_area': f"{q_config['answer_area']['x1']},{q_config['answer_area']['y1']},{q_config['answer_area']['x2']},{q_config['answer_area']['y2']}" if q_config['answer_area'] else "",
                }
                
                if q_idx_str == '1':
                    section_data['enable_three_step_scoring'] = str(q_config.get('enable_three_step_scoring', False))
                    pos1 = q_config.get('score_input_pos_step1')
                    section_data['score_input_pos_step1'] = f"{pos1[0]},{pos1[1]}" if pos1 else ""
                    pos2 = q_config.get('score_input_pos_step2')
                    section_data['score_input_pos_step2'] = f"{pos2[0]},{pos2[1]}" if pos2 else ""
                    pos3 = q_config.get('score_input_pos_step3')
                    section_data['score_input_pos_step3'] = f"{pos3[0]},{pos3[1]}" if pos3 else ""
                
                config[section_name] = section_data
            
            with open(self.config_file_path, 'w', encoding='utf-8') as f:
                config.write(f)
            return True
        except Exception as e:
            print(f"保存配置文件失败: {e}")
            return False

    def get_enabled_questions(self):
        return [i for i in range(1, self.max_questions + 1) if self.question_configs.get(str(i), {}).get('enabled', False)]

    def get_question_config(self, question_index: int) -> dict:
        """获取指定题目的配置
        
        Args:
            question_index: 题目索引（1-7）
            
        Returns:
            dict: 题目配置字典，如果不存在则返回 {'enabled': False}
        """
        return self.question_configs.get(str(question_index), {'enabled': False})
    
    def check_required_settings(self):
        if not self.first_api_key or not self.first_modelID or not self.first_api_provider:
            return False
        if self.dual_evaluation_enabled and (not self.second_api_key or not self.second_modelID or not self.second_api_provider):
            return False
        return True
