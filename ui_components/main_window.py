# --- START OF FILE main_window.py ---

import sys
import os
import datetime
import pathlib
import traceback
from typing import Union, Optional, Type, TypeVar, cast
from PyQt5.QtWidgets import (QMainWindow, QWidget, QMessageBox, QDialog,
                             QComboBox, QLineEdit, QCheckBox, QSpinBox,
                             QPlainTextEdit, QApplication, QShortcut, QLabel, QPushButton)
from PyQt5.QtCore import Qt, pyqtSignal, QEvent, QObject
from PyQt5.QtGui import QKeySequence, QFont, QKeyEvent, QCloseEvent
from PyQt5 import uic

# --- 新增导入 ---
# 从 api_service.py 导入转换函数和UI文本列表生成函数
from api_service import get_provider_id_from_ui_text, get_ui_text_from_provider_id, UI_TEXT_TO_PROVIDER_ID, PROVIDER_CONFIGS

class MainWindow(QMainWindow):
    # 日志级别定义
    LOG_LEVEL_INFO = "INFO"      # 基本信息
    LOG_LEVEL_DETAIL = "DETAIL"  # 详细处理信息
    LOG_LEVEL_RESULT = "RESULT"  # AI评分结果
    LOG_LEVEL_ERROR = "ERROR"    # 错误信息

    # ... (信号定义部分保持不变) ...
    # update_signal = pyqtSignal(str)
    log_signal = pyqtSignal(str, bool, str)  # message, is_error, level
    progress_signal = pyqtSignal(int, int)
    finished_signal = pyqtSignal()


    def __init__(self, config_manager, api_service, worker):
        super().__init__()
        self.config_manager = config_manager
        self.api_service = api_service
        self.worker = worker
        self._is_initializing = True

        # ... (UI文件加载部分保持不变) ...
        if getattr(sys, 'frozen', False):
            base_path = sys._MEIPASS  # type: ignore
        else:
            base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        ui_path = os.path.join(base_path, "setting", "七题.ui")
        uic.loadUi(ui_path, self)

        # ... (其他初始化属性保持不变) ...
        self.answer_windows = {}
        self.current_question = 1
        self.max_questions = 7  # 多题模式最多支持7道题（已移除第8题）
        self.shortcut_esc = QShortcut(QKeySequence("Escape"), self)
        self.shortcut_esc.activated.connect(self.stop_auto_thread)
        self._ui_cache = {}

        self.init_ui()



        self.show()
        self._is_initializing = False
        self.log_message("主窗口初始化完成")

    # ==========================================================================
    #  核心修改：配置处理逻辑
    # ==========================================================================

    def handle_lineEdit_save(self, field_name, value):
        if self._is_initializing: return
        self.config_manager.update_config_in_memory(field_name, value)
        self.log_message(f"配置项 '{field_name}' 更新为: {value}")

    def handle_plainTextEdit_save(self, field_name, value):
        if self._is_initializing: return
        self.config_manager.update_config_in_memory(field_name, value)
        # 答案内容较长，日志可以简洁些
        self.log_message(f"配置项 '{field_name}' 已更新")

    def handle_spinBox_save(self, field_name, value):
        if self._is_initializing: return
        self.config_manager.update_config_in_memory(field_name, value)
        self.log_message(f"配置项 '{field_name}' 更新为: {value}")
    
    # --- 统一的 ComboBox 处理函数 ---
    def handle_comboBox_save(self, combo_box_name, ui_text):
        """统一的ComboBox保存处理
        
        重要说明：
        - first_api_url/second_api_url: 只处理AI评分模型提供商
        """
        if self._is_initializing: return

        if combo_box_name in ['first_api_url', 'second_api_url']:
            # 处理AI评分模型提供商 ComboBox
            provider_id = get_provider_id_from_ui_text(ui_text)
            if not provider_id:
                self.log_message(f"错误: 无法识别的AI模型提供商 '{ui_text}'", is_error=True)
                return
            field_name = 'first_api_provider' if combo_box_name == 'first_api_url' else 'second_api_provider'
            self.config_manager.update_config_in_memory(field_name, provider_id)
            self.log_message(f"AI评分模型 '{field_name}' 更新为: {provider_id} ({ui_text})")
        else:
            # 处理普通ComboBox（如subject_text）
            field_name = combo_box_name.replace('_text', '')  # subject_text -> subject
            self.config_manager.update_config_in_memory(field_name, ui_text)
            self.log_message(f"配置项 '{field_name}' 更新为: {ui_text}")

    def handle_checkBox_save(self, field_name, state):
        if self._is_initializing: return
        value = bool(state)
        self.config_manager.update_config_in_memory(field_name, value)
        self.log_message(f"配置项 '{field_name}' 更新为: {value}")

    def _connect_direct_edit_save_signals(self):
        """连接UI控件信号到即时保存处理函数"""
        # API Key 和 Model ID 字段
        for field_name in ['first_api_key', 'first_modelID', 'second_api_key', 'second_modelID']:
            widget = self.get_ui_element(field_name, QLineEdit)
            if isinstance(widget, QLineEdit):
                widget.editingFinished.connect(
                    lambda field=field_name, w=widget: self.handle_lineEdit_save(field, w.text())
                )
        
        # --- 统一的 ComboBox 信号连接 ---
        combo_boxes = ['first_api_url', 'second_api_url', 'subject_text']
        for combo_name in combo_boxes:
            widget = self.get_ui_element(combo_name, QComboBox)
            if widget:
                widget.currentTextChanged.connect(
                    lambda text, name=combo_name: self.handle_comboBox_save(name, text)
                )

        # 其他控件的信号连接保持不变...
        for field, w_type in [('cycle_number', QSpinBox), ('wait_time', QSpinBox)]:
            widget = self.get_ui_element(field, w_type)
            if widget:
                widget.valueChanged.connect(
                    lambda val, f=field: self.handle_spinBox_save(f, val)
                )

        for i in range(1, self.max_questions + 1):
            std_answer_widget = self.get_ui_element(f'StandardAnswer_text_{i}', QPlainTextEdit)
            if std_answer_widget:
                self._connect_plain_text_edit_save_signal(std_answer_widget, i)

    # --- eventFilter 和 _connect_plain_text_edit_save_signal 保持不变 ---
    def _connect_plain_text_edit_save_signal(self, widget, question_index):
        widget.setProperty('question_index', question_index)
        widget.setProperty('needs_save_on_focus_out', True)
        widget.installEventFilter(self)

    def eventFilter(self, a0: Optional[QObject], a1: Optional[QEvent]) -> bool:
        if (a0 and a1 and a1.type() == QEvent.Type.FocusOut and
            hasattr(a0, 'property') and
            a0.property('needs_save_on_focus_out')):
            q_index = a0.property('question_index')
            field_name = f"question_{q_index}_standard_answer"
            plain_text_edit = cast(Optional[QPlainTextEdit], a0)
            if plain_text_edit:
                self.handle_plainTextEdit_save(field_name, plain_text_edit.toPlainText())
        return super().eventFilter(cast(QObject, a0), cast(QEvent, a1))

    # ==========================================================================
    #  UI初始化和加载逻辑
    # ==========================================================================

    def init_ui(self):
        """初始化UI组件和布局
        
        重要说明：
        - first_api_url 和 second_api_url 下拉框只包含AI评分模型提供商
        """
        # --- 核心修改: 动态填充 ComboBox，只包含AI评分模型 ---
        provider_ui_texts = list(UI_TEXT_TO_PROVIDER_ID.keys())
        for combo_name in ['first_api_url', 'second_api_url']:
            combo_box = self.get_ui_element(combo_name, QComboBox)
            if combo_box:
                combo_box.clear()
                combo_box.addItems(provider_ui_texts)

        # UI文件历史上包含第8题Tab；此处确保运行时只保留7题
        self._trim_question_tabs_to_max()

        self.setup_question_selector()
        # 将选中选项卡设置为高亮背景，便于视觉识别当前小题
        try:
            tab_widget = self.get_ui_element('questionTabs')
            if tab_widget:
                try:
                    tabbar = tab_widget.tabBar()
                    # 选中时黄色背景，未选中时白色，增加内边距让视觉更明显
                    tabbar.setStyleSheet(
                        "QTabBar::tab:selected { background: #FFF9C4; color: #0b3a5a; border:1px solid #FFE5B4; border-radius:4px; }"
                        "QTabBar::tab { background: #ffffff; color: #333; padding:6px 12px; margin:2px; }"
                    )
                except Exception:
                    pass
        except Exception:
            pass
        # ... 其他 setup 方法 ...
        self.setup_text_fields()
        self.setup_dual_evaluation()
        self.setup_unattended_mode()

        self.load_config_to_ui()
        self._connect_signals() # <--- 在这里统一调用

        self.log_message("UI组件初始化完成")

    def _trim_question_tabs_to_max(self) -> None:
        """确保题目Tabs数量不超过 self.max_questions。

        这样即使UI文件仍含“第8题”相关控件，运行时也会被移除，用户不可见。
        """
        tab_widget = self.get_ui_element('questionTabs')
        if not tab_widget:
            return

        try:
            while tab_widget.count() > self.max_questions:
                tab_widget.removeTab(tab_widget.count() - 1)
        except Exception:
            # UI控件异常时保持容错，不阻断主界面启动
            pass
    
    def load_config_to_ui(self):
        """将配置从ConfigManager加载到UI控件"""
        if self._is_initializing and hasattr(self, '_config_loaded_once'): return
        self.log_message("正在加载配置到UI...")
        self._is_initializing = True

        try:
            # 加载 API Key 和 Model ID
            for field in ['first_api_key', 'first_modelID', 'second_api_key', 'second_modelID']:
                widget = self.get_ui_element(field, QLineEdit)
                if widget and isinstance(widget, QLineEdit):
                    widget.setText(getattr(self.config_manager, field, ""))
            
            # --- 核心修改: 加载 Provider 并设置 ComboBox ---
            provider_map = {
                'first_api_url': self.config_manager.first_api_provider,
                'second_api_url': self.config_manager.second_api_provider,
            }
            for combo_name, provider_id in provider_map.items():
                combo_box = self.get_ui_element(combo_name, QComboBox)
                if combo_box and isinstance(combo_box, QComboBox):
                    # 将内部ID (如 "volcengine") 转换为UI文本 (如 "火山引擎 (推荐)")
                    ui_text_to_select = get_ui_text_from_provider_id(provider_id)
                    if ui_text_to_select:
                        combo_box.setCurrentText(ui_text_to_select)
                    else:
                        combo_box.setCurrentIndex(0) # 如果找不到，默认选第一个

            # 加载其他配置 (保持不变)
            subject_widget = self.get_ui_element('subject_text', QComboBox)
            if subject_widget: subject_widget.setCurrentText(self.config_manager.subject)
            
            cycle_element = self.get_ui_element('cycle_number')
            if cycle_element and isinstance(cycle_element, QSpinBox):
                cycle_element.setValue(self.config_manager.cycle_number)
            
            wait_element = self.get_ui_element('wait_time')
            if wait_element and isinstance(wait_element, QSpinBox):
                wait_element.setValue(self.config_manager.wait_time)

            dual_element = self.get_ui_element('dual_evaluation_enabled', QCheckBox)
            if dual_element and isinstance(dual_element, QCheckBox):
                dual_element.setChecked(self.config_manager.dual_evaluation_enabled)
            
            threshold_element = self.get_ui_element('score_diff_threshold')
            if threshold_element and isinstance(threshold_element, QSpinBox):
                threshold_element.setValue(self.config_manager.score_diff_threshold)

            # 加载无人模式配置
            unattended_element = self.get_ui_element('unattended_mode_enabled', QCheckBox)
            if unattended_element and isinstance(unattended_element, QCheckBox):
                unattended_element.setChecked(self.config_manager.unattended_mode_enabled)


            # 加载题目配置 (支持8道题)
            for i in range(1, self.max_questions + 1):
                q_config = self.config_manager.get_question_config(i)
                
                # 加载评分细则
                std_answer = self.get_ui_element(f'StandardAnswer_text_{i}')
                if std_answer and isinstance(std_answer, QPlainTextEdit): 
                    std_answer.setPlainText(q_config.get('standard_answer', ''))
                
                # 加载启用状态
                enable_cb = self.get_ui_element(f'enableQuestion{i}')
                if enable_cb and i > 1 and isinstance(enable_cb, QCheckBox):  # 第一题始终启用
                    enable_cb.setChecked(q_config.get('enabled', False))
                
                # 加载每题独立的步长
                step_combo = self.get_ui_element(f'score_rounding_step_{i}')
                if step_combo and isinstance(step_combo, QComboBox):
                    step_value = q_config.get('score_rounding_step', 0.5)
                    # 将步长值转为显示文本，支持 0.5, 1, 1.5, 2
                    # 整数显示为不带小数点的形式（如 1），浮点数保持小数形式（如 0.5, 1.5）
                    if step_value == int(step_value):
                        step_text = str(int(step_value))
                    else:
                        step_text = str(step_value)
                    step_combo.setCurrentText(step_text)
                

            # 加载完成后，应用所有UI约束
            self._apply_ui_constraints()
            # 强制切换到第一小题，确保每次启动默认显示第1题
            try:
                tab_widget = self.get_ui_element('questionTabs')
                if tab_widget:
                    tab_widget.setCurrentIndex(0)
                    self.current_question = 1
            except Exception:
                pass

            self.log_message("配置已成功加载到UI并应用约束。")
            self._config_loaded_once = True
        except Exception as e:
            self.log_message(f"加载配置到UI时出错: {e}\n{traceback.format_exc()}", is_error=True)
        finally:
            self._is_initializing = False

    # ==========================================================================
    #  按钮点击和事件处理 (大部分保持不变)
    # ==========================================================================

    def auto_run_but_clicked(self):
        """自动运行按钮点击事件"""
        # 先做启动前校验（包含：供应商UI文本→内部ID归一化、必要坐标检查等），避免“保存了错误配置”或“启动→秒停”。
        if not self.check_required_settings():
            return

        self.log_message("尝试在运行前保存所有配置...")
        if not self.config_manager.save_all_configs_to_file():
            self.log_message("错误：运行前保存配置失败！无法启动自动阅卷。", is_error=True)
            # 创建完整显示的保存失败提示框
            msg_box = QMessageBox(self)
            msg_box.setIcon(QMessageBox.Critical)
            msg_box.setWindowTitle("保存失败")
            msg_box.setText("保存配置失败。\n\n请检查下方日志以获取更多详细信息。")
            msg_box.setSizeGripEnabled(True)
            msg_box.setMinimumSize(500, 200)
            msg_box.setStyleSheet("QLabel{min-width: 400px;}")
            msg_box.setStandardButtons(QMessageBox.Ok)
            msg_box.exec_()
            return
        self.log_message("所有配置已成功保存。")

        # 显示提醒对话框
        msg_box = QMessageBox(self)
        msg_box.setIcon(QMessageBox.Warning)
        msg_box.setWindowTitle("重要提醒")
        msg_box.setText("AI阅卷时，务必关闭阅卷记录文件，否则阅卷记录无法被记录。\n\n请确认您已关闭Excel文件。")
        ok_button = msg_box.addButton("我确认阅卷记录Excel文件已关闭，开始自动阅卷", QMessageBox.AcceptRole)
        msg_box.setDefaultButton(ok_button)
        msg_box.exec_()

        # 用户确认后，直接启动自动阅卷（无延迟）
        self._start_auto_evaluation_after_confirmation()

    def _start_auto_evaluation_after_confirmation(self):
        """用户确认后延迟启动自动阅卷"""
        try:
            # 多题模式：获取所有启用的题目
            enabled_questions_indices = self.config_manager.get_enabled_questions()
            
            if not enabled_questions_indices:
                msg_box = QMessageBox(self)
                msg_box.setIcon(QMessageBox.Warning)
                msg_box.setWindowTitle("配置不完整")
                msg_box.setText("没有启用任何题目。\n\n请至少启用一道题目。")
                msg_box.setStandardButtons(QMessageBox.Ok)
                msg_box.exec_()
                return

            # 检查所有启用题目的答案区域配置
            missing_configs = []
            for q_idx in enabled_questions_indices:
                q_config = self.config_manager.get_question_config(q_idx)
                if not q_config or 'answer_area' not in q_config or not q_config['answer_area']:
                    missing_configs.append(f"第{q_idx}题")
            
            if missing_configs:
                msg_box = QMessageBox(self)
                msg_box.setIcon(QMessageBox.Warning)
                msg_box.setWindowTitle("配置不完整")
                msg_box.setText(f"以下题目未配置答案区域：\n{', '.join(missing_configs)}\n\n请在题目配置对话框中设置答案区域坐标。")
                msg_box.setSizeGripEnabled(True)
                msg_box.setMinimumSize(500, 150)
                msg_box.setStyleSheet("QLabel{min-width: 400px;}")
                msg_box.setStandardButtons(QMessageBox.Ok)
                msg_box.exec_()
                return

            # 准备参数给 AutoThread
            dual_evaluation = self.config_manager.dual_evaluation_enabled
            
            # 多题模式下禁用双评（只有单题时才能双评）
            if len(enabled_questions_indices) > 1 and dual_evaluation:
                dual_evaluation = False
                # 更新UI复选框状态，确保UI与实际行为一致
                dual_eval_checkbox = self.get_ui_element('dualEvaluationCheckbox')
                if dual_eval_checkbox:
                    dual_eval_checkbox.setChecked(False)
                self.log_message("多题模式下自动禁用双评功能", is_error=False)

            question_configs_for_worker = []
            for q_idx in enabled_questions_indices:
                q_config = self.config_manager.get_question_config(q_idx).copy()
                q_config['question_index'] = q_idx
                q_config['dual_eval_enabled'] = dual_evaluation
                question_configs_for_worker.append(q_config)

            params = {
                'cycle_number': self.config_manager.cycle_number,
                'wait_time': self.config_manager.wait_time,
                'question_configs': question_configs_for_worker,
                'dual_evaluation': dual_evaluation,
                'score_diff_threshold': self.config_manager.score_diff_threshold,
                'first_model_id': self.config_manager.first_modelID,
                'second_model_id': self.config_manager.second_modelID,
                'is_single_question_one_run': len(enabled_questions_indices) == 1,
                # 无人模式配置
                'unattended_mode_enabled': self.config_manager.unattended_mode_enabled,
                'unattended_retry_delay': self.config_manager.unattended_retry_delay,
                'unattended_max_retry_rounds': self.config_manager.unattended_max_retry_rounds,
            }

            self.worker.set_parameters(**params)
            
            # === 重要：在启动阅卷前，隐藏所有答题框窗口和最小化主窗口 ===
            # 1. 隐藏所有答题框窗口
            for q_idx, answer_window in list(self.answer_windows.items()):
                if answer_window and answer_window.isVisible():
                    answer_window.hide()
                    self.log_message(f"已隐藏第{q_idx}题答题框窗口")
            
            # 2. 最小化主窗口，避免遮挡答题卡
            self.showMinimized()
            self.log_message("主窗口已最小化，准备开始截图和阅卷")
            
            self.worker.start()
            self.update_ui_state(is_running=True)
            
            questions_str = ', '.join([f"第{i}题" for i in enabled_questions_indices])
            self.log_message(f"自动阅卷已启动: 批改 {questions_str}，循环 {params['cycle_number']} 次")

        except Exception as e:
            self.log_message(f"启动自动阅卷出错: {e}", is_error=True)
            traceback.print_exc()

    def check_required_settings(self):
        """检查必要的设置是否已配置"""
        errors = []
        def _resolve_provider_to_id(value: str) -> str:
            v = (value or "").strip()
            if not v:
                return ""
            if v in PROVIDER_CONFIGS:
                return v
            mapped = get_provider_id_from_ui_text(v)
            return mapped or ""

        def _is_valid_pos(pos) -> bool:
            if not pos:
                return False
            if not isinstance(pos, (tuple, list)) or len(pos) != 2:
                return False
            try:
                x, y = int(pos[0]), int(pos[1])
            except Exception:
                return False
            return not (x == 0 and y == 0)

        # --- AI供应商配置：允许用户UI文本，但启动前必须能解析为内部ID ---
        first_provider_id = _resolve_provider_to_id(getattr(self.config_manager, 'first_api_provider', ''))
        if not first_provider_id:
            errors.append("请为第一组API选择一个有效的供应商")
        else:
            # 写回内存，确保后续保存会落盘为内部ID
            self.config_manager.update_config_in_memory('first_api_provider', first_provider_id)

        if not self.config_manager.first_api_key.strip():
            errors.append("第一组API的密钥不能为空")
        if not self.config_manager.first_modelID.strip():
            errors.append("第一组API的模型ID不能为空")

        # 始终要求配置第二组API（用于故障转移）
        second_provider_id = _resolve_provider_to_id(getattr(self.config_manager, 'second_api_provider', ''))
        if not second_provider_id:
            errors.append("请为第二组API选择一个有效的供应商（用于API故障转移）")
        else:
            self.config_manager.update_config_in_memory('second_api_provider', second_provider_id)

        if not self.config_manager.second_api_key.strip():
            errors.append("第二组API的密钥不能为空（用于API故障转移）")
        if not self.config_manager.second_modelID.strip():
            errors.append("第二组API的模型ID不能为空（用于API故障转移）")

        # 检查所有启用的题目的评分细则、答案区域、以及必要坐标（分数输入/确认按钮/三步输入）
        enabled_questions = self.config_manager.get_enabled_questions()

        is_single_q1_run = (len(enabled_questions) == 1 and enabled_questions[0] == 1)
        q1_cfg = self.config_manager.get_question_config(1)
        q1_three_step = bool(q1_cfg.get('enable_three_step_scoring', False))

        for q_idx in enabled_questions:
            q_cfg = self.config_manager.get_question_config(q_idx)
            if not q_cfg.get('standard_answer', '').strip():
                errors.append(f"第{q_idx}题已启用但未设置评分细则")
            if not q_cfg.get('answer_area'):
                errors.append(f"第{q_idx}题已启用但未配置答案区域")

            # 坐标校验：减少“启动→秒停”
            confirm_pos = q_cfg.get('confirm_button_pos')
            if not _is_valid_pos(confirm_pos):
                errors.append(f"第{q_idx}题已启用但未配置确认按钮坐标")

            if q_idx == 1 and is_single_q1_run and q1_three_step:
                p1 = q_cfg.get('score_input_pos_step1')
                p2 = q_cfg.get('score_input_pos_step2')
                p3 = q_cfg.get('score_input_pos_step3')
                if not _is_valid_pos(p1):
                    errors.append("第一题启用三步打分，但未配置步骤1输入坐标")
                if not _is_valid_pos(p2):
                    errors.append("第一题启用三步打分，但未配置步骤2输入坐标")
                if not _is_valid_pos(p3):
                    errors.append("第一题启用三步打分，但未配置步骤3输入坐标")
            else:
                score_pos = q_cfg.get('score_input_pos')
                if not _is_valid_pos(score_pos):
                    errors.append(f"第{q_idx}题已启用但未配置分数输入坐标")

        if errors:
            # --- 优化错误提示 ---
            title = "启动前请完善配置"
            intro = "自动阅卷无法启动，因为缺少以下必要信息：\n"
            error_details = "\n".join([f"  - {e}" for e in errors])
            final_message = f"{intro}\n{error_details}\n\n请在主界面补充完整后再试。"

            # 创建完整显示的错误提示框
            msg_box = QMessageBox(self)
            msg_box.setIcon(QMessageBox.Warning)
            msg_box.setWindowTitle(title)
            msg_box.setText(final_message)
            msg_box.setSizeGripEnabled(True)
            msg_box.setMinimumSize(600, 300)
            msg_box.setStyleSheet("QLabel{min-width: 500px;}")
            msg_box.setStandardButtons(QMessageBox.Ok)
            msg_box.exec_()
            return False
        return True

    def check_excel_files_available(self):
        """检查阅卷记录Excel文件是否可用（未被锁定且无临时文件）"""
        try:
            # 获取当前配置
            dual_evaluation = self.config_manager.dual_evaluation_enabled
            question_config = self.config_manager.get_question_config(1)  # 单题模式只处理第一题
            full_score = question_config.get('max_score', 100) if question_config else 100

            # 手动构建文件路径（避免使用_get_excel_filepath的复杂逻辑）
            now = datetime.datetime.now()
            date_str = now.strftime('%Y年%m月%d日')
            evaluation_type = '双评' if dual_evaluation else '单评'
            excel_filename = f"此题最高{full_score}分_{evaluation_type}.xlsx"

            if getattr(sys, 'frozen', False):
                base_dir = pathlib.Path(sys.executable).parent
            else:
                base_dir = pathlib.Path(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

            record_dir = base_dir / "阅卷记录"
            date_dir = record_dir / date_str
            excel_filepath = date_dir / excel_filename

            # 检查是否存在Excel临时文件（表示文件正在被打开）
            temp_file = date_dir / f"~${excel_filename}"
            if temp_file.exists():
                # 显示提醒对话框
                from PyQt5.QtWidgets import QMessageBox
                msg_box = QMessageBox(self)
                msg_box.setIcon(QMessageBox.Warning)
                msg_box.setWindowTitle("无法开始阅卷")
                msg_box.setText("检测到阅卷记录文件正在被其他程序打开。\n\n请关闭Excel文件，然后自动阅卷才能正常进行。")
                msg_box.setStandardButtons(QMessageBox.Ok)
                msg_box.exec_()
                return False

            # 检查文件是否被锁定
            if excel_filepath.exists() and self.is_file_locked(excel_filepath):
                # 显示提醒对话框
                from PyQt5.QtWidgets import QMessageBox
                msg_box = QMessageBox(self)
                msg_box.setIcon(QMessageBox.Warning)
                msg_box.setWindowTitle("无法开始阅卷")
                msg_box.setText("阅卷记录文件被锁定，请关闭相关程序后重试。")
                msg_box.setStandardButtons(QMessageBox.Ok)
                msg_box.exec_()
                return False

            return True

        except Exception as e:
            # 显示友好的错误提示
            from PyQt5.QtWidgets import QMessageBox
            msg_box = QMessageBox(self)
            msg_box.setIcon(QMessageBox.Warning)
            msg_box.setWindowTitle("检查文件状态失败")
            msg_box.setText("请关闭阅卷记录文件，然后才能正常启动自动阅卷。")
            msg_box.setStandardButtons(QMessageBox.Ok)
            msg_box.exec_()
            self.log_message(f"检查Excel文件可用性时出错: {str(e)}", is_error=True)
            return False

    def is_file_locked(self, filepath):
        """检查文件是否被锁定（主要因被其他进程打开）"""
        try:
            with open(filepath, 'a'):
                pass
            return False
        except PermissionError:
            return True
        except Exception:
            return False
    
    def test_api_connections(self):
        """测试API连接（强制测试两个API）"""
        try:
            # 测试前无需手动更新，因为 ApiService 每次都会从 ConfigManager 获取最新配置
            self.log_message("正在测试API连接...")
            success1, message1 = self.api_service.test_api_connection("first")
            
            # 强制测试第二个API（不管双评模式是否开启）
            self.log_message("正在测试第二个API...")
            success2, message2 = self.api_service.test_api_connection("second")
            
            result_message = (
                f"【第一个API】\n{message1}\n\n"
                f"【第二个API】\n{message2}"
            )
            
            if success1 and success2: 
                self.log_message("测试完成：所有API均可正常使用")
            else: 
                self.log_message("测试完成：部分API无法正常使用", is_error=True)

            # 创建完整显示的API测试结果提示框
            msg_box = QMessageBox(self)
            msg_box.setIcon(QMessageBox.Information if (success1 and success2) else QMessageBox.Warning)
            msg_box.setWindowTitle("API连接测试")
            msg_box.setText(f"{result_message}")
            msg_box.setSizeGripEnabled(True)
            msg_box.setMinimumSize(600, 300)
            msg_box.setStyleSheet("QLabel{min-width: 500px;}")
            msg_box.setStandardButtons(QMessageBox.Ok)
            msg_box.exec_()
        except Exception as e:
            self.log_message(f"API测试出错: {str(e)}", is_error=True)
            # 创建完整显示的API测试错误提示框
            msg_box = QMessageBox(self)
            msg_box.setIcon(QMessageBox.Critical)
            msg_box.setWindowTitle("API测试错误")
            msg_box.setText(f"测试过程中发生错误。\n\n错误详情:\n{str(e)}")
            msg_box.setSizeGripEnabled(True)
            msg_box.setMinimumSize(500, 200)
            msg_box.setStyleSheet("QLabel{min-width: 400px;}")
            msg_box.setStandardButtons(QMessageBox.Ok)
            msg_box.exec_()

    def closeEvent(self, a0: Optional[QCloseEvent]) -> None:
        """窗口关闭事件（优化版）"""
        if self.worker.isRunning():
            self.worker.stop()
            self.worker.wait()  # 等待线程安全退出，这是一个好习惯

        # 遍历字典值的副本，因为我们不需要在循环中修改字典
        for window in list(self.answer_windows.values()):
            try:
                # 直接尝试关闭。
                # 1. 如果窗口还开着，它会被正常关闭。
                # 2. 如果窗口已经关闭但对象还存在，调用 close() 通常是无害的。
                # 3. 如果底层对象已被删除，这里会立即触发 RuntimeError。
                window.close()
            except RuntimeError:
                # 捕获到错误，说明这个窗口引用已经失效。
                # 我们什么都不用做，只需安静地忽略它即可。
                self.log_message("一个答案窗口在主窗口关闭前已被销毁，跳过关闭操作。")
                pass

        # 循环结束后，清空字典，确保所有（可能无效的）引用都被清除
        self.answer_windows.clear()

        # --- 保存配置的逻辑保持不变 ---
        self.log_message("尝试在关闭程序前保存所有配置...")
        if not self.config_manager.save_all_configs_to_file():
            self.log_message("警告：关闭程序前保存配置失败。", is_error=True)
        else:
            self.log_message("所有配置已在关闭前成功保存。")

        if a0:
            a0.accept()

    # --- 以下是其他未发生重大逻辑变化的函数，保持原样 ---
    # ... (包括 on_dual_evaluation_changed, _apply_ui_constraints, on_worker_finished, get_ui_element, 等等) ...
    # ... 请将您原文件中的这些函数复制到这里 ...
    # 为了完整性，我将提供这些函数的简化版或完整版
    
    def on_dual_evaluation_changed(self, state):
        if self._is_initializing: return
        is_enabled = bool(state)
        self.handle_checkBox_save('dual_evaluation_enabled', is_enabled)
        self._apply_ui_constraints()

    def _is_single_q1_mode(self):
        """检查当前是否只启用了第一题"""
        for i in range(2, self.max_questions + 1):
            cb = self.get_ui_element(f'enableQuestion{i}')
            if cb and cb.isChecked():
                return False
        return True

    def _apply_ui_constraints(self):
        is_single_q1_mode = self._is_single_q1_mode()

        dual_eval_checkbox = self.get_ui_element('dual_evaluation_enabled')
        unattended_checkbox = self.get_ui_element('unattended_mode_enabled')
        
        # 获取双评和无人模式的当前状态
        is_dual_enabled = dual_eval_checkbox and dual_eval_checkbox.isChecked() if dual_eval_checkbox else False
        is_unattended_enabled = unattended_checkbox and unattended_checkbox.isChecked() if unattended_checkbox else False
        
        if dual_eval_checkbox:
            dual_eval_checkbox.setEnabled(is_single_q1_mode)
            if not is_single_q1_mode and dual_eval_checkbox.isChecked():
                dual_eval_checkbox.blockSignals(True)
                dual_eval_checkbox.setChecked(False)
                self.handle_checkBox_save('dual_evaluation_enabled', False)
                dual_eval_checkbox.blockSignals(False)
                is_dual_enabled = False
            
            is_dual_active = dual_eval_checkbox.isChecked() and dual_eval_checkbox.isEnabled()
            self._safe_set_enabled('score_diff_threshold', is_dual_active)
        
        # 第二组API的启用逻辑：
        # 1. 双评模式启用时需要第二组API
        # 2. 无人模式启用时也需要第二组API（用于故障转移）
        # 3. 或者始终启用（因为单评模式下也需要故障转移）
        # 根据策略文档，单评模式下也强制要求配置两个API
        # 因此始终启用第二组API的配置控件
        second_api_enabled = True  # 始终启用第二组API配置
        self._safe_set_enabled('second_api_url', second_api_enabled)
        self._safe_set_enabled('second_api_key', second_api_enabled)
        self._safe_set_enabled('second_modelID', second_api_enabled)

        q1_config = self.config_manager.get_question_config(1)
        is_q1_three_step_enabled = q1_config.get('enable_three_step_scoring', False)

        # 题目依赖关系：题N只有在题1到题N-1都启用时才能启用
        can_enable_next = True
        for i in range(2, self.max_questions + 1):
            cb_i = self.get_ui_element(f'enableQuestion{i}')
            if not cb_i: continue
            
            should_be_enabled = can_enable_next and not is_q1_three_step_enabled
            cb_i.setEnabled(should_be_enabled)
            
            if not should_be_enabled and cb_i.isChecked():
                cb_i.blockSignals(True)
                cb_i.setChecked(False)
                self.handle_checkBox_save(f'question_{i}_enabled', False)
                cb_i.blockSignals(False)
            
            self.update_config_button(i, cb_i.isChecked())
            can_enable_next = cb_i.isChecked()
            
        # 更新选项卡标签显示状态
        self._update_tab_titles()
    
    def on_question_enabled_changed(self, state):
        if self._is_initializing: return
        sender = self.sender()
        if not sender: return
        try:
            q_index = int(sender.objectName().replace('enableQuestion', ''))
            self.handle_checkBox_save(f"question_{q_index}_enabled", bool(state))
            self._apply_ui_constraints()
        except (ValueError, AttributeError): pass
        
    def update_config_button(self, question_index, is_enabled):
        btn = self.get_ui_element(f'configQuestion{question_index}')
        if btn: btn.setEnabled(is_enabled)
        # 同时控制评分细则输入框和步长选择框
        std_answer = self.get_ui_element(f'StandardAnswer_text_{question_index}')
        if std_answer: std_answer.setEnabled(is_enabled)
        step_combo = self.get_ui_element(f'score_rounding_step_{question_index}')
        if step_combo: step_combo.setEnabled(is_enabled)
    
    def _update_tab_titles(self):
        """更新选项卡标题显示启用状态"""
        tab_widget = self.get_ui_element('questionTabs')
        if not tab_widget: return
        
        # 获取选项卡实际数量，避免访问不存在的索引
        tab_count = tab_widget.count()
        for i in range(1, min(tab_count, self.max_questions) + 1):
            q_config = self.config_manager.get_question_config(i)
            is_enabled = q_config.get('enabled', False) if i > 1 else True
            # 用更醒目的启用标识（✅）替代不太显眼的 ✓
            status_icon = " ✅" if is_enabled else ""
            tab_widget.setTabText(i - 1, f"题目{i}{status_icon}")
        
    def log_message(self, message, is_error=False, level=None):
        """
        显示日志消息，支持级别过滤。

        Args:
            message: 日志消息内容
            is_error: 是否为错误消息（向后兼容）
            level: 日志级别 (INFO, DETAIL, RESULT, ERROR)
        """
        # 自动确定级别（向后兼容）
        if level is None:
            level = self.LOG_LEVEL_ERROR if is_error else self.LOG_LEVEL_INFO

        # 日志过滤：只显示RESULT和ERROR级别的消息
        if level not in [self.LOG_LEVEL_RESULT, self.LOG_LEVEL_ERROR]:
            return

        log_widget = self.get_ui_element('log_text')
        if log_widget:
            if level == self.LOG_LEVEL_ERROR:
                color = "red"
                prefix = "[错误]"
            elif level == self.LOG_LEVEL_RESULT:
                color = "black"
                prefix = "【AI评分依据】"
                # 如果消息以"AI评分依据:"开头，去掉这个前缀
                if message.startswith("AI评分依据:"):
                    message = message[len("AI评分依据:"):].strip()
            else:
                color = "blue"
                prefix = "[信息]"

            # 处理消息内容，添加换行以提高可读性
            formatted_message = message.replace("；", "；<br>")
            
            # AI评分依据另起一行显示，增加空行提高视觉舒适度
            log_widget.append(f'<span style="color:{color}; font-size:14pt;">{prefix}<br>{formatted_message}</span><br>')

        # 控制台始终输出所有消息
        print(f"[{level}] {message}")

    def on_worker_finished(self):
        self.update_ui_state(is_running=False)
    
    def on_worker_error(self, error_message):
        self.log_message(f"任务中断: {error_message}", is_error=True)
        self.update_ui_state(is_running=False)
        
    def update_ui_state(self, is_running):
        self._safe_set_enabled('auto_run_but', not is_running)
        self._safe_set_enabled('stop_but', is_running)
        
        # 禁用所有配置相关控件
        config_controls = [
            'first_api_url', 'first_api_key', 'first_modelID',
            'second_api_url', 'second_api_key', 'second_modelID',
            'dual_evaluation_enabled', 'score_diff_threshold', 'subject_text',
            'cycle_number', 'wait_time', 'api_test_button', 'unattended_mode_enabled'
        ]
        # 支持7道题
        for i in range(1, self.max_questions + 1):
            config_controls.append(f'configQuestion{i}')
            config_controls.append(f'StandardAnswer_text_{i}')
            config_controls.append(f'score_rounding_step_{i}')
            if i > 1: config_controls.append(f'enableQuestion{i}')

        for name in config_controls:
            self._safe_set_enabled(name, not is_running)

        if is_running:
            if not self.isMinimized(): self.showMinimized()
        else:
            if self.isMinimized(): self.showNormal(); self.activateWindow()
            self._apply_ui_constraints() # 任务结束后恢复UI约束

    def stop_auto_thread(self):
        if self.worker.isRunning():
            self.worker.stop()
            self.log_message("已发送停止请求至自动阅卷线程。")
        else:
            self.update_ui_state(is_running=False) # 确保UI状态正确

    # ... 其他如 get_ui_element, open_question_config_dialog 等函数保持原样 ...
    # 您可以将原文件中的这些函数直接复制过来
    def get_ui_element(self, element_name: str, element_type=None) -> Optional[QWidget]:
        """获取UI元素，支持类型提示
        
        Args:
            element_name: 元素名称
            element_type: 期望的元素类型（用于类型检查）
            
        Returns:
            UI元素，如果找不到则返回None
        """
        if element_name in self._ui_cache:
            return self._ui_cache[element_name]
        
        element = cast(Optional[QWidget], self.findChild(QWidget, element_name))
        if element:
            self._ui_cache[element_name] = element
        return element
    
    def _safe_set_enabled(self, element_name: str, enabled: bool) -> None:
        """安全地设置UI元素的enabled状态"""
        element = self.get_ui_element(element_name)
        if element:
            element.setEnabled(enabled)
    
    def _safe_get_spinbox(self, element_name: str) -> Union[QSpinBox, None]:
        """获取并强制转换为QSpinBox"""
        element = self.get_ui_element(element_name)
        if element and isinstance(element, QSpinBox):
            return element
        return None
    
    def _safe_get_checkbox(self, element_name: str) -> Union[QCheckBox, None]:
        """获取并强制转换为QCheckBox"""
        element = self.get_ui_element(element_name)
        if element and isinstance(element, QCheckBox):
            return element
        return None
    
    def _safe_get_combobox(self, element_name: str) -> Union[QComboBox, None]:
        """获取并强制转换为QComboBox"""
        element = self.get_ui_element(element_name)
        if element and isinstance(element, QComboBox):
            return element
        return None
    
    def _safe_get_lineedit(self, element_name: str) -> Union[QLineEdit, None]:
        """获取并强制转换为QLineEdit"""
        element = self.get_ui_element(element_name)
        if element and isinstance(element, QLineEdit):
            return element
        return None
        
    def open_question_config_dialog(self, question_index):
        # 延迟导入以避免循环依赖
        from .question_config_dialog import QuestionConfigDialog

        dialog = QuestionConfigDialog(
            parent=self,
            config_manager=self.config_manager,
            question_index=question_index,
            is_single_q1_mode_active=self._is_single_q1_mode()
        )

        # 连接配置更新信号，确保题目配置保存到文件
        def on_config_updated():
            self.log_message(f"题目{question_index}配置已更新，正在保存到文件...")
            if self.config_manager.save_all_configs_to_file():
                self.log_message("题目配置已成功保存到文件")
            else:
                self.log_message("警告：题目配置保存到文件失败", is_error=True)

        dialog.config_updated.connect(on_config_updated)

        # 在显示配置对话框前隐藏主界面，避免遮挡改卷页面
        self.hide()
        self.log_message(f"配置第{question_index}题信息，主界面已隐藏")
        
        try:
            result = dialog.exec_()
            if result == QDialog.Accepted:
                self.load_config_to_ui()
        finally:
            # 无论保存还是取消，都恢复主界面显示
            self.show()
            self.log_message("配置对话框已关闭，主界面已恢复")

    def get_or_create_answer_window(self, question_index):
        from .question_config_dialog import MyWindow2
        if question_index not in self.answer_windows:
            window = MyWindow2(parent=self, question_index=question_index)
            # 连接窗口关闭信号，用于清理字典
            window.status_changed.connect(
                lambda status, q_idx=question_index: self._on_answer_window_status_changed(q_idx, status)
            )
            self.answer_windows[question_index] = window
        return self.answer_windows[question_index]

    def _on_answer_window_status_changed(self, question_index, status):
        """处理答案框窗口状态变化"""
        if status == "closed":
            # 当窗口关闭时，从字典中移除引用
            if question_index in self.answer_windows:
                self.log_message(f"第{question_index}题答案框窗口已关闭，从字典中移除引用")
                del self.answer_windows[question_index]

    def _get_config_safe(self, section, option, default_value):
        """安全地从配置管理器获取配置值"""
        try:
            if not self.config_manager.parser.has_section(section) or not self.config_manager.parser.has_option(section, option):
                return default_value
            return self.config_manager.parser.get(section, option)
        except Exception:
            return default_value
    
    # ... 您原有的其他辅助函数，如 connect_signals, setup_* 系列函数 ...
    # 这些函数的内部逻辑基本不需要大改，因为它们大多是连接信号或设置简单的UI属性
    # 我在这里提供简化版，您可以与您的版本对比
    def connect_signals(self):
        """连接所有UI信号的公开接口"""
        self._connect_signals()

    def setup_question_selector(self):
        # from PyQt5.QtWidgets import QButtonGroup
        # self.question_button_group = QButtonGroup(self)
        # self.question_button_group.buttonClicked.connect(self.on_question_changed)
        pass # 假设UI文件已自动连接

    def on_question_changed(self, button): pass

    def setup_text_fields(self):
        # 支持7道题
        for i in range(1, self.max_questions + 1):
            widget = self.get_ui_element(f'StandardAnswer_text_{i}')
            if widget: widget.setPlaceholderText(f"请输入第{i}题的评分细则...")

        # 设置评分细则和日志的字体为微软雅黑，继承全局字号
        font = QFont("微软雅黑")
        for i in range(1, self.max_questions + 1):
            standard_answer_widget = self.get_ui_element(f'StandardAnswer_text_{i}')
            if standard_answer_widget:
                standard_answer_widget.setFont(font)
        log_widget = self.get_ui_element('log_text')
        if log_widget:
            log_widget.setFont(font)

    def setup_dual_evaluation(self):
        cb = self.get_ui_element('dual_evaluation_enabled')
        if cb: cb.stateChanged.connect(self.on_dual_evaluation_changed)
        spin = self.get_ui_element('score_diff_threshold')
        if spin: spin.valueChanged.connect(lambda val: self.handle_spinBox_save('score_diff_threshold', val))

    def setup_unattended_mode(self):
        """设置无人模式相关控件的信号连接"""
        cb = self.get_ui_element('unattended_mode_enabled')
        if cb: 
            cb.stateChanged.connect(self.on_unattended_mode_changed)

    def on_unattended_mode_changed(self, state):
        """无人模式开关变化处理"""
        if self._is_initializing: return
        is_enabled = bool(state)
        self.handle_checkBox_save('unattended_mode_enabled', is_enabled)
        self._apply_ui_constraints()
        
        # 提示用户无人模式的含义
        if is_enabled:
            self.log_message("无人模式已启用：两个API都失败时将自动重试，直到成功或达到最大重试次数")
        else:
            self.log_message("无人模式已禁用：两个API都失败时将立即停止并等待人工介入")


    def on_subject_changed(self, index):
        # 此函数在我的重构中未直接使用，但如果您需要它，可以这样实现
        combo = self.sender()
        if combo and isinstance(combo, QComboBox): self.handle_comboBox_save('subject', combo.currentText())

    def _connect_signals(self):
        """统一连接所有UI控件的信号与槽"""
        # 连接按钮点击
        auto_btn = self.get_ui_element('auto_run_but')
        if auto_btn and isinstance(auto_btn, QPushButton):
            auto_btn.clicked.connect(self.auto_run_but_clicked)
        
        stop_btn = self.get_ui_element('stop_but')
        if stop_btn and isinstance(stop_btn, QPushButton):
            stop_btn.clicked.connect(self.stop_auto_thread)
        
        test_btn = self.get_ui_element('api_test_button')
        if test_btn and isinstance(test_btn, QPushButton):
            test_btn.clicked.connect(self.test_api_connections)
        
        # 支持7道题的配置按钮
        for i in range(1, self.max_questions + 1):
            btn = self.get_ui_element(f'configQuestion{i}')
            if btn and isinstance(btn, QPushButton):
                btn.clicked.connect(lambda checked, q=i: self.open_question_config_dialog(q))

        # 连接即时保存信号
        self._connect_direct_edit_save_signals()

        # 连接题目启用复选框（支持7道题）
        for i in range(2, self.max_questions + 1):
            checkbox = self.get_ui_element(f'enableQuestion{i}')
            if checkbox:
                checkbox.stateChanged.connect(self.on_question_enabled_changed)
        
        # 连接每题独立步长选择框的信号
        for i in range(1, self.max_questions + 1):
            step_combo = self.get_ui_element(f'score_rounding_step_{i}', QComboBox)
            if step_combo:
                step_combo.currentTextChanged.connect(
                    lambda text, q_idx=i: self._on_step_changed(q_idx, text)
                )
        

    def _on_step_changed(self, question_index, text):
        """处理每题步长选择变化"""
        if self._is_initializing: return
        try:
            step_value = float(text)
            self.config_manager.update_question_config(question_index, 'score_rounding_step', step_value)
            self.log_message(f"第{question_index}题步长更新为: {step_value}")
        except (ValueError, TypeError):
            pass  # 忽略无效的步长值
    
    



# --- END OF FILE main_window.py ---

