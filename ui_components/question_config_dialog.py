# question_config_dialog.py - 题目配置对话框模块

from typing import Optional, Any
from PyQt5.QtWidgets import (QApplication, QDialog, QVBoxLayout, QHBoxLayout, QLabel, 
                            QLineEdit, QPushButton, QCheckBox,
                            QComboBox, QTextEdit, QSpinBox, QMessageBox,
                            QMainWindow, QWidget, QSizePolicy)
from PyQt5.QtCore import Qt, QPoint, QRect, pyqtSignal, QTimer
from PyQt5.QtGui import QPainter, QColor, QPen, QFont, QMouseEvent, QPaintEvent, QCloseEvent, QShowEvent
import pyautogui
import time

class MyWindow2(QMainWindow):
    """答案框窗口类，用于框定答案区域"""
    
    # 添加信号，用于通知主窗口状态变化
    status_changed = pyqtSignal(str)
    
    def update_ui_state(self, is_running):
        """代理方法，转发到主窗口"""
        parent_window = self.parent()
        if parent_window is not None and hasattr(parent_window, 'update_ui_state'):
            parent_window.update_ui_state(is_running)  # type: ignore
    
    def __init__(self, parent=None, question_index=None):
        super(MyWindow2, self).__init__(parent)

        # 固定答案框窗口字号为 11（与题目配置信息界面一致，调大一号）
        try:
            self.setFont(QFont("微软雅黑", 11))
        except Exception:
            pass

        self.question_index = question_index

        # 根据题目编号设置不同的窗口标题
        title = f"第{question_index}题答案框" if question_index else "答案框"
        self.setWindowTitle(title)
        self.resize(400, 300)

        # 设置窗口样式 - 无边框、工具窗口（移除WindowStaysOnTopHint，避免与父窗口冲突）
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.Tool)  # type: ignore
        
        # 设置窗口属性为透明背景
        self.setAttribute(Qt.WA_TranslucentBackground)  # type: ignore
        
        # 创建中央窗口部件
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)
        
        # 设置窗口背景为高透明度
        central_widget.setStyleSheet("background-color: rgba(200, 200, 255, 30);")
        
        # 初始化拖动和调整大小相关变量
        self.drag_position = QPoint()
        self.resizing = False
        self.resize_edge = None
        self.border_width = 8  # 边框宽度，用于检测鼠标是否在边缘
        
        # 添加锁定状态变量
        self.is_locked = False
        self.is_confirmed = False
        
        # 移除窗口提示，因为用户未看到且可能不需要
        pass
        
        print(f"第{question_index}题答案框窗口已创建" if question_index else "答案框窗口已创建")
        parent_window = self.parent()
        if parent_window is not None and hasattr(parent_window, 'log_message'):
            parent_window.log_message(f"第{question_index}题答案框窗口已创建" if question_index else "答案框窗口已创建")  # type: ignore
    
    def set_confirmed_mode(self):
        """设置为已确认模式，更改提示文字和透明度"""
        self.is_confirmed = True
        # 增加透明度，使窗口更不明显但仍然可见
        widget = self.centralWidget()
        if widget is not None:
            widget.setStyleSheet("background-color: rgba(200, 200, 255, 15);")
        # 自动锁定窗口
        self.is_locked = True
        # 发送状态变化信号
        self.status_changed.emit("confirmed")
        self.update()  # 触发重绘
    
    def set_edit_mode(self):
        """设置为编辑模式，允许调整位置和大小"""
        log_msg_start = "正在设置答案框为编辑模式..."
        print(log_msg_start) # 保留 print 语句
        parent_window = self.parent()
        if parent_window is not None and hasattr(parent_window, 'log_message'):
            parent_window.log_message(log_msg_start)  # type: ignore
        
        # 彻底重置所有锁定状态
        self.is_confirmed = False
        self.is_locked = False
        self.resizing = False
        self.resize_edge = None
        
        # 恢复原来的透明度
        widget = self.centralWidget()
        if widget is not None:
            widget.setStyleSheet("background-color: rgba(200, 200, 255, 30);")
            # 确保鼠标追踪已启用
            widget.setMouseTracking(True)
        
        self.setMouseTracking(True)
        
        # 重置光标为默认
        self.setCursor(Qt.ArrowCursor)  # type: ignore
        
        # 发送状态变化信号
        self.status_changed.emit("editing")
        
        # 强制更新UI
        self.update()  # 触发重绘
        
        # 验证状态是否正确设置
        print(f"编辑模式设置完成 - 锁定状态: {self.is_locked}, 确认状态: {self.is_confirmed}")
        parent_window = self.parent()
        if parent_window is not None and hasattr(parent_window, 'log_message'):
            parent_window.log_message(f"编辑模式设置完成 - 锁定状态: {self.is_locked}, 确认状态: {self.is_confirmed}")  # type: ignore
    
    def paintEvent(self, a0: Optional[QPaintEvent]) -> None:
        """绘制窗口边框和提示文字"""
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        
        # 绘制边框
        pen = QPen(QColor(255, 0, 0), 2)
        painter.setPen(pen)
        painter.drawRect(0, 0, self.width() - 1, self.height() - 1)
        
        font = QFont("微软雅黑")
        # 让提示文字跟随本窗口字号（题目配置相关界面固定为 11）
        try:
            base_font = self.font()
            if base_font and base_font.pointSize() > 0:
                font.setPointSize(base_font.pointSize())
        except Exception:
            pass
        font.setBold(True)
        painter.setFont(font)
        
        # 根据题目显示不同的提示文字
        if self.is_confirmed:
            text1 = f"第{self.question_index}题答案区域已框定" if self.question_index else "答案区域已框定"
            text2 = "学生答案不可被遮挡"
        else:
            text1 = f"此框覆盖第{self.question_index}题学生答案区域" if self.question_index else "请使此框覆盖学生答案"
            text2 = "在配置窗口点击'确认框定'"
        
        # 计算文本宽度以便居中显示
        text1_width = painter.fontMetrics().width(text1)
        text2_width = painter.fontMetrics().width(text2)
        
        # 绘制文字
        painter.setPen(QColor(255, 0, 0, 255))
        combined_text = f"{text1} {text2}"
        combined_width = painter.fontMetrics().width(combined_text)
        combined_x = (self.width() - combined_width) // 2
        painter.drawText(combined_x, 17, combined_text)
    
    def mousePressEvent(self, a0: Optional[QMouseEvent]) -> None:
        """鼠标按下事件，用于移动窗口或调整大小"""
        # 如果窗口已锁定，则忽略鼠标事件
        if self.is_locked:
            return
        
        if a0 is None:
            return
            
        if a0.button() == Qt.LeftButton:  # type: ignore
            # 检查是否在边缘
            rect = self.rect()
            pos = a0.pos()
            
            # 检测边缘区域
            left_edge = pos.x() <= self.border_width
            right_edge = pos.x() >= rect.width() - self.border_width
            top_edge = pos.y() <= self.border_width
            bottom_edge = pos.y() >= rect.height() - self.border_width
            
            if left_edge or right_edge or top_edge or bottom_edge:
                self.resizing = True
                self.resize_edge = {
                    'left': left_edge,
                    'right': right_edge,
                    'top': top_edge,
                    'bottom': bottom_edge
                }
                a0.accept()
            else:
                # 如果不在边缘，则为拖动窗口
                self.drag_position = a0.globalPos() - self.frameGeometry().topLeft()
                a0.accept()
    
    def mouseMoveEvent(self, a0: Optional[QMouseEvent]) -> None:
        """鼠标移动事件，用于移动窗口或调整大小"""
        # 如果窗口已锁定，则只更新光标形状为默认
        if self.is_locked:
            self.setCursor(Qt.ArrowCursor)  # type: ignore
            return
        
        if a0 is None:
            return
            
        # 获取鼠标位置
        pos = a0.pos()
        rect = self.rect()
        
        # 检测边缘区域
        left_edge = pos.x() <= self.border_width
        right_edge = pos.x() >= rect.width() - self.border_width
        top_edge = pos.y() <= self.border_width
        bottom_edge = pos.y() >= rect.height() - self.border_width
        
        # 根据鼠标位置更新光标形状
        if (left_edge and top_edge) or (right_edge and bottom_edge):
            self.setCursor(Qt.SizeFDiagCursor)  # type: ignore  # 左上-右下调整
        elif (right_edge and top_edge) or (left_edge and bottom_edge):
            self.setCursor(Qt.SizeBDiagCursor)  # type: ignore  # 右上-左下调整
        elif left_edge or right_edge:
            self.setCursor(Qt.SizeHorCursor)  # type: ignore    # 水平调整
        elif top_edge or bottom_edge:
            self.setCursor(Qt.SizeVerCursor)  # type: ignore    # 垂直调整
        else:
            self.setCursor(Qt.ArrowCursor)  # type: ignore      # 默认光标
        
        # 处理拖动和调整大小
        if self.resizing and a0.buttons() & Qt.LeftButton and self.resize_edge:  # type: ignore
            # 调整大小
            global_pos = a0.globalPos()
            rect = self.geometry()
            new_rect = QRect(rect)
            
            if self.resize_edge.get('left'):  # type: ignore
                # 调整左边缘
                width_diff = rect.left() - global_pos.x()
                if rect.width() + width_diff >= 50:  # 最小宽度限制
                    new_rect.setLeft(global_pos.x())
            
            if self.resize_edge.get('right'):  # type: ignore
                # 调整右边缘
                new_rect.setRight(global_pos.x())
            
            if self.resize_edge.get('top'):  # type: ignore
                # 调整上边缘
                height_diff = rect.top() - global_pos.y()
                if rect.height() + height_diff >= 50:  # 最小高度限制
                    new_rect.setTop(global_pos.y())
            
            if self.resize_edge.get('bottom'):  # type: ignore
                # 调整下边缘
                new_rect.setBottom(global_pos.y())
            
            # 确保窗口大小不小于最小值
            if new_rect.width() >= 50 and new_rect.height() >= 50:
                self.setGeometry(new_rect)
            
            a0.accept()
        
        elif a0.buttons() == Qt.LeftButton and not self.resizing:  # type: ignore
            # 移动窗口
            self.move(a0.globalPos() - self.drag_position)
            a0.accept()
    
    def mouseReleaseEvent(self, a0: Optional[QMouseEvent]) -> None:
        """鼠标释放事件"""
        if a0 and a0.button() == Qt.LeftButton:  # type: ignore
            self.resizing = False
            self.resize_edge = None
            a0.accept()
    
    def showEvent(self, a0: Optional[QShowEvent]) -> None:
        """窗口显示事件"""
        if a0 is not None:
            super().showEvent(a0)
        # 窗口显示时，设置为编辑模式
        self.set_edit_mode()
        print("答案框窗口显示事件触发")
        parent_window = self.parent()
        if parent_window is not None and hasattr(parent_window, 'log_message'):
            parent_window.log_message("答案框窗口显示事件触发")  # type: ignore

    def closeEvent(self, a0: Optional[QCloseEvent]) -> None:
        """窗口关闭事件"""
        print("答案框窗口关闭事件触发")
        parent_window = self.parent()
        if parent_window is not None and hasattr(parent_window, 'log_message'):
            parent_window.log_message("答案框窗口关闭事件触发")  # type: ignore
        # 发送窗口关闭信号
        self.status_changed.emit("closed")
        if a0 is not None:
            a0.accept()

class QuestionConfigDialog(QDialog):
    """题目配置对话框，用于配置每个题目的评分参数"""
    config_updated = pyqtSignal() # 添加信号
    
    def __init__(self, parent: Optional[Any] = None, config_manager: Optional[Any] = None, question_index: int = 1, is_single_q1_mode_active: bool = False):
        super().__init__(parent)
        
        # 设置为非模态对话框，允许同时编辑其他窗口
        self.setModal(False)
        
        # ✨ 设置此对话框的全局字体大小为 11（覆盖应用全局的 11），调大一号
        # 便于未来调整，只需修改下面的 11 即可
        self._apply_font_size(font_size=11)
        
        self.parent_window: Optional[Any] = parent
        self.config_manager: Optional[Any] = config_manager
        self.question_index: int = question_index
        self.is_single_q1_mode_active: bool = is_single_q1_mode_active
        
        # 获取当前题目配置
        self.question_config: dict = self.config_manager.get_question_config(question_index) if self.config_manager else {}  # 确保 q_config 是字典

        # 新增三步打分相关UI元素（仅第一题使用）
        self.three_step_scoring_checkbox: Optional[QCheckBox] = None
        self.score_input_group_step1: Optional[QWidget] = None # QWidget for step1
        self.score_x_edit_step1: Optional[QSpinBox] = None
        self.score_y_edit_step1: Optional[QSpinBox] = None
        self.set_pos_button_step1: Optional[QPushButton] = None
        # ... (为 step2 和 step3 添加类似的变量)
        self.score_input_group_step2: Optional[QWidget] = None
        self.score_x_edit_step2: Optional[QSpinBox] = None
        self.score_y_edit_step2: Optional[QSpinBox] = None
        self.set_pos_button_step2: Optional[QPushButton] = None
        self.score_input_group_step3: Optional[QWidget] = None
        self.score_x_edit_step3: Optional[QSpinBox] = None
        self.score_y_edit_step3: Optional[QSpinBox] = None
        self.set_pos_button_step3: Optional[QPushButton] = None

        # 原有的单点分数输入组也需要一个引用，方便控制其启用/禁用
        self.original_score_input_group: Optional[QWidget] = None
        self.question_type_combo: Optional[QComboBox] = None # 新增题目类型下拉框成员变量

        # 在 self.init_ui() 之前添加
        self.position_capture_timer: Optional[QTimer] = None

        self.init_ui()

    def _apply_font_size(self, font_size: int = 11) -> None:
        """
        为此对话框及其所有子控件统一应用字体大小。
        
        通过 QSS 强制覆盖应用全局字体大小设置，便于未来调整维护。
        只需修改此方法的参数即可全局改变对话框字号。
        
        Args:
            font_size: 字体大小（单位：磅），默认为 11
        """
        font = QFont("微软雅黑", font_size)
        self.setFont(font)
        # 通过样式表确保所有子控件都使用相同字号，覆盖应用全局设置
        self.setStyleSheet(f"""
            QDialog, QDialog * {{
                font-size: {font_size}pt;
                font-family: '微软雅黑';
            }}
        """)

    def _compact_layout(self, layout, spacing: int = 4, margins: tuple = (6, 4, 6, 4)) -> None:
        """Apply compact spacing and margins to a QLayout."""
        try:
            layout.setSpacing(spacing)
        except Exception:
            pass
        try:
            layout.setContentsMargins(*margins)
        except Exception:
            pass

    def _log_message(self, message: str, is_error: bool = False) -> None:
        """Helper method to safely log messages to parent window"""
        if self.parent_window and hasattr(self.parent_window, 'log_message'):
            self.parent_window.log_message(message, is_error=is_error)  # type: ignore
        else:
            if is_error:
                print(f"ERROR: {message}")
            else:
                print(message)

        # 初始化UI
    def init_ui(self):
        """初始化UI组件"""
        self.setWindowTitle(f'配置第{self.question_index}题')
        # --- CRITICAL FIX: 移除 Qt.WindowStaysOnTopHint ---
        # 移除这个标志，以避免它遮挡新弹出的答案框选窗口
        self.setWindowFlags(self.windowFlags() | Qt.WindowMinimizeButtonHint)  # type: ignore
        # 缩窄对话框并减少高度以获得更紧凑的垂直布局
        self.resize(450, 440)
        # 强制最大宽度，确保对话框在不同平台上保持较窄的视觉效果
        try:
            self.setFixedWidth(450)
        except Exception:
            pass
        
        # 主布局 - 更紧凑的垂直间距
        main_layout = QVBoxLayout()
        main_layout.setSpacing(1)  # 从2减少到1，使垂直间距更紧凑
        main_layout.setContentsMargins(4, 1, 4, 1)  # 减少边距
        
        # --- 1. 分数设置组 ---
        score_group = QWidget()
        score_group_content_layout = QHBoxLayout()
        self._compact_layout(score_group_content_layout)
        
        score_group_content_layout.addWidget(QLabel("给分上限:"))
        self.max_score_edit = QSpinBox()
        self.max_score_edit.setMinimum(0)
        self.max_score_edit.setMaximum(150)
        self.max_score_edit.setValue(self.question_config.get('max_score', 150) if self.question_config else 150)
        score_group_content_layout.addWidget(self.max_score_edit)
        score_group_content_layout.addSpacing(20) # 减少水平伸缩空间以缩窄整体宽度
        
        score_group_content_layout.addWidget(QLabel("给分下限："))
        self.min_score_edit = QSpinBox()
        self.min_score_edit.setMinimum(0)
        self.min_score_edit.setMaximum(150)
        self.min_score_edit.setValue(self.question_config.get('min_score', 0) if self.question_config else 0)
        score_group_content_layout.addWidget(self.min_score_edit)
        
        score_group.setLayout(score_group_content_layout)
        main_layout.addWidget(score_group)

        # --- 新增：题目类型选择 ---
        question_type_group = QWidget() # 可以给它一个组标题
        question_type_layout = QHBoxLayout()
        self._compact_layout(question_type_layout)

        # 创建带图标与重要徽章的题型选择行
        self.question_type_combo = QComboBox()
        # 不使用占位项，直接列出题型选项
        # 定义题目类型选项 (键: 内部标识符, 值: UI显示文本)
        self.question_types_map = {
            "Subjective_PointBased_QA": "主观题  -  按点给分，强调答案命中得分点",
            "Objective_FillInTheBank": "较客观的主观题  -  填空题",
            "Formula_Proof_StepBased": "理科公式计算/证明题：步骤、逻辑、符号",
            "Holistic_Evaluation_Open": "文科开放题  -  宏观标准 综合判断（作文）"
        }
        for display_text in self.question_types_map.values():
            self.question_type_combo.addItem(display_text)

        # 加载当前配置的题目类型（查找对应显示文本）
        current_type_identifier = self.question_config.get('question_type', 'Subjective_PointBased_QA')
        current_display_text = None
        for identifier, display in self.question_types_map.items():
            if identifier == current_type_identifier:
                current_display_text = display
                break

        if current_display_text:
            current_index = self.question_type_combo.findText(current_display_text)
            if current_index != -1:
                self.question_type_combo.setCurrentIndex(current_index)
        else:
            # 如果没有配置，选择第一个题型项（索引0）
            if self.question_type_combo.count() > 0:
                self.question_type_combo.setCurrentIndex(0)

        # 右侧重要徽章
        required_badge = QLabel("重要")
        # 更紧凑的徽章样式：减小字体与垂直内边距
        required_badge.setStyleSheet("background-color: #FF5555; color: white; padding:1px 6px; border-radius:8px; font-size:11pt; font-weight:bold;")
        # 使用固定高度并设置为固定尺寸策略，保证与下拉框垂直居中对齐
        required_badge.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        try:
            combo_h = self.question_type_combo.sizeHint().height()
            required_badge.setFixedHeight(combo_h)
        except Exception:
            required_badge.setFixedHeight(20)

        # 初始样式：较大高度、内边距、橙色边框
        base_combo_style = "min-height:22px; padding:3px; border:2px solid #FF8C00; border-radius:4px;"
        self.question_type_combo.setStyleSheet(base_combo_style)

        # 将组件加入布局（下拉 + 徽章）
        question_type_layout.addWidget(self.question_type_combo)
        question_type_layout.addWidget(required_badge)
        question_type_group.setLayout(question_type_layout)
        main_layout.addWidget(question_type_group)

        # 连接信号：选择变化时更新样式（仅用于保持边框样式可扩展）
        self.question_type_combo.currentIndexChanged.connect(self.update_question_type_style)
        # 触发一次以应用初始样式
        try:
            self.update_question_type_style()
        except Exception:
            pass
        # --- 结束新增 ---

        # --- 2. 分数输入位置设置 ---
        # 使用辅助函数重构，以保持UI统一
        input_group = self._create_position_input_group(
            "设置分数输入位置",
            'score_x_edit',
            'score_y_edit',
            'set_pos_button', # 此属性将在辅助函数中动态创建
            "分数输入"
        )
        main_layout.addWidget(input_group)
        self.original_score_input_group = input_group

        # --- 3. 三步分数输入模式 (仅第一题) ---
        if self.question_index == 1:
            three_step_group = QWidget()
            three_step_group_main_layout = QVBoxLayout() 
            self._compact_layout(three_step_group_main_layout)

            self.three_step_scoring_checkbox = QCheckBox("60'作文专用，分数分3份输入3个位置 (仅当第1题）")
            enable_three_step_from_config = self.question_config.get('enable_three_step_scoring', False)
            self.three_step_scoring_checkbox.setChecked(enable_three_step_from_config)

            self.three_step_scoring_checkbox.stateChanged.connect(self.toggle_three_step_mode_ui)
            three_step_group_main_layout.addWidget(self.three_step_scoring_checkbox)

            # 使用新的辅助函数创建三个步骤的输入组
            self.score_input_group_step1 = self._create_position_input_group(
                "设置位置 1", 'score_x_edit_step1', 'score_y_edit_step1', 'set_pos_button_step1', "位置 1"
            )
            three_step_group_main_layout.addWidget(self.score_input_group_step1)

            self.score_input_group_step2 = self._create_position_input_group(
                "设置位置 2", 'score_x_edit_step2', 'score_y_edit_step2', 'set_pos_button_step2', "位置 2"
            )
            three_step_group_main_layout.addWidget(self.score_input_group_step2)

            self.score_input_group_step3 = self._create_position_input_group(
                "设置位置 3", 'score_x_edit_step3', 'score_y_edit_step3', 'set_pos_button_step3', "位置 3"
            )
            three_step_group_main_layout.addWidget(self.score_input_group_step3)

            # 初始化时，三个位置输入组默认隐藏（未启用三步打分时）
            self.score_input_group_step1.setVisible(False)
            self.score_input_group_step2.setVisible(False)
            self.score_input_group_step3.setVisible(False)

            if not self.is_single_q1_mode_active:
                self.three_step_scoring_checkbox.setChecked(False)
                self.three_step_scoring_checkbox.setEnabled(False)

            # 确保在所有相关UI元素创建并赋值后调用
            if enable_three_step_from_config:
                self.toggle_three_step_mode_ui(True)

            three_step_group.setLayout(three_step_group_main_layout)
            
            original_input_group_index = -1
            for i in range(main_layout.count()):
                widget_item = main_layout.itemAt(i)
                if widget_item and widget_item.widget() == self.original_score_input_group:
                    original_input_group_index = i
                    break
            if original_input_group_index != -1:
                main_layout.insertWidget(original_input_group_index + 1, three_step_group)
            else: 
                main_layout.addWidget(three_step_group)


        # --- 4. 提交按钮位置设置 ---
        submit_group = QWidget()
        submit_group_content_layout = QHBoxLayout() 
        self._compact_layout(submit_group_content_layout)
        
        submit_group_content_layout.addWidget(QLabel("坐标:"))
        submit_group_content_layout.addWidget(QLabel("X:"))
        self.submit_x_edit = QLineEdit()
        submit_pos_val = self.question_config.get('confirm_button_pos') if self.question_config else None
        submit_x, submit_y = submit_pos_val if submit_pos_val is not None else (0, 0)
        self.submit_x_edit.setText(str(submit_x))
        self.submit_x_edit.setReadOnly(True)
        self.submit_x_edit.setStyleSheet("background-color: #F2F7FF; color: #5E6F80; border: 1px dashed #AFCBFF; border-radius: 4px; padding: 4px 6px;")
        submit_group_content_layout.addWidget(self.submit_x_edit)
        
        submit_group_content_layout.addWidget(QLabel("Y:"))
        self.submit_y_edit = QLineEdit()
        self.submit_y_edit.setText(str(submit_y))
        self.submit_y_edit.setReadOnly(True)
        self.submit_y_edit.setStyleSheet("background-color: #F2F7FF; color: #5E6F80; border: 1px dashed #AFCBFF; border-radius: 4px; padding: 4px 6px;")
        submit_group_content_layout.addWidget(self.submit_y_edit)
        
        set_submit_button = QPushButton("设置提交按钮位置")
        set_submit_button.clicked.connect(lambda: self.set_position('submit_x_edit', 'submit_y_edit', "提交按钮"))
        submit_group_content_layout.addWidget(set_submit_button)
        
        submit_group.setLayout(submit_group_content_layout)
        main_layout.addWidget(submit_group)
        
        # --- 5. 翻页按钮配置 ---
        next_group = QWidget()
        next_group_main_layout = QVBoxLayout() 
        self._compact_layout(next_group_main_layout)
        
        self.enable_next_check = QCheckBox("启用翻页按钮")
        enable_next_button_for_current_q = self.question_config.get('enable_next_button', False) if self.question_config else False
        self.enable_next_check.setChecked(enable_next_button_for_current_q)
        self.enable_next_check.stateChanged.connect(self.toggle_next_button_fields)
        next_group_main_layout.addWidget(self.enable_next_check) 
        
        next_coord_and_button_layout = QHBoxLayout() 
        self._compact_layout(next_coord_and_button_layout)
        next_coord_and_button_layout.addWidget(QLabel("坐标:"))
        next_coord_and_button_layout.addWidget(QLabel("X:"))
        self.next_x_edit = QLineEdit()
        next_pos_val = self.question_config.get('next_button_pos') if self.question_config else None
        current_next_x, current_next_y = next_pos_val if next_pos_val is not None else (0, 0)
        self.next_x_edit.setText(str(current_next_x))
        self.next_x_edit.setReadOnly(True)
        self.next_x_edit.setStyleSheet("background-color: #F2F7FF; color: #5E6F80; border: 1px dashed #AFCBFF; border-radius: 4px; padding: 4px 6px;")
        next_coord_and_button_layout.addWidget(self.next_x_edit)
        
        next_coord_and_button_layout.addWidget(QLabel("Y:"))
        self.next_y_edit = QLineEdit()
        self.next_y_edit.setText(str(current_next_y))
        self.next_y_edit.setReadOnly(True)
        self.next_y_edit.setStyleSheet("background-color: #F2F7FF; color: #5E6F80; border: 1px dashed #AFCBFF; border-radius: 4px; padding: 4px 6px;")
        next_coord_and_button_layout.addWidget(self.next_y_edit)
        
        self.set_next_button = QPushButton("设置翻页按钮位置")
        self.set_next_button.clicked.connect(lambda: self.set_position('next_x_edit', 'next_y_edit', "翻页按钮"))
        next_coord_and_button_layout.addWidget(self.set_next_button)
        
        next_group_main_layout.addLayout(next_coord_and_button_layout)
        next_group.setLayout(next_group_main_layout)
        main_layout.addWidget(next_group)
        
        self.toggle_next_button_fields(self.enable_next_check.isChecked()) 
        
        # --- 5.5. 异常卷按钮配置 ---
        anomaly_group = QWidget()
        anomaly_group_main_layout = QVBoxLayout()
        self._compact_layout(anomaly_group_main_layout)
        
        # 启用异常卷按钮复选框
        self.enable_anomaly_check = QCheckBox("启用异常卷按钮（否则阅卷停止 人工介入）")
        enable_anomaly_for_current_q = self.question_config.get('enable_anomaly_button', False) if self.question_config else False
        self.enable_anomaly_check.setChecked(enable_anomaly_for_current_q)
        self.enable_anomaly_check.stateChanged.connect(self.toggle_anomaly_fields)
        anomaly_group_main_layout.addWidget(self.enable_anomaly_check)
        
        # 异常卷按钮坐标设置
        anomaly_coord_layout = QHBoxLayout()
        self._compact_layout(anomaly_coord_layout)
        anomaly_coord_layout.addWidget(QLabel("坐标:"))
        anomaly_coord_layout.addWidget(QLabel("X:"))
        self.anomaly_x_edit = QLineEdit()
        anomaly_pos_val = self.question_config.get('anomaly_button_pos') if self.question_config else None
        anomaly_x, anomaly_y = anomaly_pos_val if anomaly_pos_val is not None else (0, 0)
        self.anomaly_x_edit.setText(str(anomaly_x))
        self.anomaly_x_edit.setReadOnly(True)
        self.anomaly_x_edit.setStyleSheet("background-color: #F2F7FF; color: #5E6F80; border: 1px dashed #AFCBFF; border-radius: 4px; padding: 4px 6px;")
        anomaly_coord_layout.addWidget(self.anomaly_x_edit)
        
        anomaly_coord_layout.addWidget(QLabel("Y:"))
        self.anomaly_y_edit = QLineEdit()
        self.anomaly_y_edit.setText(str(anomaly_y))
        self.anomaly_y_edit.setReadOnly(True)
        self.anomaly_y_edit.setStyleSheet("background-color: #F2F7FF; color: #5E6F80; border: 1px dashed #AFCBFF; border-radius: 4px; padding: 4px 6px;")
        anomaly_coord_layout.addWidget(self.anomaly_y_edit)
        
        self.set_anomaly_button = QPushButton("设置异常卷按钮位置")
        self.set_anomaly_button.clicked.connect(lambda: self.set_position('anomaly_x_edit', 'anomaly_y_edit', "异常卷按钮"))
        anomaly_coord_layout.addWidget(self.set_anomaly_button)
        
        anomaly_group_main_layout.addLayout(anomaly_coord_layout)
        anomaly_group.setLayout(anomaly_group_main_layout)
        main_layout.addWidget(anomaly_group)
        
        # 根据启用状态切换异常卷按钮相关字段的启用状态
        self.toggle_anomaly_fields(self.enable_anomaly_check.isChecked())
        
        # --- 6. 答案区域配置 (恢复到旧版对称布局) ---
        answer_group = QWidget() # 确保 answer_group 已定义
        answer_group_main_layout = QVBoxLayout()
        self._compact_layout(answer_group_main_layout)

        # 先获取 answer_area，如果为 None 则使用空字典作为后备
        answer_area_config = self.question_config.get('answer_area') if self.question_config else None
        if answer_area_config is None:
            answer_area_config = {} # 使用空字典，这样后续的 .get 调用不会出错

        # 坐标信息合并到一行 - 左上角和右下角
        answer_coord_layout = QHBoxLayout() 
        self._compact_layout(answer_coord_layout, spacing=2, margins=(2, 0, 2, 0))
        
        # 左上角坐标
        answer_coord_layout.addWidget(QLabel("左上角 X:"))
        self.answer_x1_edit = QLineEdit()
        answer_x1 = answer_area_config.get('x1', 0)
        self.answer_x1_edit.setText(str(answer_x1))
        self.answer_x1_edit.setReadOnly(True)
        self.answer_x1_edit.setFixedWidth(50)  # 设置固定宽度以节省空间
        self.answer_x1_edit.setStyleSheet("background-color: #F2F7FF; color: #5E6F80; border: 1px dashed #AFCBFF; border-radius: 4px; padding: 2px 4px;")
        answer_coord_layout.addWidget(self.answer_x1_edit)
        
        answer_coord_layout.addWidget(QLabel("Y:"))
        self.answer_y1_edit = QLineEdit()
        answer_y1 = answer_area_config.get('y1', 0)
        self.answer_y1_edit.setText(str(answer_y1))
        self.answer_y1_edit.setReadOnly(True)
        self.answer_y1_edit.setFixedWidth(50)
        self.answer_y1_edit.setStyleSheet("background-color: #F2F7FF; color: #5E6F80; border: 1px dashed #AFCBFF; border-radius: 4px; padding: 2px 4px;")
        answer_coord_layout.addWidget(self.answer_y1_edit)
        
        # 添加分隔符
        answer_coord_layout.addWidget(QLabel(" | "))
        
        # 右下角坐标
        answer_coord_layout.addWidget(QLabel("右下角 X:"))
        self.answer_x2_edit = QLineEdit()
        answer_x2 = answer_area_config.get('x2', 0)
        self.answer_x2_edit.setText(str(answer_x2))
        self.answer_x2_edit.setReadOnly(True)
        self.answer_x2_edit.setFixedWidth(50)
        self.answer_x2_edit.setStyleSheet("background-color: #F2F7FF; color: #5E6F80; border: 1px dashed #AFCBFF; border-radius: 4px; padding: 2px 4px;")
        answer_coord_layout.addWidget(self.answer_x2_edit)
        
        answer_coord_layout.addWidget(QLabel("Y:"))
        self.answer_y2_edit = QLineEdit()
        answer_y2 = answer_area_config.get('y2', 0)
        self.answer_y2_edit.setText(str(answer_y2))
        self.answer_y2_edit.setReadOnly(True)
        self.answer_y2_edit.setFixedWidth(50)
        self.answer_y2_edit.setStyleSheet("background-color: #F2F7FF; color: #5E6F80; border: 1px dashed #AFCBFF; border-radius: 4px; padding: 2px 4px;")
        answer_coord_layout.addWidget(self.answer_y2_edit)
        
        answer_coord_layout.addStretch()  # 添加弹性空间
        answer_group_main_layout.addLayout(answer_coord_layout)

        # 按钮占据可用水平空间 (这部分不变)
        answer_button_layout = QHBoxLayout()
        self._compact_layout(answer_button_layout)
        self.set_answer_button = QPushButton("框定答案区域")
        self.set_answer_button.clicked.connect(self.start_answer_area_selection)
        answer_button_layout.addWidget(self.set_answer_button)
        answer_group_main_layout.addLayout(answer_button_layout)
        
        answer_group.setLayout(answer_group_main_layout)
        main_layout.addWidget(answer_group) # 确保 answer_group 被添加到主布局
        
        # --- 7. 按钮区域 (Save, Cancel) ---
        button_layout = QHBoxLayout()
        self._compact_layout(button_layout)
        save_button = QPushButton("保存")
        save_button.clicked.connect(self.save_config)
        cancel_button = QPushButton("取消")
        cancel_button.clicked.connect(self.reject)
        
        button_layout.addWidget(save_button)
        button_layout.addWidget(cancel_button)
        main_layout.addLayout(button_layout)
        
        self.setLayout(main_layout)

        # 由于设置了固定的 resize，adjustSize() 和 setMinimumWidth() 通常不再需要，
        # 或者其效果会被 resize 覆盖。
        # self.adjustSize() 
        # current_width = self.width()
        # self.setMinimumWidth(max(600, current_width))

    def update_question_type_style(self):
        """根据当前选择更新题型下拉框的边框颜色与样式。

        - 当选择为占位项（索引0）时，显示红色边框以提示用户必须选择。
        - 否则显示橙色/默认边框。
        """
        try:
            # 始终使用橙色高对比边框以保持视觉一致性
            border_color = '#FF8C00'
            style = f"min-height:28px; padding:4px; border:2px solid {border_color}; border-radius:4px;"
            # Guard against optional member being None to satisfy static type checkers
            if self.question_type_combo is not None:
                self.question_type_combo.setStyleSheet(style)
        except Exception:
            pass

    def _create_position_input_group(self, title, x_edit_attr, y_edit_attr, button_attr, pos_name):
        """辅助函数：创建一个用于设置坐标的位置输入UI组。"""
        # 创建UI组件
        group_box = QWidget()  # 组的标题由外部QLabel提供，保持UI一致性
        layout = QHBoxLayout()
        # 更紧凑的布局间距与外边距
        try:
            self._compact_layout(layout)
        except Exception:
            pass
        
        x_edit = QLineEdit()
        y_edit = QLineEdit()
        set_button = QPushButton(title)

        # 将创建的控件设置为类的成员变量，以便后续访问
        setattr(self, x_edit_attr, x_edit)
        setattr(self, y_edit_attr, y_edit)
        setattr(self, button_attr, set_button)
        
        # 设置坐标输入框为只读，并移除加粗蓝边
        x_edit.setReadOnly(True)
        x_edit.setStyleSheet("background-color: #F2F7FF; color: #5E6F80; border: 1px dashed #AFCBFF; border-radius: 4px; padding: 4px 6px;")
        y_edit.setReadOnly(True)
        y_edit.setStyleSheet("background-color: #F2F7FF; color: #5E6F80; border: 1px dashed #AFCBFF; border-radius: 4px; padding: 4px 6px;")

        # 布局
        layout.addWidget(QLabel("X:"))
        layout.addWidget(x_edit)
        layout.addWidget(QLabel("Y:"))
        layout.addWidget(y_edit)
        layout.addWidget(set_button)
        group_box.setLayout(layout)

        # 从配置加载初始值
        # 处理不同的位置输入组
        if pos_name == "分数输入":
            config_key = "score_input_pos"
        elif "位置" in pos_name:  # 例如 "位置 1"、"位置 2"、"位置 3"
            step_num = pos_name.split()[-1]  # 提取"1"、"2"或"3"
            config_key = f"score_input_pos_step{step_num}"
        else:
            config_key = f"score_input_pos_{pos_name}"
            
        pos_val = self.question_config.get(config_key) if self.question_config else None
        pos_x, pos_y = pos_val if pos_val is not None else (0, 0)
        x_edit.setText(str(pos_x))
        y_edit.setText(str(pos_y))

        # 连接信号
        # 使用 setattr 来动态地引用成员变量名
        set_button.clicked.connect(lambda: self.set_position(x_edit_attr, y_edit_attr, f"{pos_name}"))

        return group_box

    def toggle_next_button_fields(self, checked):
        """切换翻页按钮字段的启用状态"""
        self.next_x_edit.setEnabled(checked)
        self.next_y_edit.setEnabled(checked)
        self.set_next_button.setEnabled(checked)

    def toggle_anomaly_fields(self, checked):
        """切换异常卷按钮字段的启用状态"""
        self.anomaly_x_edit.setEnabled(checked)
        self.anomaly_y_edit.setEnabled(checked)
        self.set_anomaly_button.setEnabled(checked)

    def start_answer_area_selection(self):
        """开始框定答案区域"""
        try:
            if not self.parent_window:
                raise Exception("无法获取父窗口对象")
            
            # 获取当前题目的专用答案框窗口
            answer_window = self.parent_window.get_or_create_answer_window(self.question_index)  # type: ignore

            # 确保答案框是配置框的子窗口，并设置适当的窗口关系
            if answer_window.parent() != self:
                # 如果不是子窗口，设置为子窗口
                answer_window.setParent(self, Qt.Window)  # type: ignore
                # 重新设置窗口标志，确保没有置顶冲突
                answer_window.setWindowFlags(Qt.FramelessWindowHint | Qt.Tool)  # type: ignore

            # 显示窗口并确保其可见性
            answer_window.show()
            answer_window.set_edit_mode()

            # 使用QTimer延迟执行，确保窗口完全显示后再提升
            QTimer.singleShot(100, lambda: self._ensure_answer_window_visible(answer_window))

            self._log_message(f"请调整第{self.question_index}题的答案框位置和大小")

            # 更新按钮状态
            try:
                self.set_answer_button.clicked.disconnect()
            except TypeError:
                # 可能尚未连接任何槽，忽略
                pass
            self.set_answer_button.clicked.connect(
                lambda: self.confirm_answer_area_selection(answer_window)
            )
            self.set_answer_button.setText("确认框定")
            self.set_answer_button.setStyleSheet("background-color: #FF5555; color: white; font-weight: bold;")
        except Exception as e:
            self._log_message(f"框定第{self.question_index}题答案区域出错: {str(e)}", is_error=True)

    def _ensure_answer_window_visible(self, answer_window):
        """确保答案框窗口可见并获得焦点"""
        try:
            if answer_window and answer_window.isVisible():
                # 提升窗口到前面
                answer_window.raise_()
                answer_window.activateWindow()
                # 设置焦点到答案框
                answer_window.setFocus()
                print("答案框窗口已提升到前面并获得焦点")
        except Exception as e:
            print(f"提升答案框窗口时出错: {str(e)}")

    def confirm_answer_area_selection(self, answer_window):
        """确认框定答案区域"""
        try:
            screen_pos = answer_window.geometry().topLeft()
            x1 = int(screen_pos.x())
            y1 = int(screen_pos.y())
            x2 = answer_window.geometry().width() + x1
            y2 = answer_window.geometry().height() + y1

            # 更新坐标显示
            self.answer_x1_edit.setText(str(x1))
            self.answer_y1_edit.setText(str(y1))
            self.answer_x2_edit.setText(str(x2))
            self.answer_y2_edit.setText(str(y2))

            self.capture_answer_area(x1, y1, x2, y2)

            self._log_message(f"第{self.question_index}题答案区域已框定: ({x1}, {y1}) - ({x2}, {y2})")

            # 恢复按钮状态
            try:
                self.set_answer_button.clicked.disconnect()
            except TypeError:
                pass
            self.set_answer_button.clicked.connect(self.start_answer_area_selection)
            self.set_answer_button.setText("重新框定答案区域")
            self.set_answer_button.setStyleSheet("")

            # === 立即隐藏答案框窗口（在set_confirmed_mode之前） ===
            # 这样可以确保用户点击确认后，答题框立即消失
            answer_window.hide()
            self._log_message(f"第{self.question_index}题答题框窗口已立即隐藏")
            
            # 设置为已确认模式
            answer_window.set_confirmed_mode()

            # 关闭答案框窗口（cleanup）
            answer_window.close()
        except Exception as e:
            self._log_message(f"确认第{self.question_index}题答案区域出错: {str(e)}", is_error=True)

    def capture_answer_area(self, x1, y1, x2, y2):
        """捕获答案区域的屏幕截图或处理坐标"""
        try:
            # 这里可以添加逻辑来捕获屏幕截图或其他处理
            # 目前作为占位符，确保方法存在
            self._log_message(f"答案区域坐标已处理: ({x1}, {y1}) - ({x2}, {y2})")
        except Exception as e:
            self._log_message(f"处理答案区域坐标出错: {str(e)}", is_error=True)

    def set_position(self, x_edit_name, y_edit_name, position_name):
        """设置位置坐标"""
        try:
            # 创建一个提示标签，显示在对话框顶部
            if not hasattr(self, 'instruction_label'):
                self.instruction_label = QLabel("")
                # 将标签插入到布局的顶部（索引为0，确保在最顶端）
                self.layout().insertWidget(0, self.instruction_label)  # type: ignore

            self.instruction_label.setText(f"请将鼠标移动到{position_name}位置，5秒后将自动捕获位置...")
            self.instruction_label.setStyleSheet("color: red; font-weight: bold;")
            
            self._log_message(f"请将鼠标移动到{position_name}位置，5秒后将自动捕获位置...")

            # 使用定时器模拟等待5秒
            # 创建一个全新的定时器实例以确保信号连接正确
            self.position_capture_timer = QTimer(self)
            self.position_capture_timer.setSingleShot(True)
            self.position_capture_timer.timeout.connect(lambda: self.capture_position(x_edit_name, y_edit_name, position_name))
            self.position_capture_timer.start(5000)

        except Exception as e:
            self._log_message(f"设置{position_name}位置出错: {str(e)}", is_error=True)
            if hasattr(self, 'instruction_label'):
                self.instruction_label.setText("")

    def capture_position(self, x_edit_name, y_edit_name, position_name):
        """捕获鼠标位置并更新文本框"""
        try:
            # 获取当前鼠标位置
            x, y = pyautogui.position()

            # 设置坐标到文本框
            x_edit = getattr(self, x_edit_name)
            y_edit = getattr(self, y_edit_name)
            x_edit.setText(str(x))
            y_edit.setText(str(y))

            # 更新提示信息
            if hasattr(self, 'instruction_label'):
                self.instruction_label.setText(f"{position_name}位置已捕获: ({x}, {y})")
            
            self._log_message(f"{position_name}位置已设置为: ({x}, {y})")
            
            self.position_capture_timer = None  # 任务完成，清理引用
        except Exception as e:
            self._log_message(f"捕获{position_name}位置出错: {str(e)}", is_error=True)
            if hasattr(self, 'instruction_label'):
                self.instruction_label.setText("")

    def save_config(self):
        """保存配置"""
        try:
            # 获取分数设置
            max_score = self.max_score_edit.value()
            min_score = self.min_score_edit.value()

            # 获取分数输入位置
            score_x = int(self.score_x_edit.text())
            score_y = int(self.score_y_edit.text())

            # 获取提交按钮位置
            submit_x = int(self.submit_x_edit.text())
            submit_y = int(self.submit_y_edit.text())

            # 获取答案区域位置
            answer_x1 = int(self.answer_x1_edit.text())
            answer_y1 = int(self.answer_y1_edit.text())
            answer_x2 = int(self.answer_x2_edit.text())
            answer_y2 = int(self.answer_y2_edit.text())

            # --- 新增：获取并保存题目类型 ---
            selected_display_text = self.question_type_combo.currentText() if self.question_type_combo else 'Subjective_PointBased_QA'
            selected_type_identifier = 'Subjective_PointBased_QA'  # 默认值
            if self.question_type_combo:
                for identifier, display in self.question_types_map.items():  # self.question_types_map 在 init_ui 中定义
                    if display == selected_display_text:
                        selected_type_identifier = identifier
                        break
            # --- 结束新增 ---

            # 更新题目配置
            if self.config_manager:
                self.config_manager.update_question_config(str(self.question_index), 'max_score', max_score)
                self.config_manager.update_question_config(str(self.question_index), 'min_score', min_score)
                self.config_manager.update_question_config(str(self.question_index), 'score_input_pos', (score_x, score_y))
                self.config_manager.update_question_config(str(self.question_index), 'confirm_button_pos', (submit_x, submit_y))
                self.config_manager.update_question_config(str(self.question_index), 'question_type', selected_type_identifier)  # 保存题目类型
                self.config_manager.update_question_config(str(self.question_index), 'answer_area', {
                    'x1': answer_x1,
                    'y1': answer_y1,
                    'x2': answer_x2,
                    'y2': answer_y2
                })

            # 更新当前小题的翻页按钮配置
            enable_next_for_current_q = self.enable_next_check.isChecked()
            if self.config_manager:
                self.config_manager.update_question_config(str(self.question_index), 'enable_next_button', enable_next_for_current_q)

            if enable_next_for_current_q and self.config_manager:
                try:
                    current_q_next_x = int(self.next_x_edit.text())
                    current_q_next_y = int(self.next_y_edit.text())
                    self.config_manager.update_question_config(str(self.question_index), 'next_button_pos', (current_q_next_x, current_q_next_y))
                except ValueError:
                    # 如果坐标无效，则保存为 None，并记录日志或提示用户
                    self.config_manager.update_question_config(str(self.question_index), 'next_button_pos', None)
                    self._log_message(f"警告: 第{self.question_index}题的翻页按钮坐标无效，未保存。", is_error=True)
            elif self.config_manager:
                self.config_manager.update_question_config(str(self.question_index), 'next_button_pos', None)

            # --- 更新当前小题的异常卷按钮配置 ---
            enable_anomaly_for_current_q = self.enable_anomaly_check.isChecked()
            if self.config_manager:
                self.config_manager.update_question_config(str(self.question_index), 'enable_anomaly_button', enable_anomaly_for_current_q)

            if enable_anomaly_for_current_q and self.config_manager:
                try:
                    anomaly_x = int(self.anomaly_x_edit.text())
                    anomaly_y = int(self.anomaly_y_edit.text())
                    self.config_manager.update_question_config(str(self.question_index), 'anomaly_button_pos', (anomaly_x, anomaly_y))
                except ValueError:
                    self.config_manager.update_question_config(str(self.question_index), 'anomaly_button_pos', None)
                    self._log_message(f"警告: 第{self.question_index}题的异常卷按钮坐标无效，未保存。", is_error=True)
            elif self.config_manager:
                self.config_manager.update_question_config(str(self.question_index), 'anomaly_button_pos', None)

            # --- 开始保存三步打分相关配置 (仅第一题) ---
            if self.question_index == 1 and self.three_step_scoring_checkbox and self.config_manager:  # 确保复选框和config_manager已创建
                is_three_step_enabled = self.three_step_scoring_checkbox.isChecked()
                self.config_manager.update_question_config(str(self.question_index), 'enable_three_step_scoring', is_three_step_enabled)  # 使用 str(self.question_index)

                if is_three_step_enabled and self.score_x_edit_step1 and self.score_y_edit_step1:
                    try:
                        pos_s1_x = int(self.score_x_edit_step1.text())
                        pos_s1_y = int(self.score_y_edit_step1.text())
                        self.config_manager.update_question_config(str(self.question_index), 'score_input_pos_step1', (pos_s1_x, pos_s1_y))
                    except ValueError:
                        self.config_manager.update_question_config(str(self.question_index), 'score_input_pos_step1', None)
                        self._log_message(f"警告: 第{self.question_index}题三步打分位置1坐标无效", is_error=True)

                    if self.score_x_edit_step2 and self.score_y_edit_step2:
                        try:
                            pos_s2_x = int(self.score_x_edit_step2.text())
                            pos_s2_y = int(self.score_y_edit_step2.text())
                            self.config_manager.update_question_config(str(self.question_index), 'score_input_pos_step2', (pos_s2_x, pos_s2_y))
                        except ValueError:
                            self.config_manager.update_question_config(str(self.question_index), 'score_input_pos_step2', None)
                            self._log_message(f"警告: 第{self.question_index}题三步打分位置2坐标无效", is_error=True)

                    if self.score_x_edit_step3 and self.score_y_edit_step3:
                        try:
                            pos_s3_x = int(self.score_x_edit_step3.text())
                            pos_s3_y = int(self.score_y_edit_step3.text())
                            self.config_manager.update_question_config(str(self.question_index), 'score_input_pos_step3', (pos_s3_x, pos_s3_y))
                        except ValueError:
                            self.config_manager.update_question_config(str(self.question_index), 'score_input_pos_step3', None)
                            self._log_message(f"警告: 第{self.question_index}题三步打分位置3坐标无效", is_error=True)
                else:
                    # 如果未启用，可以选择清除这些位置或保持原样
                    # 为了简单，可以不清除，ConfigManager在加载时会处理None
                    # 或者明确设为None:
                    self.config_manager.update_question_config(str(self.question_index), 'score_input_pos_step1', None)
                    self.config_manager.update_question_config(str(self.question_index), 'score_input_pos_step2', None)
                    self.config_manager.update_question_config(str(self.question_index), 'score_input_pos_step3', None)
            # --- 结束保存三步打分相关配置 ---

            if self.parent_window and hasattr(self.parent_window, 'log_message') and self.config_manager:
                self._log_message(f"第{self.question_index}题配置已保存")
                log_enable_next_status = self.config_manager.get_question_config(self.question_index).get('enable_next_button', False)
                self._log_message(f"第{self.question_index}题翻页按钮状态: {'启用' if log_enable_next_status else '禁用'}")
                self._log_message(f"答案区域已配置: ({answer_x1}, {answer_y1}) - ({answer_x2}, {answer_y2})")

            # 发射信号通知配置已更新，由MainWindow负责保存到文件
            self.config_updated.emit()

            self.accept()
        except Exception as e:
            self._log_message(f"保存配置出错: {str(e)}", is_error=True)

    def toggle_three_step_mode_ui(self, checked):
        """根据三步打分模式复选框状态，切换UI元素的启用/禁用和显示/隐藏"""
        if self.question_index != 1:  # 确保只对第一题操作
            return
        # 当启用三步打分时，确保给分上限不超过60（但不盲目覆盖用户原有较小的设置）
        if checked:
            # 限制用户可设置的最大值为60
            try:
                self.max_score_edit.setMaximum(60)
            except Exception:
                pass
            # 如果当前用户设置的上限超过60，则将其降到60并记录日志
            if self.max_score_edit.value() > 60:
                self.max_score_edit.setValue(60)
                self._log_message("启用三步打分：给分上限超过60，已自动限制为60。")
        else:
            # 取消三步打分时，恢复上限为默认的较大值（例如150）以便用户调整
            try:
                self.max_score_edit.setMaximum(150)
            except Exception:
                pass

        # 切换原有单点分数输入组的启用/显示状态
        if self.original_score_input_group:
            self.original_score_input_group.setEnabled(not checked)
            self.original_score_input_group.setVisible(not checked)  # 未启用时隐藏

        # 切换三个新分数输入位置组的启用/显示状态
        # 启用时显示并启用，未启用时隐藏并禁用
        if self.score_input_group_step1:
            self.score_input_group_step1.setEnabled(checked)
            self.score_input_group_step1.setVisible(checked)  # 启用时显示
        if self.score_input_group_step2:
            self.score_input_group_step2.setEnabled(checked)
            self.score_input_group_step2.setVisible(checked)  # 启用时显示
        if self.score_input_group_step3:
            self.score_input_group_step3.setEnabled(checked)
            self.score_input_group_step3.setVisible(checked)  # 启用时显示

    def closeEvent(self, a0: Optional[QCloseEvent]) -> None:
        """窗口关闭事件，确保停止任何活动的定时器"""
        if self.position_capture_timer and self.position_capture_timer.isActive():
            self.position_capture_timer.stop()
        if a0 is not None:
            super().closeEvent(a0)

# --- END OF FILE question_config_dialog.py (Corrected) ---
