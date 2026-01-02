import sys
import os
import datetime
import pathlib
import warnings

# 设置 Qt 插件路径，避免平台插件加载失败
def _set_qt_platform_plugin_path() -> None:
    """Set QT_QPA_PLATFORM_PLUGIN_PATH for both dev and frozen runs.

    This reduces "could not load the Qt platform plugin" startup failures.
    """
    candidate_paths = []

    if getattr(sys, 'frozen', False):
        exe_dir = os.path.dirname(sys.executable)
        candidate_paths.append(os.path.join(exe_dir, 'PyQt5', 'Qt5', 'plugins', 'platforms'))
        # Common PyInstaller layouts
        meipass = getattr(sys, '_MEIPASS', None)
        if meipass:
            candidate_paths.append(os.path.join(meipass, 'PyQt5', 'Qt5', 'plugins', 'platforms'))
            candidate_paths.append(os.path.join(meipass, 'Qt5', 'plugins', 'platforms'))
            candidate_paths.append(os.path.join(meipass, 'Qt', 'plugins', 'platforms'))
    else:
        try:
            import PyQt5
            candidate_paths.append(os.path.join(os.path.dirname(PyQt5.__file__), 'Qt5', 'plugins', 'platforms'))
        except Exception:
            return

    for path in candidate_paths:
        if path and os.path.isdir(path):
            os.environ['QT_QPA_PLATFORM_PLUGIN_PATH'] = path
            return


try:
    _set_qt_platform_plugin_path()
except Exception:
    pass

# 过滤PyQt5的弃用警告
warnings.filterwarnings('ignore', category=DeprecationWarning, module='PyQt5')

# 设置Windows控制台输出编码为UTF-8，解决中文乱码问题
if sys.platform == 'win32':
    try:
        import io
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
    except Exception:
        pass  # 如果设置失败，继续使用默认编码

from PyQt5.QtWidgets import QApplication, QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QFont
from ui_components.main_window import MainWindow
from api_service import ApiService
from config_manager import ConfigManager
from auto_thread import GradingThread
import winsound
import csv
import traceback
import pandas as pd

class SimpleNotificationDialog(QDialog):
    def __init__(self, title, message, sound_type='info', parent=None):
        super().__init__(parent)
        self.sound_type = sound_type
        self.setup_ui(title, message)
        self.setup_sound_timer()

    def setup_ui(self, title, message):
        self.setWindowTitle(title)
        self.setMinimumSize(300, 100)
        self.setMaximumSize(600, 400)
        # Set the WindowStaysOnTopHint flag when available (guarded to satisfy static type checkers)
        try:
            flags = self.windowFlags()
            ws = getattr(Qt, 'WindowStaysOnTopHint', None)
            if ws is not None:
                flags |= ws
            self.setWindowFlags(flags)
        except Exception:
            # Fallback: silently ignore if window flags API is not available
            pass

        layout = QVBoxLayout()

        # 消息标签
        msg_label = QLabel(message)
        msg_label.setWordWrap(True)
        msg_label.setStyleSheet("padding: 20px;")
        layout.addWidget(msg_label)

        # 确定按钮
        button_layout = QHBoxLayout()
        close_btn = QPushButton("确定")
        close_btn.clicked.connect(self.accept)
        close_btn.setDefault(True)  # 支持回车键确认
        button_layout.addStretch()
        button_layout.addWidget(close_btn)
        button_layout.addStretch()

        layout.addLayout(button_layout)
        self.setLayout(layout)

    def setup_sound_timer(self):
        # 立即播放一次
        self.play_system_sound()

        # 设置30秒重复定时器，确保用户不会错过提醒
        self.sound_timer = QTimer()
        self.sound_timer.timeout.connect(self.play_system_sound)
        self.sound_timer.start(30000)  # 30秒重复一次

    def play_system_sound(self):
        """播放系统默认提示音，错误情况使用更清晰的警告音"""
        try:
            if self.sound_type == 'error':
                # 错误声音：连续两次beep以吸引用户注意
                winsound.Beep(1000, 300)  # 较高音调，300ms
                winsound.Beep(1000, 300)  # 重复，增强存在感
            else:
                # 信息声音：单次beep
                winsound.Beep(800, 200)
        except Exception:
            # 如果系统声音不可用，回退到系统消息提示音
            try:
                winsound.MessageBeep(-1)
            except Exception:
                pass  # 完全静默失败

    def closeEvent(self, a0):
        """窗口关闭时停止定时器"""
        if hasattr(self, 'sound_timer'):
            self.sound_timer.stop()
        super().closeEvent(a0) 

    def accept(self):
        """点击确定时停止定时器"""
        if hasattr(self, 'sound_timer'):
            self.sound_timer.stop()
        super().accept()


class ManualInterventionDialog(QDialog):
    """专用于人工介入提示的模态对话框，带重复提示音和明确的继续/停止按钮"""
    def __init__(self, title, message, raw_feedback=None, sound_type='error', parent=None):
        super().__init__(parent)
        self.sound_type = sound_type
        self.raw_feedback = raw_feedback or ''
        self.setup_ui(title, message)
        self.setup_sound_timer()

    def setup_ui(self, title, message):
        self.setWindowTitle(title)
        self.setMinimumSize(420, 220)
        self.setMaximumSize(900, 600)
        try:
            flags = self.windowFlags()
            ws = getattr(Qt, 'WindowStaysOnTopHint', None)
            if ws is not None:
                flags |= ws
            self.setWindowFlags(flags)
        except Exception:
            pass

        layout = QVBoxLayout()

        # 主消息
        msg_label = QLabel(message)
        msg_label.setWordWrap(True)
        msg_label.setStyleSheet("padding: 12px;")
        layout.addWidget(msg_label)

        # 原始反馈（可折叠/展示）
        fb_label = QLabel("检测到的AI原始反馈（仅供参考）：\n" + (self.raw_feedback[:1000] or "(无)") )
        fb_label.setWordWrap(True)
        fb_label.setStyleSheet("padding: 6px; color: #333333; background: #f7f7f7; border-radius:4px;")
        layout.addWidget(fb_label)

        # 按钮区域
        button_layout = QHBoxLayout()
        continue_btn = QPushButton("我已人工处理，继续")
        stop_btn = QPushButton("暂停并关闭")
        continue_btn.clicked.connect(self.accept)
        stop_btn.clicked.connect(self.reject)
        button_layout.addStretch()
        button_layout.addWidget(continue_btn)
        button_layout.addWidget(stop_btn)
        button_layout.addStretch()

        layout.addLayout(button_layout)
        self.setLayout(layout)

    def setup_sound_timer(self):
        # 立即播放并每30秒重复一次，确保用户能及时注意到需要人工介入
        self.play_system_sound()
        self.sound_timer = QTimer()
        self.sound_timer.timeout.connect(self.play_system_sound)
        self.sound_timer.start(30000)  # 30秒重复一次

    def play_system_sound(self):
        """播放需要人工介入的警告音"""
        try:
            # 使用三次连续beep制造更清晰的警告效果
            for _ in range(3):
                winsound.Beep(1000, 250)  # 较高音调，每次250ms
        except Exception:
            try:
                winsound.MessageBeep(-1)  # 回退到系统错误音
            except Exception:
                pass

    def accept(self):
        if hasattr(self, 'sound_timer'):
            self.sound_timer.stop()
        super().accept()

    def reject(self):
        if hasattr(self, 'sound_timer'):
            self.sound_timer.stop()
        super().reject()


class SignalConnectionManager:
    def __init__(self):
        self.connections = []

    def connect(self, signal, slot, connection_type=None):
        """安全地连接信号，避免重复"""
        # 检查是否已经存在相同的连接，避免重复添加
        connection_key = (id(signal), id(slot))
        if connection_key in [(id(s), id(sl)) for s, sl in self.connections]:
            return  # 已存在，不重复连接
        
        # 先尝试断开可能存在的连接
        try:
            signal.disconnect(slot)
        except (TypeError, RuntimeError):
            pass

        # 建立新连接
        try:
            signal.connect(slot)
            self.connections.append((signal, slot))
        except Exception as e:
            print(f"[警告] 信号连接失败: {e}")

    def disconnect_all(self):
        """断开所有管理的连接"""
        disconnected = 0
        failed = 0
        
        for signal, slot in self.connections:
            try:
                signal.disconnect(slot)
                disconnected += 1
            except (TypeError, RuntimeError):
                failed += 1
        
        self.connections.clear()
        
        if failed > 0:
            print(f"[信号管理] 成功断开 {disconnected} 个连接，{failed} 个连接断开失败（可能已断开）")

class Application:
    def __init__(self):
        self.app = QApplication(sys.argv)
        # 先加载配置管理器，以便应用字体可由配置控制
        self.config_manager: ConfigManager = ConfigManager()
        # 固定主界面字号为 11（不提供用户自行调整字号的入口）
        try:
            self.app.setFont(QFont("微软雅黑", 11))
        except Exception:
            pass
        self.api_service = ApiService(self.config_manager)
        self.worker = GradingThread(self.api_service, self.config_manager)
        self.main_window = MainWindow(self.config_manager, self.api_service, self.worker)
        self.signal_manager = SignalConnectionManager()



        self._setup_application()

    def _setup_global_exception_hook(self):
        """设置全局异常钩子"""
        def handle_exception(exc_type, exc_value, exc_traceback):
            if issubclass(exc_type, KeyboardInterrupt):
                sys.__excepthook__(exc_type, exc_value, exc_traceback)
                return

            error_msg = "".join(traceback.format_exception(exc_type, exc_value, exc_traceback))
            
            # 尝试记录到UI
            if hasattr(self, 'main_window') and hasattr(self.main_window, 'log_message'):
                self.main_window.log_message(f"全局异常捕获:\n{error_msg}", is_error=True)
            
            # 尝试记录到文件
            try:
                # 确定日志目录的绝对路径
                if getattr(sys, 'frozen', False):
                    # 打包后，相对于exe文件
                    base_dir = pathlib.Path(sys.executable).parent
                else:
                    # 开发时，相对于main.py
                    base_dir = pathlib.Path(__file__).parent

                log_dir = base_dir / "logs"
                log_dir.mkdir(exist_ok=True)
                current_time = datetime.datetime.now()
                formatted_time = current_time.strftime('%H点%M分%S秒')
                log_file = log_dir / f"global_error_{current_time.strftime('%Y%m%d')}_{formatted_time}.log"
                with open(log_file, 'w', encoding='utf-8') as f:
                    f.write(error_msg)
            except Exception as e:
                print(f"写入全局异常日志失败: {e}")

            # 显示一个简单的错误对话框
            try:
                dialog = SimpleNotificationDialog(
                    title="严重错误",
                    message=f"发生了一个意外的严重错误，应用程序可能需要关闭。\n\n错误: {exc_value}",
                    sound_type='error'
                )
                dialog.exec_()
            except Exception:
                # 如果对话框创建失败，至少打印错误信息
                print(f"严重错误: {exc_value}")

        sys.excepthook = handle_exception

    def _setup_application(self):
        """初始化应用程序设置"""
        try:
            self._setup_global_exception_hook()
            self.connect_worker_signals()
            self.load_config()
            self._create_record_directory()
        except Exception as e:
            print(f"应用程序初始化失败: {str(e)}")
            sys.exit(1)

    def _create_record_directory(self):
        """创建记录目录"""
        try:
            if getattr(sys, 'frozen', False):
                # 如果是打包后的exe，使用exe所在的实际目录
                base_dir = pathlib.Path(sys.executable).parent
            else:
                # 否则，使用当前文件所在的目录
                base_dir = pathlib.Path(__file__).parent
            record_dir = base_dir / "阅卷记录"
            record_dir.mkdir(exist_ok=True)
        except OSError as e:
            self.main_window.log_message(f"创建记录目录失败: {str(e)}", is_error=True)

    def connect_worker_signals(self):
        """连接工作线程信号"""
        try:
            self.signal_manager.disconnect_all() # 断开旧连接

            # self.signal_manager.connect(
            #     self.worker.update_signal,
            #     self.main_window.update_suggestion_text
            # )
            self.signal_manager.connect(
                self.worker.log_signal,
                self.main_window.log_message
            )
            # self.signal_manager.connect(
            #     self.worker.progress_signal,
            #     self.main_window.update_progress
            # )
            self.signal_manager.connect(
                self.worker.record_signal,
                self.save_grading_record
            )

            # 任务正常完成
            self.signal_manager.connect(
                self.worker.finished_signal,
                self.show_completion_notification # 这个方法内部会调用 main_window.on_worker_finished
            )

            # 任务因错误中断
            if hasattr(self.worker, 'error_signal'): # 确保 GradingThread 有 error_signal
                self.signal_manager.connect(
                    self.worker.error_signal,
                    self.show_error_notification # 这个方法内部需要调用 main_window.on_worker_error
                )

            # 双评分差超过阈值中断
            if hasattr(self.worker, 'threshold_exceeded_signal'):
                self.signal_manager.connect(
                    self.worker.threshold_exceeded_signal,
                    self.show_threshold_exceeded_notification # 这个方法内部需要调用 main_window.on_worker_error
                )

            # 人工介入信号：当AI明确请求人工复核时触发
            if hasattr(self.worker, 'manual_intervention_signal'):
                self.signal_manager.connect(
                    self.worker.manual_intervention_signal,
                    self.show_manual_intervention_notification
                )



        except Exception as e:
            # 避免在 main_window 可能还未完全初始化时调用其 log_message
            print(f"[CRITICAL_ERROR] 连接工作线程信号时出错: {str(e)}")
            if hasattr(self.main_window, 'log_message'):
                 self.main_window.log_message(f"连接工作线程信号时出错: {str(e)}", is_error=True)

    def show_completion_notification(self):
        """显示任务完成通知"""
        # 先调用原有的完成处理
        self.main_window.on_worker_finished()

        # 显示简洁的完成通知
        dialog = SimpleNotificationDialog(
            title="批次完成",
            message="✅ 本次自动阅卷已完成！\n\n请复查AI阅卷结果，人工审核0分、满分",
            sound_type='info',
            parent=self.main_window
        )
        dialog.exec_()
        
        # 对话框关闭后，确保主窗口恢复并显示在前台
        if self.main_window.isMinimized():
            self.main_window.showNormal()
        self.main_window.raise_()  # 将窗口提升到最前
        self.main_window.activateWindow()  # 激活窗口

    def show_error_notification(self, error_message):
        """显示错误通知并恢复主窗口状态"""
        if hasattr(self.main_window, 'on_worker_error'):
            self.main_window.on_worker_error(error_message)
        else:
            print(f"[ERROR] MainWindow missing on_worker_error. Error: {error_message}")
            # 基本的后备恢复
            if self.main_window.isMinimized(): self.main_window.showNormal(); self.main_window.activateWindow()
            if hasattr(self.main_window, 'update_ui_state'): self.main_window.update_ui_state(is_running=False)

        dialog = SimpleNotificationDialog(
            title="阅卷中断",
            message=f"⚠️ 自动阅卷因发生错误而停止！\n\n错误: {error_message}\n请检查界面下方日志。",
            sound_type='error',
            parent=self.main_window
        )
        dialog.exec_()
        
        # 对话框关闭后，确保主窗口恢复并显示在前台
        if self.main_window.isMinimized():
            self.main_window.showNormal()
        self.main_window.raise_()  # 将窗口提升到最前
        self.main_window.activateWindow()  # 激活窗口

    def show_threshold_exceeded_notification(self, reason):
        """显示双评分差超过阈值的通知并恢复主窗口状态"""
        specific_error_message = f"双评分差过大 ({reason})"
        if hasattr(self.main_window, 'on_worker_error'):
            self.main_window.on_worker_error(specific_error_message)
        else:
            print(f"[ERROR] MainWindow missing on_worker_error. Reason: {specific_error_message}")
            # 基本的后备恢复
            if self.main_window.isMinimized(): self.main_window.showNormal(); self.main_window.activateWindow()
            if hasattr(self.main_window, 'update_ui_state'): self.main_window.update_ui_state(is_running=False)

        dialog = SimpleNotificationDialog(
            title="双评分差过大",
            message=f"⚠️ {specific_error_message}，自动阅卷已中断！\n\n请检查日志并手动处理。",
            sound_type='error',
            parent=self.main_window
        )
        dialog.exec_()
        
        # 对话框关闭后，确保主窗口恢复并显示在前台
        if self.main_window.isMinimized():
            self.main_window.showNormal()
        self.main_window.raise_()  # 将窗口提升到最前
        self.main_window.activateWindow()  # 激活窗口

    def show_manual_intervention_notification(self, message, raw_feedback):
        """当工作线程请求人工介入时调用，展示更明显的模态对话框并播放提示音。"""
        specific_error_message = f"需要人工介入: {message}"
        if hasattr(self.main_window, 'on_worker_error'):
            self.main_window.on_worker_error(specific_error_message)
        else:
            if self.main_window.isMinimized(): self.main_window.showNormal(); self.main_window.activateWindow()
            if hasattr(self.main_window, 'update_ui_state'): self.main_window.update_ui_state(is_running=False)

        # 显示模态对话框，提供“继续”或“暂停并关闭”选项
        dialog = ManualInterventionDialog(
            title="人工介入 - 需要人工核查",
            message=(f"检测到需要人工核查的情况：{message}\n\n请人工检查对应答题图片并处理。"),
            raw_feedback=raw_feedback,
            sound_type='error',
            parent=self.main_window
        )
        result = dialog.exec_()
        if result == QDialog.Accepted:
            # 用户选择继续：提醒用户手动重启自动阅卷（不自动重启）
            self.main_window.log_message("用户确认已人工处理，若要继续，请手动重新启动自动阅卷。")
        else:
            self.main_window.log_message("用户选择暂停并手动处理当前项。请在处理完成后手动重启阅卷。")
        
        # 对话框关闭后，确保主窗口恢复并显示在前台
        if self.main_window.isMinimized():
            self.main_window.showNormal()
        self.main_window.raise_()  # 将窗口提升到最前
        self.main_window.activateWindow()  # 激活窗口

    def load_config(self):
        """加载配置并设置到主窗口"""
        # 加载配置到内存
        self.config_manager.load_config()
        # 将配置加载到UI
        self.main_window.load_config_to_ui()

        # 更新API服务的配置
        self.api_service.update_config_from_manager()

        self.main_window.log_message("配置已成功加载并应用。")

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



    def _get_excel_filepath(self, record_data, worker=None):
        """获取Excel文件路径的辅助函数"""
        timestamp_str = record_data.get('timestamp', datetime.datetime.now().strftime('%Y年%m月%d日_%H点%M分%S秒'))

        # 处理日期字符串，支持中文格式
        if '_' in timestamp_str:
            date_str = timestamp_str.split('_')[0]
        else:
            # 如果没有下划线，使用当前时间
            now = datetime.datetime.now()
            date_str = now.strftime('%Y年%m月%d日')

        # 转换日期格式：从中文格式提取数字部分用于目录命名
        if '年' in date_str and '月' in date_str and '日' in date_str:
            # 中文格式：2025年09月20日 -> 20250920
            try:
                year = date_str.split('年')[0]
                month = date_str.split('年')[1].split('月')[0].zfill(2)
                day = date_str.split('月')[1].split('日')[0].zfill(2)
                numeric_date_str = f"{year}{month}{day}"
            except (IndexError, ValueError):
                # 如果解析失败，使用当前日期
                numeric_date_str = datetime.datetime.now().strftime('%Y%m%d')
        else:
            # 假设已经是数字格式或使用当前日期
            numeric_date_str = date_str if date_str.isdigit() and len(date_str) == 8 else datetime.datetime.now().strftime('%Y%m%d')

        if getattr(sys, 'frozen', False):
            base_dir = pathlib.Path(sys.executable).parent
        else:
            base_dir = pathlib.Path(__file__).parent

        record_dir = base_dir / "阅卷记录"
        record_dir.mkdir(exist_ok=True)

        date_dir = record_dir / date_str
        date_dir.mkdir(exist_ok=True)

        if worker:
            dual_evaluation = worker.parameters.get('dual_evaluation', False)
            question_configs = worker.parameters.get('question_configs', [])
            question_count = len(question_configs)
            full_score = question_configs[0].get('max_score', 100) if question_configs else 100
        else:
            dual_evaluation = record_data.get('is_dual_evaluation_run', False)
            question_count = record_data.get('total_questions_in_run', 1)
            full_score = 100  # 默认值

        if question_count == 0:
            question_count = 1

        evaluation_type = '双评' if dual_evaluation else '单评'

        if question_count == 1:
            excel_filename = f"此题最高{full_score}分_{evaluation_type}.xlsx"
        else:
            excel_filename = f"共阅{question_count}题_{evaluation_type}.xlsx"

        excel_filepath = date_dir / excel_filename

        return excel_filepath

    def _save_summary_record(self, record_data):
        """保存汇总记录到对应的Excel文件

        Args:
            record_data: 汇总记录数据
        """
        try:
            excel_filepath = self._get_excel_filepath(record_data, self.worker)
            excel_filename = excel_filepath.name

            # 从 record_data 构建汇总行
            status_map = {
                "completed": "正常完成",
                "error": "因错误中断",
                "threshold_exceeded": "因双评分差过大中断"
            }
            status_text = status_map.get(record_data.get('completion_status', 'unknown'), "未知状态")

            interrupt_reason = record_data.get('interrupt_reason')
            if interrupt_reason:
                status_text += f" ({interrupt_reason})"

            # 格式化汇总时间戳
            timestamp_raw = record_data.get('timestamp', '未提供_未提供')
            if '_' in timestamp_raw:
                time_part = timestamp_raw.split('_')[1]
                if len(time_part) == 6:
                    formatted_summary_time = f"{time_part[:2]}点{time_part[2:4]}分{time_part[4:6]}秒"
                else:
                    formatted_summary_time = time_part
            else:
                formatted_summary_time = timestamp_raw

            summary_data = [
                f"--- 批次阅卷汇总 ({formatted_summary_time}) ---",
                f"状态: {status_text}",
                f"计划/完成: {record_data.get('total_questions_attempted', '未提供')} / {record_data.get('questions_completed', '未提供')} 个",
                f"总用时: {record_data.get('total_elapsed_time_seconds', 0):.2f} 秒",
                f"模式: {'双评' if record_data.get('dual_evaluation_enabled') else '单评'}",
            ]

            if record_data.get('dual_evaluation_enabled'):
                summary_data.append(f"模型: {record_data.get('first_model_id', '未指定')} vs {record_data.get('second_model_id', '未指定')}")
            else:
                summary_data.append(f"模型: {record_data.get('first_model_id', '未指定')}")

            # 读取现有Excel文件或创建新的
            if excel_filepath.exists():
                try:
                    existing_df = pd.read_excel(excel_filepath, header=0)
                    # 检查是否是汇总记录格式（只有一列）
                    if len(existing_df.columns) == 1 and existing_df.columns[0] == "汇总信息":
                        # 如果是汇总格式，直接添加
                        summary_df = pd.DataFrame([summary_data], columns=["汇总信息"])
                        combined_df = pd.concat([existing_df, summary_df], ignore_index=True)
                    else:
                        # 如果是详细记录格式，添加到末尾
                        # 添加空白行
                        blank_rows = pd.DataFrame([[""] * len(existing_df.columns)] * 2)
                        # 创建汇总行，填充到与现有列数相同
                        summary_row = summary_data[:len(existing_df.columns)] if len(summary_data) >= len(existing_df.columns) else summary_data + [""] * (len(existing_df.columns) - len(summary_data))
                        summary_df = pd.DataFrame([summary_row], columns=existing_df.columns)
                        more_blank_rows = pd.DataFrame([[""] * len(existing_df.columns)] * 4)
                        combined_df = pd.concat([existing_df, blank_rows, summary_df, more_blank_rows], ignore_index=True)
                except Exception as e:
                    self.main_window.log_message(f"读取现有Excel文件失败: {str(e)}，将创建新汇总文件", True)
                    combined_df = pd.DataFrame([summary_data], columns=["汇总信息"])
            else:
                combined_df = pd.DataFrame([summary_data], columns=["汇总信息"])

            # 写入Excel文件
            with pd.ExcelWriter(excel_filepath, engine='openpyxl') as writer:
                combined_df.to_excel(writer, index=False, sheet_name='阅卷记录')

                # 获取工作簿和工作表
                workbook = writer.book
                worksheet = writer.sheets['阅卷记录']

                # 设置列宽
                column_widths = {
                    'A': 80,  # 汇总信息列
                }

                for col, width in column_widths.items():
                    if col in worksheet.column_dimensions:
                        worksheet.column_dimensions[col].width = width

                # 设置自动换行
                from openpyxl.styles import Alignment
                wrap_alignment = Alignment(wrap_text=True, vertical='top')

                for row in worksheet.iter_rows():
                    for cell in row:
                        cell.alignment = wrap_alignment

            self.main_window.log_message(f"已保存汇总记录到: {excel_filename}")
            return excel_filepath

        except Exception as e:
            self.main_window.log_message(f"保存汇总记录失败: {str(e)}", is_error=True)
            return None

    def save_grading_record(self, record_data):
        """
        重构后的保存阅卷记录到Excel文件的方法。
        - 动态构建Excel表头和行数据，支持单评和双评模式。
        - 设置列宽和格式，便于在Excel中查看。
        - 简化错误处理，直接缓存无法写入的记录。
        """
        # Prevent potential 'possibly unbound' references in except/ finally blocks by initializing variables
        excel_filepath = None
        excel_filename = ""
        try:
            # 记录汇总信息
            if record_data.get('record_type') == 'summary':
                return self._save_summary_record(record_data)

            # --- 1. 准备文件路径 ---
            excel_filepath = self._get_excel_filepath(record_data, self.worker)
            excel_filename = excel_filepath.name
            file_exists = excel_filepath.exists()

            # --- 2. 动态构建表头和行 ---
            is_dual = record_data.get('is_dual_evaluation', False)
            question_index_str = f"题目{record_data.get('question_index', 0)}"
            final_total_score_str = str(record_data.get('total_score', 0))

            headers = ["题目编号"]
            rows_to_write = []

            if is_dual:
                headers.extend(["API标识", "分差阈值", "学生答案摘要", "AI分项得分", "AI原始回复", "AI原始总分", "双评分差", "最终得分", "评分细则(前50字)"])

                rubric_str = record_data.get('scoring_rubric_summary', '未配置')
                
                row1 = [question_index_str,
                       "API-1",
                       str(record_data.get('score_diff_threshold', "未提供")),
                       record_data.get('api1_scoring_basis', '未提供'),
                       str(record_data.get('api1_itemized_scores', [])),
                       record_data.get('api1_raw_response', '未提供'),
                       str(record_data.get('api1_raw_score', 0.0)),
                       f"{record_data.get('score_difference', 0.0):.2f}",
                       final_total_score_str,
                       rubric_str]
                row2 = [question_index_str,
                       "API-2",
                       str(record_data.get('score_diff_threshold', "未提供")),
                       record_data.get('api2_scoring_basis', '未提供'),
                       str(record_data.get('api2_itemized_scores', [])),
                       record_data.get('api2_raw_response', '未提供'),
                       str(record_data.get('api2_raw_score', 0.0)),
                       f"{record_data.get('score_difference', 0.0):.2f}",
                       final_total_score_str,
                       rubric_str]
                rows_to_write.extend([row1, row2])
            else: # 单评模式
                headers.extend(["学生答案摘要", "AI分项得分", "AI原始回复", "最终得分", "评分细则(前50字)"])

                single_row = [question_index_str,
                             record_data.get('reasoning_basis', '无法提取'),
                             str(record_data.get('sub_scores', '未提供')),
                             record_data.get('raw_ai_response', '无法提取'),
                             final_total_score_str,
                             record_data.get('scoring_rubric_summary', '未配置')]
                rows_to_write.append(single_row)

            # --- 3. 写入Excel文件 ---
            if file_exists:
                # 如果文件存在，读取现有数据并追加
                try:
                    existing_df = pd.read_excel(excel_filepath, header=0)
                    new_df = pd.DataFrame(rows_to_write, columns=headers)
                    combined_df = pd.concat([existing_df, new_df], ignore_index=True)
                except Exception as e:
                    self.main_window.log_message(f"读取现有Excel文件失败: {str(e)}，将覆盖文件", True)
                    combined_df = pd.DataFrame(rows_to_write, columns=headers)
            else:
                combined_df = pd.DataFrame(rows_to_write, columns=headers)

            # 使用openpyxl引擎写入并设置格式
            with pd.ExcelWriter(excel_filepath, engine='openpyxl') as writer:
                combined_df.to_excel(writer, index=False, sheet_name='阅卷记录')

                # 获取工作簿和工作表
                workbook = writer.book
                worksheet = writer.sheets['阅卷记录']

                # 设置列宽
                column_widths = {
                    'A': 10,  # 题目编号
                    'B': 10,  # API标识 / 学生答案摘要
                    'C': 10,  # 分差阈值 / AI分项得分
                    'D': 80,  # 学生答案摘要（增加宽度以容纳较长的AI回答）
                    'E': 20,  # AI分项得分
                    'F': 200, # AI原始回复（增加宽度以容纳完整的JSON回复）
                    'G': 15,  # AI原始总分/最终得分
                    'H': 12,  # 双评分差
                    'I': 12,  # 最终得分

                    'L': 50   # 评分细则(前50字)
                }

                for col, width in column_widths.items():
                    if col in worksheet.column_dimensions:
                        worksheet.column_dimensions[col].width = width

                # 设置自动换行
                from openpyxl.styles import Alignment
                wrap_alignment = Alignment(wrap_text=True, vertical='top')

                for row in worksheet.iter_rows():
                    for cell in row:
                        cell.alignment = wrap_alignment

                # 设置标题行格式
                from openpyxl.styles import Font
                header_font = Font(bold=True)
                for cell in worksheet[1]:
                    cell.font = header_font

            self.main_window.log_message(f"已保存阅卷记录到: {excel_filename}")
            return excel_filepath

        except PermissionError as e:
            # 文件被占用，直接报错
            self.main_window.log_message(f"保存阅卷记录失败: Excel文件被占用，请关闭文件后重试。文件路径: {excel_filepath}", True)
            return None

        except Exception as e:
            error_detail_full = traceback.format_exc()
            self.main_window.log_message(f"保存阅卷记录失败: {str(e)}\n详细错误:\n{error_detail_full}", True)
            return None

    def check_excel_files_available(self):
        """检查阅卷记录Excel文件是否可用（未被锁定且无临时文件）"""
        try:
            # 获取当前配置
            dual_evaluation = self.config_manager.dual_evaluation_enabled
            question_config = self.config_manager.get_question_config(1)  # type: ignore[attr-defined]  # 单题模式只处理第一题
            full_score = question_config.get('max_score', 100) if question_config else 100

            # 手动构建文件路径（避免使用_get_excel_filepath的复杂逻辑）
            now = datetime.datetime.now()
            date_str = now.strftime('%Y年%m月%d日')
            evaluation_type = '双评' if dual_evaluation else '单评'
            excel_filename = f"此题最高{full_score}分_{evaluation_type}.xlsx"

            if getattr(sys, 'frozen', False):
                base_dir = pathlib.Path(sys.executable).parent
            else:
                base_dir = pathlib.Path(__file__).parent

            record_dir = base_dir / "阅卷记录"
            date_dir = record_dir / date_str
            excel_filepath = date_dir / excel_filename

            # 检查是否存在Excel临时文件（表示文件正在被打开）
            temp_file = date_dir / f"~${excel_filename}"
            if temp_file.exists():
                # 显示提醒对话框
                from PyQt5.QtWidgets import QMessageBox
                msg_box = QMessageBox(self.main_window)
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
                msg_box = QMessageBox(self.main_window)
                msg_box.setIcon(QMessageBox.Warning)
                msg_box.setWindowTitle("无法开始阅卷")
                msg_box.setText("阅卷记录文件被锁定，请关闭相关程序后重试。")
                msg_box.setStandardButtons(QMessageBox.Ok)
                msg_box.exec_()
                return False

            return True

        except Exception as e:
            self.main_window.log_message(f"检查Excel文件可用性时出错: {str(e)}", is_error=True)
            return False

    def start_auto_evaluation(self):
        """开始自动阅卷"""
        try:
            # 检查必要设置
            if not self.main_window.check_required_settings():
                return

            self.worker.start()
        except Exception as e:
            self.main_window.log_message(f"运行自动阅卷失败: {str(e)}", is_error=True)
            # 如果启动失败，确保UI状态正确
            self.main_window.update_ui_state(is_running=False)

    def run(self):
        """运行应用程序"""
        # 显示主窗口
        self.main_window.show()

        # 运行应用程序事件循环
        result = self.app.exec_()
        return result

if __name__ == "__main__":
    # 创建应用程序实例
    app = Application()

    # 运行应用程序
    sys.exit(app.run())
