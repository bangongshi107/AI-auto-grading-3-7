import sys
import os
import datetime
import pathlib
import warnings
import ctypes
import threading
from ctypes import wintypes
from typing import Callable

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
import traceback
import pandas as pd
import time

# 可选：用于“全局监听 Esc（不拦截按键）”以便最小化/切到其它窗口时也能中止阅卷。
# 若环境未安装 pynput，则会自动降级为仅支持窗口内快捷键（原行为）。
try:
    from pynput import keyboard as _pynput_keyboard  # type: ignore
except Exception:
    _pynput_keyboard = None

# 可选：用于“全局拦截并吞掉 Esc（仅在阅卷进行中启用）”，确保 Esc 只对本程序生效。
# 依赖第三方库 keyboard；未安装则会降级为不吞键的监听模式。
try:
    import keyboard as _keyboard  # type: ignore
except Exception:
    _keyboard = None


class _WindowsExclusiveEscHook:
    """Windows 低级键盘钩子：在启用时吞掉 Esc，并回调停止。

    目标：不引入新依赖；仅在“阅卷进行中”独占 Esc。
    注意：这是全局键盘钩子，可能被安全软件关注；已尽量降低副作用：
    - 默认禁用，仅在 worker 运行期间启用
    - 只处理 VK_ESCAPE
    """

    _WH_KEYBOARD_LL = 13
    _WM_KEYDOWN = 0x0100
    _WM_SYSKEYDOWN = 0x0104
    _VK_ESCAPE = 0x1B
    _HC_ACTION = 0
    _WM_QUIT = 0x0012

    class _KBDLLHOOKSTRUCT(ctypes.Structure):
        _fields_ = [
            ("vkCode", ctypes.c_uint32),
            ("scanCode", ctypes.c_uint32),
            ("flags", ctypes.c_uint32),
            ("time", ctypes.c_uint32),
            ("dwExtraInfo", ctypes.c_void_p),
        ]

    def __init__(self, should_swallow: Callable[[], bool], on_esc: Callable[[], None]):
        self._should_swallow = should_swallow
        self._on_esc = on_esc

        self._enabled = False
        self._lock = threading.Lock()
        self._last_hit_ts = 0.0

        self._hook_handle = None
        self._thread = None
        self._thread_id = 0
        self._proc_ref = None

    def set_enabled(self, enabled: bool) -> None:
        with self._lock:
            self._enabled = bool(enabled)

    def start(self) -> bool:
        if sys.platform != 'win32':
            return False
        if self._thread is not None:
            return True

        def _thread_main() -> None:
            try:
                user32 = ctypes.windll.user32
                kernel32 = ctypes.windll.kernel32

                LRESULT = ctypes.c_longlong if ctypes.sizeof(ctypes.c_void_p) == 8 else ctypes.c_long
                LowLevelProcType = ctypes.WINFUNCTYPE(LRESULT, ctypes.c_int, wintypes.WPARAM, wintypes.LPARAM)

                def _hook_proc(nCode, wParam, lParam):
                    try:
                        if nCode == self._HC_ACTION and wParam in (self._WM_KEYDOWN, self._WM_SYSKEYDOWN):
                            kb = ctypes.cast(lParam, ctypes.POINTER(self._KBDLLHOOKSTRUCT)).contents
                            if kb.vkCode == self._VK_ESCAPE:
                                with self._lock:
                                    enabled = self._enabled
                                if enabled and bool(self._should_swallow()):
                                    # 限流：避免长按重复触发
                                    now = time.time()
                                    if now - float(self._last_hit_ts) >= 0.25:
                                        self._last_hit_ts = now
                                        self._on_esc()
                                    return 1  # 吞掉 Esc
                        return user32.CallNextHookEx(self._hook_handle, nCode, wParam, lParam)
                    except Exception:
                        return user32.CallNextHookEx(self._hook_handle, nCode, wParam, lParam)

                self._proc_ref = LowLevelProcType(_hook_proc)
                self._thread_id = kernel32.GetCurrentThreadId()
                h_mod = kernel32.GetModuleHandleW(None)
                self._hook_handle = user32.SetWindowsHookExW(self._WH_KEYBOARD_LL, self._proc_ref, h_mod, 0)
                if not self._hook_handle:
                    return

                msg = wintypes.MSG()
                while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) != 0:
                    user32.TranslateMessage(ctypes.byref(msg))
                    user32.DispatchMessageW(ctypes.byref(msg))
            finally:
                try:
                    if self._hook_handle:
                        ctypes.windll.user32.UnhookWindowsHookEx(self._hook_handle)
                except Exception:
                    pass
                self._hook_handle = None

        t = threading.Thread(target=_thread_main, name="ExclusiveEscHook", daemon=True)
        t.start()
        self._thread = t
        return True

    def stop(self) -> None:
        if sys.platform != 'win32':
            return
        try:
            if self._thread_id:
                ctypes.windll.user32.PostThreadMessageW(self._thread_id, self._WM_QUIT, 0, 0)
        except Exception:
            pass


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
        # 按要求：不论任何原因，只要弹窗存在就每30秒提示一次
        self.play_system_sound()

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

        # 原始反馈（摘要）：避免长文本遮挡主要操作
        preview = (self.raw_feedback or "").strip()
        if preview:
            preview = preview.replace("\r\n", "\n").replace("\r", "\n")
            preview = preview[:200] + ("…" if len(preview) > 200 else "")
        else:
            preview = "(无)"

        fb_label = QLabel("AI反馈摘要（供参考）：\n" + preview)
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

        # 人工介入/阈值弹窗会先于 error_signal 到达。
        # 为避免紧接着再弹“阅卷中断”导致重复提示，做一个短时间的屏蔽窗口。
        self._suppress_error_dialog_until = 0.0

        # 全局 Esc 监听器（不拦截键盘输入）：仅在“正在阅卷”时响应，用于减少副作用。
        self._global_esc_listener = None
        self._global_esc_last_hit_ts = 0.0

        # “独占 Esc（吞键）”热键句柄：只在阅卷中启用，停止后立即解除。
        self._exclusive_esc_hotkey_id = None

        # 纯内置实现（ctypes）独占 Esc：不需要任何新依赖
        self._win_exclusive_esc = _WindowsExclusiveEscHook(
            should_swallow=lambda: bool(getattr(self, 'worker', None) and self.worker.isRunning()),
            on_esc=lambda: QTimer.singleShot(0, self.main_window.stop_auto_thread),
        )



        self._setup_application()

        # 在应用就绪后启动全局 Esc 支持：优先“吞键独占（仅阅卷中启用）”，不可用则降级为“监听不吞键”。
        self._setup_global_esc_support()

    def _simplify_for_teacher(self, text: str) -> str:
        """把底层错误压缩成老师能看懂的一句话 + 建议。"""
        t = (text or "").strip()
        low = t.lower()
        if any(k in low for k in ["timed out", "timeout"]):
            return "网络可能不稳定（连接超时）。建议：检查网络，稍等再试。"
        if any(k in low for k in ["401", "unauthorized", "invalid api key"]):
            return "密钥可能不正确或已失效。建议：重新复制密钥再试。"
        if any(k in low for k in ["403", "forbidden", "quota", "余额", "payment", "insufficient"]):
            return "账号可能没有权限或余额/额度不足。建议：检查账号余额/额度。"
        if any(k in low for k in ["429", "rate limit", "too many"]):
            return "请求太频繁，平台临时限制。建议：等10~30秒再试。"
        if any(k in low for k in ["502", "503", "504", "service unavailable", "bad gateway"]):
            return "平台服务繁忙或临时不可用。建议：稍后再试或换备用平台。"
        if any(k in low for k in ["permission", "access is denied", "被占用", "正在使用"]):
            return "文件可能被占用或没有写入权限。建议：关闭Excel后再试。"
        if not t:
            return "发生了错误，但没有收到具体原因。"
        return f"发生了错误：{t[:80]}{'…' if len(t) > 80 else ''}"

    def _setup_global_exception_hook(self):
        """设置全局异常钩子"""
        def handle_exception(exc_type, exc_value, exc_traceback):
            if issubclass(exc_type, KeyboardInterrupt):
                sys.__excepthook__(exc_type, exc_value, exc_traceback)
                return

            error_msg = "".join(traceback.format_exception(exc_type, exc_value, exc_traceback))
            
            # 记录到文件（堆栈仅写入日志，避免刷屏）
            log_file = None
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

            # 尝试记录到UI（只给一句人话 + 可选日志文件名）
            try:
                if hasattr(self, 'main_window') and hasattr(self.main_window, 'log_message'):
                    ui_msg = "程序内部错误，已停止当前操作。"
                    if log_file is not None:
                        ui_msg += f"（已保存日志：{log_file.name}）"
                    self.main_window.log_message(ui_msg, is_error=True)
            except Exception:
                pass

            # 显示一个简单的错误对话框
            try:
                user_tip = self._simplify_for_teacher(str(exc_value))
                dialog = SimpleNotificationDialog(
                    title="严重错误",
                    message=(
                        "程序遇到严重问题，可能需要关闭并重新打开。\n\n"
                        f"原因（简要）：{user_tip}\n\n"
                        "如果反复出现：请把程序目录 logs 文件夹里的最新日志发给技术人员。"
                    ),
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

    def _setup_global_esc_support(self) -> None:
        """配置 Esc 中止：优先“吞键独占（仅阅卷中启用）”。

        1) 首选：ctypes + Windows 低级键盘钩子（无需新依赖，且可吞键）
        2) 退化：keyboard 库（若用户已安装）
        3) 再退化：pynput 监听（不吞键）
        """
        if sys.platform != 'win32':
            return

        # 退出时清理（无论使用哪种实现）
        try:
            self.app.aboutToQuit.connect(self._disable_exclusive_esc)
        except Exception:
            pass
        try:
            self.app.aboutToQuit.connect(self._stop_global_esc_listener)
        except Exception:
            pass

        # 退出时停止 ctypes hook
        try:
            self.app.aboutToQuit.connect(self._stop_windows_exclusive_esc)
        except Exception:
            pass

        # 优先启用“纯内置”独占 Esc 钩子（默认禁用，阅卷开始时才启用吞键）
        try:
            if self._win_exclusive_esc.start():
                try:
                    if hasattr(self, 'main_window') and hasattr(self.main_window, 'log_message'):
                        self.main_window.log_message(
                            "已启用独占 Esc（内置实现）：阅卷进行中按 Esc 会停止，且不会传给其它程序。",
                            True,
                            "INFO",
                        )
                except Exception:
                    pass
                return
        except Exception as e:
            print(f"[WARN] 启动内置独占 Esc 失败: {e}")

        if _keyboard is not None:
            # 独占吞键模式：只在阅卷线程启动时才注册热键；此处只做提示。
            try:
                if hasattr(self, 'main_window') and hasattr(self.main_window, 'log_message'):
                    self.main_window.log_message(
                        "已启用独占 Esc 中止：阅卷进行中按 Esc 会停止，且不会传给其它程序。",
                        True,
                        "INFO",
                    )
            except Exception:
                pass
            return

        # 降级：仅监听不吞键
        self._setup_global_esc_to_stop_if_running()

    def _stop_windows_exclusive_esc(self) -> None:
        try:
            if getattr(self, '_win_exclusive_esc', None) is not None:
                self._win_exclusive_esc.stop()
        except Exception:
            pass

    def _enable_exclusive_esc_if_possible(self) -> None:
        """仅在阅卷开始时启用“吞键独占 Esc”。"""
        if sys.platform != 'win32':
            return

        # 优先：内置 hook
        try:
            if getattr(self, '_win_exclusive_esc', None) is not None:
                self._win_exclusive_esc.set_enabled(True)
                return
        except Exception:
            pass

        # 退化：keyboard 库
        if _keyboard is None:
            return
        if self._exclusive_esc_hotkey_id is not None:
            return

        def _on_esc() -> None:
            try:
                # 先解除吞键，尽快恢复系统默认行为（减少副作用）
                self._disable_exclusive_esc()
                QTimer.singleShot(0, self.main_window.stop_auto_thread)
            except Exception:
                pass

        try:
            # suppress=True: 吞掉 Esc，让 Esc 不再传递给其它程序
            self._exclusive_esc_hotkey_id = _keyboard.add_hotkey(
                'esc',
                _on_esc,
                suppress=True,
                trigger_on_release=False,
            )
        except Exception as e:
            print(f"[WARN] 启用独占 Esc 失败: {e}")
            self._exclusive_esc_hotkey_id = None

    def _disable_exclusive_esc(self) -> None:
        """阅卷结束/停止后解除“吞键独占 Esc”。"""
        # 先关内置 hook 的吞键开关
        try:
            if getattr(self, '_win_exclusive_esc', None) is not None:
                self._win_exclusive_esc.set_enabled(False)
        except Exception:
            pass

        # 再处理 keyboard 库（如果用到了）
        if _keyboard is None:
            return
        try:
            hid = self._exclusive_esc_hotkey_id
            if hid is not None:
                _keyboard.remove_hotkey(hid)
        except Exception:
            pass
        finally:
            self._exclusive_esc_hotkey_id = None

    def _setup_global_esc_to_stop_if_running(self) -> None:
        """全局监听 Esc，在阅卷中时中止（尽量减少副作用）。

        设计目标：
        - 不拦截/不吞掉 Esc（不影响其它软件正常接收 Esc）
        - 仅当 worker 正在运行时才触发停止
        - 通过 Qt 主线程安全调用 stop_auto_thread
        """
        if sys.platform != 'win32':
            return

        if _pynput_keyboard is None:
            # 不强制依赖；保持原有 UI 内快捷键可用。
            try:
                if hasattr(self, 'main_window') and hasattr(self.main_window, 'log_message'):
                    self.main_window.log_message(
                        "提示：未安装 keyboard/pynput，无法在切到其它窗口时用 Esc 中止；仍可点击“中止”按钮。",
                        True,
                        "INFO",
                    )
            except Exception:
                pass
            return

        kb = _pynput_keyboard

        def _on_press(key) -> None:
            try:
                if key != kb.Key.esc:
                    return

                # 限流：避免长按 Esc 或系统重复触发导致多次 stop
                now = time.time()
                if now - float(getattr(self, '_global_esc_last_hit_ts', 0.0)) < 0.35:
                    return
                self._global_esc_last_hit_ts = now

                # 只有阅卷线程运行时才响应，减少误触副作用
                if hasattr(self, 'worker') and self.worker and self.worker.isRunning():
                    # 跨线程：必须切回 Qt 主线程调用 UI/线程控制
                    QTimer.singleShot(0, self.main_window.stop_auto_thread)
            except Exception:
                # 全局监听器里不要抛异常，避免监听线程退出
                pass

        try:
            listener = kb.Listener(on_press=_on_press)
            listener.daemon = True
            listener.start()
            self._global_esc_listener = listener

            # 退出时尝试停止监听，避免后台线程残留
            try:
                self.app.aboutToQuit.connect(self._stop_global_esc_listener)
            except Exception:
                pass

            try:
                if hasattr(self, 'main_window') and hasattr(self.main_window, 'log_message'):
                    self.main_window.log_message(
                        "已启用全局 Esc 中止（降级模式）：阅卷中按 Esc 可停止，但 Esc 仍会传给其它程序。",
                        True,
                        "INFO",
                    )
            except Exception:
                pass
        except Exception as e:
            # 启动失败则静默降级
            print(f"[WARN] 启动全局 Esc 监听失败: {e}")

    def _stop_global_esc_listener(self) -> None:
        try:
            if self._global_esc_listener is not None:
                self._global_esc_listener.stop()
        except Exception:
            pass

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

            # 线程启动/结束：控制“独占 Esc（吞键）”是否启用（仅阅卷期间生效）
            try:
                self.worker.started.connect(self._enable_exclusive_esc_if_possible)
            except Exception:
                pass
            try:
                self.worker.finished_signal.connect(self._disable_exclusive_esc)
            except Exception:
                pass
            self.signal_manager.connect(
                self.worker.log_signal,
                self.main_window.log_message
            )
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
                try:
                    self.signal_manager.connect(self.worker.error_signal, lambda *_args: self._disable_exclusive_esc())
                except Exception:
                    pass

            # 双评分差超过阈值中断
            if hasattr(self.worker, 'threshold_exceeded_signal'):
                self.signal_manager.connect(
                    self.worker.threshold_exceeded_signal,
                    self.show_threshold_exceeded_notification # 这个方法内部需要调用 main_window.on_worker_error
                )
                try:
                    self.signal_manager.connect(self.worker.threshold_exceeded_signal, lambda *_args: self._disable_exclusive_esc())
                except Exception:
                    pass

            # 人工介入信号：当AI明确请求人工复核时触发
            if hasattr(self.worker, 'manual_intervention_signal'):
                self.signal_manager.connect(
                    self.worker.manual_intervention_signal,
                    self.show_manual_intervention_notification
                )
                try:
                    self.signal_manager.connect(self.worker.manual_intervention_signal, lambda *_args: self._disable_exclusive_esc())
                except Exception:
                    pass



        except Exception as e:
            # 避免在 main_window 可能还未完全初始化时调用其 log_message
            print(f"[CRITICAL_ERROR] 连接工作线程信号时出错: {str(e)}")
            if hasattr(self.main_window, 'log_message'):
                 self.main_window.log_message(f"连接工作线程信号时出错: {str(e)}", is_error=True)

    def show_completion_notification(self):
        """显示任务完成通知"""
        self._disable_exclusive_esc()
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
        self._disable_exclusive_esc()
        # 若刚刚触发了“人工介入”弹窗，则不再重复弹“阅卷中断”
        try:
            if time.time() < float(getattr(self, '_suppress_error_dialog_until', 0.0)):
                if hasattr(self.main_window, 'update_ui_state'):
                    self.main_window.update_ui_state(is_running=False)
                return
        except Exception:
            pass

        # 兜底：如果错误原因本身就是“需人工介入/异常试卷”，也不弹通用中断框
        try:
            msg_str = str(error_message)
            if any(k in msg_str for k in ["需人工介入", "需要人工介入", "人工介入", "异常试卷"]):
                if hasattr(self.main_window, 'update_ui_state'):
                    self.main_window.update_ui_state(is_running=False)
                return
        except Exception:
            pass

        # 用户主动停止：不弹“错误中断”，只恢复UI状态
        try:
            if "用户手动停止" in str(error_message) or "手动停止" in str(error_message):
                if hasattr(self.main_window, 'on_worker_error'):
                    self.main_window.on_worker_error("用户手动停止")
                else:
                    if hasattr(self.main_window, 'update_ui_state'):
                        self.main_window.update_ui_state(is_running=False)
                return
        except Exception:
            pass

        if hasattr(self.main_window, 'on_worker_error'):
            self.main_window.on_worker_error(error_message)
        else:
            print(f"[ERROR] MainWindow missing on_worker_error. Error: {error_message}")
            # 基本的后备恢复
            if self.main_window.isMinimized(): self.main_window.showNormal(); self.main_window.activateWindow()
            if hasattr(self.main_window, 'update_ui_state'): self.main_window.update_ui_state(is_running=False)

        # 给老师看的简要提示（不把英文/堆栈塞进弹窗）
        try:
            user_tip = self._simplify_for_teacher(str(error_message))
        except Exception:
            user_tip = "发生错误，自动阅卷已停止。"

        dialog = SimpleNotificationDialog(
            title="阅卷中断",
            message=(
                f"原因：{user_tip}\n\n"
                "建议：检查网络/密钥/模型ID；确认Excel已关闭；必要时切换备用AI平台。"
            ),
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
        self._disable_exclusive_esc()
        if hasattr(self.main_window, 'on_worker_error'):
            self.main_window.on_worker_error(reason)
        else:
            print(f"[ERROR] MainWindow missing on_worker_error. Reason: {reason}")
            # 基本的后备恢复
            if self.main_window.isMinimized(): self.main_window.showNormal(); self.main_window.activateWindow()
            if hasattr(self.main_window, 'update_ui_state'): self.main_window.update_ui_state(is_running=False)

        dialog = SimpleNotificationDialog(
            title="双评分差过大",
            message=(
                "两次评分差距过大，需要人工复核。\n\n"
                "建议：人工查看该题答题截图，确认分数后再继续下一份。"
            ),
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
        self._disable_exclusive_esc()
        # 标记：接下来短时间内如果收到 error_signal，不再重复弹“阅卷中断”
        try:
            self._suppress_error_dialog_until = time.time() + 2.0
        except Exception:
            self._suppress_error_dialog_until = 0.0

        # 只恢复UI状态，不重复走 on_worker_error（避免日志/建议堆叠）
        if self.main_window.isMinimized():
            self.main_window.showNormal()
            self.main_window.activateWindow()
        if hasattr(self.main_window, 'update_ui_state'):
            self.main_window.update_ui_state(is_running=False)

        # 显示模态对话框
        dialog = ManualInterventionDialog(
            title="人工介入",
            message=(f"{message}\n\n请人工检查并处理。"),
            raw_feedback=raw_feedback,
            sound_type='error',
            parent=self.main_window
        )
        dialog.exec_()
        
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
                headers.extend(["API标识", "分差阈值", "学生答案摘要", "AI分项得分", "AI评分依据", "AI原始总分", "双评分差", "最终得分", "评分细则(前50字)"])

                rubric_str = record_data.get('scoring_rubric_summary', '未配置')
                
                row1 = [question_index_str,
                       "API-1",
                       str(record_data.get('score_diff_threshold', "未提供")),
                       record_data.get('api1_student_answer_summary', '未提供'),
                       str(record_data.get('api1_itemized_scores', [])),
                       record_data.get('api1_scoring_basis', '未提供'),
                       str(record_data.get('api1_raw_score', 0.0)),
                       f"{record_data.get('score_difference', 0.0):.2f}",
                       final_total_score_str,
                       rubric_str]
                row2 = [question_index_str,
                       "API-2",
                       str(record_data.get('score_diff_threshold', "未提供")),
                       record_data.get('api2_student_answer_summary', '未提供'),
                       str(record_data.get('api2_itemized_scores', [])),
                       record_data.get('api2_scoring_basis', '未提供'),
                       str(record_data.get('api2_raw_score', 0.0)),
                       f"{record_data.get('score_difference', 0.0):.2f}",
                       final_total_score_str,
                       rubric_str]
                rows_to_write.extend([row1, row2])
            else: # 单评模式
                headers.extend(["学生答案摘要", "AI分项得分", "AI评分依据", "最终得分", "评分细则(前50字)"])

                single_row = [question_index_str,
                             record_data.get('student_answer', '无法提取'),
                             str(record_data.get('sub_scores', '未提供')),
                             record_data.get('reasoning_basis', '无法提取'),
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
                    'D': 80,  # 学生答案摘要
                    'E': 20,  # AI分项得分
                    'F': 200, # AI评分依据（增加宽度以容纳完整的评分依据）
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
