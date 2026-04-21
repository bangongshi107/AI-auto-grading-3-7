# main_window.py - 主窗口UI模块

import sys
import os
import traceback
import datetime
import pathlib
import re
from typing import Union, Optional, Type, TypeVar, cast, Tuple
from PyQt5.QtWidgets import (QMainWindow, QWidget, QMessageBox, QDialog,
                             QComboBox, QLineEdit, QCheckBox, QSpinBox, QDoubleSpinBox,
                             QPlainTextEdit, QApplication, QShortcut, QLabel, QPushButton)
from PyQt5.QtCore import Qt, pyqtSignal, QEvent, QObject, QTimer
from PyQt5.QtGui import QKeySequence, QFont, QKeyEvent, QCloseEvent, QIcon
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

    log_signal = pyqtSignal(str, bool, str)  # message, is_error, level
    progress_signal = pyqtSignal(int, int)
    finished_signal = pyqtSignal()


    def __init__(self, config_manager, api_service, worker):
        super().__init__()
        self.config_manager = config_manager
        self.api_service = api_service
        self.worker = worker
        self._is_initializing = True
        self._pending_save = False
        self._last_save_reason = ""
        self._autosave_timer: Optional[QTimer] = None

        # 加载UI文件
        if getattr(sys, 'frozen', False):
            base_path = sys._MEIPASS  # type: ignore
        else:
            base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        ui_path = os.path.join(base_path, "setting", "七题.ui")
        uic.loadUi(ui_path, self)

        # 设置窗口图标
        try:
            icon_path = None
            if getattr(sys, 'frozen', False):
                meipass = getattr(sys, '_MEIPASS', None)
                if meipass:
                    candidate = os.path.join(meipass, 'AI阅卷助手.ico')
                    if os.path.exists(candidate):
                        icon_path = candidate
                if not icon_path:
                    candidate = os.path.join(os.path.dirname(sys.executable), 'AI阅卷助手.ico')
                    if os.path.exists(candidate):
                        icon_path = candidate
            else:
                candidate = os.path.join(base_path, 'AI阅卷助手.ico')
                if os.path.exists(candidate):
                    icon_path = candidate

            if icon_path:
                self.setWindowIcon(QIcon(icon_path))
        except Exception:
            pass  # 图标加载失败不影响程序运行

        # 初始化属性
        self.answer_windows = {}
        self.current_question = 1
        self.max_questions = 7  # 多题模式最多支持7道题
        self._ui_cache = {}

        self.init_ui()
        self._setup_autosave_timer()



        self.show()
        self._is_initializing = False
        self.log_message("主窗口初始化完成")

    # ======================================================================
    #  面向老师的“人话提示”工具
    # ======================================================================

    def _mask_secret(self, value: str) -> str:
        s = (value or "").strip()
        if not s:
            return "(空)"
        if len(s) <= 8:
            return "***"
        return f"{s[:4]}****{s[-4:]}"

    def _display_name_for_field(self, field_name: str) -> str:
        """将内部字段名转换为老师可读的中文标签。"""
        f = (field_name or "").strip()
        mapping = {
            "first_api_provider": "第一组AI平台",
            "first_api_key": "第一组密钥",
            "first_modelID": "第一组模型ID",
            "second_api_provider": "第二组AI平台",
            "second_api_key": "第二组密钥",
            "second_modelID": "第二组模型ID",
            "dual_evaluation_enabled": "双评模式",
            "score_diff_threshold": "分差阈值",
            "subject": "学科",
            "cycle_number": "循环次数",
            "wait_time": "间隔时间(秒)",
            "unattended_mode_enabled": "无人模式",
        }
        if f in mapping:
            return mapping[f]

        m = re.match(r"^question_(\d+)_enabled$", f)
        if m:
            return f"第{m.group(1)}题启用"

        m = re.match(r"^question_(\d+)_standard_answer$", f)
        if m:
            return f"第{m.group(1)}题评分细则"

        m = re.match(r"^question_(\d+)_work_mode$", f)
        if m:
            return f"第{m.group(1)}题工作模式"

        return f

    def _get_base_dir(self) -> pathlib.Path:
        """获取可写日志目录的基准路径（兼容打包/源码运行）。"""
        try:
            if getattr(sys, 'frozen', False):
                return pathlib.Path(sys.executable).parent
        except Exception:
            pass
        return pathlib.Path(__file__).resolve().parent.parent

    def _write_debug_log(self, title: str, detail: str) -> Optional[pathlib.Path]:
        """写入调试日志（给技术人员/开发者看），不打扰普通用户。"""
        try:
            base_dir = self._get_base_dir()
            log_dir = base_dir / "logs"
            log_dir.mkdir(exist_ok=True)
            now = datetime.datetime.now()
            filename = f"ui_{title}_{now.strftime('%Y%m%d_%H%M%S')}.log"
            file_path = log_dir / filename
            with open(file_path, "w", encoding="utf-8") as f:
                f.write(detail or "")
            return file_path
        except Exception:
            return None

    def _simplify_message_for_teacher(self, text: str) -> Tuple[str, str]:
        """把复杂/英文/堆栈信息压缩成老师能看懂的提示。

        Returns:
            (summary, detail)
            - summary: 给用户看的简短说明 + 建议操作
            - detail: 原始信息（可放到“详细信息”或日志文件）
        """
        original = (text or "").strip()
        if not original:
            return "发生了问题，但没有收到具体原因。", ""

        detail = original
        low = original.lower()

        # ==================================================================
        # 先做“业务前缀/建议行”去噪：避免 UI 堆叠同一句话
        # ==================================================================
        # 去掉常见错误前缀
        cleaned_for_parse = re.sub(r"^\s*\[(错误|业务错误|网络错误|配置错误|资源错误|系统错误)\]\s*", "", original).strip()

        # 如果包含“→ 建议: ...”，UI主消息只保留第一句原因；建议由本函数统一给出
        # （detail 仍保留原始文本，便于排查）
        if "→" in cleaned_for_parse and "建议" in cleaned_for_parse:
            cleaned_for_parse = re.split(r"\n\s*→\s*建议\s*:\s*", cleaned_for_parse, maxsplit=1)[0].strip()

        # 用户主动停止：不算错误，也不需要“检查密钥/网络”等建议
        if any(k in low for k in ["用户手动停止", "手动停止", "user stopped", "user stop"]):
            return "已停止（用户手动停止）。", ""

        # 若包含 traceback，正文只给一句“程序内部出错”，细节进日志
        if "traceback (most recent call last)" in low:
            return "程序内部出现了错误，已停止当前操作。\n建议：关闭软件重新打开后再试一次。", detail

        # ==================================================================
        # 关键场景：异常试卷 / 无有效内容 / 需要人工介入
        # 目标：只给老师一句“发生了什么 + 下一步做什么”，不再堆叠多条来源信息。
        # ==================================================================
        if any(k in cleaned_for_parse for k in ["异常试卷", "无有效内容"]):
            # 尝试提取题号
            q_match = re.search(r"题目\s*(\d+)", cleaned_for_parse) or re.search(r"第\s*(\d+)\s*题", cleaned_for_parse)
            q = q_match.group(1) if q_match else ""

            # 提取括号内原因：例如 (无有效内容)
            reason = ""
            m = re.search(r"异常试卷\s*\(?\s*([^\)\n]+?)\s*\)?", cleaned_for_parse)
            if m:
                reason = m.group(1).strip()
            if not reason and "无有效内容" in cleaned_for_parse:
                reason = "无有效内容"

            reason_part = f"（{reason}）" if reason else ""
            head = f"题目{q}：" if q else ""

            # 是否提示“启用异常卷按钮”
            need_button_tip = any(k in cleaned_for_parse for k in ["未启用异常卷按钮", "未配置坐标"])
            tip = "可选：在题目配置里启用“异常卷按钮”，下次可自动跳过。" if need_button_tip else ""

            summary = f"{head}检测到异常试卷{reason_part}。已暂停，请人工处理后继续。"
            if tip:
                summary += f"\n{tip}"
            return summary, detail

        if any(k in cleaned_for_parse for k in ["需人工介入", "需要人工介入", "人工介入"]):
            # 【优化】尝试提取AI给出的具体原因（去掉"需人工介入:"前缀）
            reason_text = ""
            for line in cleaned_for_parse.split('\n'):
                line = line.strip()
                # 跳过纯标记行
                if line in ["需人工介入", "人工介入", "需要人工介入"]:
                    continue
                # 去掉常见前缀，提取实际原因
                for prefix in ["需人工介入:", "需人工介入：", "需要人工介入:", "需要人工介入：", "[需人工介入]"]:
                    if line.startswith(prefix):
                        line = line[len(prefix):].strip()
                        break
                if line and len(line) > 5:
                    reason_text = line
                    break
            
            # 尝试保留题号信息
            q_match = re.search(r"题目\s*(\d+)", cleaned_for_parse) or re.search(r"第\s*(\d+)\s*题", cleaned_for_parse)
            q = q_match.group(1) if q_match else ""
            head = f"题目{q}：" if q else ""
            
            # 如果提取到了具体原因，显示它；否则用通用提示
            if reason_text:
                return f"{head}{reason_text}", detail
            else:
                return f"{head}需要人工介入处理。已暂停，请处理后继续。", detail

        # 去掉常见 emoji/符号，减少干扰
        cleaned = re.sub(r"[✅❌⚠️💡]", "", original).strip()

        # 统一术语为更口语的中文
        replacements = {
            "api": "AI接口",
            "key": "密钥",
            "model": "模型",
            "model id": "模型ID",
            "unauthorized": "未授权",
            "forbidden": "无权限",
            "rate limit": "请求太频繁",
            "timeout": "网络超时",
        }
        simplified = cleaned
        for k, v in replacements.items():
            simplified = re.sub(k, v, simplified, flags=re.IGNORECASE)

        # ==================================================================
        # 成功场景：连接测试通过
        # 说明：test_api_connection() 成功时会返回类似“火山引擎 (推荐)：连接成功”。
        # 这里要直接按成功展示，避免被默认分支包装成“操作未成功：...”。
        # ==================================================================
        success_markers = [
            "连接成功",
            "测试通过",
            "可正常使用",
        ]
        if any(m in simplified for m in success_markers) or any(m in cleaned_for_parse for m in success_markers):
            # 保留平台名等信息；只做最基础的去噪
            ok_text = (simplified or cleaned).strip()
            return ok_text, detail

        # JSON/响应格式问题：通常是模型输出不符合要求（不要提示“检查密钥”）
        if any(k in low for k in ["json解析", "json parse", "响应格式", "api响应格式异常", "format" ]):
            return (
                "AI接口返回格式异常，已停止当前操作。\n"
                "建议：切换模型或更换AI平台后再试。",
                detail,
            )

        # 典型错误归因（尽量“原因 + 怎么办”）
        if any(k in low for k in ["timed out", "timeout", "read timed out"]):
            return (
                "网络可能不稳定，连接超时。\n"
                "建议：1）检查网络是否能上网  2）稍等1分钟再点一次“测试/开始”。",
                detail,
            )

        if any(k in low for k in ["401", "unauthorized", "invalid api key", "api key"]):
            return (
                "AI平台提示“密钥不正确或已失效”。\n"
                "建议：到平台后台重新复制密钥，粘贴到软件里再测试。",
                detail,
            )

        if any(k in low for k in ["403", "forbidden", "insufficient", "quota", "余额", "payment"]):
            return (
                "AI平台账号可能没有权限或余额不足。\n"
                "建议：检查账号余额/额度；必要时更换一个可用的AI平台。",
                detail,
            )

        if any(k in low for k in ["429", "请求太频繁", "rate limit", "too many"]):
            return (
                "请求太频繁，AI平台暂时不让访问。\n"
                "建议：等10~30秒再试；或开启/使用第二组AI作为备用。",
                detail,
            )

        if any(k in low for k in ["502", "503", "504", "service unavailable", "bad gateway"]):
            return (
                "AI平台当前服务繁忙或临时不可用。\n"
                "建议：稍后再试；或切换到第二组AI平台。",
                detail,
            )

        if any(k in low for k in ["permission", "permissionerror", "access is denied", "被占用", "正在使用"]):
            return (
                "文件可能正在被占用，或没有写入权限。\n"
                "建议：1）关闭所有Excel文件  2）把软件放到桌面/D盘再运行  3）再试一次。",
                detail,
            )

        # 默认：给一个稳妥的通用说明（保持简短，不堆叠括号/前后缀）
        short_reason = f"{simplified[:80]}{'…' if len(simplified) > 80 else ''}".strip()
        return (f"操作未成功：{short_reason}。建议：检查密钥/模型ID/网络后再试。", detail)

    def _normalize_log_text(self, text: str, preserve_newlines: bool = False) -> str:
        """对日志文本做去噪与去重（面向主界面日志区/弹窗）。"""
        t = (text or "").strip()
        if not t:
            return ""

        # 去掉常见重复前缀
        prefixes = [
            "[提示]",
            "[信息]",
            "[错误]",
            "错误:",
            "错误：",
            "操作未成功：",
            "操作未成功:",
            "任务已停止：",
            "任务已停止:",
            "任务已停止",
            "需要人工介入:",
            "需要人工介入：",
        ]
        changed = True
        while changed:
            changed = False
            for p in prefixes:
                if t.startswith(p):
                    t = t[len(p):].strip()
                    changed = True

        # 清理奇怪的冒号/括号堆叠
        t = re.sub(r"[:：]{2,}", "：", t)

        # 默认会把所有空白（含换行）压成单个空格，避免日志区刷屏。
        # 但 RESULT 需要保留换行（例如：标题行 + 评分依据明细）。
        if preserve_newlines:
            lines = [re.sub(r"[\t\f\v ]+", " ", line).strip() for line in t.splitlines()]
            # 去掉空行（避免出现很多空白段落）
            lines = [line for line in lines if line]
            t = "\n".join(lines).strip()
        else:
            t = re.sub(r"\s+", " ", t).strip()

        # 统一一些“重复来源”表述（避免同一句话出现多种开头）
        t = re.sub(r"^API\s*[12]\s*检测到异常试卷\s*[:：]\s*", "检测到异常试卷：", t)
        t = re.sub(r"^检测到异常试卷\s*[:：]\s*", "检测到异常试卷：", t)

        # 若包含多段“建议：...建议：...”，只保留第一段（UI不刷屏，细节在logs）
        if t.count("建议：") >= 2:
            first, _, rest = t.partition("建议：")
            # first 里可能还带一段内容，把第一个“建议：xxx”拼回去
            second = rest.split("建议：", 1)[0].strip()
            t = (first + "建议：" + second).strip()
        return t

    def _escape_html(self, text: str) -> str:
        return (
            (text or "")
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
        )

    def _show_message(self, title: str, summary: str, icon=QMessageBox.Warning, detail: str = "") -> None:
        """统一的消息框：主文本简单易懂，技术细节放到详细信息。"""
        msg_box = QMessageBox(self)
        msg_box.setIcon(icon)
        msg_box.setWindowTitle(title)
        msg_box.setText(summary)
        if detail:
            msg_box.setDetailedText(detail)
        msg_box.setSizeGripEnabled(True)
        msg_box.setMinimumSize(680, 320)
        msg_box.setStyleSheet("QLabel{min-width: 560px;}")
        msg_box.setStandardButtons(QMessageBox.Ok)
        msg_box.exec_()

    # ==========================================================================
    #  核心修改：配置处理逻辑
    # ==========================================================================

    def handle_lineEdit_save(self, field_name, value):
        if self._is_initializing: return
        self.config_manager.update_config_in_memory(field_name, value)

        label = self._display_name_for_field(str(field_name))
        # 密钥类内容不在UI里展示明文
        if str(field_name) in ["first_api_key", "second_api_key"]:
            self.log_message(f"{label} 已更新（已隐藏）：{self._mask_secret(str(value))}")
        else:
            self.log_message(f"{label} 已更新：{value}")
        self._mark_config_dirty(f"lineedit:{field_name}")

    def handle_plainTextEdit_save(self, field_name, value):
        if self._is_initializing: return
        self.config_manager.update_config_in_memory(field_name, value)
        # 答案内容较长，日志可以简洁些
        label = self._display_name_for_field(str(field_name))
        self.log_message(f"{label} 已更新")
        self._mark_config_dirty(f"plaintext:{field_name}")

    def handle_spinBox_save(self, field_name, value):
        if self._is_initializing: return
        self.config_manager.update_config_in_memory(field_name, value)
        label = self._display_name_for_field(str(field_name))
        self.log_message(f"{label} 已更新：{value}")
        self._mark_config_dirty(f"spinbox:{field_name}")
    
    def handle_doubleSpinBox_save(self, field_name, value):
        """处理 QDoubleSpinBox 控件的保存"""
        if self._is_initializing: return
        self.config_manager.update_config_in_memory(field_name, value)
        label = self._display_name_for_field(str(field_name))
        self.log_message(f"{label} 已更新：{value}")
        self._mark_config_dirty(f"doublespin:{field_name}")
    
    # --- 统一的 ComboBox 处理函数 ---
    def handle_comboBox_save(self, combo_box_name, ui_text):
        """统一的ComboBox保存处理
        
        重要说明：
        - first_api_url/second_api_url: 只处理AI评分模型提供商
        """
        if self._is_initializing: return

        if combo_box_name in ['first_api_url', 'second_api_url']:
            # 处理AI评分模型提供商 ComboBox
            combo_box = self.get_ui_element(combo_box_name, QComboBox)
            provider_id = None
            if combo_box and isinstance(combo_box, QComboBox):
                provider_id = combo_box.currentData()
            if not provider_id:
                provider_id = get_provider_id_from_ui_text(ui_text)
            if not provider_id:
                self.log_message(f"错误: 无法识别的AI模型提供商 '{ui_text}'", is_error=True)
                return
            field_name = 'first_api_provider' if combo_box_name == 'first_api_url' else 'second_api_provider'
            self.config_manager.update_config_in_memory(field_name, provider_id)
            label = self._display_name_for_field(str(field_name))
            display_text = combo_box.currentText() if combo_box else ui_text
            self.log_message(f"{label} 已更新为：{display_text}")
            self._mark_config_dirty(f"combobox:{field_name}")
        elif combo_box_name.startswith('work_mode_'):
            q_index = combo_box_name.replace('work_mode_', '')
            combo_box = self.get_ui_element(combo_box_name, QComboBox)
            work_mode = None
            if combo_box and isinstance(combo_box, QComboBox):
                work_mode = combo_box.currentData()
            if not work_mode:
                normalized_text = self._normalize_work_mode_ui_text(ui_text)
                mode_map = {
                    '识图直评': 'direct_grade',
                    '直评+推理': 'direct_grade_thinking',
                    '识评分离': 'ocr_then_grade',
                    '分离+推理': 'ocr_then_grade_thinking',
                    '分离+双推理': 'ocr_then_grade_dual_thinking'
                }
                work_mode = mode_map.get(normalized_text, 'direct_grade')
            field_name = f"question_{q_index}_work_mode"
            self.config_manager.update_config_in_memory(field_name, work_mode)
            label = self._display_name_for_field(str(field_name))
            display_text = combo_box.currentText() if combo_box else ui_text
            self.log_message(f"{label} 已更新为：{display_text}")
            self._apply_ui_constraints()
            self._mark_config_dirty(f"combobox:{field_name}")
        else:
            # 处理普通ComboBox（如subject_text）
            field_name = combo_box_name.replace('_text', '')  # subject_text -> subject
            self.config_manager.update_config_in_memory(field_name, ui_text)
            label = self._display_name_for_field(str(field_name))
            self.log_message(f"{label} 已更新为：{ui_text}")
            self._mark_config_dirty(f"combobox:{field_name}")

    def handle_checkBox_save(self, field_name, state):
        if self._is_initializing: return
        value = bool(state)
        self.config_manager.update_config_in_memory(field_name, value)
        label = self._display_name_for_field(str(field_name))
        self.log_message(f"{label} 已更新为：{'开启' if value else '关闭'}")
        self._mark_config_dirty(f"checkbox:{field_name}")

    # ======================================================================
    #  统一的保存/加载体系：内存更新 + 定时/关键操作落盘
    # ======================================================================

    def _setup_autosave_timer(self) -> None:
        """启动自动保存计时器（只在有变更时落盘）。"""
        try:
            if self._autosave_timer is None:
                self._autosave_timer = QTimer(self)
                self._autosave_timer.setInterval(60000)  # 60秒一次
                self._autosave_timer.timeout.connect(self._autosave_handler)
            if not self._autosave_timer.isActive():
                self._autosave_timer.start()
        except Exception:
            # 定时器异常不影响主流程
            pass
    
    def _autosave_handler(self) -> None:
        """自动保存定时器的槽函数，确保返回类型为 None。"""
        self._flush_config_to_file("autosave")

    def _mark_config_dirty(self, reason: str = "") -> None:
        """标记当前配置有变更，仅保存到内存，等待定时/关键操作落盘。"""
        if self._is_initializing:
            return
        self._pending_save = True
        if reason:
            self._last_save_reason = reason

    def _flush_config_to_file(self, reason: str = "") -> bool:
        """将内存中的配置统一保存到文件。"""
        try:
            if not self._pending_save and reason == "autosave":
                return True
            # 保存前同步工作模式（避免UI未失焦导致内存未更新）
            self._sync_work_mode_from_ui()

            ok = self.config_manager.save_all_configs_to_file()
            if ok:
                self._pending_save = False
                if reason:
                    self._last_save_reason = reason
            return ok
        except Exception:
            return False

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

        # cycle_number 使用 QSpinBox
        cycle_widget = self.get_ui_element('cycle_number', QSpinBox)
        if cycle_widget:
            cycle_widget.valueChanged.connect(
                lambda val: self.handle_spinBox_save('cycle_number', val)
            )
        
        # wait_time 使用 QDoubleSpinBox
        wait_widget = self.get_ui_element('wait_time', QDoubleSpinBox)
        if wait_widget:
            wait_widget.valueChanged.connect(
                lambda val: self.handle_doubleSpinBox_save('wait_time', val)
            )

        for i in range(1, self.max_questions + 1):
            std_answer_widget = self.get_ui_element(f'StandardAnswer_text_{i}', QPlainTextEdit)
            if std_answer_widget:
                self._connect_plain_text_edit_save_signal(std_answer_widget, i)

        for i in range(1, self.max_questions + 1):
            work_mode_combo = self.get_ui_element(f'work_mode_{i}', QComboBox)
            if work_mode_combo:
                work_mode_combo.currentTextChanged.connect(
                    lambda text, name=f'work_mode_{i}': self.handle_comboBox_save(name, text)
                )

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
        for combo_name in ['first_api_url', 'second_api_url']:
            combo_box = self.get_ui_element(combo_name, QComboBox)
            if combo_box and isinstance(combo_box, QComboBox):
                combo_box.clear()
                for ui_text, provider_id in UI_TEXT_TO_PROVIDER_ID.items():
                    combo_box.addItem(ui_text, provider_id)

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

        # 初始化每题工作模式下拉框
        for i in range(1, self.max_questions + 1):
            work_mode_combo = self.get_ui_element(f'work_mode_{i}', QComboBox)
            if work_mode_combo and isinstance(work_mode_combo, QComboBox):
                work_mode_combo.clear()
                work_mode_combo.addItem("一 识图直评", "direct_grade")
                work_mode_combo.addItem("二 直评+推理", "direct_grade_thinking")
                work_mode_combo.addItem("三 识评分离", "ocr_then_grade")
                work_mode_combo.addItem("四 分离+单推理", "ocr_then_grade_thinking")
                work_mode_combo.addItem("五 分离+双推理", "ocr_then_grade_dual_thinking")
                work_mode_combo.setToolTip("识图直评：AI看图直接评分；直评+推理：看图评分并开启推理；识评分离：AI识别文字后评分；分离+推理：识别不推理、评分开启推理；分离+双推理：识别与评分均开启推理")

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
                    index = combo_box.findData(provider_id)
                    if index >= 0:
                        combo_box.setCurrentIndex(index)
                    else:
                        # 兼容旧配置：尝试通过显示文本定位
                        ui_text_to_select = get_ui_text_from_provider_id(provider_id)
                        if ui_text_to_select:
                            combo_box.setCurrentText(ui_text_to_select)
                        else:
                            combo_box.setCurrentIndex(0)

            # 加载其他配置
            subject_widget = self.get_ui_element('subject_text', QComboBox)
            if subject_widget: subject_widget.setCurrentText(self.config_manager.subject)
            
            cycle_element = self.get_ui_element('cycle_number')
            if cycle_element and isinstance(cycle_element, QSpinBox):
                cycle_element.setValue(self.config_manager.cycle_number)
            
            wait_element = self.get_ui_element('wait_time', QDoubleSpinBox)
            if wait_element and isinstance(wait_element, QDoubleSpinBox):
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
                unattended_element.setToolTip("启用后，AI评分失败时自动给保守分并标记待复核，系统不停机继续阅卷。\n• 标准模式：按填充率给分（<25%给0分，否则给步长最小分）\n• 三步打分：每步各给1分（共3分），便于快速定位回评\n禁用：遇到问题时立即停止，等待人工处理")

            # 加载题目配置
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

                # 加载每题工作模式
                work_mode_combo = self.get_ui_element(f'work_mode_{i}', QComboBox)
                if work_mode_combo and isinstance(work_mode_combo, QComboBox):
                    work_mode_value = q_config.get('work_mode', 'direct_grade')
                    index = work_mode_combo.findData(work_mode_value)
                    if index >= 0:
                        work_mode_combo.setCurrentIndex(index)
                    else:
                        display_text_map = {
                            'direct_grade': '一 识图直评',
                            'direct_grade_thinking': '二 直评+推理',
                            'ocr_then_grade': '三 识评分离',
                            'ocr_then_grade_thinking': '四 分离+单推理',
                            'ocr_then_grade_dual_thinking': '五 分离+双推理'
                        }
                        display_text = display_text_map.get(work_mode_value, '一 识图直评')
                        work_mode_combo.setCurrentText(display_text)
                

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
            self._pending_save = False
            self._config_loaded_once = True
        except Exception as e:
            detail = traceback.format_exc()
            log_path = self._write_debug_log("load_config", detail)
            msg = "读取设置时出错，但不影响打开主界面。\n建议：关闭软件重新打开；如果反复出现，请把 logs 里的日志发给技术人员。"
            if log_path:
                msg += f"\n（已保存日志：{log_path.name}）"
            self.log_message(msg, is_error=True)
        finally:
            self._is_initializing = False

    def auto_run_but_clicked(self):
        """自动运行按钮点击事件"""
        # 先做启动前校验（包含：供应商UI文本→内部ID归一化、必要坐标检查等），避免“保存了错误配置”或“启动→秒停”。
        if not self.check_required_settings():
            return

        # 保存前同步工作模式选择
        self._sync_work_mode_from_ui()

        self.log_message("尝试在运行前保存所有配置...")
        if not self._flush_config_to_file("before_run"):
            self.log_message("保存设置失败，自动阅卷无法开始。", is_error=True)
            self._show_message(
                title="保存设置失败",
                icon=QMessageBox.Critical,
                summary=(
                    "保存设置失败，自动阅卷无法开始。\n\n"
                    "常见原因：\n"
                    "1）Excel（阅卷记录）还开着，导致文件被占用\n"
                    "2）软件所在文件夹没有写入权限\n\n"
                    "建议：先关闭所有Excel文件；把软件放到桌面或D盘；再点一次“开始自动阅卷”。"
                ),
            )
            return
        self.log_message("所有配置已成功保存。")

        # 显示提醒对话框
        msg_box = QMessageBox(self)
        msg_box.setIcon(QMessageBox.Warning)
        msg_box.setWindowTitle("开始前请确认")
        msg_box.setText(
            "开始自动阅卷前，请先把所有Excel表格关闭。\n"
            "（尤其是‘阅卷记录’相关的Excel文件）\n\n"
            "否则：可能保存不了阅卷记录，甚至中途报错。"
        )
        ok_button = msg_box.addButton("我已关闭Excel，开始自动阅卷", QMessageBox.AcceptRole)
        cancel_button = msg_box.addButton("取消", QMessageBox.RejectRole)
        msg_box.setDefaultButton(ok_button)
        msg_box.setSizeGripEnabled(True)
        msg_box.setMinimumSize(680, 260)
        msg_box.setStyleSheet("QLabel{min-width: 560px;}")
        result = msg_box.exec_()

        # 检查用户是否点击了"开始自动阅卷"按钮（而不是点击X或取消）
        if msg_box.clickedButton() == ok_button:
            # 用户确认后，直接启动自动阅卷（无延迟）
            self._start_auto_evaluation_after_confirmation()
        else:
            # 用户点击了取消或X关闭窗口
            self.log_message("用户取消了自动阅卷操作")
            return

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

            # 识评分离模式下禁用双评
            if dual_evaluation:
                for q_idx in enabled_questions_indices:
                    q_cfg = self.config_manager.get_question_config(q_idx)
                    if q_cfg.get('work_mode') in {'ocr_then_grade', 'ocr_then_grade_thinking', 'ocr_then_grade_dual_thinking'}:
                        dual_evaluation = False
                        dual_eval_checkbox = self.get_ui_element('dualEvaluationCheckbox')
                        if dual_eval_checkbox:
                            dual_eval_checkbox.setChecked(False)
                        self.log_message("识评分离模式下自动禁用双评功能", is_error=False)
                        break

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
            detail = traceback.format_exc()
            summary, _ = self._simplify_message_for_teacher(str(e))
            log_path = self._write_debug_log("start_run", detail)
            if log_path:
                summary += f"\n（已保存日志：{log_path.name}）"
            self.log_message("启动自动阅卷失败。" + summary, is_error=True)
            self._show_message("启动失败", summary, icon=QMessageBox.Critical, detail=detail)

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
            errors.append("第一组：请选择一个AI平台（下拉框里选）")
        else:
            # 写回内存，确保后续保存会落盘为内部ID
            self.config_manager.update_config_in_memory('first_api_provider', first_provider_id)

        if not self.config_manager.first_api_key.strip():
            errors.append("第一组：密钥不能为空（在平台后台复制粘贴）")
        if not self.config_manager.first_modelID.strip():
            errors.append("第一组：模型ID不能为空（例如模型名称/ID）")

        # 始终要求配置第二组API（用于故障转移）
        second_provider_id = _resolve_provider_to_id(getattr(self.config_manager, 'second_api_provider', ''))
        if not second_provider_id:
            errors.append("第二组：请选择一个AI平台（用于备用/故障切换）")
        else:
            self.config_manager.update_config_in_memory('second_api_provider', second_provider_id)

        if not self.config_manager.second_api_key.strip():
            errors.append("第二组：密钥不能为空（用于备用/故障切换）")
        if not self.config_manager.second_modelID.strip():
            errors.append("第二组：模型ID不能为空（用于备用/故障切换）")

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
            title = "还差几项设置，先补齐"
            intro = "自动阅卷现在不能开始，请按下面清单补齐：\n"
            error_details = "\n".join([f"  - {e}" for e in errors])
            final_message = f"{intro}\n{error_details}\n\n补齐后，再点一次“开始自动阅卷”。"

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

    def test_api_connections(self):
        """测试API连接（强制测试两个API）"""
        try:
            # 测试前无需手动更新，因为 ApiService 每次都会从 ConfigManager 获取最新配置
            self.log_message("正在测试API连接...")
            success1, message1 = self.api_service.test_api_connection("first")
            
            # 强制测试第二个API（不管双评模式是否开启）
            self.log_message("正在测试第二个API...")
            success2, message2 = self.api_service.test_api_connection("second")
            
            s1, d1 = self._simplify_message_for_teacher(message1)
            s2, d2 = self._simplify_message_for_teacher(message2)
            result_message = (
                f"【第一组AI平台】\n{s1}\n\n"
                f"【第二组AI平台】\n{s2}"
            )
            
            if success1 and success2: 
                self.log_message("测试完成：所有API均可正常使用")
            else: 
                self.log_message("测试完成：部分API无法正常使用", is_error=True)

            # 创建完整显示的API测试结果提示框
            details = "\n\n".join([
                "[第一组-原始信息]",
                d1,
                "\n[第二组-原始信息]",
                d2,
            ]).strip()
            self._show_message(
                title="AI平台连接测试",
                icon=QMessageBox.Information if (success1 and success2) else QMessageBox.Warning,
                summary=result_message,
                detail=details,
            )
        except Exception as e:
            detail = traceback.format_exc()
            summary, _ = self._simplify_message_for_teacher(str(e))
            self.log_message("AI平台连接测试失败：" + summary, is_error=True)
            self._show_message("AI平台连接测试失败", summary, icon=QMessageBox.Critical, detail=detail)

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

        # 循环结束后，清空字典
        self.answer_windows.clear()

        # 保存前同步工作模式选择
        self._sync_work_mode_from_ui()

        # 保存配置
        self.log_message("尝试在关闭程序前保存所有配置...")
        if not self._flush_config_to_file("on_close"):
            self.log_message("警告：关闭程序前保存配置失败。", is_error=True)
        else:
            self.log_message("所有配置已在关闭前成功保存。")

        if a0:
            a0.accept()

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

        has_ocr_then_grade = False
        try:
            enabled_questions = self.config_manager.get_enabled_questions()
            for q_idx in enabled_questions:
                work_mode_combo = self.get_ui_element(f'work_mode_{q_idx}', QComboBox)
                if work_mode_combo and isinstance(work_mode_combo, QComboBox):
                    mode_value = work_mode_combo.currentData()
                    if not mode_value:
                        normalized = self._normalize_work_mode_ui_text(work_mode_combo.currentText())
                        fallback_map = {
                            '识图直评': 'direct_grade',
                            '直评+推理': 'direct_grade_thinking',
                            '识评分离': 'ocr_then_grade',
                            '分离+推理': 'ocr_then_grade_thinking',
                            '分离+双推理': 'ocr_then_grade_dual_thinking'
                        }
                        mode_value = fallback_map.get(normalized, 'direct_grade')
                    if mode_value in {'ocr_then_grade', 'ocr_then_grade_thinking', 'ocr_then_grade_dual_thinking'}:
                        has_ocr_then_grade = True
                        break
                else:
                    q_cfg = self.config_manager.get_question_config(q_idx)
                    if q_cfg.get('work_mode') in {'ocr_then_grade', 'ocr_then_grade_thinking', 'ocr_then_grade_dual_thinking'}:
                        has_ocr_then_grade = True
                        break
        except Exception:
            has_ocr_then_grade = False

        dual_eval_checkbox = self.get_ui_element('dual_evaluation_enabled')
        unattended_checkbox = self.get_ui_element('unattended_mode_enabled')
        
        # 获取双评和无人模式的当前状态
        is_dual_enabled = dual_eval_checkbox and dual_eval_checkbox.isChecked() if dual_eval_checkbox else False
        is_unattended_enabled = unattended_checkbox and unattended_checkbox.isChecked() if unattended_checkbox else False
        
        # ===========================================================
        # 双评与无人模式互斥逻辑
        # ===========================================================
        if dual_eval_checkbox:
            # 双评允许条件：单题模式 + 非识评分离 + 无人模式未开启
            dual_allowed = is_single_q1_mode and not has_ocr_then_grade and not is_unattended_enabled
            dual_eval_checkbox.setEnabled(dual_allowed)
            if (not dual_allowed) and dual_eval_checkbox.isChecked():
                dual_eval_checkbox.blockSignals(True)
                dual_eval_checkbox.setChecked(False)
                self.handle_checkBox_save('dual_evaluation_enabled', False)
                dual_eval_checkbox.blockSignals(False)
                is_dual_enabled = False
                if is_unattended_enabled:
                    self.log_message("无人模式与双评互斥，已自动关闭双评", False, "INFO")
            
            is_dual_active = dual_eval_checkbox.isChecked() and dual_eval_checkbox.isEnabled()
            self._safe_set_enabled('score_diff_threshold', is_dual_active)
        
        if unattended_checkbox:
            # 无人模式允许条件：双评未开启
            unattended_allowed = not is_dual_enabled
            unattended_checkbox.setEnabled(unattended_allowed)
            if (not unattended_allowed) and unattended_checkbox.isChecked():
                unattended_checkbox.blockSignals(True)
                unattended_checkbox.setChecked(False)
                self.handle_checkBox_save('unattended_mode_enabled', False)
                unattended_checkbox.blockSignals(False)
                self.log_message("双评与无人模式互斥，已自动关闭无人模式", False, "INFO")
        
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

    def _sync_work_mode_from_ui(self) -> None:
        """将UI中的工作模式同步回配置内存，确保落盘一致。"""
        try:
            mode_map = {
                '识图直评': 'direct_grade',
                '直评+推理': 'direct_grade_thinking',
                '识评分离': 'ocr_then_grade',
                '分离+推理': 'ocr_then_grade_thinking',
                '分离+双推理': 'ocr_then_grade_dual_thinking'
            }
            for i in range(1, self.max_questions + 1):
                work_mode_combo = self.get_ui_element(f'work_mode_{i}', QComboBox)
                if work_mode_combo and isinstance(work_mode_combo, QComboBox):
                    work_mode = work_mode_combo.currentData()
                    if not work_mode:
                        ui_text = self._normalize_work_mode_ui_text(work_mode_combo.currentText())
                        work_mode = mode_map.get(ui_text, 'direct_grade')
                    field_name = f'question_{i}_work_mode'
                    self.config_manager.update_config_in_memory(field_name, work_mode)
        except Exception:
            pass

    def _normalize_work_mode_ui_text(self, ui_text: str) -> str:
        """将工作模式下拉框显示文本标准化为核心关键字。"""
        text = str(ui_text).strip() if ui_text is not None else ""
        if text and text[0] in "一二三四五":
            text = text[1:].strip()
        text = text.replace(" ", "")
        return text
    
    def on_question_enabled_changed(self, state):
        if self._is_initializing: return
        sender = self.sender()
        if not sender: return
        try:
            q_index = int(sender.objectName().replace('enableQuestion', ''))
            self.handle_checkBox_save(f"question_{q_index}_enabled", bool(state))
            self._apply_ui_constraints()
            self._mark_config_dirty(f"question_enabled:{q_index}")
        except (ValueError, AttributeError): pass
        
    def update_config_button(self, question_index, is_enabled):
        btn = self.get_ui_element(f'configQuestion{question_index}')
        if btn: btn.setEnabled(is_enabled)
        # 同时控制评分细则输入框和步长选择框
        std_answer = self.get_ui_element(f'StandardAnswer_text_{question_index}')
        if std_answer: std_answer.setEnabled(is_enabled)
        step_combo = self.get_ui_element(f'score_rounding_step_{question_index}')
        if step_combo: step_combo.setEnabled(is_enabled)
        work_mode_combo = self.get_ui_element(f'work_mode_{question_index}')
        if work_mode_combo: work_mode_combo.setEnabled(is_enabled)
    
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
        # 兼容：worker 发来的第二个参数在多数情况下表示“重要/需要展示”
        is_important = bool(is_error)

        # 自动确定级别（向后兼容）
        if level is None:
            level = self.LOG_LEVEL_ERROR if is_error else self.LOG_LEVEL_INFO

        # 日志过滤：始终显示 RESULT/ERROR；INFO/WARNING 仅显示重要消息；DETAIL/DEBUG 不显示
        level_upper = str(level).upper()
        if level_upper in ["DETAIL", "DEBUG"]:
            return
        if level_upper not in ["ERROR", "RESULT"] and not is_important:
            return

        # 统一做去噪
        message = self._normalize_log_text(str(message), preserve_newlines=(level_upper == "RESULT"))
        if not message:
            return

        log_widget = self.get_ui_element('log_text')
        if log_widget:
            if level_upper == "ERROR":
                color = "red"
                prefix = "[错误]"
            elif level_upper == "RESULT":
                color = "black"
                # RESULT默认标题
                prefix = "【AI评分依据】"

                # 兼容旧格式：如果消息以"AI评分依据:"开头，去掉这个前缀
                if message.startswith("AI评分依据:"):
                    message = message[len("AI评分依据:"):].strip()

                # 新格式：如果第一行是【总分 xx 分 - AI评分依据如下】，则将其作为标题
                # 其余行作为正文，避免 UI 出现重复标题块。
                first_line, sep, rest = message.partition("\n")
                if first_line.strip().startswith("【总分") and first_line.strip().endswith("】"):
                    prefix = first_line.strip()
                    message = rest.strip() if sep else ""
            else:
                color = "blue"
                prefix = "[信息]" if level_upper == "INFO" else "[提示]"

            # 处理消息内容：HTML转义 + 规范换行
            formatted_message = self._escape_html(message)
            formatted_message = formatted_message.replace("\r\n", "\n").replace("\r", "\n")
            formatted_message = formatted_message.replace("\n", "<br>")
            formatted_message = formatted_message.replace("；", "；<br>")
            
            # AI评分依据另起一行显示，增加空行提高视觉舒适度
            log_widget.append(f'<span style="color:{color}; font-size:14pt;">{prefix}<br>{formatted_message}</span><br>')

        # 控制台始终输出所有消息
        print(f"[{level_upper}] {message}")

    def on_worker_finished(self):
        self.update_ui_state(is_running=False)
    
    def on_worker_error(self, error_message):
        summary, detail = self._simplify_message_for_teacher(str(error_message))
        if detail and detail != summary:
            self._write_debug_log("worker_error", detail)

        # 用户手动停止：用信息级别，不走错误模板
        if "已停止（用户手动停止" in summary:
            self.log_message(summary, True, "INFO")
        else:
            self.log_message(summary, True, "ERROR")

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
            # 重要信息：让用户确认“确实停了”
            self.log_message("已停止（用户手动停止）。", True, "INFO")
        else:
            self.update_ui_state(is_running=False)

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
            if self._flush_config_to_file("question_dialog_save"):
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
    
    def connect_signals(self):
        """连接所有UI信号的公开接口"""
        self._connect_signals()

    def setup_question_selector(self):
        pass  # UI文件已自动连接

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
            self.log_message("无人模式已启用：AI失败时自动给保守分并标记待复核（三步打分模式下每步各给1分，共3分）")
        else:
            self.log_message("无人模式已禁用：遇到问题时将立即停止并等待人工处理")

    def _connect_signals(self):
        """统一连接所有UI控件的信号与槽"""
        # 连接按钮点击
        auto_btn = self.get_ui_element('auto_run_but')
        if auto_btn and isinstance(auto_btn, QPushButton):
            auto_btn.clicked.connect(self.auto_run_but_clicked)
        
        stop_btn = self.get_ui_element('stop_but')
        if stop_btn and isinstance(stop_btn, QPushButton):
            stop_btn.setToolTip("中止快捷键 Ctrl+Alt+Shift+Z")
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
            self._mark_config_dirty(f"step:{question_index}")
        except (ValueError, TypeError):
            pass  # 忽略无效的步长值