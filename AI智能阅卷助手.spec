# Future-proofing instructions for building this application.
#
# To build the executable, follow these steps in your terminal:
#
# 1. (Recommended) Clean up previous builds to avoid conflicts:
#    rm -r -fo dist build
#
# 2. Run PyInstaller with this spec file:
#    pyinstaller AI智能阅卷助手.spec
#
# The final single-file executable will be located in the 'dist' folder.

# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_submodules, collect_data_files

# 收集可能遗漏的子模块（仅在确有隐式导入时才需要），和 PyQt5 的 Qt 插件
hiddenimports = collect_submodules('numpy') + collect_submodules('cv2')

# 包含 PyQt5 的 Qt 平台插件，避免运行时 "could not load the Qt platform plugin" 错误
# 并保留程序的 setting 目录作为数据文件
datas = collect_data_files('PyQt5', subdir='Qt/plugins') + [('setting', 'setting')]

a = Analysis(
    ['main.py'],
    pathex=['.'],  # 项目根路径，帮助 PyInstaller 定位模块和资源
    binaries=[],
    datas=datas,
    hiddenimports=['PyQt5.sip', 'PyQt5.QtCore', 'PyQt5.QtGui', 'PyQt5.QtWidgets', 'api_service', 'auto_thread', 'config_manager', 'ui_components.main_window', 'ui_components.question_config_dialog', 'pyautogui', 'PIL', 'PIL.ImageGrab', 'PIL.Image', 'PIL.ImageDraw', 'appdirs', 'requests', 'winsound', 'pandas', 'openpyxl'] + hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['tkinter', 'unittest', 'doctest'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='AI智能阅卷助手',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    # 注意: 启用 UPX 可以减小可执行文件体积，但在某些环境下可能导致兼容性或杀软误报问题。
    # 如果遇到构建或运行异常，可以尝试将 `upx=False`。
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='AI阅卷助手.ico',
    onefile=True,
)
