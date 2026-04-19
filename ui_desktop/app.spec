# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller 打包配置
打包命令（在 ui_desktop 目录下执行）：
    pyinstaller app.spec --noconfirm
生成目录：dist/跨表核对/   内含 跨表核对.exe 及依赖
"""
from PyInstaller.utils.hooks import collect_submodules, collect_data_files

# 显式收集 pandas/openpyxl/xlrd 的子模块，避免运行时 ImportError
hidden = (
    collect_submodules("pandas")
    + collect_submodules("openpyxl")
    + collect_submodules("xlrd")
    + ["_cffi_backend"]
)

# 收集 openpyxl 的样式 XML 等数据
datas = collect_data_files("openpyxl")

a = Analysis(
    ["app.py"],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hidden,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # 剔除不必要的大包（显著减小体积）
    excludes=[
        "tkinter", "matplotlib", "scipy", "IPython", "jupyter",
        "PyQt5", "PyQt6", "PySide2", "notebook", "sphinx",
        "pytest", "pydoc_data", "test", "unittest",
        # 重型依赖（非本应用所需，但会被 pandas/streamlit 间接拉入）
        "torch", "torchvision", "torchaudio",
        "cv2", "opencv-python",
        "pyarrow", "numba", "numba_special",
        "sklearn", "sympy", "tensorflow", "keras",
        "streamlit", "altair", "pydeck", "narwhals",
        "plotly", "bokeh", "seaborn",
        "PIL.ImageQt", "PIL.ImageTk",
    ],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="跨表核对",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,                 # 若机器上有 upx 会自动压缩
    console=False,            # GUI 应用，不显示命令行窗口
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="跨表核对",
)
