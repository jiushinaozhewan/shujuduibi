# -*- mode: python ; coding: utf-8 -*-
"""打包 Streamlit Web 版为独立 exe
命令：pyinstaller app_web.spec --noconfirm
产物：dist/跨表核对Web/
"""
from PyInstaller.utils.hooks import collect_data_files, collect_submodules, copy_metadata

hidden = (
    collect_submodules("streamlit")
    + collect_submodules("pandas")
    + collect_submodules("openpyxl")
    + collect_submodules("xlrd")
    + collect_submodules("altair")
    + ["pyarrow", "importlib_metadata"]
)

datas = []
datas += collect_data_files("streamlit")
datas += collect_data_files("altair")
datas += collect_data_files("openpyxl")
# Streamlit 运行依赖包的 metadata
for pkg in [
    "streamlit", "altair", "numpy", "pandas", "pyarrow", "tornado",
    "gitpython", "pydeck", "narwhals", "rich", "click", "packaging",
    "protobuf", "tenacity", "toml", "watchdog", "blinker", "cachetools",
    "jsonschema", "jinja2", "requests", "typing_extensions",
]:
    try:
        datas += copy_metadata(pkg)
    except Exception:
        pass
# 把 app.py 放进包内（启动器会从 _MEIPASS 读取）
datas += [("app.py", ".")]

a = Analysis(
    ["launcher.py"],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hidden,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "tkinter", "matplotlib", "scipy", "IPython", "jupyter",
        "PyQt5", "PyQt6", "PySide2", "PySide6",
        "notebook", "sphinx", "pytest", "pydoc_data",
        "torch", "torchvision", "cv2",
        "sklearn", "sympy", "tensorflow",
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
    name="跨表核对Web",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,      # 保留控制台，显示启动信息/关闭即退出
    disable_windowed_traceback=False,
    argv_emulation=False,
    icon=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="跨表核对Web",
)
