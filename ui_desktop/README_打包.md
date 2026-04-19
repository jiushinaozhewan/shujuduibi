# 跨表核对 - 独立桌面版

## 用户使用（无需安装 Python）

1. 下载 `跨表核对.zip` 并解压到任意文件夹
2. 双击文件夹内的 **`跨表核对.exe`** 即可启动
3. 在应用中通过『浏览』按钮选择本地 Excel 文件开始工作

首次启动需几秒（解压运行时），之后很快。

**重要：** `_internal` 文件夹必须与 `跨表核对.exe` 保持在同一目录，不能单独移动 exe。

## 功能

- **① 自定义分组聚合** — 任意 Excel 按键列分组，每列独立选聚合方式
- **② 跨表查询及核对** — 两表按关联字段比对（核对模式）或回填数据（查询模式）
- **③ 带运算核对指定列** — 行内公式核对：某列 = X [+-×÷] Y

## 开发者：重新打包

依赖：
```
pip install pyinstaller PySide6 pandas openpyxl xlrd
```

打包：
```
cd ui_desktop
双击 build_exe.bat    或    python -m PyInstaller app.spec --noconfirm
```

生成产物：`dist\跨表核对\`（约 225MB）

## 体积优化

`app.spec` 中已排除 torch、cv2、pyarrow、scipy、streamlit 等大型依赖，
如发现仍被拉入其他重型包，在 `excludes=[...]` 中追加即可。

如需生成单个 exe（启动更慢，约 +5 秒解压）：
修改 `app.spec` 将 `exclude_binaries=True` 改为 `False`，
删除 `COLLECT(...)` 段落，重新运行打包。
