# 跨表核对 · Web 版（独立发布包）

## 用户使用（无需装 Python / Streamlit）

1. 解压 `跨表核对Web.zip` 到任意目录
2. 双击 **`跨表核对Web.exe`**
3. 程序会：
   - 弹出黑色命令窗（保留打开，显示运行信息）
   - 自动找到可用端口（默认 8501）
   - 2 秒后自动用默认浏览器打开 `http://localhost:8501`
4. 在浏览器中使用所有功能；关闭黑窗即退出程序

**注意事项：**
- `_internal` 文件夹必须与 `跨表核对Web.exe` 保持同目录，不可单独移动 exe
- 如果 Windows 防火墙弹窗询问，选择『允许访问』（仅本机访问，不开外网）
- 首次启动约 5-10 秒（加载运行时），之后秒开

## 文件目录约定

解压后的目录结构：
```
跨表核对Web/
├── 跨表核对Web.exe        ← 双击启动
└── _internal/              ← 运行时（勿动）
```

用户若把需要处理的 Excel 文件放在 **`跨表核对Web.exe`同目录**，侧栏「文件管家」会自动列出它们。也可以在浏览器内直接上传任意位置的文件。

## 开发者：重新打包

```
cd ui_web
双击 build_web_exe.bat
```
或
```
python -m PyInstaller app_web.spec --noconfirm
```

产物：`dist\跨表核对Web\`（约 369MB，压缩后 138MB）

## 体积说明

Streamlit 自带大量前端/Web 静态资源（CSS/JS/字体），加上 pandas/numpy/pyarrow/altair，
总大小约 370MB，是 PySide6 桌面版（225MB）的 1.6 倍。

如果追求更小体积，请使用桌面版（`ui_desktop\dist\跨表核对.zip`）。
