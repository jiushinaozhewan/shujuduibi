<div align="center">

# 📊 跨表核对

**一个直观、可视化的 Excel 跨表查询与核对工具**

*把 VLOOKUP / 差异对比 / 分组聚合 这些枯燥活，变成三下五除二的事*

[![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?logo=python&logoColor=white)](https://www.python.org/)
[![PySide6](https://img.shields.io/badge/Desktop-PySide6-41CD52?logo=qt&logoColor=white)](https://www.qt.io/)
[![Streamlit](https://img.shields.io/badge/Web-Streamlit-FF4B4B?logo=streamlit&logoColor=white)](https://streamlit.io/)
[![License](https://img.shields.io/badge/License-MIT-yellow.svg)](#-许可证)

**[⬇ 直接下载](#-下载与使用) · [✨ 功能亮点](#-功能亮点) · [📖 使用说明](#-使用说明) · [🛠 开发者](#-开发者指南)**

</div>

---

## 💡 是什么

当你手里有好几张 Excel 表，需要：

- 把若干张表里的某个字段 **对齐核对**（VLOOKUP 的升级版）
- 把一张表里的记录 **按某个键分组聚合**（求和 / 取最新 / 去重拼接……）
- 按一条公式逐行核对（如 `合计 = 缴费 - 退费`）

这些事放进 Excel 里做，公式一长就难以维护；写代码又要开发环境。**本工具用可视化界面把这些操作做成了三个按钮。**

## 📦 两个版本

<table>
<tr>
<th width="50%">🖥 桌面版（推荐）</th>
<th width="50%">🌐 Web 版</th>
</tr>
<tr>
<td valign="top">

**`跨表核对.exe`** · PySide6 原生窗口

- ⚡ 秒开（启动 ~1s）
- 🖼 更紧凑的原生 Qt 界面
- 📦 解压 225MB / zip 99MB
- 💼 **适合日常使用，推荐**

</td>
<td valign="top">

**`跨表核对Web.exe`** · Streamlit 浏览器端

- 🌐 浏览器内操作
- 🔀 可改端口多实例运行
- 📦 解压 369MB / zip 138MB
- 🧪 适合喜欢 Web 界面的用户

</td>
</tr>
</table>

> **两个版本功能完全一致，UI 风格不同。** 都已独立打包为 Windows exe，**无需安装 Python 或任何依赖**。

## ⬇ 下载与使用

到 **[📦 Releases 页面](https://github.com/jiushinaozhewan/shujuduibi/releases/latest)** 下载最新版：

| 下载 | 大小 | 适用 | 操作 |
|---|---|---|---|
| **[kuabiao-desktop-v1.0.0.zip](https://github.com/jiushinaozhewan/shujuduibi/releases/download/v1.0.0/kuabiao-desktop-v1.0.0.zip)** | 98 MB | 🖥 桌面版 **（推荐）** | 解压 → 双击 `跨表核对.exe` |
| [kuabiao-web-v1.0.0.zip](https://github.com/jiushinaozhewan/shujuduibi/releases/download/v1.0.0/kuabiao-web-v1.0.0.zip) | 137 MB | 🌐 Web 版 | 解压 → 双击 `跨表核对Web.exe` → 浏览器自动打开 |

> ⚠️ `_internal` 文件夹必须与 exe 保持同目录，不能单独移动 exe。

## ✨ 功能亮点

### ① 自定义分组聚合

选任意 Excel 文件 → 选一列作主索引 → 为其他每列单独选聚合方式。

支持的聚合方式（**每个选项都带详细中文解释**）：

| 方式 | 说明 |
|---|---|
| `sum` | 同组所有数值相加（如金额合计） |
| `max` | 同组取最大；日期列即取最新日期 |
| `min` | 同组取最小；日期列即取最早日期 |
| `first` / `last` | 首条 / 末条 |
| `count` | 同组非空记录条数 |
| `concat去重` | 把同组所有不同文本用逗号串起来 |
| `mean` | 算术平均 |
| `—跳过—` | 该列不参与聚合 |

智能默认：`金额/缴费/退费/合计/总` → sum、`时间/日期` → max、`姓名/部门/班级/性别` → last。

### ② 跨表查询及核对

一个对话框，**两种模式任选**：

<table>
<tr>
<td width="50%" valign="top">

**🔍 核对模式**

比对 A、B 两表『目标数据』是否一致。

输出三张表：
- **汇总** — 一致/不一致/仅A有/仅B有/合计/差额
- **差异明细** — 只列出有问题的行
- **全量对比** — 所有行带状态标签

</td>
<td width="50%" valign="top">

**🔁 查询模式**

从 B 表查到目标数据，**回填到 A 表副本**（不改原文件）。

副本里会额外追加 4 列辅助信息：
- `_原值` — A 原本的值
- `B_来源数据值` — B 表命中的原始数据（多条用 `\|` 分隔）
- `B_来源记录数` — B 表同键命中几条
- `匹配状态` — 已匹配 / 未匹配

</td>
</tr>
</table>

两侧都支持：
- **多 Sheet 合并**（B 侧可勾选多个 sheet 自动拼接）
- **过滤表达式**（pandas query 语法，如 `金额>1000` / `部门=='教职工'`）
- **键列规范化**（自动处理工号前导零、浮点化等常见坑）
- **差额容差**（小于此值视为一致）

### ③ 带运算核对指定列

行内公式核对：`实际值列 = X 列 [+ - × ÷] Y 列`。

逐行计算『应为』值，与实际值比对，超差视为不一致。可指定『携带列』（用于定位差异行，多个列名用英文逗号隔开）。

## 📖 使用说明

### 公共操作

- 左侧（或顶部）**文件管家**：可从本地目录一键加载文件，或上传外部 Excel
- **表头行**可调（不少 Excel 表头不在第 1 行）
- 所有字段（文件/Sheet/过滤/键列/目标数据/聚合方式/容差）均可随时修改后重新执行
- 结果自动预览，点击 `💾 保存 xlsx` 输出文件

### 桌面版快捷操作

- 工具栏 `📁 打开数据目录` — 直接在资源管理器中打开
- 工具栏 `ℹ 使用说明` — 弹窗查看内置帮助

### Web 版特别说明

双击 `跨表核对Web.exe` 后：
1. 弹出黑色命令窗口（保留打开，关闭即退出程序）
2. 自动寻找可用端口（8501 → 8502 → 8503 → 8504）
3. 2 秒后自动用默认浏览器打开

如果 Windows 防火墙弹窗询问，选择「允许访问」（仅本机访问，不开外网）。

## 🛠 开发者指南

### 环境

```bash
pip install pandas openpyxl xlrd PySide6 streamlit pyinstaller
```

### 从源码运行

```bash
# 桌面版
cd ui_desktop
python app.py

# Web 版
cd ui_web
python -m streamlit run app.py
```

### 重新打包

```bash
# 桌面版 → dist/跨表核对/
cd ui_desktop
双击 build_exe.bat      # 或：python -m PyInstaller app.spec --noconfirm

# Web 版 → dist/跨表核对Web/
cd ui_web
双击 build_web_exe.bat  # 或：python -m PyInstaller app_web.spec --noconfirm
```

### 项目结构

```
shujuduibi/
├── scripts/                命令行脚本（固定规则批处理示例）
│   ├── common.py
│   └── task*_*.py
├── ui_desktop/             桌面版 PySide6 应用
│   ├── app.py              主程序（~1300 行）
│   ├── app.spec            PyInstaller 配置
│   ├── build_exe.bat       一键打包
│   ├── run_desktop.bat     一键启动（开发）
│   └── README_打包.md
├── ui_web/                 Web 版 Streamlit 应用
│   ├── app.py              主程序
│   ├── launcher.py         PyInstaller 入口（起服务 + 开浏览器）
│   ├── app_web.spec
│   ├── build_web_exe.bat
│   ├── run_web.bat
│   └── README_打包.md
└── README.md               本文件
```

### 关键技术点

- **聚合层**：`pandas.groupby.agg` 支持多规则映射
- **规范化键**：自定义 `norm_id()` 消除 `'123'` / `'123.0'` / `123` 混用问题
- **过滤器**：`df.query(expr, engine='python')` 接受 pandas query 语法
- **表头可调**：`pd.read_excel(header=N)` 支持非首行表头
- **Streamlit 打包**：通过 `launcher.py` + PyInstaller，修复 `importlib.metadata` 等运行时依赖

## 🔒 隐私

- **所有处理都在本地完成**，不上传任何数据到任何服务器
- Web 版只监听 `localhost`，不对外暴露
- 建议在含敏感数据的 Excel 上使用前，确认 `.gitignore` 生效

## 📝 许可证

MIT License — 可自由使用、修改、分发。

---

<div align="center">

**如果对你有帮助，欢迎点个 ⭐ Star · [🐛 报告问题](https://github.com/jiushinaozhewan/shujuduibi/issues)**

</div>
