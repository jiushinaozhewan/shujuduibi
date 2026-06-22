<div align="center">

# 跨表核对

**一个面向 Excel 的本地跨表查询、核对、聚合工具**

把 `VLOOKUP`、差异对比、分组聚合、公式核对这些重复工作，整理成可视化操作。

[![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?logo=python&logoColor=white)](https://www.python.org/)
[![Desktop](https://img.shields.io/badge/Desktop-PySide6-41CD52?logo=qt&logoColor=white)](https://www.qt.io/)
[![Web](https://img.shields.io/badge/Web-Streamlit-FF4B4B?logo=streamlit&logoColor=white)](https://streamlit.io/)
[![License](https://img.shields.io/badge/License-MIT-yellow.svg)](#许可证)

[下载使用](#下载使用) · [功能亮点](#功能亮点) · [使用教程](#使用教程) · [开发者指南](#开发者指南)

</div>

---

## 适合谁

当你经常需要处理这些 Excel 场景时，这个工具会很顺手：

- 以一张表为基表，去多个参考表里查找、回填、核对数据。
- 同一个工号、学号、客户编号出现多行时，先求和、取最大、取最新，再做核对。
- 同时核对多组目标数据，例如 `目标数据1` 对应退费金额，`目标数据2` 对应缴费金额。
- 按一条公式逐行核对，例如 `合计 = 缴费 - 退费`。
- 不想写脚本，也不想维护一堆复杂 Excel 公式。

所有处理都在本机完成。桌面版是原生窗口，Web 版只监听 `localhost`。

## 下载使用

前往 [Releases 最新版](https://github.com/jiushinaozhewan/shujuduibi/releases/latest) 下载发布包。

| 版本 | 发布包 | 启动方式 | 适合场景 |
|---|---|---|---|
| 桌面版（推荐） | `跨表核对.zip` | 解压后双击 `跨表核对.exe` | 日常使用，启动快，原生窗口更紧凑 |
| Web 版 | `跨表核对Web.zip` | 解压后双击 `跨表核对Web.exe` | 喜欢浏览器界面，或需要本机多端口实例 |

> `_internal` 文件夹必须和 exe 保持在同一个目录。不要只移动 exe 文件。

### 启动后怎么做

1. 在文件管家中加载本地 Excel，或直接上传 Excel。
2. 选择要执行的任务：分组聚合、跨表查询及核对、公式核对。
3. 选择 Sheet、表头行、关联字段、目标数据和聚合方式。
4. 点击执行，预览结果。
5. 保存或下载生成的 xlsx 结果文件。

## 功能亮点

### 1. 自定义分组聚合

选择任意 Excel 文件，指定主索引列，再给其他列分别设置聚合方式。

| 聚合方式 | 用途 |
|---|---|
| `sum` | 同组数值求和，例如金额合计 |
| `max` / `min` | 同组取最大/最小，日期列可用于取最新/最早 |
| `first` / `last` | 同组取首条/末条 |
| `count` | 统计同组非空记录数 |
| `concat去重` | 同组文本去重后拼接 |
| `mean` | 同组数值取平均 |
| `—跳过—` | 当前列不参与输出 |

程序会根据列名和数据类型推荐默认聚合方式。比如金额类列默认 `sum`，日期类列默认 `max`。

### 2. 跨表查询及核对

这是本项目的核心功能。新版支持 **一个基表 A 表 + 多个参考表 B/C/D... + 多组目标数据**。

#### 核对模式

用于检查 A 表与参考表的数据是否一致。

- A 表作为基表。
- B/C/D... 表作为参考表，可继续点击 `+` 添加。
- 每个参考表右上角可点击 `-` 删除。
- A 表可以添加 `目标数据1`、`目标数据2`、`目标数据3`。
- 参考表的目标数据编号以 A 表为准。
- 同一个目标数据编号只能被一个参考表占用，避免对应关系混乱。

核对结果包含：

- `汇总`：按参考表、目标数据编号统计一致、不一致、仅 A 有、仅参考表有、合计和差额。
- `差异明细`：只列出有问题的记录。
- `全量对比`：所有键值的完整核对结果。

#### 查询模式

用于从参考表查到数据后，回填到 A 表副本。

- 不修改原始 A 表文件。
- 可从多个参考表查询不同目标数据。
- 命中后回填到 A 表对应目标列。
- 保留 A 表原值，便于复核。

查询结果会追加这些辅助列：

- `*_原值`：A 表回填前的原值。
- `B/C/D_来源数据值`：参考表命中的原始数据，多条用 `|` 分隔。
- `B/C/D_来源记录数`：各参考表同键命中的记录数。
- `匹配参考表`：哪些参考表命中了当前行。
- `匹配状态`：已匹配 / 未匹配。

#### 通用能力

- A 表支持单 Sheet。
- 参考表支持多 Sheet 合并。
- 每张表都支持过滤表达式，例如 `金额 > 1000`、`部门 == '教职工'`。
- 键列可按字符串规范化，减少 `123`、`123.0`、`00123` 这类 Excel 常见问题。
- 同键多行可按 `sum`、`first`、`last`、`max`、`min`、`mean` 聚合后再查询或核对。
- 可设置差额容差，避免小数精度造成误判。

### 3. 带运算核对指定列

逐行验证一个公式关系：

```text
实际值列 = X 列 [+ - × ÷] Y 列
```

适合核对：

- `合计 = 缴费 - 退费`
- `金额 = 单价 × 数量`
- `单价 = 总额 ÷ 数量`

你可以指定携带列，例如部门、工号、姓名，让差异结果更容易定位。

## 使用教程

### 桌面版

- `打开数据目录`：在资源管理器中打开默认数据目录。
- `使用说明`：查看程序内置帮助。
- 结果区域会自动预览汇总、差异明细、全量对比或查询结果。
- 点击 `保存 xlsx` 输出结果文件。

### Web 版

双击 `跨表核对Web.exe` 后，程序会：

1. 打开一个本地命令窗口。
2. 自动寻找可用端口，默认从 `8501` 开始。
3. 用默认浏览器打开 `http://localhost:<端口>`。

如果 Windows 防火墙弹窗询问，选择允许访问即可。Web 版只服务本机 `localhost`。

## 开发者指南

### 安装依赖

```bash
python -m pip install pandas openpyxl xlrd PySide6 streamlit pyinstaller
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
# 桌面版
cd ui_desktop
python -m PyInstaller app.spec --noconfirm

# Web 版
cd ui_web
python -m PyInstaller app_web.spec --noconfirm
```

本地发布构建可输出到 `dist_release/`，用于整理发布包。发布包体积较大，建议上传到 GitHub Releases，不建议直接提交到 git 仓库。

### 项目结构

```text
shujuduibi/
├── scripts/                命令行批处理脚本示例
├── ui_desktop/             桌面版 PySide6 应用
│   ├── app.py              桌面版主程序
│   ├── app.spec            PyInstaller 配置
│   ├── build_exe.bat       一键打包脚本
│   └── README_打包.md
├── ui_web/                 Web 版 Streamlit 应用
│   ├── app.py              Web 版主程序
│   ├── launcher.py         独立 exe 启动器
│   ├── app_web.spec        PyInstaller 配置
│   ├── build_web_exe.bat   一键打包脚本
│   └── README_打包.md
├── RELEASE_NOTES.md        发布说明
└── README.md               项目主页
```

### 关键实现

- `pandas.groupby.agg`：负责同键多行聚合。
- `norm_id()`：把工号、学号等键值规范成字符串。
- `df.query(..., engine="python")`：提供过滤表达式能力。
- `pd.read_excel(header=N)`：支持非首行表头。
- `PySide6`：桌面版窗口和对话框。
- `Streamlit + launcher.py`：Web 版本机服务和浏览器启动。

## 隐私与数据安全

- 所有 Excel 处理均在本机完成。
- Web 版只监听 `localhost`，不会主动对外提供服务。
- `.gitignore` 已排除常见 Excel 数据文件和结果文件，避免误提交真实数据。
- 处理敏感数据时，请不要把原始 Excel 放进 git 提交。

## 许可证

MIT License，可自由使用、修改、分发。

---

<div align="center">

如果这个工具帮你少写了几个小时 Excel 公式，欢迎点个 Star。

[报告问题](https://github.com/jiushinaozhewan/shujuduibi/issues) · [查看发布包](https://github.com/jiushinaozhewan/shujuduibi/releases/latest)

</div>
