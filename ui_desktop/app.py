"""桌面版：7 个固定任务一键执行
- 左侧：任务按钮 + 全部执行
- 右上：实时日志
- 右下：当前任务"汇总"表预览
- 工具栏：打开数据目录 / 打开结果文件
"""
from __future__ import annotations
import sys, io, traceback, subprocess, os
from pathlib import Path
from contextlib import redirect_stdout

import pandas as pd
from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QAction, QFont, QColor
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QHBoxLayout, QVBoxLayout, QPushButton,
    QPlainTextEdit, QTableView, QSplitter, QLabel, QFileDialog, QMessageBox,
    QToolBar, QStatusBar, QHeaderView, QAbstractItemView,
    QDialog, QComboBox, QSpinBox, QLineEdit, QGroupBox, QScrollArea, QFormLayout,
    QDialogButtonBox, QGridLayout,
)
from PySide6.QtCore import QAbstractTableModel, QModelIndex

# 版本标识（显示在所有窗口标题栏靠右侧靠近操作按钮处）
VERSION_TAG = "女神专享版"
def _t(title: str) -> str:
    """给窗口标题追加版本标识，用大量空格把标识推到右侧靠近 [−][□][×] 三个按钮"""
    return f"{title}                                                                    ✦  {VERSION_TAG}  ✦"

if getattr(sys, "frozen", False):
    # PyInstaller 打包后运行：数据目录 = exe 所在目录
    ROOT = Path(sys.executable).resolve().parent
    SCRIPTS = ROOT  # 打包后不再依赖 scripts/
else:
    # 开发环境
    ROOT = Path(__file__).resolve().parent.parent  # D:/meta/shujuduibi
    SCRIPTS = ROOT / "scripts"
    if SCRIPTS.exists():
        sys.path.insert(0, str(SCRIPTS))

# 任务定义：(显示名, 预设key或None, 输出文件, 交互类型)
# 交互类型：'dialog_aggregate' | 'dialog_check' | 'dialog_formula'
TASKS = [
    ("① 自定义分组聚合（选文件/键/规则）", None, "aggregate_custom.xlsx", "dialog_aggregate"),
    ("② 跨表查询及核对", "t2", "jieguo1.xlsx", "dialog_check"),
    ("③ 带运算核对指定列", "t4", "jieguo3.xlsx", "dialog_formula"),
]


class PandasModel(QAbstractTableModel):
    def __init__(self, df: pd.DataFrame):
        super().__init__()
        self._df = df.reset_index(drop=True)

    def rowCount(self, parent=QModelIndex()):
        return len(self._df)

    def columnCount(self, parent=QModelIndex()):
        return len(self._df.columns)

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None
        v = self._df.iat[index.row(), index.column()]
        if role == Qt.DisplayRole:
            if pd.isna(v):
                return ""
            if isinstance(v, float):
                return f"{v:,.2f}" if v != int(v) else f"{int(v):,}"
            return str(v)
        if role == Qt.TextAlignmentRole:
            if isinstance(v, (int, float)):
                return int(Qt.AlignRight | Qt.AlignVCenter)
            return int(Qt.AlignLeft | Qt.AlignVCenter)
        return None

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role != Qt.DisplayRole:
            return None
        if orientation == Qt.Horizontal:
            return str(self._df.columns[section])
        return str(section + 1)


class TaskRunner(QThread):
    log_signal = Signal(str)
    done_signal = Signal(bool, str, str)  # ok, output_file, error_msg

    def __init__(self, module_name: str, output_file: str):
        super().__init__()
        self.module_name = module_name
        self.output_file = output_file

    def run(self):
        buf = io.StringIO()
        ok = True
        err = ""
        try:
            # 重新导入以保证脚本内运行
            import importlib
            with redirect_stdout(buf):
                if self.module_name in sys.modules:
                    importlib.reload(sys.modules[self.module_name])
                else:
                    importlib.import_module(self.module_name)
        except Exception:
            ok = False
            err = traceback.format_exc()
        self.log_signal.emit(buf.getvalue())
        if not ok:
            self.log_signal.emit(f"\n[ERROR]\n{err}")
        self.done_signal.emit(ok, self.output_file, err)


# ================ 通用工具 ================
def _norm_id(v):
    if pd.isna(v):
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def _to_num(v):
    if pd.isna(v):
        return 0.0
    if isinstance(v, str):
        s = v.strip().replace(",", "")
        if not s:
            return 0.0
        try:
            return float(s)
        except ValueError:
            return 0.0
    return float(v)


def _save_xlsx(path: Path, sheets: dict):
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        for name, df in sheets.items():
            df.to_excel(w, sheet_name=name[:31], index=False)
        from openpyxl.utils import get_column_letter
        for name, df in sheets.items():
            ws = w.sheets[name[:31]]
            ws.freeze_panes = "A2"
            for i, col in enumerate(df.columns, 1):
                try:
                    mx = max([len(str(col))] + [len(str(x)) for x in df[col].head(200).tolist()])
                except Exception:
                    mx = 12
                ws.column_dimensions[get_column_letter(i)].width = min(max(mx + 2, 8), 40)


AGG_OPTS = [
    "—跳过— (该列不参与聚合，不出现在结果中)",
    "sum — 求和 (同组所有数值相加，如金额合计)",
    "max — 最大值 (同组取最大；日期列即取最新日期)",
    "min — 最小值 (同组取最小；日期列即取最早日期)",
    "first — 首条 (同组第一次出现的那一行的值)",
    "last — 末条 (同组最后一次出现的那一行的值)",
    "count — 计数 (同组非空记录的条数)",
    "concat去重 — 拼接去重 (把同组所有不同文本用逗号串起来)",
    "mean — 平均值 (同组所有数值的算术平均)",
]


def _agg_key(opt_text: str) -> str:
    """从带说明的下拉文字中取出关键字：'sum — 求和 (...)' → 'sum'"""
    s = opt_text.strip()
    if s.startswith("—跳过—"):
        return "—跳过—"
    return s.split(" ")[0]


class AggregateDialog(QDialog):
    """自定义分组聚合对话框：选文件→选sheet/表头→选主键→为每列选规则→执行"""

    def __init__(self, parent=None, default_dir: Path = None, default_out: Path = None):
        super().__init__(parent)
        self.setWindowTitle(_t("自定义分组聚合"))
        self.resize(1000, 720)
        self.default_dir = default_dir or Path.home()
        self.default_out = default_out or (self.default_dir / "aggregate_custom.xlsx")
        self._df: pd.DataFrame | None = None
        self._file_path: Path | None = None
        self._result: pd.DataFrame | None = None

        root = QVBoxLayout(self)

        # Step 1：文件 + sheet + 表头行
        gb1 = QGroupBox("① 选择文件")
        g1 = QGridLayout(gb1)
        self.ed_file = QLineEdit()
        self.ed_file.setPlaceholderText("点击右侧『浏览』选择 xlsx/xls 文件")
        self.ed_file.setReadOnly(True)
        btn_browse = QPushButton("浏览…")
        btn_browse.clicked.connect(self._pick_file)
        self.cb_sheet = QComboBox()
        self.cb_sheet.setEnabled(False)
        self.sp_header = QSpinBox()
        self.sp_header.setRange(0, 50)
        self.sp_header.setValue(0)
        self.sp_header.setPrefix("表头在第 ")
        self.sp_header.setSuffix(" 行 (0=第1行)")
        btn_load = QPushButton("加载 ▶")
        btn_load.clicked.connect(self._load)
        g1.addWidget(QLabel("文件:"), 0, 0)
        g1.addWidget(self.ed_file, 0, 1)
        g1.addWidget(btn_browse, 0, 2)
        g1.addWidget(QLabel("Sheet:"), 1, 0)
        g1.addWidget(self.cb_sheet, 1, 1)
        g1.addWidget(self.sp_header, 1, 2)
        g1.addWidget(btn_load, 1, 3)
        root.addWidget(gb1)

        # Step 2：主键 + 每列规则
        gb2 = QGroupBox("② 聚合规则（加载后填写）")
        v2 = QVBoxLayout(gb2)
        top = QHBoxLayout()
        top.addWidget(QLabel("主索引列:"))
        self.cb_key = QComboBox()
        self.cb_key.setMinimumWidth(220)
        self.cb_key.currentIndexChanged.connect(self._rebuild_rules)
        top.addWidget(self.cb_key)
        top.addStretch(1)
        top.addWidget(QLabel("输出文件:"))
        self.ed_out = QLineEdit(str(self.default_out))
        btn_out = QPushButton("…")
        btn_out.setFixedWidth(30)
        btn_out.clicked.connect(self._pick_out)
        top.addWidget(self.ed_out, 2)
        top.addWidget(btn_out)
        v2.addLayout(top)
        hint = QLabel("提示：每列下拉框内包含聚合方式的详细说明；默认按列名/类型智能推荐，可随意修改")
        hint.setStyleSheet("color:#666;")
        v2.addWidget(hint)
        # 规则区（滚动）
        self.rules_scroll = QScrollArea()
        self.rules_scroll.setWidgetResizable(True)
        self.rules_host = QWidget()
        self.rules_form = QFormLayout(self.rules_host)
        self.rules_form.setLabelAlignment(Qt.AlignRight)
        self.rules_scroll.setWidget(self.rules_host)
        self.rules_scroll.setMinimumHeight(220)
        v2.addWidget(self.rules_scroll)
        root.addWidget(gb2, 1)

        # Step 3：结果预览
        gb3 = QGroupBox("③ 结果预览")
        v3 = QVBoxLayout(gb3)
        self.result_info = QLabel("尚未执行")
        v3.addWidget(self.result_info)
        self.preview = QTableView()
        self.preview.setAlternatingRowColors(True)
        self.preview.setEditTriggers(QAbstractItemView.NoEditTriggers)
        v3.addWidget(self.preview)
        root.addWidget(gb3, 1)

        # 按钮
        btnbox = QDialogButtonBox()
        self.btn_run = btnbox.addButton("▶ 执行", QDialogButtonBox.ActionRole)
        self.btn_save = btnbox.addButton("💾 保存 xlsx", QDialogButtonBox.ActionRole)
        self.btn_close = btnbox.addButton(QDialogButtonBox.Close)
        self.btn_run.clicked.connect(self._run)
        self.btn_save.clicked.connect(self._save)
        self.btn_save.setEnabled(False)
        self.btn_close.clicked.connect(self.accept)
        root.addWidget(btnbox)

        self._rule_combos: dict[str, QComboBox] = {}

    def _pick_file(self):
        p, _ = QFileDialog.getOpenFileName(
            self, "选择 Excel", str(self.default_dir), "Excel (*.xlsx *.xls *.XLSX)"
        )
        if not p:
            return
        self.ed_file.setText(p)
        self._file_path = Path(p)
        # 枚举 sheet
        self.cb_sheet.clear()
        try:
            engine = "xlrd" if self._file_path.suffix.lower() == ".xls" else "openpyxl"
            names = pd.ExcelFile(p, engine=engine).sheet_names
            self.cb_sheet.addItems(names)
            self.cb_sheet.setEnabled(True)
        except Exception as e:
            QMessageBox.critical(self, "错误", f"读取 sheet 失败：{e}")

    def _pick_out(self):
        p, _ = QFileDialog.getSaveFileName(
            self, "保存为", self.ed_out.text() or str(self.default_dir),
            "Excel (*.xlsx)"
        )
        if p:
            if not p.lower().endswith(".xlsx"):
                p += ".xlsx"
            self.ed_out.setText(p)

    def _load(self):
        if not self._file_path or not self.cb_sheet.currentText():
            QMessageBox.warning(self, "提示", "请先选择文件和 sheet")
            return
        try:
            engine = "xlrd" if self._file_path.suffix.lower() == ".xls" else "openpyxl"
            df = pd.read_excel(
                self._file_path, sheet_name=self.cb_sheet.currentText(),
                header=self.sp_header.value(), engine=engine,
            )
            df.columns = [str(c).strip() for c in df.columns]
            self._df = df
            # 填充主键下拉
            self.cb_key.clear()
            self.cb_key.addItems(list(df.columns))
            self._rebuild_rules()
            self.result_info.setText(f"已加载：{len(df)} 行 × {df.shape[1]} 列")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"加载失败：{e}")

    def _rebuild_rules(self):
        # 清空
        while self.rules_form.count():
            item = self.rules_form.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        self._rule_combos.clear()
        if self._df is None:
            return
        key = self.cb_key.currentText()
        # 默认规则推断
        for col in self._df.columns:
            if col == key:
                continue
            cb = QComboBox()
            cb.addItems(AGG_OPTS)
            cb.setToolTip("选择此列如何聚合；鼠标悬停选项可看说明")
            # 启发式默认值（按关键字匹配选项）
            s = self._df[col]
            name = str(col)
            def _find(prefix: str) -> int:
                for i, o in enumerate(AGG_OPTS):
                    if o.startswith(prefix):
                        return i
                return 0
            if pd.api.types.is_numeric_dtype(s) or any(k in name for k in ["金额", "缴费", "退费", "合计", "总", "数量"]):
                cb.setCurrentIndex(_find("sum"))
            elif any(k in name for k in ["时间", "日期"]) or pd.api.types.is_datetime64_any_dtype(s):
                cb.setCurrentIndex(_find("max"))
            elif any(k in name for k in ["姓名", "名称", "部门", "班级", "性别"]):
                cb.setCurrentIndex(_find("last"))
            else:
                cb.setCurrentIndex(0)  # —跳过—
            self._rule_combos[col] = cb
            self.rules_form.addRow(f"{col}:", cb)

    def _run(self):
        if self._df is None:
            QMessageBox.warning(self, "提示", "请先加载文件")
            return
        key = self.cb_key.currentText()
        if not key:
            return
        df = self._df.copy()
        # 去掉键列为空的行
        df = df[df[key].notna()].copy()
        if df.empty:
            QMessageBox.warning(self, "提示", "键列全部为空")
            return
        # 如果键看起来像工号（字符串+数字混合），规范化
        df["__key"] = df[key].map(_norm_id)
        df = df[df["__key"] != ""]

        agg_dict = {}
        for col, cb in self._rule_combos.items():
            mode = _agg_key(cb.currentText())
            if mode == "—跳过—":
                continue
            if mode in ("sum", "mean", "min", "max"):
                # 尝试数值化；max 对时间列尝试转datetime
                if mode == "max":
                    try:
                        dt = pd.to_datetime(df[col], errors="coerce")
                        if dt.notna().sum() > 0:
                            df[col] = dt
                        else:
                            df[col] = df[col].map(_to_num)
                    except Exception:
                        df[col] = df[col].map(_to_num)
                else:
                    df[col] = df[col].map(_to_num)
                agg_dict[col] = mode
            elif mode == "concat去重":
                agg_dict[col] = lambda s: ",".join(sorted({str(x) for x in s.dropna()}))
            else:
                agg_dict[col] = mode  # first / last / count
        if not agg_dict:
            QMessageBox.warning(self, "提示", "请至少为一列选聚合方式")
            return

        try:
            out = df.groupby("__key", as_index=False).agg(agg_dict)
            # 加一列记录笔数
            cnt = df.groupby("__key", as_index=False).size().rename(columns={"size": "记录笔数"})
            out = out.merge(cnt, on="__key")
            out = out.rename(columns={"__key": key})
            # 金额列四舍五入
            for c, m in agg_dict.items():
                if m in ("sum", "mean") and pd.api.types.is_numeric_dtype(out[c]):
                    out[c] = out[c].round(2)
            out = out.sort_values(key).reset_index(drop=True)
        except Exception as e:
            QMessageBox.critical(self, "错误", f"聚合失败：{e}")
            return

        self._result = out
        self.preview.setModel(PandasModel(out))
        self.preview.resizeColumnsToContents()
        self.result_info.setText(
            f"✓ 源 {len(df)} 行 → 聚合后 {len(out)} 组  |  主键：{key}  |  规则：{len(agg_dict)} 列"
        )
        self.btn_save.setEnabled(True)

    def _save(self):
        if self._result is None:
            return
        out_path = Path(self.ed_out.text())
        if not out_path.suffix:
            out_path = out_path.with_suffix(".xlsx")
        try:
            _save_xlsx(out_path, {"聚合结果": self._result})
            QMessageBox.information(self, "完成", f"已保存：{out_path}")
            self.saved_path = out_path
        except Exception as e:
            QMessageBox.critical(self, "错误", f"保存失败：{e}")


# ================ 跨表核对对话框 ================
CHECK_AGG_OPTS = [
    "sum — 求和 (同键多行求和，如退费金额合计)",
    "first — 首条 (同键取第一次出现的那行)",
    "last — 末条 (同键取最后一次出现的那行)",
    "max — 最大值 (同键取最大数值或最新日期)",
    "min — 最小值 (同键取最小)",
    "mean — 平均值 (同键数值平均)",
]


def _parse_agg(text: str) -> str:
    return text.split(" ")[0]


class _SideEditor(QWidget):
    """对话框内的一侧（A 或 B）编辑器：文件+sheet(多选)+表头+过滤+键列+值列+聚合方式"""

    def __init__(self, title: str, default_dir: Path, allow_multi_sheet: bool = False):
        super().__init__()
        self.default_dir = default_dir
        self.allow_multi_sheet = allow_multi_sheet
        self._file_path: Path | None = None
        self._cols: list[str] = []

        g = QGridLayout(self)
        g.setContentsMargins(6, 6, 6, 6)
        t = QLabel(f"【{title}】")
        t.setStyleSheet("font-weight:bold;")
        g.addWidget(t, 0, 0, 1, 4)

        self.ed_file = QLineEdit(); self.ed_file.setReadOnly(True)
        self.ed_file.setPlaceholderText("点击『浏览』选择文件")
        btn_browse = QPushButton("浏览…"); btn_browse.clicked.connect(self._pick_file)
        g.addWidget(QLabel("文件:"), 1, 0); g.addWidget(self.ed_file, 1, 1, 1, 2); g.addWidget(btn_browse, 1, 3)

        self.cb_sheet = QComboBox(); self.cb_sheet.setEnabled(False)
        self.ed_sheets_multi = QLineEdit(); self.ed_sheets_multi.setReadOnly(True)
        self.ed_sheets_multi.setPlaceholderText("多选后显示，如 1,2,3")
        btn_pick_sheets = QPushButton("选 sheet…"); btn_pick_sheets.clicked.connect(self._pick_sheets)
        btn_pick_sheets.setEnabled(False)
        self.btn_pick_sheets = btn_pick_sheets
        self.sp_header = QSpinBox(); self.sp_header.setRange(0, 50); self.sp_header.setValue(0)
        self.sp_header.setPrefix("表头行 "); self.sp_header.setSuffix(" (0=第1行)")
        if allow_multi_sheet:
            g.addWidget(QLabel("Sheets:"), 2, 0)
            g.addWidget(self.ed_sheets_multi, 2, 1)
            g.addWidget(btn_pick_sheets, 2, 2)
            g.addWidget(self.sp_header, 2, 3)
        else:
            g.addWidget(QLabel("Sheet:"), 2, 0)
            g.addWidget(self.cb_sheet, 2, 1, 1, 2)
            g.addWidget(self.sp_header, 2, 3)

        btn_load = QPushButton("加载列 ▶"); btn_load.clicked.connect(self._load)
        self.ed_filter = QLineEdit()
        self.ed_filter.setPlaceholderText("过滤表达式(可选)，如: 退费金额>1000   或   `客户部门`.str.contains('教职工')")
        g.addWidget(QLabel("过滤:"), 3, 0); g.addWidget(self.ed_filter, 3, 1, 1, 2); g.addWidget(btn_load, 3, 3)

        self.cb_key = QComboBox(); self.cb_val = QComboBox()
        self.cb_agg = QComboBox(); self.cb_agg.addItems(CHECK_AGG_OPTS)
        self.cb_agg.setToolTip("当同一键在该表出现多次时如何合并")
        g.addWidget(QLabel("关联相同字段:"), 4, 0); g.addWidget(self.cb_key, 4, 1)
        g.addWidget(QLabel("目标数据:"), 4, 2); g.addWidget(self.cb_val, 4, 3)
        g.addWidget(QLabel("同键聚合:"), 5, 0); g.addWidget(self.cb_agg, 5, 1, 1, 3)

        self.status = QLabel("（未加载）")
        self.status.setStyleSheet("color:#666;")
        g.addWidget(self.status, 6, 0, 1, 4)

        self._all_sheets: list[str] = []
        self._selected_sheets: list[str] = []

    def _pick_file(self):
        p, _ = QFileDialog.getOpenFileName(self, "选择 Excel", str(self.default_dir), "Excel (*.xlsx *.xls *.XLSX)")
        if not p:
            return
        self.ed_file.setText(p)
        self._file_path = Path(p)
        try:
            engine = "xlrd" if self._file_path.suffix.lower() == ".xls" else "openpyxl"
            names = pd.ExcelFile(p, engine=engine).sheet_names
            self._all_sheets = names
            self.cb_sheet.clear(); self.cb_sheet.addItems(names); self.cb_sheet.setEnabled(True)
            self.btn_pick_sheets.setEnabled(True)
            if self.allow_multi_sheet:
                self._selected_sheets = []
                self.ed_sheets_multi.setText("")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"读取 sheet 失败：{e}")

    def _pick_sheets(self):
        if not self._all_sheets:
            return
        # 用一个简易对话框做多选
        dlg = QDialog(self)
        dlg.setWindowTitle(_t("选择多个 Sheet"))
        v = QVBoxLayout(dlg)
        v.addWidget(QLabel("勾选要合并的 sheet（所有勾选的 sheet 会被纵向拼接后再参与核对）："))
        from PySide6.QtWidgets import QCheckBox
        checks = []
        for s in self._all_sheets:
            cb = QCheckBox(s); cb.setChecked(s in self._selected_sheets)
            v.addWidget(cb); checks.append(cb)
        bb = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        bb.accepted.connect(dlg.accept); bb.rejected.connect(dlg.reject)
        v.addWidget(bb)
        if dlg.exec():
            self._selected_sheets = [c.text() for c in checks if c.isChecked()]
            self.ed_sheets_multi.setText(",".join(self._selected_sheets))

    def _load(self):
        if not self._file_path:
            QMessageBox.warning(self, "提示", "请先选择文件")
            return
        try:
            df = self._read()
            self._cols = list(df.columns)
            self.cb_key.clear(); self.cb_key.addItems(self._cols)
            self.cb_val.clear(); self.cb_val.addItems(self._cols)
            self.status.setText(f"✓ 已加载 {len(df)} 行 × {df.shape[1]} 列")
        except Exception as e:
            self.status.setText(f"✗ {e}")
            QMessageBox.critical(self, "错误", f"加载失败：{e}")

    def _read(self) -> pd.DataFrame:
        engine = "xlrd" if self._file_path.suffix.lower() == ".xls" else "openpyxl"
        hdr = self.sp_header.value()
        if self.allow_multi_sheet:
            sheets = self._selected_sheets or ([self._all_sheets[0]] if self._all_sheets else [])
            dfs = []
            for s in sheets:
                d = pd.read_excel(self._file_path, sheet_name=s, header=hdr, engine=engine)
                d.columns = [str(c).strip() for c in d.columns]
                d["__src_sheet"] = s
                dfs.append(d)
            df = pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()
        else:
            sheet = self.cb_sheet.currentText()
            df = pd.read_excel(self._file_path, sheet_name=sheet, header=hdr, engine=engine)
            df.columns = [str(c).strip() for c in df.columns]
        # 应用过滤
        flt = self.ed_filter.text().strip()
        if flt:
            try:
                df = df.query(flt, engine="python")
            except Exception as e:
                raise RuntimeError(f"过滤表达式错误：{e}")
        return df

    # 供外部：取 (df, key列, 值列, 聚合方式)
    def get_data(self) -> tuple[pd.DataFrame, str, str, str]:
        df = self._read()
        return df, self.cb_key.currentText(), self.cb_val.currentText(), _parse_agg(self.cb_agg.currentText())

    # 预设
    def apply_preset(self, cfg: dict, data_dir: Path):
        if "file" in cfg:
            p = data_dir / cfg["file"]
            if p.exists():
                self.ed_file.setText(str(p)); self._file_path = p
                engine = "xlrd" if p.suffix.lower() == ".xls" else "openpyxl"
                try:
                    names = pd.ExcelFile(p, engine=engine).sheet_names
                    self._all_sheets = names
                    self.cb_sheet.clear(); self.cb_sheet.addItems(names); self.cb_sheet.setEnabled(True)
                    self.btn_pick_sheets.setEnabled(True)
                except Exception:
                    pass
        if "sheet" in cfg and not self.allow_multi_sheet:
            idx = self.cb_sheet.findText(cfg["sheet"])
            if idx >= 0:
                self.cb_sheet.setCurrentIndex(idx)
        if "sheets" in cfg and self.allow_multi_sheet:
            self._selected_sheets = [s for s in cfg["sheets"] if s in self._all_sheets]
            self.ed_sheets_multi.setText(",".join(self._selected_sheets))
        if "header" in cfg:
            self.sp_header.setValue(cfg["header"])
        if "filter" in cfg:
            self.ed_filter.setText(cfg["filter"])
        # 自动加载以填充列下拉
        try:
            self._load()
        except Exception:
            pass
        if "key" in cfg:
            idx = self.cb_key.findText(cfg["key"])
            if idx >= 0:
                self.cb_key.setCurrentIndex(idx)
        if "val" in cfg:
            idx = self.cb_val.findText(cfg["val"])
            if idx >= 0:
                self.cb_val.setCurrentIndex(idx)
        if "agg" in cfg:
            for i, o in enumerate(CHECK_AGG_OPTS):
                if o.startswith(cfg["agg"]):
                    self.cb_agg.setCurrentIndex(i)
                    break


class CrossTableCheckDialog(QDialog):
    """跨表核对：A 表 vs B 表（B 支持多 sheet 合并），按键列比对值列"""

    def __init__(self, parent=None, title: str = "跨表核对",
                 default_dir: Path = None, default_out: Path = None,
                 preset: dict | None = None):
        super().__init__(parent)
        self.setWindowTitle(_t(title))
        self.resize(1180, 820)
        self.default_dir = default_dir or Path.home()
        self.default_out = default_out or (self.default_dir / "check_result.xlsx")
        self._summary = self._diff = self._full = None

        root = QVBoxLayout(self)

        # 模式选择：核对 / 查询
        from PySide6.QtWidgets import QRadioButton, QButtonGroup
        mode_box = QGroupBox("模式")
        mh = QHBoxLayout(mode_box)
        self.rb_check = QRadioButton("核对  (比对两表目标数据是否一致，输出差异报告)")
        self.rb_lookup = QRadioButton("查询  (从B表查到目标数据回填到A表，输出带结果的A表副本)")
        self.rb_check.setChecked(True)
        self._mode_group = QButtonGroup(self)
        self._mode_group.addButton(self.rb_check, 0)
        self._mode_group.addButton(self.rb_lookup, 1)
        self._mode_group.idClicked.connect(self._on_mode_change)
        mh.addWidget(self.rb_check); mh.addWidget(self.rb_lookup); mh.addStretch(1)
        root.addWidget(mode_box)

        # A / B 并排
        side_box = QHBoxLayout()
        self.sideA = _SideEditor("A 表 (基准/待核对)", self.default_dir, allow_multi_sheet=False)
        self.sideB = _SideEditor("B 表 (参照/权威，可多 sheet)", self.default_dir, allow_multi_sheet=True)
        wrapA = QGroupBox(); wrapA.setLayout(QVBoxLayout()); wrapA.layout().addWidget(self.sideA)
        wrapB = QGroupBox(); wrapB.setLayout(QVBoxLayout()); wrapB.layout().addWidget(self.sideB)
        side_box.addWidget(wrapA, 1); side_box.addWidget(wrapB, 1)
        root.addLayout(side_box)

        # 公共参数
        common = QGroupBox("核对参数")
        cl = QGridLayout(common)
        self.sp_tol = QSpinBox(); self.sp_tol.setRange(0, 10000); self.sp_tol.setValue(0)
        self.sp_tol.setSuffix(" (差额容差,单位:分)")
        self.sp_tol.setToolTip("差额绝对值 ≤ 此值 视为一致，0 表示必须完全相等")
        from PySide6.QtWidgets import QCheckBox
        self.cb_norm = QCheckBox("键列按字符串规范化（推荐，支持工号前导0等）"); self.cb_norm.setChecked(True)
        cl.addWidget(QLabel("容差:"), 0, 0); cl.addWidget(self.sp_tol, 0, 1)
        cl.addWidget(self.cb_norm, 0, 2, 1, 2)
        cl.addWidget(QLabel("输出:"), 1, 0)
        self.ed_out = QLineEdit(str(self.default_out))
        btn_out = QPushButton("…"); btn_out.setFixedWidth(30); btn_out.clicked.connect(self._pick_out)
        cl.addWidget(self.ed_out, 1, 1, 1, 2); cl.addWidget(btn_out, 1, 3)
        root.addWidget(common)

        # 结果
        gb3 = QGroupBox("核对结果")
        v3 = QVBoxLayout(gb3)
        self.result_info = QLabel("尚未执行")
        v3.addWidget(self.result_info)
        self.preview = QTableView(); self.preview.setAlternatingRowColors(True)
        self.preview.setEditTriggers(QAbstractItemView.NoEditTriggers)
        v3.addWidget(self.preview)
        # 切换预览 sheet
        prev_row = QHBoxLayout()
        self.btn_show_sum = QPushButton("显示 汇总"); self.btn_show_sum.clicked.connect(lambda: self._show("sum"))
        self.btn_show_diff = QPushButton("显示 差异明细"); self.btn_show_diff.clicked.connect(lambda: self._show("diff"))
        self.btn_show_all = QPushButton("显示 全量对比"); self.btn_show_all.clicked.connect(lambda: self._show("full"))
        self.btn_show_lookup = QPushButton("显示 查询结果"); self.btn_show_lookup.clicked.connect(lambda: self._show("lookup"))
        for b in (self.btn_show_sum, self.btn_show_diff, self.btn_show_all, self.btn_show_lookup):
            b.setEnabled(False); prev_row.addWidget(b)
        self.btn_show_lookup.setVisible(False)
        prev_row.addStretch(1)
        v3.addLayout(prev_row)
        root.addWidget(gb3, 1)

        # 按钮
        btnbox = QDialogButtonBox()
        self.btn_run = btnbox.addButton("▶ 执行核对", QDialogButtonBox.ActionRole)
        self.btn_save = btnbox.addButton("💾 保存 xlsx", QDialogButtonBox.ActionRole)
        self.btn_close = btnbox.addButton(QDialogButtonBox.Close)
        self.btn_run.clicked.connect(self._run)
        self.btn_save.clicked.connect(self._save); self.btn_save.setEnabled(False)
        self.btn_close.clicked.connect(self.accept)
        root.addWidget(btnbox)

        self._lookup = None  # 查询模式结果
        self._gb_result = gb3  # 用于改标题

        if preset:
            if preset.get("A"):
                self.sideA.apply_preset(preset["A"], self.default_dir)
            if preset.get("B"):
                self.sideB.apply_preset(preset["B"], self.default_dir)
            if "tol" in preset:
                self.sp_tol.setValue(int(preset["tol"] * 100))  # 元→分

    def _pick_out(self):
        p, _ = QFileDialog.getSaveFileName(self, "保存为", self.ed_out.text() or str(self.default_dir), "Excel (*.xlsx)")
        if p:
            if not p.lower().endswith(".xlsx"): p += ".xlsx"
            self.ed_out.setText(p)

    def _on_mode_change(self, idx):
        is_check = (idx == 0)
        self.btn_run.setText("▶ 执行核对" if is_check else "▶ 执行查询")
        self._gb_result.setTitle("核对结果" if is_check else "查询结果")
        self.btn_show_lookup.setVisible(not is_check)
        for b in (self.btn_show_sum, self.btn_show_diff, self.btn_show_all):
            b.setVisible(is_check)
        # 清空之前的结果
        self._summary = self._diff = self._full = self._lookup = None
        self.result_info.setText("尚未执行")
        self.preview.setModel(None)
        for b in (self.btn_show_sum, self.btn_show_diff, self.btn_show_all,
                  self.btn_show_lookup, self.btn_save):
            b.setEnabled(False)

    def _run(self):
        if self.rb_lookup.isChecked():
            return self._run_lookup()
        return self._run_check()

    def _run_check(self):
        try:
            dfA, kA, vA, aggA = self.sideA.get_data()
            dfB, kB, vB, aggB = self.sideB.get_data()
        except Exception as e:
            QMessageBox.critical(self, "错误", f"读取失败：{e}")
            return
        if not kA or not vA or not kB or not vB:
            QMessageBox.warning(self, "提示", "请确认关联相同字段和目标数据都已选择")
            return
        tol = self.sp_tol.value() / 100.0
        norm = self.cb_norm.isChecked()

        A = dfA[dfA[kA].notna()].copy()
        B = dfB[dfB[kB].notna()].copy()
        A["__k"] = A[kA].map(_norm_id) if norm else A[kA]
        B["__k"] = B[kB].map(_norm_id) if norm else B[kB]
        A["__v"] = A[vA].map(_to_num)
        B["__v"] = B[vB].map(_to_num)
        Ag = A.groupby("__k", as_index=False).agg(A值=("__v", aggA))
        Bg = B.groupby("__k", as_index=False).agg(B值=("__v", aggB))
        m = Ag.merge(Bg, on="__k", how="outer", indicator=True)
        m["A值"] = m["A值"].fillna(0).round(2); m["B值"] = m["B值"].fillna(0).round(2)
        m["差额(A-B)"] = (m["A值"] - m["B值"]).round(2)

        def cls(r):
            if r["_merge"] == "left_only":
                return "仅A有"
            if r["_merge"] == "right_only":
                return "仅B有(A遗漏)"
            return "一致" if abs(r["差额(A-B)"]) <= tol else "金额不一致"
        m["核对状态"] = m.apply(cls, axis=1)
        m = m.drop(columns=["_merge"]).rename(columns={"__k": "键值"})

        cnt = m["核对状态"].value_counts().to_dict()
        summary = pd.DataFrame({
            "指标": ["合集", "一致", "金额不一致", "仅A有", "仅B有(A遗漏)",
                     "A合计", "B合计", "差额合计", f"A表({aggA})行数", f"B表({aggB})行数"],
            "值": [len(m), cnt.get("一致", 0), cnt.get("金额不一致", 0),
                   cnt.get("仅A有", 0), cnt.get("仅B有(A遗漏)", 0),
                   round(m["A值"].sum(), 2), round(m["B值"].sum(), 2),
                   round(m["差额(A-B)"].sum(), 2), len(A), len(B)],
        })
        diff = m[m["核对状态"] != "一致"].reset_index(drop=True)
        self._summary, self._diff, self._full = summary, diff, m
        self._lookup = None
        self.result_info.setText(
            f"✓ 合集 {len(m)} | 一致 {cnt.get('一致',0)} | 不一致 {cnt.get('金额不一致',0)}"
            f" | 仅A {cnt.get('仅A有',0)} | 仅B {cnt.get('仅B有(A遗漏)',0)}"
        )
        for b in (self.btn_show_sum, self.btn_show_diff, self.btn_show_all, self.btn_save):
            b.setEnabled(True)
        self._show("sum")

    def _run_lookup(self):
        """查询模式：B 表的目标数据 → 按关联字段回填到 A 表的目标数据列
        不修改原A表；输出A表副本。"""
        try:
            dfA, kA, vA, aggA = self.sideA.get_data()
            dfB, kB, vB, aggB = self.sideB.get_data()
        except Exception as e:
            QMessageBox.critical(self, "错误", f"读取失败：{e}")
            return
        if not kA or not vA or not kB or not vB:
            QMessageBox.warning(self, "提示", "请确认关联相同字段和目标数据都已选择")
            return
        norm = self.cb_norm.isChecked()

        A = dfA.copy()
        B = dfB[dfB[kB].notna()].copy()
        A["__k"] = A[kA].map(_norm_id) if norm else A[kA]
        B["__k"] = B[kB].map(_norm_id) if norm else B[kB]
        # B 按关联字段聚合 + 保留来源数据值（多条用 | 分隔）
        B["__v"] = B[vB].map(_to_num)
        def _list_vals(s):
            items = []
            for v in s.dropna():
                if isinstance(v, float) and v.is_integer():
                    items.append(str(int(v)))
                else:
                    items.append(str(v))
            return " | ".join(items)
        Bg = B.groupby("__k", as_index=False).agg(
            __查询值=("__v", aggB),
            __B来源记录数=("__k", "count"),
            __B来源数据值=(vB, _list_vals),
        )

        merged = A.merge(Bg, on="__k", how="left")
        # 标识匹配
        merged["匹配状态"] = merged["__查询值"].apply(
            lambda x: "未匹配" if pd.isna(x) else "已匹配"
        )
        # 保留A原值作对比
        orig_col = f"{vA}_原值"
        merged[orig_col] = merged[vA]
        # 回填：A 的目标列 <- 查询值（未匹配保留原值）
        merged[vA] = merged["__查询值"].where(merged["__查询值"].notna(), merged[orig_col])
        # 整理输出列：保留A原有列 + 原值列 + B来源数据值 + B来源记录数 + 匹配状态
        out_cols = (
            [c for c in A.columns if c != "__k"]
            + [orig_col, "__B来源数据值", "__B来源记录数", "匹配状态"]
        )
        out = merged[out_cols].rename(columns={
            "__B来源数据值": "B_来源数据值",
            "__B来源记录数": "B_来源记录数",
        })
        # 四舍五入数值
        if pd.api.types.is_numeric_dtype(out[vA]):
            out[vA] = out[vA].round(2)

        matched = (out["匹配状态"] == "已匹配").sum()
        total = len(out)
        # 汇总
        summary = pd.DataFrame({
            "指标": [
                "A表总行数", "已匹配(B有数据)", "未匹配(B无数据)",
                f"A原『{vA}』合计", f"回填后『{vA}』合计", "变化量",
                "B聚合方式", "关联字段(A→B)",
            ],
            "值": [
                total, int(matched), int(total - matched),
                round(pd.to_numeric(out[orig_col], errors="coerce").fillna(0).sum(), 2),
                round(pd.to_numeric(out[vA], errors="coerce").fillna(0).sum(), 2),
                round(
                    pd.to_numeric(out[vA], errors="coerce").fillna(0).sum()
                    - pd.to_numeric(out[orig_col], errors="coerce").fillna(0).sum(), 2
                ),
                aggB, f"{kA} ↔ {kB}",
            ],
        })
        self._lookup = out
        self._summary = summary
        self._diff = self._full = None
        self.result_info.setText(
            f"✓ A 表 {total} 行 | 已匹配 {matched} | 未匹配 {total - matched}  →  "
            f"回填列『{vA}』，同时保留原值列『{orig_col}』"
        )
        self.btn_show_lookup.setEnabled(True)
        self.btn_show_sum.setEnabled(True)
        self.btn_save.setEnabled(True)
        self._show("lookup")

    def _show(self, which: str):
        df = {
            "sum": self._summary,
            "diff": self._diff,
            "full": self._full,
            "lookup": self._lookup,
        }.get(which)
        if df is None:
            return
        self.preview.setModel(PandasModel(df))
        self.preview.resizeColumnsToContents()

    def _save(self):
        p = Path(self.ed_out.text())
        if not p.suffix:
            p = p.with_suffix(".xlsx")
        try:
            if self.rb_lookup.isChecked() and self._lookup is not None:
                sheets = {"查询结果(A表副本)": self._lookup, "汇总": self._summary}
            elif self._summary is not None:
                sheets = {"汇总": self._summary, "差异明细": self._diff, "全量对比": self._full}
            else:
                return
            _save_xlsx(p, sheets)
            self.saved_path = p
            QMessageBox.information(self, "完成", f"已保存：{p}")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"保存失败：{e}")


# ================ 行内公式核对对话框 ================
class FormulaCheckDialog(QDialog):
    """行内公式核对：如 3月合计 = 3月缴费 - 3月退费"""

    def __init__(self, parent=None, default_dir: Path = None, default_out: Path = None, preset: dict | None = None):
        super().__init__(parent)
        self.setWindowTitle(_t("行内公式核对"))
        self.resize(980, 720)
        self.default_dir = default_dir or Path.home()
        self.default_out = default_out or (self.default_dir / "formula_check.xlsx")
        self._file_path: Path | None = None
        self._df = None
        self._result = None

        root = QVBoxLayout(self)
        gb1 = QGroupBox("① 选择文件")
        g1 = QGridLayout(gb1)
        self.ed_file = QLineEdit(); self.ed_file.setReadOnly(True)
        btn = QPushButton("浏览…"); btn.clicked.connect(self._pick_file)
        self.cb_sheet = QComboBox(); self.cb_sheet.setEnabled(False)
        self.sp_hdr = QSpinBox(); self.sp_hdr.setRange(0, 50); self.sp_hdr.setPrefix("表头行 ")
        btn_load = QPushButton("加载 ▶"); btn_load.clicked.connect(self._load)
        g1.addWidget(QLabel("文件:"), 0, 0); g1.addWidget(self.ed_file, 0, 1); g1.addWidget(btn, 0, 2)
        g1.addWidget(QLabel("Sheet:"), 1, 0); g1.addWidget(self.cb_sheet, 1, 1)
        g1.addWidget(self.sp_hdr, 1, 2); g1.addWidget(btn_load, 1, 3)
        root.addWidget(gb1)

        gb2 = QGroupBox("② 公式  (实际值列) = (X 列) [+ - * /] (Y 列)")
        g2 = QGridLayout(gb2)
        self.cb_actual = QComboBox(); self.cb_x = QComboBox(); self.cb_y = QComboBox()
        self.cb_op = QComboBox(); self.cb_op.addItems([
            "- — 减法 (验算 合计=缴-退 这类)",
            "+ — 加法 (验算 合计=a+b)",
            "* — 乘法 (验算 金额=单价×数量)",
            "/ — 除法 (验算 单价=总额÷数量)",
        ])
        g2.addWidget(QLabel("实际值列:"), 0, 0); g2.addWidget(self.cb_actual, 0, 1)
        g2.addWidget(QLabel("  = "), 0, 2); g2.addWidget(self.cb_x, 0, 3)
        g2.addWidget(self.cb_op, 0, 4); g2.addWidget(self.cb_y, 0, 5)
        self.sp_tol = QSpinBox(); self.sp_tol.setRange(0, 10000); self.sp_tol.setValue(1)
        self.sp_tol.setSuffix(" 分 (容差)")
        g2.addWidget(QLabel("容差:"), 1, 0); g2.addWidget(self.sp_tol, 1, 1)
        self.ed_extra = QLineEdit()
        self.ed_extra.setPlaceholderText("携带列（用于定位差异行，逗号分隔），如: 客户部门,工号,姓名")
        g2.addWidget(QLabel("携带列:填写输出列名称以逗号隔开"), 1, 2); g2.addWidget(self.ed_extra, 1, 3, 1, 3)
        g2.addWidget(QLabel("输出:"), 2, 0)
        self.ed_out = QLineEdit(str(self.default_out))
        btn_out = QPushButton("…"); btn_out.setFixedWidth(30); btn_out.clicked.connect(self._pick_out)
        g2.addWidget(self.ed_out, 2, 1, 1, 4); g2.addWidget(btn_out, 2, 5)
        root.addWidget(gb2)

        gb3 = QGroupBox("③ 核对结果")
        v3 = QVBoxLayout(gb3)
        self.info = QLabel("尚未执行")
        v3.addWidget(self.info)
        self.preview = QTableView(); self.preview.setAlternatingRowColors(True)
        self.preview.setEditTriggers(QAbstractItemView.NoEditTriggers)
        v3.addWidget(self.preview)
        root.addWidget(gb3, 1)

        bb = QDialogButtonBox()
        self.btn_run = bb.addButton("▶ 执行", QDialogButtonBox.ActionRole)
        self.btn_save = bb.addButton("💾 保存 xlsx", QDialogButtonBox.ActionRole)
        cls_btn = bb.addButton(QDialogButtonBox.Close)
        self.btn_run.clicked.connect(self._run)
        self.btn_save.clicked.connect(self._save); self.btn_save.setEnabled(False)
        cls_btn.clicked.connect(self.accept)
        root.addWidget(bb)

        if preset:
            self._apply_preset(preset)

    def _pick_file(self):
        p, _ = QFileDialog.getOpenFileName(self, "选择 Excel", str(self.default_dir), "Excel (*.xlsx *.xls *.XLSX)")
        if not p:
            return
        self.ed_file.setText(p)
        self._file_path = Path(p)
        try:
            engine = "xlrd" if self._file_path.suffix.lower() == ".xls" else "openpyxl"
            names = pd.ExcelFile(p, engine=engine).sheet_names
            self.cb_sheet.clear(); self.cb_sheet.addItems(names); self.cb_sheet.setEnabled(True)
        except Exception as e:
            QMessageBox.critical(self, "错误", str(e))

    def _pick_out(self):
        p, _ = QFileDialog.getSaveFileName(self, "保存为", self.ed_out.text(), "Excel (*.xlsx)")
        if p:
            if not p.lower().endswith(".xlsx"): p += ".xlsx"
            self.ed_out.setText(p)

    def _load(self):
        if not self._file_path or not self.cb_sheet.currentText():
            return
        try:
            engine = "xlrd" if self._file_path.suffix.lower() == ".xls" else "openpyxl"
            df = pd.read_excel(self._file_path, sheet_name=self.cb_sheet.currentText(),
                              header=self.sp_hdr.value(), engine=engine)
            df.columns = [str(c).strip() for c in df.columns]
            self._df = df
            cols = list(df.columns)
            for cb in (self.cb_actual, self.cb_x, self.cb_y):
                cb.clear(); cb.addItems(cols)
            self.info.setText(f"已加载：{len(df)} 行 × {df.shape[1]} 列")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"加载失败：{e}")

    def _apply_preset(self, cfg: dict):
        if "file" in cfg:
            p = self.default_dir / cfg["file"]
            if p.exists():
                self.ed_file.setText(str(p)); self._file_path = p
                engine = "xlrd" if p.suffix.lower() == ".xls" else "openpyxl"
                try:
                    names = pd.ExcelFile(p, engine=engine).sheet_names
                    self.cb_sheet.clear(); self.cb_sheet.addItems(names); self.cb_sheet.setEnabled(True)
                except Exception:
                    pass
        if "sheet" in cfg:
            idx = self.cb_sheet.findText(cfg["sheet"])
            if idx >= 0: self.cb_sheet.setCurrentIndex(idx)
        if "header" in cfg:
            self.sp_hdr.setValue(cfg["header"])
        try:
            self._load()
        except Exception:
            pass
        for key, widget in (("actual", self.cb_actual), ("x", self.cb_x), ("y", self.cb_y)):
            if key in cfg:
                idx = widget.findText(cfg[key])
                if idx >= 0: widget.setCurrentIndex(idx)
        if "op" in cfg:
            for i in range(self.cb_op.count()):
                if self.cb_op.itemText(i).startswith(cfg["op"]):
                    self.cb_op.setCurrentIndex(i); break
        if "extra" in cfg:
            self.ed_extra.setText(",".join(cfg["extra"]))
        if "tol" in cfg:
            self.sp_tol.setValue(int(cfg["tol"] * 100))

    def _run(self):
        if self._df is None:
            QMessageBox.warning(self, "提示", "请先加载文件")
            return
        col_a = self.cb_actual.currentText()
        col_x = self.cb_x.currentText()
        col_y = self.cb_y.currentText()
        op = self.cb_op.currentText().split(" ")[0]
        tol = self.sp_tol.value() / 100.0
        d = self._df.copy()
        d["__a"] = d[col_a].map(_to_num)
        d["__x"] = d[col_x].map(_to_num)
        d["__y"] = d[col_y].map(_to_num)
        if op == "+":   d["__exp"] = d["__x"] + d["__y"]
        elif op == "-": d["__exp"] = d["__x"] - d["__y"]
        elif op == "*": d["__exp"] = d["__x"] * d["__y"]
        else:           d["__exp"] = d["__x"] / d["__y"].replace(0, pd.NA)
        d["差额"] = (d["__a"] - d["__exp"]).round(4)
        d["核对"] = d["差额"].abs().apply(lambda x: "一致" if pd.notna(x) and x <= tol else "不一致")
        extras = [c.strip() for c in self.ed_extra.text().split(",") if c.strip() and c.strip() in d.columns]
        out_cols = extras + [col_a, col_x, col_y, "差额", "核对"]
        out = d[out_cols].copy()
        cnt = out["核对"].value_counts().to_dict()
        self._result = out
        self.info.setText(
            f"✓ 共 {len(out)} 行 | 一致 {cnt.get('一致', 0)} | 不一致 {cnt.get('不一致', 0)} | 差额合计 {out['差额'].sum():.2f}"
        )
        diff = out[out["核对"] == "不一致"]
        self.preview.setModel(PandasModel(diff if len(diff) else out))
        self.preview.resizeColumnsToContents()
        self.btn_save.setEnabled(True)

    def _save(self):
        if self._result is None: return
        p = Path(self.ed_out.text())
        if not p.suffix: p = p.with_suffix(".xlsx")
        diff = self._result[self._result["核对"] == "不一致"]
        try:
            _save_xlsx(p, {"差异明细": diff, "全量": self._result})
            self.saved_path = p
            QMessageBox.information(self, "完成", f"已保存：{p}")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"保存失败：{e}")


# ================ 任务预设 ================
PRESETS = {
    "t2": {
        "title": "② 核对 huizong 3月退费 vs xuesheng 退费金额",
        "dialog": "check",
        "out": "jieguo1.xlsx",
        "A": {"file": "huizong.XLSX", "sheet": "Sheet", "header": 0,
              "key": "工号", "val": "3月退费", "agg": "sum", "filter": ""},
        "B": {"file": "xuesheng.xlsx", "sheets": ["1", "2", "3"], "header": 2,
              "key": "工号", "val": "退费金额", "agg": "sum", "filter": ""},
        "tol": 0.01,
    },
    "t4": {
        "title": "④ 核对 huizong 3月合计 = 3月缴费 - 3月退费",
        "dialog": "formula",
        "out": "jieguo3.xlsx",
        "file": "huizong.XLSX", "sheet": "Sheet", "header": 0,
        "actual": "3月合计", "x": "3月缴费", "op": "-", "y": "3月退费",
        "extra": ["客户部门", "工号", "姓名"], "tol": 0.01,
    },
}


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(_t("跨表核对 · 桌面版（交互式对话框）"))
        self.resize(1280, 800)
        self.runner: TaskRunner | None = None

        # 工具栏
        tb = QToolBar("操作")
        tb.setMovable(False)
        self.addToolBar(tb)
        act_open_dir = QAction("📁 打开数据目录", self)
        act_open_dir.triggered.connect(lambda: os.startfile(str(ROOT)))
        tb.addAction(act_open_dir)
        act_open_scripts = QAction("📜 打开脚本目录", self)
        act_open_scripts.triggered.connect(lambda: os.startfile(str(SCRIPTS)))
        tb.addAction(act_open_scripts)
        tb.addSeparator()
        help_act = QAction("ℹ 使用说明", self)
        help_act.triggered.connect(lambda: QMessageBox.information(
            self, "使用说明",
            "左侧 3 个任务，每项点击后打开对应的交互式对话框：\n\n"
            "① 自定义分组聚合\n"
            "   选任意文件/Sheet/表头行 → 选主索引列 → 为其他每列独立选聚合方式\n"
            "   （sum/max/min/first/last/count/concat去重/mean，支持跳过）。\n\n"
            "② 跨表查询及核对\n"
            "   顶部先选模式：\n"
            "     • 核对 — 比对 A/B 两表『目标数据』是否一致，输出差异报告。\n"
            "     • 查询 — 从 B 表查目标数据回填到 A 表副本（不改原文件），\n"
            "              附带『_原值 / B_来源数据值 / B_来源记录数 / 匹配状态』四列。\n"
            "   A 表单 Sheet；B 表可勾选多个 Sheet 合并。\n"
            "   两侧都支持过滤表达式（pandas query，如  金额>1000  或  部门=='教职工'）。\n\n"
            "③ 带运算核对指定列\n"
            "   行内公式核对：实际值列 = X 列 [+ - × ÷] Y 列。\n"
            "   逐行用 X 与 Y 计算一个『应为』值，与『实际值列』比对，超差视为不一致。\n"
            "   『携带列』填写希望一同输出到结果中的列名，多个用英文逗号隔开（便于定位差异行）。\n\n"
            "所有对话框：文件/Sheet/表头行/关联字段/目标数据/聚合方式/过滤/容差均可自由修改；\n"
            "结果自动预览，点击『💾 保存 xlsx』生成结果文件。",
        ))
        tb.addAction(help_act)

        # 主体：左按钮区 + 右侧分割器
        central = QWidget()
        self.setCentralWidget(central)
        h = QHBoxLayout(central)
        h.setContentsMargins(8, 8, 8, 8)

        # 左：任务按钮
        left = QWidget()
        left_v = QVBoxLayout(left)
        left_v.setSpacing(6)
        title = QLabel("📋 任务列表")
        title.setFont(QFont("", 11, QFont.Bold))
        left_v.addWidget(title)
        for name, mod, outf, kind in TASKS:
            btn = QPushButton(name)
            btn.setMinimumHeight(40)
            btn.setStyleSheet(
                "QPushButton{text-align:left;padding:6px 10px;}"
                "QPushButton:hover{background:#e8f0fe;}"
            )
            btn.clicked.connect(
                lambda _=False, m=mod, o=outf, n=name, k=kind: self._on_task_click(m, o, n, k)
            )
            left_v.addWidget(btn)
        left_v.addStretch(1)
        btn_open_results = QPushButton("📂 打开结果文件夹")
        btn_open_results.clicked.connect(lambda: os.startfile(str(ROOT)))
        left_v.addWidget(btn_open_results)
        left.setFixedWidth(320)
        h.addWidget(left)

        # 右：日志(上) + 表格预览(下)
        split = QSplitter(Qt.Vertical)
        self.log = QPlainTextEdit()
        self.log.setReadOnly(True)
        self.log.setFont(QFont("Consolas", 10))
        self.log.setPlaceholderText("点击左侧任务按钮开始执行…")
        split.addWidget(self.log)

        bottom_box = QWidget()
        bv = QVBoxLayout(bottom_box)
        bv.setContentsMargins(0, 0, 0, 0)
        self.preview_label = QLabel("📊 结果预览（执行后显示当前任务的『汇总』表）")
        self.preview_label.setFont(QFont("", 10, QFont.Bold))
        bv.addWidget(self.preview_label)
        self.table = QTableView()
        self.table.setAlternatingRowColors(True)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        bv.addWidget(self.table)

        btn_row = QHBoxLayout()
        self.btn_open_xlsx = QPushButton("用 Excel 打开当前结果")
        self.btn_open_xlsx.setEnabled(False)
        self.btn_open_xlsx.clicked.connect(self._open_current)
        btn_row.addWidget(self.btn_open_xlsx)
        btn_row.addStretch(1)
        bv.addLayout(btn_row)
        split.addWidget(bottom_box)
        split.setSizes([280, 420])
        h.addWidget(split, 1)

        self.setStatusBar(QStatusBar())
        self.statusBar().showMessage(f"数据目录：{ROOT}")
        self._current_output: Path | None = None

        # 任务队列（用于"全部执行"）
        self._queue: list[tuple[str, str, str]] = []

    # ---------- 执行控制 ----------
    def _on_task_click(self, module_or_preset, output_file: str, name: str, kind: str):
        """按钮点击分发：打开对应对话框"""
        if kind == "dialog_aggregate":
            self._open_aggregate_dialog(name, output_file)
        elif kind == "dialog_check":
            self._open_check_dialog(name, output_file, module_or_preset)
        elif kind == "dialog_formula":
            self._open_formula_dialog(name, output_file, module_or_preset)

    def _open_aggregate_dialog(self, name: str, output_file: str):
        dlg = AggregateDialog(self, default_dir=ROOT, default_out=ROOT / output_file)
        self.log.appendPlainText(f"\n{'=' * 60}\n▶ 打开：{name}\n{'=' * 60}")
        dlg.exec()
        self._after_dialog(dlg, name)

    def _open_check_dialog(self, name: str, output_file: str, preset_key):
        preset = PRESETS.get(preset_key) if preset_key else None
        dlg = CrossTableCheckDialog(
            self, title=name, default_dir=ROOT,
            default_out=ROOT / output_file, preset=preset,
        )
        self.log.appendPlainText(f"\n{'=' * 60}\n▶ 打开：{name}\n{'=' * 60}")
        if preset:
            self.log.appendPlainText("（已按预设自动填充，可任意修改后执行）")
        dlg.exec()
        self._after_dialog(dlg, name)

    def _open_formula_dialog(self, name: str, output_file: str, preset_key):
        preset = PRESETS.get(preset_key) if preset_key else None
        dlg = FormulaCheckDialog(
            self, default_dir=ROOT, default_out=ROOT / output_file, preset=preset,
        )
        dlg.setWindowTitle(_t(name))
        self.log.appendPlainText(f"\n{'=' * 60}\n▶ 打开：{name}\n{'=' * 60}")
        if preset:
            self.log.appendPlainText("（已按预设自动填充，可任意修改后执行）")
        dlg.exec()
        self._after_dialog(dlg, name)

    def _after_dialog(self, dlg, name: str):
        saved = getattr(dlg, "saved_path", None)
        if saved and Path(saved).exists():
            self._current_output = Path(saved)
            self.btn_open_xlsx.setEnabled(True)
            self._load_summary(Path(saved))
            self.log.appendPlainText(f"✓ 完成：{name} → {saved}")
            self.statusBar().showMessage(f"已生成：{Path(saved).name}")
        else:
            self.log.appendPlainText(f"（{name} 关闭，未保存）")

    def run_one(self, module: str, output_file: str, name: str):
        """保留（兼容 TaskRunner），当前流程不再使用"""
        if self.runner and self.runner.isRunning():
            return
        if module is None:
            return
        self.log.appendPlainText(f"\n{'=' * 60}\n▶ 开始：{name}\n{'=' * 60}")
        self.runner = TaskRunner(module, output_file)
        self.runner.log_signal.connect(self._on_log)
        self.runner.done_signal.connect(lambda ok, o, e: self._on_done(ok, o, e, name))
        self.runner.start()

    def run_all(self):
        QMessageBox.information(
            self, "提示",
            "现在每个任务都是交互式对话框，请依次点开任务按钮手动执行。\n"
            "若需批量固定规则，可用 scripts/run_all.py 命令行脚本。"
        )

    def _next_in_queue(self):
        # 保留但不使用
        return

    def _on_log(self, text: str):
        if text.strip():
            self.log.appendPlainText(text.rstrip())

    def _on_done(self, ok: bool, output_file: str, err: str, name: str):
        if ok:
            out = ROOT / output_file
            self._current_output = out
            self.btn_open_xlsx.setEnabled(out.exists())
            self.log.appendPlainText(f"✓ 完成：{name} → {output_file}")
            self._load_summary(out)
            self.statusBar().showMessage(f"已生成：{out.name}")
        else:
            self.log.appendPlainText(f"✗ 失败：{name}")
            self.statusBar().showMessage(f"失败：{name}")
        # 队列继续
        if self._queue:
            self._next_in_queue()

    def _load_summary(self, xlsx: Path):
        try:
            xl = pd.ExcelFile(xlsx, engine="openpyxl")
            sheet = "汇总" if "汇总" in xl.sheet_names else xl.sheet_names[0]
            df = pd.read_excel(xlsx, sheet_name=sheet, engine="openpyxl")
            self.preview_label.setText(f"📊 {xlsx.name} · sheet=[{sheet}]  ({len(df)} 行)")
            self.table.setModel(PandasModel(df))
            self.table.resizeColumnsToContents()
        except Exception as e:
            self.log.appendPlainText(f"[预览失败] {e}")

    def _open_current(self):
        if self._current_output and self._current_output.exists():
            os.startfile(str(self._current_output))


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    w = MainWindow()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
