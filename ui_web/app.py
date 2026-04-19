"""Web 版 · 跨表核对 (Streamlit)
功能与桌面版完全一致：
  ① 自定义分组聚合
  ② 跨表查询及核对（含 核对 / 查询 两种模式）
  ③ 带运算核对指定列
"""
from __future__ import annotations
import io
from pathlib import Path
import pandas as pd
import streamlit as st

st.set_page_config(page_title="跨表核对 · Web 版", layout="wide", page_icon="📊")

import sys
if getattr(sys, "frozen", False):
    DATA_DIR = Path(sys.executable).resolve().parent
else:
    DATA_DIR = Path(r"D:/meta/shujuduibi")


# ================ 通用工具 ================
def norm_id(v):
    if pd.isna(v):
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def to_num(v):
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


def list_vals(s: pd.Series) -> str:
    items = []
    for v in s.dropna():
        if isinstance(v, float) and v.is_integer():
            items.append(str(int(v)))
        else:
            items.append(str(v))
    return " | ".join(items)


@st.cache_data(show_spinner=False)
def list_sheets(file_bytes: bytes, ext: str) -> list[str]:
    engine = "xlrd" if ext.lower() == ".xls" else "openpyxl"
    return pd.ExcelFile(io.BytesIO(file_bytes), engine=engine).sheet_names


@st.cache_data(show_spinner=False)
def load_sheet(file_bytes: bytes, ext: str, sheet: str, header_row: int) -> pd.DataFrame:
    engine = "xlrd" if ext.lower() == ".xls" else "openpyxl"
    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet, header=header_row, engine=engine)
    df.columns = [str(c).strip() for c in df.columns]
    return df


@st.cache_data(show_spinner=False)
def load_concat(file_bytes: bytes, ext: str, sheets: tuple[str, ...], header_row: int) -> pd.DataFrame:
    engine = "xlrd" if ext.lower() == ".xls" else "openpyxl"
    parts = []
    for s in sheets:
        d = pd.read_excel(io.BytesIO(file_bytes), sheet_name=s, header=header_row, engine=engine)
        d.columns = [str(c).strip() for c in d.columns]
        d["__src_sheet"] = s
        parts.append(d)
    return pd.concat(parts, ignore_index=True) if parts else pd.DataFrame()


def apply_filter(df: pd.DataFrame, expr: str) -> pd.DataFrame:
    expr = (expr or "").strip()
    if not expr:
        return df
    return df.query(expr, engine="python")


def to_xlsx_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
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
    return buf.getvalue()


# ================ 聚合方式选项（与桌面版完全一致） ================
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
CHECK_AGG_OPTS = [
    "sum — 求和 (同键多行求和，如退费金额合计)",
    "first — 首条 (同键取第一次出现的那行)",
    "last — 末条 (同键取最后一次出现的那行)",
    "max — 最大值 (同键取最大数值或最新日期)",
    "min — 最小值 (同键取最小)",
    "mean — 平均值 (同键数值平均)",
]


def agg_key(opt_text: str) -> str:
    s = opt_text.strip()
    if s.startswith("—跳过—"):
        return "—跳过—"
    return s.split(" ")[0]


def suggest_agg(col_name: str, series: pd.Series) -> int:
    """按列名/类型推荐 AGG_OPTS 下标"""
    name = str(col_name)
    def _find(prefix):
        for i, o in enumerate(AGG_OPTS):
            if o.startswith(prefix):
                return i
        return 0
    if pd.api.types.is_numeric_dtype(series) or any(k in name for k in ["金额", "缴费", "退费", "合计", "总", "数量"]):
        return _find("sum")
    if any(k in name for k in ["时间", "日期"]) or pd.api.types.is_datetime64_any_dtype(series):
        return _find("max")
    if any(k in name for k in ["姓名", "名称", "部门", "班级", "性别"]):
        return _find("last")
    return 0


# ================ 预设（与桌面版对齐） ================
PRESETS = {
    "t2": {
        "title": "② 跨表查询及核对（默认：huizong / xuesheng 退费）",
        "A": {"file": "huizong.XLSX", "sheet": "Sheet", "header": 0,
              "key": "工号", "val": "3月退费", "agg": "sum", "filter": ""},
        "B": {"file": "xuesheng.xlsx", "sheets": ["1", "2", "3"], "header": 2,
              "key": "工号", "val": "退费金额", "agg": "sum", "filter": ""},
        "tol": 0.01,
    },
    "t4": {
        "title": "③ 带运算核对指定列（默认：huizong 3月合计=3月缴费-3月退费）",
        "file": "huizong.XLSX", "sheet": "Sheet", "header": 0,
        "actual": "3月合计", "x": "3月缴费", "op": "-", "y": "3月退费",
        "extra": ["客户部门", "工号", "姓名"], "tol": 0.01,
    },
}


# ================ 会话状态 ================
if "files" not in st.session_state:
    st.session_state.files = {}  # name -> {'bytes':..., 'ext':...}


def add_file(name: str, data: bytes):
    st.session_state.files[name] = {"bytes": data, "ext": Path(name).suffix}


# ================ 侧栏：文件管家 ================
with st.sidebar:
    st.title("📁 文件管家")

    st.markdown("**从本地数据目录加载**")
    if DATA_DIR.exists():
        local_files = sorted(
            [f for f in DATA_DIR.iterdir() if f.suffix.lower() in (".xlsx", ".xls")]
        )
        chosen = st.multiselect(f"目录: {DATA_DIR}", [f.name for f in local_files])
        if st.button("加载选中文件", use_container_width=True):
            for fn in chosen:
                add_file(fn, (DATA_DIR / fn).read_bytes())
            st.success(f"已加载 {len(chosen)} 个")
    else:
        st.info(f"默认目录不存在：{DATA_DIR}")

    st.markdown("**或上传文件**")
    up = st.file_uploader("选择 Excel", type=["xlsx", "xls"], accept_multiple_files=True)
    if up:
        for f in up:
            add_file(f.name, f.read())
        st.success(f"已上传 {len(up)} 个")

    st.markdown("---")
    st.markdown(f"**已加载：{len(st.session_state.files)} 个**")
    for n in list(st.session_state.files):
        c1, c2 = st.columns([5, 1])
        c1.write(f"• {n}")
        if c2.button("✕", key=f"del_{n}"):
            del st.session_state.files[n]
            st.rerun()
    if st.session_state.files and st.button("🗑 清空全部", use_container_width=True):
        st.session_state.files = {}
        st.rerun()


# ================ 文件/Sheet 选择器（小组件） ================
def file_sheet_picker(key: str, multi_sheet: bool = False, default_file: str | None = None,
                     default_sheet: str | None = None, default_sheets: list[str] | None = None,
                     default_header: int = 0):
    """返回 (filename, sheet_or_sheets, header, df_or_None)"""
    files = list(st.session_state.files.keys())
    if not files:
        st.info("请在左侧『文件管家』先加载或上传文件")
        return None, None, 0, None
    # 文件
    idx = files.index(default_file) if (default_file and default_file in files) else 0
    fn = st.selectbox("文件", files, index=idx, key=f"{key}_file")
    info = st.session_state.files[fn]
    sheets = list_sheets(info["bytes"], info["ext"])
    # Sheet
    if multi_sheet:
        defaults = [s for s in (default_sheets or []) if s in sheets]
        sel_sheets = st.multiselect("Sheets（可多选合并）", sheets, default=defaults, key=f"{key}_sheets")
    else:
        sidx = sheets.index(default_sheet) if (default_sheet and default_sheet in sheets) else 0
        sel_sheets = st.selectbox("Sheet", sheets, index=sidx, key=f"{key}_sheet")
    hdr = st.number_input("表头所在行（0=第1行）", 0, 50, default_header, key=f"{key}_hdr")
    # 读
    try:
        if multi_sheet:
            if not sel_sheets:
                return fn, sel_sheets, int(hdr), None
            df = load_concat(info["bytes"], info["ext"], tuple(sel_sheets), int(hdr))
        else:
            df = load_sheet(info["bytes"], info["ext"], sel_sheets, int(hdr))
    except Exception as e:
        st.error(f"读取失败：{e}")
        return fn, sel_sheets, int(hdr), None
    return fn, sel_sheets, int(hdr), df


# ================ 主页 ================
st.title("📊 跨表核对 · Web 版")
st.caption("功能与桌面版一致：① 自定义分组聚合 ② 跨表查询及核对 ③ 带运算核对指定列")

tab1, tab2, tab3 = st.tabs([
    "① 自定义分组聚合",
    "② 跨表查询及核对",
    "③ 带运算核对指定列",
])


# ---------------- Tab 1：自定义分组聚合 ----------------
with tab1:
    st.subheader("① 自定义分组聚合")
    st.caption("选任意文件 → 选主索引列 → 为每列独立选聚合方式（支持跳过）")

    with st.expander("① 选择文件", expanded=True):
        fn, sheet, hdr, df = file_sheet_picker("agg")
    if df is not None:
        with st.expander("② 聚合规则", expanded=True):
            cols = list(df.columns)
            key_col = st.selectbox("主索引列", cols, key="agg_key")
            st.caption("提示：每列下拉框内包含聚合方式的详细说明；默认按列名/类型智能推荐，可随意修改")
            rules: dict[str, str] = {}
            other_cols = [c for c in cols if c != key_col]
            # 3 列网格
            ncols = 3
            for i in range(0, len(other_cols), ncols):
                row = st.columns(ncols)
                for j, c in enumerate(other_cols[i:i + ncols]):
                    with row[j]:
                        default_idx = suggest_agg(c, df[c])
                        rules[c] = st.selectbox(c, AGG_OPTS, index=default_idx, key=f"agg_{c}")
            out_name = st.text_input("输出文件名", value="aggregate_custom.xlsx", key="agg_outname")

        if st.button("▶ 执行聚合", type="primary", key="agg_run"):
            work = df[df[key_col].notna()].copy()
            work["__k"] = work[key_col].map(norm_id)
            work = work[work["__k"] != ""]
            agg_dict: dict = {}
            for col, opt in rules.items():
                mode = agg_key(opt)
                if mode == "—跳过—":
                    continue
                if mode in ("sum", "mean", "min", "max"):
                    if mode == "max":
                        try:
                            dt = pd.to_datetime(work[col], errors="coerce")
                            if dt.notna().sum() > 0:
                                work[col] = dt
                            else:
                                work[col] = work[col].map(to_num)
                        except Exception:
                            work[col] = work[col].map(to_num)
                    else:
                        work[col] = work[col].map(to_num)
                    agg_dict[col] = mode
                elif mode == "concat去重":
                    agg_dict[col] = lambda s: ",".join(sorted({str(x) for x in s.dropna()}))
                else:
                    agg_dict[col] = mode
            if not agg_dict:
                st.warning("请至少为一列选聚合方式")
            else:
                out = work.groupby("__k", as_index=False).agg(agg_dict)
                cnt = work.groupby("__k", as_index=False).size().rename(columns={"size": "记录笔数"})
                out = out.merge(cnt, on="__k").rename(columns={"__k": key_col})
                for c, m in agg_dict.items():
                    if m in ("sum", "mean") and pd.api.types.is_numeric_dtype(out[c]):
                        out[c] = out[c].round(2)
                out = out.sort_values(key_col).reset_index(drop=True)
                st.success(f"✓ 源 {len(work)} 行 → 聚合后 {len(out)} 组 | 主键：{key_col} | 规则：{len(agg_dict)} 列")
                st.dataframe(out, use_container_width=True, height=420)
                st.download_button("⬇ 下载结果", data=to_xlsx_bytes({"聚合结果": out}),
                                  file_name=out_name, key="agg_dl")


# ---------------- Tab 2：跨表查询及核对 ----------------
with tab2:
    st.subheader("② 跨表查询及核对")
    c1, c2 = st.columns([3, 1])
    with c1:
        mode = st.radio(
            "模式",
            [
                "核对  (比对两表目标数据是否一致，输出差异报告)",
                "查询  (从B表查到目标数据回填到A表，输出带结果的A表副本)",
            ],
            horizontal=True, key="ck_mode",
        )
    with c2:
        if st.button("应用默认预设 (t2)", key="ck_preset"):
            # 加载默认文件到文件管家
            for name in (PRESETS["t2"]["A"]["file"], PRESETS["t2"]["B"]["file"]):
                p = DATA_DIR / name
                if p.exists() and name not in st.session_state.files:
                    add_file(name, p.read_bytes())
            # 设置初始选择值（通过删除已存在的 widget key 再设置）
            for k in list(st.session_state.keys()):
                if k.startswith(("ckA_", "ckB_")):
                    del st.session_state[k]
            pa = PRESETS["t2"]["A"]; pb = PRESETS["t2"]["B"]
            st.session_state["_preset_A"] = pa
            st.session_state["_preset_B"] = pb
            st.session_state["_preset_tol"] = PRESETS["t2"]["tol"]
            st.rerun()

    is_check = mode.startswith("核对")

    pa = st.session_state.get("_preset_A", {})
    pb = st.session_state.get("_preset_B", {})

    colA, colB = st.columns(2)
    with colA:
        st.markdown("#### 【A 表 · 基准/待核对】")
        fnA, shA, hdrA, dfA = file_sheet_picker(
            "ckA",
            multi_sheet=False,
            default_file=pa.get("file"),
            default_sheet=pa.get("sheet"),
            default_header=pa.get("header", 0),
        )
        filtA = st.text_input("过滤表达式(可选)", value=pa.get("filter", ""),
                             placeholder="如  金额>1000   或   部门=='教职工'", key="ckA_filter")
        if dfA is not None:
            try:
                dfA_filt = apply_filter(dfA, filtA)
            except Exception as e:
                st.error(f"A 过滤表达式错误：{e}")
                dfA_filt = None
            if dfA_filt is not None:
                cols = list(dfA_filt.columns)
                ki = cols.index(pa["key"]) if pa.get("key") in cols else 0
                vi = cols.index(pa["val"]) if pa.get("val") in cols else 0
                kA = st.selectbox("关联相同字段", cols, index=ki, key="ckA_key")
                vA = st.selectbox("目标数据", cols, index=vi, key="ckA_val")
                ai = 0
                if pa.get("agg"):
                    for i, o in enumerate(CHECK_AGG_OPTS):
                        if o.startswith(pa["agg"]):
                            ai = i; break
                aggA_text = st.selectbox("同键聚合", CHECK_AGG_OPTS, index=ai, key="ckA_agg")
                st.caption(f"✓ A 加载 {len(dfA_filt)} 行")
        else:
            dfA_filt = None
            kA = vA = aggA_text = None

    with colB:
        st.markdown("#### 【B 表 · 参照/权威，可多 Sheet】")
        fnB, shB, hdrB, dfB = file_sheet_picker(
            "ckB",
            multi_sheet=True,
            default_file=pb.get("file"),
            default_sheets=pb.get("sheets", []),
            default_header=pb.get("header", 0),
        )
        filtB = st.text_input("过滤表达式(可选)", value=pb.get("filter", ""),
                             placeholder="如  退费金额>1000", key="ckB_filter")
        if dfB is not None:
            try:
                dfB_filt = apply_filter(dfB, filtB)
            except Exception as e:
                st.error(f"B 过滤表达式错误：{e}")
                dfB_filt = None
            if dfB_filt is not None:
                cols = list(dfB_filt.columns)
                ki = cols.index(pb["key"]) if pb.get("key") in cols else 0
                vi = cols.index(pb["val"]) if pb.get("val") in cols else 0
                kB = st.selectbox("关联相同字段", cols, index=ki, key="ckB_key")
                vB = st.selectbox("目标数据", cols, index=vi, key="ckB_val")
                ai = 0
                if pb.get("agg"):
                    for i, o in enumerate(CHECK_AGG_OPTS):
                        if o.startswith(pb["agg"]):
                            ai = i; break
                aggB_text = st.selectbox("同键聚合", CHECK_AGG_OPTS, index=ai, key="ckB_agg")
                st.caption(f"✓ B 加载 {len(dfB_filt)} 行")
        else:
            dfB_filt = None
            kB = vB = aggB_text = None

    c1, c2, c3 = st.columns(3)
    default_tol = st.session_state.get("_preset_tol", 0.01)
    tol = c1.number_input("差额容差", 0.0, 10000.0, float(default_tol), step=0.01, key="ck_tol")
    norm = c2.checkbox("键列按字符串规范化（推荐）", value=True, key="ck_norm")
    out_name = c3.text_input("输出文件名", value="jieguo1.xlsx", key="ck_outname")

    can_run = (dfA_filt is not None) and (dfB_filt is not None) and all([kA, vA, kB, vB])
    run_label = "▶ 执行核对" if is_check else "▶ 执行查询"
    if st.button(run_label, type="primary", disabled=not can_run, key="ck_run"):
        aggA = agg_key(aggA_text); aggB = agg_key(aggB_text)

        if is_check:
            # ---------- 核对 ----------
            A = dfA_filt[dfA_filt[kA].notna()].copy()
            B = dfB_filt[dfB_filt[kB].notna()].copy()
            A["__k"] = A[kA].map(norm_id) if norm else A[kA]
            B["__k"] = B[kB].map(norm_id) if norm else B[kB]
            A["__v"] = A[vA].map(to_num); B["__v"] = B[vB].map(to_num)
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

            k1, k2, k3, k4 = st.columns(4)
            k1.metric("一致", cnt.get("一致", 0))
            k2.metric("不一致", cnt.get("金额不一致", 0))
            k3.metric("仅A有", cnt.get("仅A有", 0))
            k4.metric("仅B有", cnt.get("仅B有(A遗漏)", 0))
            st.dataframe(summary, use_container_width=True, hide_index=True)
            with st.expander(f"差异明细（{len(diff)} 条）", expanded=True):
                st.dataframe(diff, use_container_width=True, height=320)
            with st.expander("全量对比", expanded=False):
                st.dataframe(m, use_container_width=True, height=320)
            st.download_button("⬇ 下载核对结果",
                              data=to_xlsx_bytes({"汇总": summary, "差异明细": diff, "全量对比": m}),
                              file_name=out_name, key="ck_dl")
        else:
            # ---------- 查询：B 回填到 A 副本 ----------
            A = dfA_filt.copy()
            B = dfB_filt[dfB_filt[kB].notna()].copy()
            A["__k"] = A[kA].map(norm_id) if norm else A[kA]
            B["__k"] = B[kB].map(norm_id) if norm else B[kB]
            B["__v"] = B[vB].map(to_num)
            Bg = B.groupby("__k", as_index=False).agg(
                __查询值=("__v", aggB),
                __B来源记录数=("__k", "count"),
                __B来源数据值=(vB, list_vals),
            )
            merged = A.merge(Bg, on="__k", how="left")
            merged["匹配状态"] = merged["__查询值"].apply(lambda x: "未匹配" if pd.isna(x) else "已匹配")
            orig_col = f"{vA}_原值"
            merged[orig_col] = merged[vA]
            merged[vA] = merged["__查询值"].where(merged["__查询值"].notna(), merged[orig_col])
            out_cols = (
                [c for c in A.columns if c != "__k"]
                + [orig_col, "__B来源数据值", "__B来源记录数", "匹配状态"]
            )
            out = merged[out_cols].rename(columns={
                "__B来源数据值": "B_来源数据值",
                "__B来源记录数": "B_来源记录数",
            })
            if pd.api.types.is_numeric_dtype(out[vA]):
                out[vA] = out[vA].round(2)

            matched = int((out["匹配状态"] == "已匹配").sum())
            total = len(out)
            summary = pd.DataFrame({
                "指标": [
                    "A表总行数", "已匹配(B有数据)", "未匹配(B无数据)",
                    f"A原『{vA}』合计", f"回填后『{vA}』合计", "变化量",
                    "B聚合方式", "关联字段(A→B)",
                ],
                "值": [
                    total, matched, total - matched,
                    round(pd.to_numeric(out[orig_col], errors="coerce").fillna(0).sum(), 2),
                    round(pd.to_numeric(out[vA], errors="coerce").fillna(0).sum(), 2),
                    round(
                        pd.to_numeric(out[vA], errors="coerce").fillna(0).sum()
                        - pd.to_numeric(out[orig_col], errors="coerce").fillna(0).sum(), 2
                    ),
                    aggB, f"{kA} ↔ {kB}",
                ],
            })
            k1, k2, k3 = st.columns(3)
            k1.metric("A 表总行数", total)
            k2.metric("已匹配", matched)
            k3.metric("未匹配", total - matched)
            st.dataframe(summary, use_container_width=True, hide_index=True)
            st.markdown(f"**查询结果 (A 表副本，已回填『{vA}』列)**")
            st.dataframe(out, use_container_width=True, height=420)
            st.download_button("⬇ 下载查询结果",
                              data=to_xlsx_bytes({"查询结果(A表副本)": out, "汇总": summary}),
                              file_name=out_name, key="ck_dl_lookup")


# ---------------- Tab 3：带运算核对指定列 ----------------
with tab3:
    st.subheader("③ 带运算核对指定列")
    st.caption("行内公式核对：实际值列 = X 列 [+ - × ÷] Y 列")

    c_top1, c_top2 = st.columns([3, 1])
    with c_top2:
        if st.button("应用默认预设 (t4)", key="fm_preset"):
            p = PRESETS["t4"]
            if p["file"] not in st.session_state.files:
                src = DATA_DIR / p["file"]
                if src.exists():
                    add_file(p["file"], src.read_bytes())
            for k in list(st.session_state.keys()):
                if k.startswith("fm_"):
                    del st.session_state[k]
            st.session_state["_preset_t4"] = p
            st.rerun()

    pt = st.session_state.get("_preset_t4", {})
    with st.expander("① 选择文件", expanded=True):
        fn, sheet, hdr, df = file_sheet_picker(
            "fm",
            default_file=pt.get("file"),
            default_sheet=pt.get("sheet"),
            default_header=pt.get("header", 0),
        )

    if df is not None:
        with st.expander("② 公式  (实际值列) = (X 列) [+ - × ÷] (Y 列)", expanded=True):
            cols = list(df.columns)
            c1, c2, c3, c4 = st.columns([3, 3, 2, 3])

            def _idx(name):
                return cols.index(name) if (name and name in cols) else 0

            actual = c1.selectbox("实际值列", cols, index=_idx(pt.get("actual")), key="fm_actual")
            x_col = c2.selectbox("X 列", cols, index=_idx(pt.get("x")), key="fm_x")
            op_opts = [
                "- — 减法 (验算 合计=缴-退 这类)",
                "+ — 加法 (验算 合计=a+b)",
                "* — 乘法 (验算 金额=单价×数量)",
                "/ — 除法 (验算 单价=总额÷数量)",
            ]
            op_idx = 0
            if pt.get("op"):
                for i, o in enumerate(op_opts):
                    if o.startswith(pt["op"]):
                        op_idx = i; break
            op_text = c3.selectbox("运算", op_opts, index=op_idx, key="fm_op")
            y_col = c4.selectbox("Y 列", cols, index=_idx(pt.get("y")), key="fm_y")

            c1, c2, c3 = st.columns([1, 3, 3])
            tol = c1.number_input("容差", 0.0, 10000.0, float(pt.get("tol", 0.01)), step=0.01, key="fm_tol")
            default_extra = ",".join(pt.get("extra", [])) if pt.get("extra") else ""
            extra_text = c2.text_input("携带列:填写输出列名称以逗号隔开",
                                       value=default_extra,
                                       placeholder="如: 客户部门,工号,姓名", key="fm_extra")
            out_name = c3.text_input("输出文件名", value="jieguo3.xlsx", key="fm_outname")

        if st.button("▶ 执行核对", type="primary", key="fm_run"):
            op = op_text.split(" ")[0]
            d = df.copy()
            d["__a"] = d[actual].map(to_num)
            d["__x"] = d[x_col].map(to_num)
            d["__y"] = d[y_col].map(to_num)
            if op == "+":   d["__exp"] = d["__x"] + d["__y"]
            elif op == "-": d["__exp"] = d["__x"] - d["__y"]
            elif op == "*": d["__exp"] = d["__x"] * d["__y"]
            else:           d["__exp"] = d["__x"] / d["__y"].replace(0, pd.NA)
            d["差额"] = (d["__a"] - d["__exp"]).round(4)
            d["核对"] = d["差额"].abs().apply(lambda x: "一致" if pd.notna(x) and x <= tol else "不一致")
            extras = [c.strip() for c in (extra_text or "").split(",") if c.strip() and c.strip() in d.columns]
            out_cols = extras + [actual, x_col, y_col, "差额", "核对"]
            out = d[out_cols].copy()
            cnt = out["核对"].value_counts().to_dict()
            k1, k2, k3 = st.columns(3)
            k1.metric("一致", cnt.get("一致", 0))
            k2.metric("不一致", cnt.get("不一致", 0))
            k3.metric("差额合计", f"{out['差额'].sum():.2f}")
            diff = out[out["核对"] == "不一致"]
            st.markdown("**差异明细**" if len(diff) else "**全部一致；下表为全量**")
            st.dataframe(diff if len(diff) else out, use_container_width=True, height=420)
            st.download_button("⬇ 下载核对结果",
                              data=to_xlsx_bytes({"差异明细": diff, "全量": out}),
                              file_name=out_name, key="fm_dl")
