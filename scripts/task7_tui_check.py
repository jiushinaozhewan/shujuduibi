"""任务7：核对 tui.xlsx 的"金额" vs 3.xlsx 三个sheet 的退费金额>1000
以学工号为主索引：
- 筛选 3.xlsx 三个 sheet 中 退费金额 > 1000 的记录（按工号聚合）
- 对比 tui.xlsx 的金额
→ jieguo6.xlsx
"""
import pandas as pd
from common import DATA_DIR, norm_id, to_num, save_excel

TUI = DATA_DIR / "tui.xlsx"
SRC = DATA_DIR / "3.xlsx"
OUT = DATA_DIR / "jieguo6.xlsx"

# tui.xlsx：表头在 row 0，但列名是 班级/_/学工号/姓名/性别/金额
tui = pd.read_excel(TUI, sheet_name="Sheet1", engine="openpyxl")
tui.columns = [str(c).strip() for c in tui.columns]
col_tui_id = "学工号"
col_tui_amt = "金额"
tui = tui[tui[col_tui_id].notna()].copy()
tui["__id"] = tui[col_tui_id].map(norm_id)
tui["__amt"] = tui[col_tui_amt].map(to_num)

# 3.xlsx：三个 sheet，表头在 row 2
xs_all = []
for s in ["3", "2", "1"]:
    df = pd.read_excel(SRC, sheet_name=s, header=2, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]
    df["__sheet"] = s
    xs_all.append(df)
xs = pd.concat(xs_all, ignore_index=True)
xs = xs[xs["工号"].notna()].copy()
xs["__id"] = xs["工号"].map(norm_id)
xs["__amt"] = xs["退费金额"].map(to_num)

# 仅保留退费金额 > 1000 的行，再按工号聚合
xs_big = xs[xs["__amt"] > 1000].copy()
xs_agg = xs_big.groupby("__id", as_index=False).agg(
    xsh退费金额_gt1000=("__amt", "sum"),
    出现sheet=("__sheet", lambda s: ",".join(sorted(set(s)))),
    xsh记录数=("__id", "count"),
)

tui_slim = tui[["__id", "班级", "学工号", "姓名", "性别", "__amt"]].rename(columns={"__amt": "tui_金额"})

merged = tui_slim.merge(xs_agg, on="__id", how="outer", indicator=True)
merged["tui_金额"] = merged["tui_金额"].fillna(0).round(2)
merged["xsh退费金额_gt1000"] = merged["xsh退费金额_gt1000"].fillna(0).round(2)
merged["差额(tui-xsh)"] = (merged["tui_金额"] - merged["xsh退费金额_gt1000"]).round(2)

def classify(r):
    if r["_merge"] == "left_only":
        return "仅tui有(3.xlsx无>1000记录)"
    if r["_merge"] == "right_only":
        return "仅3.xlsx有(tui遗漏)"
    return "一致" if abs(r["差额(tui-xsh)"]) < 0.01 else "金额不一致"

merged["核对状态"] = merged.apply(classify, axis=1)
merged = merged.drop(columns=["_merge"]).rename(columns={"__id": "工号(索引)"})

diff = merged[merged["核对状态"] != "一致"]
summary = pd.DataFrame({
    "指标": [
        "tui 行数", "3.xlsx 总行数", "3.xlsx 退费>1000 行数", "3.xlsx >1000 聚合工号数",
        "全集合数", "一致", "金额不一致", "仅tui有", "仅3.xlsx有(遗漏)",
        "tui 金额合计", "3.xlsx >1000 合计", "差额",
    ],
    "值": [
        len(tui), len(xs), len(xs_big), len(xs_agg), len(merged),
        (merged["核对状态"] == "一致").sum(),
        (merged["核对状态"] == "金额不一致").sum(),
        (merged["核对状态"] == "仅tui有(3.xlsx无>1000记录)").sum(),
        (merged["核对状态"] == "仅3.xlsx有(tui遗漏)").sum(),
        round(merged["tui_金额"].sum(), 2),
        round(merged["xsh退费金额_gt1000"].sum(), 2),
        round(merged["差额(tui-xsh)"].sum(), 2),
    ],
})

save_excel(OUT, {"汇总": summary, "差异明细": diff, "全量对比": merged, "3xlsx大于1000源": xs_big[["__sheet","__id","工号","姓名","__amt"]].rename(columns={"__sheet":"sheet","__id":"工号规范化","__amt":"退费金额"})})
print(summary.to_string(index=False))
print(f"[OK] 输出：{OUT}")
