"""任务6：对比 daikou.xls vs daikou0.xls
- daikou0.xls 声称是 daikou 中删除扣费金额为0后的结果
- 验证：
  1) daikou - daikou0 的差集 是否等价于 "daikou中扣费=0"
  2) daikou0 中非0记录 是否与 daikou 非0记录完全一致
→ jieguo5.xlsx
"""
import pandas as pd
from common import DATA_DIR, norm_id, to_num, save_excel

DK = DATA_DIR / "daikou.xls"
DK0 = DATA_DIR / "daikou0.xls"
OUT = DATA_DIR / "jieguo5.xlsx"


def load(path):
    df = pd.read_excel(path, sheet_name="导入更新模板", engine="xlrd")
    df.columns = [str(c).strip() for c in df.columns]
    cid = next(c for c in df.columns if "学" in c and "号" in c)
    camt = next(c for c in df.columns if "扣费" in c)
    df = df[df[cid].notna()].copy()
    df["__id"] = df[cid].map(norm_id)
    df["__amt"] = df[camt].map(to_num)
    return df


dk = load(DK)
dk0 = load(DK0)

# 以学工号为键，对 daikou 聚合（若有重复）
dk_agg = dk.groupby("__id", as_index=False).agg(daikou金额=("__amt", "sum"), daikou次数=("__id", "count"))
dk0_agg = dk0.groupby("__id", as_index=False).agg(daikou0金额=("__amt", "sum"), daikou0次数=("__id", "count"))

merged = dk_agg.merge(dk0_agg, on="__id", how="outer", indicator=True)
merged["daikou金额"] = merged["daikou金额"].fillna(0).round(2)
merged["daikou0金额"] = merged["daikou0金额"].fillna(0).round(2)
merged["daikou次数"] = merged["daikou次数"].fillna(0).astype(int)
merged["daikou0次数"] = merged["daikou0次数"].fillna(0).astype(int)

def classify(r):
    if r["_merge"] == "left_only":
        # daikou 有，daikou0 没有 → 期望：daikou金额==0
        return "仅daikou(金额为0,已删除✓)" if r["daikou金额"] == 0 else "仅daikou(金额非0,不应删除✗)"
    if r["_merge"] == "right_only":
        return "仅daikou0(新增)"
    if abs(r["daikou金额"] - r["daikou0金额"]) < 0.01:
        return "一致"
    return "两边都有但金额不同"

merged["核对状态"] = merged.apply(classify, axis=1)
merged = merged.drop(columns=["_merge"]).rename(columns={"__id": "学工号"})

# daikou 中金额为 0 的行
dk_zero = dk[dk["__amt"] == 0][["__id", "__amt"]].rename(columns={"__id": "学工号", "__amt": "扣费金额"})

summary = pd.DataFrame({
    "指标": [
        "daikou 行数", "daikou0 行数", "daikou 金额=0 行数",
        "预期daikou0行数 = daikou-零行", "全集合数",
        "一致", "仅daikou(零,删除✓)", "仅daikou(非零,异常✗)", "仅daikou0(新增)", "两边金额不同",
        "daikou 金额合计", "daikou0 金额合计", "差额",
    ],
    "值": [
        len(dk), len(dk0), len(dk_zero),
        len(dk) - len(dk_zero), len(merged),
        (merged["核对状态"] == "一致").sum(),
        (merged["核对状态"] == "仅daikou(金额为0,已删除✓)").sum(),
        (merged["核对状态"] == "仅daikou(金额非0,不应删除✗)").sum(),
        (merged["核对状态"] == "仅daikou0(新增)").sum(),
        (merged["核对状态"] == "两边都有但金额不同").sum(),
        round(dk["__amt"].sum(), 2), round(dk0["__amt"].sum(), 2),
        round(dk["__amt"].sum() - dk0["__amt"].sum(), 2),
    ],
})

diff = merged[merged["核对状态"] != "一致"]

save_excel(OUT, {"汇总": summary, "差异明细": diff, "全量对比": merged, "daikou中零金额行": dk_zero})
print(summary.to_string(index=False))
print(f"[OK] 输出：{OUT}")
