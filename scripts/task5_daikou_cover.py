"""任务5：核对 daikou.xls 是否完整、准确地包含 huizong(学工号,3月合计) 组合
- 遗漏：huizong 有但 daikou 没有
- 多余：daikou 有但 huizong 没有
- 金额不符：两边都有但扣费金额 ≠ 3月合计
→ jieguo4.xlsx
"""
import pandas as pd
from common import DATA_DIR, norm_id, to_num, save_excel

HUIZONG = DATA_DIR / "huizong.XLSX"
DAIKOU = DATA_DIR / "daikou.xls"
OUT = DATA_DIR / "jieguo4.xlsx"

hz = pd.read_excel(HUIZONG, sheet_name="Sheet", engine="openpyxl")
hz.columns = [str(c).strip() for c in hz.columns]
hz = hz[hz["工号"].notna()].copy()
hz["__id"] = hz["工号"].map(norm_id)
hz["__heji"] = hz["3月合计"].map(to_num)

dk = pd.read_excel(DAIKOU, sheet_name="导入更新模板", engine="xlrd")
dk.columns = [str(c).strip() for c in dk.columns]
col_id = next(c for c in dk.columns if "学" in c and "号" in c)
col_amt = next(c for c in dk.columns if "扣费" in c)
dk = dk[dk[col_id].notna()].copy()
dk["__id"] = dk[col_id].map(norm_id)
dk["__amt"] = dk[col_amt].map(to_num)

# daikou 内部是否有重复工号
dk_dup = dk[dk.duplicated("__id", keep=False)].sort_values("__id")

hz_slim = hz[["__id", "客户部门", "工号", "姓名", "__heji"]].rename(columns={"__heji": "huizong_3月合计"})
dk_slim = dk[["__id", "__amt"]].rename(columns={"__amt": "daikou_扣费金额"})

# 先聚合 daikou（防重复）
dk_agg = dk.groupby("__id", as_index=False).agg(
    daikou_扣费金额=("__amt", "sum"),
    daikou出现次数=("__id", "count"),
)

merged = hz_slim.merge(dk_agg, on="__id", how="outer", indicator=True)
merged["huizong_3月合计"] = merged["huizong_3月合计"].fillna(0).round(2)
merged["daikou_扣费金额"] = merged["daikou_扣费金额"].fillna(0).round(2)
merged["差额(daikou-huizong)"] = (merged["daikou_扣费金额"] - merged["huizong_3月合计"]).round(2)

def classify(r):
    if r["_merge"] == "left_only":
        return "daikou遗漏(huizong有)"
    if r["_merge"] == "right_only":
        return "daikou多余(huizong无)"
    return "一致" if abs(r["差额(daikou-huizong)"]) < 0.01 else "金额不符"

merged["核对状态"] = merged.apply(classify, axis=1)
merged = merged.drop(columns=["_merge"]).rename(columns={"__id": "学工号"})

diff = merged[merged["核对状态"] != "一致"]
summary = pd.DataFrame({
    "指标": [
        "huizong学工号数", "daikou学工号数(去重后)", "daikou总行数", "daikou重复工号行",
        "全集合数", "完全一致", "金额不符", "daikou遗漏", "daikou多余",
        "huizong_3月合计", "daikou_扣费金额合计", "差额合计",
    ],
    "值": [
        len(hz_slim), len(dk_agg), len(dk), len(dk_dup),
        len(merged),
        (merged["核对状态"] == "一致").sum(),
        (merged["核对状态"] == "金额不符").sum(),
        (merged["核对状态"] == "daikou遗漏(huizong有)").sum(),
        (merged["核对状态"] == "daikou多余(huizong无)").sum(),
        round(merged["huizong_3月合计"].sum(), 2),
        round(merged["daikou_扣费金额"].sum(), 2),
        round(merged["差额(daikou-huizong)"].sum(), 2),
    ],
})

save_excel(OUT, {"汇总": summary, "差异明细": diff, "全量对比": merged, "daikou重复工号": dk_dup})
print(summary.to_string(index=False))
print(f"[OK] 输出：{OUT}")
