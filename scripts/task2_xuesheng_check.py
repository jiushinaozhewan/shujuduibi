"""任务2：核对 huizong.XLSX 的"3月退费" vs xuesheng.xlsx 三个sheet的退费金额
以工号为主索引：
- 汇总 xuesheng 三个 sheet 中每个工号的退费金额之和（同一人可能在多行/多sheet出现）
- 对比 huizong 的 3月退费
- 差异/遗漏/多出 → jieguo1.xlsx
"""
import pandas as pd
from common import DATA_DIR, norm_id, to_num, save_excel

HUIZONG = DATA_DIR / "huizong.XLSX"
XUESHENG = DATA_DIR / "xuesheng.xlsx"
OUT = DATA_DIR / "jieguo1.xlsx"

# 读 huizong 主表，表头在 row 0
hz = pd.read_excel(HUIZONG, sheet_name="Sheet", engine="openpyxl")
hz.columns = [str(c).strip() for c in hz.columns]
hz = hz[hz["工号"].notna()].copy()
hz["__id"] = hz["工号"].map(norm_id)
hz["__huizong_3退"] = hz["3月退费"].map(to_num)

# 读 xuesheng 3个 sheet，表头在 row 2
xs_all = []
for s in ["3", "2", "1"]:
    df = pd.read_excel(XUESHENG, sheet_name=s, header=2, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]
    df["__sheet"] = s
    xs_all.append(df)
xs = pd.concat(xs_all, ignore_index=True)
xs = xs[xs["工号"].notna()].copy()
xs["__id"] = xs["工号"].map(norm_id)
xs["__amt"] = xs["退费金额"].map(to_num)

# 按工号聚合 xuesheng
xs_agg = xs.groupby("__id", as_index=False).agg(
    xuesheng退费合计=("__amt", "sum"),
    出现sheet=("__sheet", lambda s: ",".join(sorted(set(s)))),
    xuesheng记录数=("__id", "count"),
)

# 对比
merged = hz[["__id", "客户部门", "工号", "姓名", "__huizong_3退"]].merge(
    xs_agg, on="__id", how="outer", indicator=True
)
merged = merged.rename(columns={"__huizong_3退": "huizong_3月退费"})
merged["huizong_3月退费"] = merged["huizong_3月退费"].fillna(0).round(2)
merged["xuesheng退费合计"] = merged["xuesheng退费合计"].fillna(0).round(2)
merged["差额(huizong-xuesheng)"] = (merged["huizong_3月退费"] - merged["xuesheng退费合计"]).round(2)

def classify(r):
    if r["_merge"] == "left_only":
        return "仅huizong有" if r["huizong_3月退费"] != 0 else "huizong无退费"
    if r["_merge"] == "right_only":
        return "仅xuesheng有(huizong遗漏)"
    return "一致" if abs(r["差额(huizong-xuesheng)"]) < 0.01 else "金额不一致"

merged["核对状态"] = merged.apply(classify, axis=1)
merged = merged.drop(columns=["_merge"])
merged = merged.rename(columns={"__id": "工号(规范化)"})

# 拆分三张表：全量 + 差异 + 汇总
diff = merged[merged["核对状态"].isin(["金额不一致", "仅huizong有", "仅xuesheng有(huizong遗漏)"])]
summary = pd.DataFrame({
    "指标": ["总工号数(合集)", "完全一致", "金额不一致", "仅huizong有(且金额≠0)", "仅xuesheng有(遗漏)",
             "huizong 3月退费合计", "xuesheng 退费合计", "差额合计"],
    "值": [
        len(merged),
        (merged["核对状态"] == "一致").sum(),
        (merged["核对状态"] == "金额不一致").sum(),
        (merged["核对状态"] == "仅huizong有").sum(),
        (merged["核对状态"] == "仅xuesheng有(huizong遗漏)").sum(),
        round(merged["huizong_3月退费"].sum(), 2),
        round(merged["xuesheng退费合计"].sum(), 2),
        round(merged["差额(huizong-xuesheng)"].sum(), 2),
    ],
})

save_excel(OUT, {"汇总": summary, "差异明细": diff, "全量对比": merged})
print(f"[OK] huizong 学生行 {len(hz)} 条 | xuesheng 汇总 {len(xs_agg)} 个工号")
print(summary.to_string(index=False))
print(f"[OK] 输出：{OUT}")
