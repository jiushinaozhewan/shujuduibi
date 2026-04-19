"""任务3：核对 huizong 部门=教职工 的"3月缴费" vs jiaoshi.xlsx 男教师/女教师的"总计"
以姓名为主索引：
- 合并 jiaoshi 两个 sheet 取 col1=姓名, col78=总计
- 对比 huizong 中客户部门=="教职工" 的 3月缴费
- 差异 → jieguo2.xlsx
"""
import pandas as pd
from common import DATA_DIR, to_num, save_excel

HUIZONG = DATA_DIR / "huizong.XLSX"
JIAOSHI = DATA_DIR / "jiaoshi.xlsx"
OUT = DATA_DIR / "jieguo2.xlsx"

# huizong 主表
hz = pd.read_excel(HUIZONG, sheet_name="Sheet", engine="openpyxl")
hz.columns = [str(c).strip() for c in hz.columns]
hz_js = hz[hz["客户部门"].astype(str).str.contains("教职工", na=False)].copy()
hz_js["__name"] = hz_js["姓名"].astype(str).str.strip()
hz_js["__huizong_3缴"] = hz_js["3月缴费"].map(to_num)

# jiaoshi：无表头读取，取 col0=序号, col1=姓名, col2=性别, col78=总计
js_all = []
for s in ["男教师", "女教师"]:
    raw = pd.read_excel(JIAOSHI, sheet_name=s, header=None, engine="openpyxl")
    # 数据从 row 2 开始
    sub = raw.iloc[2:, [1, 2, 78]].copy()
    sub.columns = ["姓名", "性别", "总计"]
    sub["__sheet"] = s
    # 去除空行
    sub = sub[sub["姓名"].notna()]
    sub["__name"] = sub["姓名"].astype(str).str.strip()
    sub["__zongji"] = sub["总计"].map(to_num)
    js_all.append(sub)
js = pd.concat(js_all, ignore_index=True)

# 按姓名聚合 jiaoshi（应是一一对应，但可能同名或重复）
js_agg = js.groupby("__name", as_index=False).agg(
    jiaoshi总计=("__zongji", "sum"),
    来源sheet=("__sheet", lambda s: ",".join(sorted(set(s)))),
    jiaoshi记录数=("__name", "count"),
)

# 对比
merged = hz_js[["__name", "客户部门", "工号", "姓名", "__huizong_3缴"]].merge(
    js_agg, on="__name", how="outer", indicator=True
)
merged = merged.rename(columns={"__huizong_3缴": "huizong_3月缴费", "__name": "姓名(索引)"})
merged["huizong_3月缴费"] = merged["huizong_3月缴费"].fillna(0).round(2)
merged["jiaoshi总计"] = merged["jiaoshi总计"].fillna(0).round(2)
merged["差额(huizong-jiaoshi)"] = (merged["huizong_3月缴费"] - merged["jiaoshi总计"]).round(2)

def classify(r):
    if r["_merge"] == "left_only":
        return "仅huizong教职工有"
    if r["_merge"] == "right_only":
        return "仅jiaoshi有(huizong遗漏)"
    return "一致" if abs(r["差额(huizong-jiaoshi)"]) < 0.01 else "金额不一致"

merged["核对状态"] = merged.apply(classify, axis=1)
merged = merged.drop(columns=["_merge"])

diff = merged[merged["核对状态"] != "一致"]
summary = pd.DataFrame({
    "指标": ["教职工姓名(合集)", "完全一致", "金额不一致", "仅huizong有", "仅jiaoshi有",
             "huizong教职工3月缴费合计", "jiaoshi总计合计", "差额合计"],
    "值": [
        len(merged),
        (merged["核对状态"] == "一致").sum(),
        (merged["核对状态"] == "金额不一致").sum(),
        (merged["核对状态"] == "仅huizong教职工有").sum(),
        (merged["核对状态"] == "仅jiaoshi有(huizong遗漏)").sum(),
        round(merged["huizong_3月缴费"].sum(), 2),
        round(merged["jiaoshi总计"].sum(), 2),
        round(merged["差额(huizong-jiaoshi)"].sum(), 2),
    ],
})

save_excel(OUT, {"汇总": summary, "差异明细": diff, "全量对比": merged})
print(f"[OK] huizong教职工 {len(hz_js)} 人 | jiaoshi 聚合 {len(js_agg)} 人")
print(summary.to_string(index=False))
print(f"[OK] 输出：{OUT}")
