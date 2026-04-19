"""任务1：3yuequanbu.XLSX 按学工号聚合
- 交易金额求和
- 交易时间取最新
- 每个学工号一条 → jieguo.xlsx
"""
import pandas as pd
from common import DATA_DIR, norm_id, to_num, save_excel

SRC = DATA_DIR / "3yuequanbu.XLSX"
OUT = DATA_DIR / "jieguo.xlsx"

# 表头在 row 7（0-based），数据从 row 8
df = pd.read_excel(SRC, header=7, engine="openpyxl")
df.columns = [str(c).strip() for c in df.columns]

# 精确识别列
col_id = next(c for c in df.columns if "学工号" in c)
col_amt = next(c for c in df.columns if "交易金额" in c)
col_time = next(c for c in df.columns if "交易时间" in c)
col_name = next(c for c in df.columns if "客户姓名" in c)
col_dept = next(c for c in df.columns if c == "部门")

# 清洗：去掉空学工号
df = df[df[col_id].notna()].copy()
df["__id"] = df[col_id].map(norm_id)
df["__amt"] = df[col_amt].map(to_num)
df["__time"] = pd.to_datetime(df[col_time], errors="coerce")
df = df[df["__id"] != ""].copy()

# 分组聚合：先取每组第一条作为姓名/部门基准（最新时间那条）
df_sorted = df.sort_values("__time")
agg = df_sorted.groupby("__id", as_index=False).agg(
    客户姓名=(col_name, "last"),
    部门=(col_dept, "last"),
    交易金额=("__amt", "sum"),
    交易时间=("__time", "max"),
    记录笔数=("__id", "count"),
)
agg = agg.rename(columns={"__id": "学工号"})
agg["学工号"] = agg["__id"] if "__id" in agg.columns else agg["学工号"]
agg = agg[["学工号", "客户姓名", "部门", "交易金额", "交易时间", "记录笔数"]]
agg["交易金额"] = agg["交易金额"].round(2)
agg = agg.sort_values("学工号").reset_index(drop=True)

save_excel(OUT, {"聚合结果": agg})
print(f"[OK] 源记录 {len(df)} 条 → 聚合后 {len(agg)} 个学工号")
print(f"[OK] 输出：{OUT}")
print(f"[汇总] 交易金额合计：{agg['交易金额'].sum():.2f}  |  记录笔数合计：{agg['记录笔数'].sum()}")
