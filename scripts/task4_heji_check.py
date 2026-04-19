"""任务4：核对 huizong 的"3月合计" 是否等于 "3月缴费"-"3月退费"
全量逐行核对 → jieguo3.xlsx
"""
import pandas as pd
from common import DATA_DIR, to_num, save_excel

HUIZONG = DATA_DIR / "huizong.XLSX"
OUT = DATA_DIR / "jieguo3.xlsx"

hz = pd.read_excel(HUIZONG, sheet_name="Sheet", engine="openpyxl")
hz.columns = [str(c).strip() for c in hz.columns]
hz = hz[hz["工号"].notna()].copy()

hz["3月缴费_数值"] = hz["3月缴费"].map(to_num)
hz["3月退费_数值"] = hz["3月退费"].map(to_num)
hz["3月合计_数值"] = hz["3月合计"].map(to_num)
hz["应为"] = (hz["3月缴费_数值"] - hz["3月退费_数值"]).round(2)
hz["差额(实际-应为)"] = (hz["3月合计_数值"] - hz["应为"]).round(2)
hz["核对状态"] = hz["差额(实际-应为)"].abs().apply(lambda x: "一致" if x < 0.01 else "不一致")

out = hz[["客户部门", "工号", "姓名", "3月缴费_数值", "3月退费_数值", "3月合计_数值", "应为", "差额(实际-应为)", "核对状态"]]
out = out.rename(columns={"3月缴费_数值": "3月缴费", "3月退费_数值": "3月退费", "3月合计_数值": "3月合计"})

diff = out[out["核对状态"] == "不一致"]
summary = pd.DataFrame({
    "指标": ["总行数", "一致", "不一致", "3月缴费合计", "3月退费合计", "3月合计(原)合计", "应为合计", "差额合计"],
    "值": [
        len(out), (out["核对状态"] == "一致").sum(), (out["核对状态"] == "不一致").sum(),
        round(out["3月缴费"].sum(), 2), round(out["3月退费"].sum(), 2),
        round(out["3月合计"].sum(), 2), round(out["应为"].sum(), 2),
        round(out["差额(实际-应为)"].sum(), 2),
    ],
})

save_excel(OUT, {"汇总": summary, "差异明细": diff, "全量对比": out})
print(summary.to_string(index=False))
print(f"[OK] 输出：{OUT}")
