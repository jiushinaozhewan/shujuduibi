"""通用工具：路径、工号规范化、写 Excel 带样式"""
from pathlib import Path
import pandas as pd

DATA_DIR = Path(__file__).resolve().parent.parent


def norm_id(v):
    """把工号统一成字符串：纯数字的浮点/整数都转成整数字符串；NaN → ''"""
    if pd.isna(v):
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def to_num(v):
    """把金额统一成浮点，NaN/None/空字符串返回 0.0"""
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


def save_excel(path: Path, sheets: dict):
    """写入多sheet，冻结首行并自动列宽"""
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        for name, df in sheets.items():
            df.to_excel(w, sheet_name=name, index=False)
        from openpyxl.utils import get_column_letter
        for name, df in sheets.items():
            ws = w.sheets[name]
            ws.freeze_panes = "A2"
            for i, col in enumerate(df.columns, 1):
                try:
                    max_len = max(
                        [len(str(col))] + [len(str(x)) for x in df[col].head(200).tolist()]
                    )
                except Exception:
                    max_len = 12
                ws.column_dimensions[get_column_letter(i)].width = min(max(max_len + 2, 8), 40)
