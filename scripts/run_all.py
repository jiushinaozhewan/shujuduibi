"""一键重跑全部 7 个任务"""
import subprocess, sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
tasks = [
    "task1_aggregate.py",
    "task2_xuesheng_check.py",
    "task3_jiaoshi_check.py",
    "task4_heji_check.py",
    "task5_daikou_cover.py",
    "task6_daikou_diff.py",
    "task7_tui_check.py",
]
for t in tasks:
    print(f"\n{'='*60}\n▶ {t}\n{'='*60}")
    r = subprocess.run([sys.executable, str(HERE / t)], cwd=HERE)
    if r.returncode != 0:
        sys.exit(f"[FAIL] {t}")
print("\n✅ 全部 7 个任务完成")
