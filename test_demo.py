
from pathlib import Path
from reporting import load_history, calculate_kpis, deterministic_summary, generate_powerpoint

base = Path(__file__).resolve().parent
history, status = load_history(base / "data")
latest = history["Week Ending"].max()
kpis = calculate_kpis(history, latest)
summary = deterministic_summary(kpis)
ppt = generate_powerpoint(kpis, summary)

print(status[["file", "valid", "week_ending", "rows"]].to_string(index=False))
print()
print(summary)
print()
print(f"Generated PPTX bytes: {len(ppt):,}")
