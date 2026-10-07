import sys, io, os
sys.path.insert(0, os.getcwd())
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
import yaml
from src.transactional_detail_reader import TransactionalDetailReader

with open("configs/config_project.yaml") as f:
    cfg = yaml.safe_load(f)
t = cfg["transactional_detail_reader"]

print(f"shift_months from config: {t.get('shift_months', True)}")

reader = TransactionalDetailReader(
    file_path=t["file_path"],
    required_cols=t["required_cols"],
    valid_types=t["valid_types"],
    colmap=t["colmap"],
    shift_months=t.get("shift_months", True),   # explicitly pass it
)
print(f"reader.shift_months = {reader.shift_months}")

_ = reader.get_transactional_data()
td = reader.get_transactional_data()
matches = {k: v for k, v in td.items() if isinstance(k, tuple) and "9501182061" in str(k[0])}
print("\nAggregated month buckets for 9501182061:")
for k, v in matches.items():
    for m, vals in v.items():
        if m not in ("cost_center", "wbs", "gross_ber_total"):
            print(f"  {m}: {vals}")
