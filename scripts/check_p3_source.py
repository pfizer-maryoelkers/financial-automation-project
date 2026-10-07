import sys, io, os
sys.path.insert(0, os.getcwd())
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
import yaml
from src.transactional_detail_reader import TransactionalDetailReader

with open("configs/config_project.yaml") as f:
    cfg = yaml.safe_load(f)
t = cfg["transactional_detail_reader"]

reader = TransactionalDetailReader(
    file_path=t["file_path"],
    required_cols=t["required_cols"],
    valid_types=t["valid_types"],
    colmap=t["colmap"],
    shift_months=t.get("shift_months", True),
)
_ = reader.get_transactional_data()

df = reader.data
cc_col = t["colmap"]["cost_center"]  # "Cost Center*" / "P3 ID"

# Find all unique P3 IDs in the transactional file
all_p3s = df[cc_col].dropna().astype(str).str.strip().unique()
all_p3s = sorted(p for p in all_p3s if p.startswith("P") and "-" in p)
print(f"All P3 IDs in transactional file ({len(all_p3s)} total):")
for p in all_p3s:
    print(f"  {p}")

# Check the empty ones
empty = ["P325-0016238", "P325-0016303", "P326-0017401", "P326-0017427", "P326-0017038", "P326-0017865"]
print()
print("Checking empty P3 IDs against raw data:")
for p3 in empty:
    rows = df[df[cc_col].astype(str).str.strip() == p3]
    print(f"  {p3}: {len(rows)} rows in source file")
    if len(rows) > 0:
        print(f"    Sample: {rows.head(2)[cc_col].tolist()}")
