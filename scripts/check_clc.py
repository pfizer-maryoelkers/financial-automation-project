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
cc_col = t["colmap"]["cost_center"]
po_col = t["colmap"]["po"]
month_col = t["colmap"]["month"]
type_col  = t["colmap"]["type"]
amt_col   = "GL BER Corp Amount"

p3 = "P326-0017038"
rows = df[df[cc_col].astype(str).str.strip() == p3]
print(f"Raw rows for {p3} ({len(rows)} rows):")
show = [c for c in [po_col, month_col, type_col, amt_col, cc_col] if c in df.columns]
print(rows[show].to_string())

# Also check transactional_data for this P3
td = reader.get_transactional_data()
matches = [(k, v) for k, v in td.items() if isinstance(k, tuple) and k[1] == p3]
print(f"\ntransactional_data keys for {p3}: {[k for k, v in matches]}")

# And hierarchy_map
hmap = reader.get_hierarchy_map()
hmap_rows = [(i, r) for i, r in hmap.items() if str(r.get("cost_center", "")).strip() == p3]
print(f"\nhierarchy_map rows for {p3}: {len(hmap_rows)}")
for i, r in hmap_rows[:5]:
    print(f"  [{i}] po={r.get('po')} wbs={r.get('wbs')} type via df: {df.loc[i, type_col] if i in df.index else '?'}")
