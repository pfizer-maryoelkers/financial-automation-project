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

cls_col  = t["colmap"]["classifier"]
type_col = t["colmap"]["type"]
co_col   = "CO Doc Line Item Txt"
po_col   = t["colmap"]["po"]
cc_col   = t["colmap"]["cost_center"]
amt_col  = "GL Transaction Amount"
month_col = t["colmap"]["month"]

df["_prefix"] = df[cls_col].astype(str).str.strip().str[:1]

# Show ALL transactions for each P3 ID that has data, with full classification info
target_p3s = ["P324-0013439", "P325-0016261", "P325-0016899", "P326-0017089", "P326-0017431"]
show = [c for c in [po_col, month_col, cls_col, type_col, co_col, amt_col] if c in df.columns]

for p3 in target_p3s:
    rows = df[df[cc_col].astype(str).str.strip() == p3]
    if rows.empty:
        continue
    print(f"\n{'='*70}")
    print(f"P3 {p3}  ({len(rows)} rows)")
    print(rows[show].to_string())
