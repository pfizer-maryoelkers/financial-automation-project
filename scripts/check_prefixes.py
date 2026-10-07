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
amt_col  = "GL Transaction Amount"

df["_prefix"] = df[cls_col].astype(str).str.strip().str[:1]

# Sample each unusual prefix
for prefix in ["1", "6", "8"]:
    rows = df[df["_prefix"] == prefix]
    print(f"\n=== Prefix '{prefix}' — {len(rows)} rows, type distribution: ===")
    print(rows[type_col].value_counts().to_string())
    show = [c for c in [po_col, cls_col, type_col, co_col, amt_col] if c in df.columns]
    print(rows[show].head(5).to_string())
