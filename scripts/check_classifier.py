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

cls_col  = t["colmap"]["classifier"]  # "Vendor Invoice #"
type_col = t["colmap"]["type"]        # "Type"
po_col   = t["colmap"]["po"]          # "PO Number"
co_col   = "CO Doc Line Item Txt"

print(f"Classifier column in data: '{cls_col}' — present: {cls_col in df.columns}")
print(f"Type col: '{type_col}' — present: {type_col in df.columns}")
print(f"CO Doc col present: {co_col in df.columns}")
print()

# Sample the classifier values for a few P3 IDs that have transactions
for p3 in ["P324-0013439", "P325-0016261", "P325-0016899"]:
    cc_col = t["colmap"]["cost_center"]
    rows = df[df[cc_col].astype(str).str.strip() == p3].head(10)
    if rows.empty:
        continue
    print(f"\n=== P3 {p3} — sample rows ===")
    show = [c for c in [po_col, cls_col, type_col, co_col, "GL Transaction Amount"] if c in df.columns]
    print(rows[show].to_string())

# Distribution of Type values overall
print("\n\nType value distribution:")
print(df[type_col].value_counts().to_string())

# Check classifier prefix distribution for each type
if cls_col in df.columns:
    print("\nClassifier prefix by type:")
    df["_prefix"] = df[cls_col].astype(str).str.strip().str[:1]
    print(df.groupby([type_col, "_prefix"]).size().to_string())
