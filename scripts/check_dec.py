import sys, io, os, yaml
sys.path.insert(0, os.getcwd())
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
from src.transactional_detail_reader import TransactionalDetailReader

with open("configs/config_project.yaml") as f:
    cfg = yaml.safe_load(f)
t = cfg["transactional_detail_reader"]

reader = TransactionalDetailReader(
    file_path=t["file_path"], required_cols=t["required_cols"],
    valid_types=t["valid_types"], colmap=t["colmap"],
    shift_months=t.get("shift_months", True),
)
td = reader.get_transactional_data()

# Find any P3 IDs that have Dec or Dec (25) month buckets
print("P3 IDs with Dec / Dec (25) bucket:")
for k, v in td.items():
    if not isinstance(k, tuple):
        continue
    for m in v:
        if m not in ("cost_center", "wbs", "gross_ber_total") and "dec" in m.lower():
            print(f"  PO={k[0]}  P3={k[1]}  month='{m}'  data={v[m]}")

# Also check raw rows for Period 12 2025
df = reader.data
mc = t["colmap"]["month"]
p12 = df[df[mc].astype(str).str.contains("Period 12 2025", na=False)]
print(f"\nRaw Period 12 2025 rows: {len(p12)}")
if len(p12):
    show = [c for c in [t["colmap"]["po"], mc, t["colmap"]["type"],
                        "Vendor Invoice #", "CO Doc Line Item Txt",
                        t["colmap"]["cost_center"]] if c in df.columns]
    print(p12[show].head(10).to_string())
