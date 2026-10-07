import sys, io, os
sys.path.insert(0, os.getcwd())
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
import yaml
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
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
td = reader.get_transactional_data()

# Read the config sheet to get all P3 IDs
wb = load_workbook("data/templates/P3 Testing.xlsx", data_only=True)
config_ws = wb["Insert P3 HERE"]

p3_tab = {}
for r in range(5, 30):
    p3  = config_ws[f"B{r}"].value
    tab = config_ws[f"C{r}"].value
    if not p3:
        break
    p3_tab[str(p3).strip()] = str(tab).strip() if tab else str(p3).strip()

print(f"{'P3 ID':<20} {'Tab (truncated)':<32} {'POs in TD':<10} {'Months'}")
print("-" * 90)

for p3_id, tab_name in p3_tab.items():
    # Find all transactional data keys matching this P3 ID
    po_keys = [(k, v) for k, v in td.items()
               if isinstance(k, tuple) and k[1] == p3_id]
    # Also check plain-key matches (non-tuple)
    plain_keys = [(k, v) for k, v in td.items()
                  if not isinstance(k, tuple) and td.get((k, p3_id))]

    months_with_data = set()
    for k, v in po_keys:
        for m in v:
            if m not in ("cost_center", "wbs", "gross_ber_total"):
                months_with_data.add(m)

    tab_short = tab_name[:31]
    print(f"{p3_id:<20} {tab_short:<32} {len(po_keys):<10} {', '.join(sorted(months_with_data)) or '(none)'}")

# Now check what's in the output file for each tab
print()
print("Output file data rows per tab:")
print("-" * 60)
for p3_id, tab_name in p3_tab.items():
    tab_short = tab_name[:31]
    # sanitise same way as pipeline
    for ch in ('\\', '/', '?', '*', '[', ']', ':'):
        tab_short = tab_short.replace(ch, '-')
    tab_short = tab_short[:31].strip()

    if tab_short not in wb.sheetnames:
        print(f"  {tab_short:<32} -- TAB NOT FOUND in output")
        continue

    ws = wb[tab_short]
    data_rows = 0
    for r in range(13, 100):
        a = ws.cell(row=r, column=1).value
        if a and "Previous Period" in str(a):
            break
        b = ws.cell(row=r, column=2).value
        if b is not None:
            data_rows += 1
    # Count non-empty numeric cells in monthly cols (L=12 onwards)
    numeric_vals = 0
    for r in range(13, 50):
        for col in range(11, 75):
            v = ws.cell(row=r, column=col).value
            if isinstance(v, (int, float)) and v != 0:
                numeric_vals += 1
    print(f"  {tab_short:<32} PO rows: {data_rows:<4} numeric cells: {numeric_vals}")
