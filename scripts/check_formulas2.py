import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula

wb = load_workbook("data/templates/P3 Testing.xlsx")
# Check tab 3 (first P3 tab that had rows inserted - GSC EMEA Dx)
# Tab order: 0=Insert P3 HERE, 1=Project Template, 2=GSC Market..., ...17=GSC EMEA Dx
tab = wb.sheetnames[17]
print(f"Checking tab: {tab}")
ws = wb[tab]

# Check all R9 array formulas
print("\nRow 9 formulas (P3 Actuals summary):")
for col in range(1, 75):
    v = ws.cell(row=9, column=col).value
    if v is not None:
        from openpyxl.utils import get_column_letter
        col_letter = get_column_letter(col)
        if isinstance(v, ArrayFormula):
            print(f"  {col_letter}9: ARRAY: {v.text[:100]}")
        else:
            print(f"  {col_letter}9: {repr(v)}")

print("\nRow 11 (BS total formula):", ws["BS11"].value)
print("Previous Period Invoices row:")
for r in range(13, 45):
    if ws.cell(row=r, column=1).value == "Previous Period Invoices":
        print(f"  Found at row {r}")
        break
