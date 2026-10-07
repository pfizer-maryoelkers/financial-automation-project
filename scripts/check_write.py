import sys, io, os
sys.path.insert(0, os.getcwd())
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
import yaml
from src.transactional_detail_reader import TransactionalDetailReader
from src.project_template_reader import ProjectTemplateReader
from src.template_writer import TemplateWriter
from src.project_utils import build_project_hierarchy
from src.models import ExceptionLog
from src.forecast_reader import ForecastReader

with open("configs/config_project.yaml") as f:
    cfg = yaml.safe_load(f)

t   = cfg["template"]
ptw = cfg["template_writer"]

fr = ForecastReader(
    file_paths=cfg["forecast_reader"]["file_paths"],
    po_col=cfg["forecast_reader"]["po_col"],
)
forecast_data = fr.get_forecast_data()

tr = TransactionalDetailReader(
    file_path=cfg["transactional_detail_reader"]["file_path"],
    required_cols=cfg["transactional_detail_reader"]["required_cols"],
    valid_types=cfg["transactional_detail_reader"]["valid_types"],
    colmap=cfg["transactional_detail_reader"]["colmap"],
    shift_months=cfg["transactional_detail_reader"].get("shift_months", True),
)
td      = tr.get_transactional_data()
hmap    = tr.get_hierarchy_map()
reclass = tr.get_reclass_notes()
intl    = tr.get_intl_po_set()
df      = tr.data

template_reader = ProjectTemplateReader(
    file_path=t["file_path"],
    header_row=t["header_row"],
    po_col=t["po_col"],
    po_stop_marker=t.get("po_stop_marker", "Previous Period Invoices"),
    wbs_col=t.get("wbs_col", "A"),
    p3_id_col=t.get("p3_id_col", "B"),
    wbs_start_row=t.get("wbs_start_row", 2),
    template_sheet_name=t.get("template_sheet_name"),
)

p3_id   = "P325-0016899"
tab_name = "GL&NS- L-Seagen Products to be"

el = ExceptionLog()
hier = build_project_hierarchy(
    projects=template_reader.projects,
    hierarchy_map=hmap,
    transactional_data=td,
    forecast_data=forecast_data,
    exception_log=el,
    transactional_df=df,
    p3_wbs_map={p3_id: []},
    reclass_notes=reclass,
    template_pos={},
    intl_po_set=intl,
    p3_ids=[p3_id],
)

print(f"Hierarchy for {p3_id}:")
for p3k, p3v in hier.items():
    for wbs_code, wbs in p3v.wbs_codes.items():
        for po_num, po in wbs.pos.items():
            print(f"  PO in hierarchy: '{po_num}'  monthly: {list(po.monthly_data.keys())}")

# Now simulate what insert_missing_po_rows does
tw = TemplateWriter(
    file_path=t["file_path"],
    header_row=t["header_row"],
    po_column=t["po_col"],
    output_path=ptw["output_path"],
    overwrite=ptw["overwrite"],
    dec_acc_reversal_col=ptw["dec_acc_reversal_col"],
    forecast_source_cols=ptw["forecast_source_cols"],
    transactional_source_cols=ptw["transactional_source_cols"],
    p3_id_column=t.get("p3_id_col"),
    template_sheet_name=t.get("template_sheet_name"),
)
tw.clone_template_sheet(tab_name)
tw.sheet["B2"] = p3_id

pos = tw.insert_missing_po_rows(hier, pos={}, blank_po_rows=[], exception_log=el)
print(f"\npos dict after insert_missing_po_rows: {pos}")

# Check what write_hierarchy will look up
print("\nChecking write_hierarchy PO lookup:")
for p3k, p3v in hier.items():
    for wbs_code, wbs in p3v.wbs_codes.items():
        for po_num, po in wbs.pos.items():
            norm = tw._norm_po(po_num)
            found = norm in pos
            print(f"  PO '{po_num}' -> norm '{norm}' -> in pos: {found}  (pos keys: {list(pos.keys())})")
