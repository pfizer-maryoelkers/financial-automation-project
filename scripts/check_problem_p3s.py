import sys, io, os
sys.path.insert(0, os.getcwd())
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
import yaml
from src.transactional_detail_reader import TransactionalDetailReader
from src.project_template_reader import ProjectTemplateReader
from src.template_writer import TemplateWriter
from src.project_utils import build_project_hierarchy
from src.models import ExceptionLog

with open("configs/config_project.yaml") as f:
    cfg = yaml.safe_load(f)

t   = cfg["template"]
ptw = cfg["template_writer"]

# Load data once
tr = TransactionalDetailReader(
    file_path=cfg["transactional_detail_reader"]["file_path"],
    required_cols=cfg["transactional_detail_reader"]["required_cols"],
    valid_types=cfg["transactional_detail_reader"]["valid_types"],
    colmap=cfg["transactional_detail_reader"]["colmap"],
    shift_months=cfg["transactional_detail_reader"].get("shift_months", True),
)
td         = tr.get_transactional_data()
hmap       = tr.get_hierarchy_map()
reclass    = tr.get_reclass_notes()
intl       = tr.get_intl_po_set()
df         = tr.data

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

# Examine a P3 ID that has POs but few numbers
problem_p3s = ["P325-0016899", "P326-0017089", "P326-0017431"]

for p3_id in problem_p3s:
    print(f"\n{'='*60}")
    print(f"P3 ID: {p3_id}")

    # Show transactional data keys
    po_data = [(k, v) for k, v in td.items()
               if isinstance(k, tuple) and k[1] == p3_id]
    print(f"  POs in transactional_data: {[k[0] for k, v in po_data]}")
    for k, v in po_data:
        for m, vals in v.items():
            if m not in ("cost_center", "wbs", "gross_ber_total"):
                print(f"    {k[0]} / {m}: {vals}")

    # Build hierarchy for this P3 only
    el = ExceptionLog()
    hier = build_project_hierarchy(
        projects=template_reader.projects,
        hierarchy_map=hmap,
        transactional_data=td,
        forecast_data={},
        exception_log=el,
        transactional_df=df,
        p3_wbs_map={p3_id: []},
        reclass_notes=reclass,
        template_pos={},
        intl_po_set=intl,
        p3_ids=[p3_id],
    )

    print(f"  Hierarchy keys: {list(hier.keys())}")
    for p3_key, p3_obj in hier.items():
        print(f"  P3 '{p3_key}' wbs_codes: {list(p3_obj.wbs_codes.keys())}")
        for wbs_code, wbs in p3_obj.wbs_codes.items():
            for po_num, po in wbs.pos.items():
                months = list(po.monthly_data.keys())
                print(f"    PO {po_num}: monthly_data months = {months}")
                for m, met in po.monthly_data.items():
                    print(f"      {m}: actual={met.actual} accrual={met.accrual} reversal={met.accrual_reversal}")
