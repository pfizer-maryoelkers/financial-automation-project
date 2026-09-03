"""project_main.py
Entry point for the CapEx / Project pipeline.

Mirrors main.py but uses:
  - ProjectTemplateReader  (groups by WBS root instead of Cost Center)
  - build_project_hierarchy (from src/project_utils.py)
  - configs/config_project.yaml

Usage:
    py project_main.py
"""

from src.utils import load_config, convert_base64
from src.forecast_reader import ForecastReader
from src.transactional_detail_reader import TransactionalDetailReader
from src.project_template_reader import ProjectTemplateReader
from src.template_writer import TemplateWriter
from src.project_utils import build_project_hierarchy
from src.models import ExceptionLog

config_path = 'configs/config_project.yaml'
config = load_config(config_path)

forecast_reader = ForecastReader(
    file_paths=config['forecast_reader']['file_paths'],
    po_col=config['forecast_reader']['po_col'],
)
transactional_reader = TransactionalDetailReader(
    file_path=config['transactional_detail_reader']['file_path'],
    required_cols=config['transactional_detail_reader']['required_cols'],
    valid_types=config['transactional_detail_reader']['valid_types'],
    colmap=config['transactional_detail_reader']['colmap'],
)

t = config['template']
template_reader = ProjectTemplateReader(
    file_path=t['file_path'],
    header_row=t['header_row'],
    po_col=t['po_col'],
    po_stop_marker=t['po_stop_marker'],
    wbs_col=t.get('wbs_col', 'A'),
    p3_id_col=t.get('p3_id_col', 'B'),
    wbs_start_row=t.get('wbs_start_row', 9),
)

tw = config['template_writer']
template_writer = TemplateWriter(
    file_path=t['file_path'],
    header_row=t['header_row'],
    po_column=t['po_col'],
    output_path=tw['output_path'],
    overwrite=tw['overwrite'],
    dec_acc_reversal_col=tw['dec_acc_reversal_col'],
    forecast_source_cols=tw['forecast_source_cols'],
    transactional_source_cols=tw['transactional_source_cols'],
    p3_id_column=t.get('p3_id_col'),
)


def main():
    print("============  PROJECT PIPELINE  ============")
    exception_log = ExceptionLog()

    # ── Step 1: Load data ─────────────────────────────────────────────────
    print("Step 1: Loading data\n")
    forecast_data   = forecast_reader.get_forecast_data()
    print("Loaded forecast data\n")

    transactional_data = transactional_reader.get_transactional_data()
    reclass_data       = transactional_reader.get_reclass_data()
    reclass_notes      = transactional_reader.get_reclass_notes()
    hierarchy_map      = transactional_reader.get_hierarchy_map()
    intl_po_set        = transactional_reader.get_intl_po_set()
    print("Loaded transactional data\n")

    # ── Step 2: Build project hierarchy ──────────────────────────────────
    print("Step 2: Building project hierarchy\n")
    assert transactional_reader.data is not None, "Transactional data should be loaded"

    hierarchy = build_project_hierarchy(
        projects=template_reader.projects,
        hierarchy_map=hierarchy_map,
        transactional_data=transactional_data,
        forecast_data=forecast_data,
        exception_log=exception_log,
        transactional_df=transactional_reader.data,
        p3_wbs_map=template_reader.p3_wbs_map,
        reclass_data=reclass_data,
        reclass_notes=reclass_notes,
        template_pos=template_reader.pos,
        intl_po_set=intl_po_set,
        p3_ids=config.get('template', {}).get('p3_ids'),
    )

    # ── Step 3: Write to template ─────────────────────────────────────────
    print("Step 3: Writing template output\n")
    pos = template_writer.insert_missing_po_rows(
        hierarchy,
        pos=template_reader.pos,
        blank_po_rows=template_reader.blank_po_rows,
    )
    template_writer.write_hierarchy(hierarchy, pos=pos)

    # ── Step 4: Exception reporting ───────────────────────────────────────
    print("Step 4: Writing exception reports\n")
    exception_log.summary()
    template_writer.write_exception_data_sheet(exception_log)
    template_writer.write_exception_sheet(exception_log, transactional_reader.data)
    template_writer.save()


if __name__ == "__main__":
    main()
