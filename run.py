"""run.py
Unified entry point for the financial automation pipeline.

Automatically detects whether the configured template is an OpEx (Cost-Center)
template or a CapEx/Project (P3-ID / WBS) template by inspecting the top
corner of the workbook, then runs the appropriate pipeline.

Usage
-----
    py run.py                          # uses configs/config_base.yaml  (OpEx)
    py run.py configs/config_project.yaml   # uses project config explicitly

The config file to use can also be changed by editing CONFIG_PATH below.
If no argument is passed the script looks for configs/config_base.yaml.
"""

import sys
from src.utils import load_config, convert_base64, detect_template_type
from src.models import ExceptionLog


def _run_opex(config: dict) -> None:
    """Run the OpEx (Cost-Center) pipeline."""
    from src.forecast_reader import ForecastReader
    from src.transactional_detail_reader import TransactionalDetailReader
    from src.template_reader import TemplateReader
    from src.template_writer import TemplateWriter
    from src.utils import build_hierarchy

    t  = config["template"]
    tw = config["template_writer"]

    forecast_reader = ForecastReader(
        file_paths=config["forecast_reader"]["file_paths"],
        po_col=config["forecast_reader"]["po_col"],
    )
    transactional_reader = TransactionalDetailReader(
        file_path=config["transactional_detail_reader"]["file_path"],
        required_cols=config["transactional_detail_reader"]["required_cols"],
        valid_types=config["transactional_detail_reader"]["valid_types"],
        colmap=config["transactional_detail_reader"]["colmap"],
    )
    template_reader = TemplateReader(
        file_path=t["file_path"],
        header_row=t["header_row"],
        po_col=t["po_col"],
        po_stop_marker=t["po_stop_marker"],
        cost_center_col=t["cost_center_col"],
        cost_center_start_row=t["cost_center_start_row"],
    )
    template_writer = TemplateWriter(
        file_path=t["file_path"],
        header_row=t["header_row"],
        po_column=t["po_col"],
        output_path=tw["output_path"],
        overwrite=tw["overwrite"],
        dec_acc_reversal_col=tw["dec_acc_reversal_col"],
        forecast_source_cols=tw["forecast_source_cols"],
        transactional_source_cols=tw["transactional_source_cols"],
    )

    print("============  OPEX PIPELINE  ============")
    exception_log = ExceptionLog()

    print("Step 1: Loading data\n")
    forecast_data      = forecast_reader.get_forecast_data()
    transactional_data = transactional_reader.get_transactional_data()
    reclass_data       = transactional_reader.get_reclass_data()
    reclass_notes      = transactional_reader.get_reclass_notes()
    hierarchy_map      = transactional_reader.get_hierarchy_map()
    intl_po_set        = transactional_reader.get_intl_po_set()
    print("Loaded data\n")

    print("Step 2: Building hierarchy\n")
    assert transactional_reader.data is not None
    hierarchy = build_hierarchy(
        cost_centers=template_reader.cost_centers,
        hierarchy_map=hierarchy_map,
        transactional_data=transactional_data,
        forecast_data=forecast_data,
        exception_log=exception_log,
        transactional_df=transactional_reader.data,
        reclass_data=reclass_data,
        reclass_notes=reclass_notes,
        template_pos=template_reader.pos,
        intl_po_set=intl_po_set,
    )

    print("Step 3: Writing template output\n")
    pos = template_writer.insert_missing_po_rows(hierarchy, pos=template_reader.pos)
    template_writer.write_hierarchy(hierarchy, pos=pos)
    template_writer.write_forecast_source_sheet(forecast_reader.data, pos=pos)
    template_writer.write_transactional_source_sheet(transactional_reader.data, pos=pos)

    print("Step 4: Writing exception reports\n")
    exception_log.summary()
    template_writer.write_exception_data_sheet(exception_log)
    template_writer.write_exception_sheet(exception_log, transactional_reader.data, pos=pos)
    template_writer.write_exception_summary_sheet(exception_log)
    template_writer.save()


def _run_project(config: dict) -> None:
    """Run the CapEx / Project (P3-ID / WBS) pipeline."""
    from src.forecast_reader import ForecastReader
    from src.transactional_detail_reader import TransactionalDetailReader
    from src.project_template_reader import ProjectTemplateReader
    from src.template_writer import TemplateWriter
    from src.project_utils import build_project_hierarchy

    t  = config["template"]
    tw = config["template_writer"]

    forecast_reader = ForecastReader(
        file_paths=config["forecast_reader"]["file_paths"],
        po_col=config["forecast_reader"]["po_col"],
    )
    transactional_reader = TransactionalDetailReader(
        file_path=config["transactional_detail_reader"]["file_path"],
        required_cols=config["transactional_detail_reader"]["required_cols"],
        valid_types=config["transactional_detail_reader"]["valid_types"],
        colmap=config["transactional_detail_reader"]["colmap"],
    )
    template_reader = ProjectTemplateReader(
        file_path=t["file_path"],
        header_row=t["header_row"],
        po_col=t["po_col"],
        po_stop_marker=t["po_stop_marker"],
        wbs_col=t.get("wbs_col", "A"),
        p3_id_col=t.get("p3_id_col", "B"),
        wbs_start_row=t.get("wbs_start_row", 2),
    )
    template_writer = TemplateWriter(
        file_path=t["file_path"],
        header_row=t["header_row"],
        po_column=t["po_col"],
        output_path=tw["output_path"],
        overwrite=tw["overwrite"],
        dec_acc_reversal_col=tw["dec_acc_reversal_col"],
        forecast_source_cols=tw["forecast_source_cols"],
        transactional_source_cols=tw["transactional_source_cols"],
    )

    print("============  PROJECT PIPELINE  ============")
    exception_log = ExceptionLog()

    print("Step 1: Loading data\n")
    forecast_data      = forecast_reader.get_forecast_data()
    transactional_data = transactional_reader.get_transactional_data()
    reclass_data       = transactional_reader.get_reclass_data()
    reclass_notes      = transactional_reader.get_reclass_notes()
    hierarchy_map      = transactional_reader.get_hierarchy_map()
    intl_po_set        = transactional_reader.get_intl_po_set()
    print("Loaded data\n")

    print("Step 2: Building project hierarchy\n")
    assert transactional_reader.data is not None
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
    )

    print("Step 3: Writing template output\n")
    pos = template_writer.insert_missing_po_rows(hierarchy, pos=template_reader.pos)
    template_writer.write_hierarchy(hierarchy, pos=pos)
    template_writer.write_forecast_source_sheet(forecast_reader.data, pos=pos)
    template_writer.write_transactional_source_sheet(transactional_reader.data, pos=pos)

    print("Step 4: Writing exception reports\n")
    exception_log.summary()
    template_writer.write_exception_data_sheet(exception_log)
    template_writer.write_exception_sheet(exception_log, transactional_reader.data, pos=pos)
    template_writer.write_exception_summary_sheet(exception_log)
    template_writer.save()


def main() -> None:
    config_path = sys.argv[1] if len(sys.argv) > 1 else "configs/config_base.yaml"
    config = load_config(config_path)

    template_file = config["template"]["file_path"]
    template_type = detect_template_type(template_file)
    print(f"Detected template type: {template_type.upper()}  ({template_file})\n")

    if template_type == "project":
        _run_project(config)
    else:
        _run_opex(config)


if __name__ == "__main__":
    main()
