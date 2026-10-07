"""project_main.py
Entry point for the CapEx / Project pipeline.

New layout (v2):
    The template workbook contains a sheet called "Enter All Your P3 IDs":
        B4  = label "Enter All Your P3 IDs"
        B5+ = P3 ID
        C5+ = tab name for that P3 ID

    For each P3 ID → tab pair the pipeline:
      1. Clones the master template sheet into a fresh tab named by column C.
      2. Runs build_project_hierarchy for that single P3 ID.
      3. Writes the hierarchy into the cloned tab.
      4. Writes exception data for that P3 ID.
    All tabs are saved together in one output workbook at the end.

Legacy layout (no "Enter All Your P3 IDs" sheet) still works as before.

Usage:
    py project_main.py
"""

import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

from src.utils import load_config
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
    shift_months=config['transactional_detail_reader'].get('shift_months', True),
)

t  = config['template']
tw = config['template_writer']

template_reader = ProjectTemplateReader(
    file_path=t['file_path'],
    header_row=t['header_row'],
    po_col=t['po_col'],
    po_stop_marker=t['po_stop_marker'],
    wbs_col=t.get('wbs_col', 'A'),
    p3_id_col=t.get('p3_id_col', 'B'),
    wbs_start_row=t.get('wbs_start_row', 9),
    template_sheet_name=t.get('template_sheet_name'),
)

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
    template_sheet_name=t.get('template_sheet_name'),
)


def main():
    print("============  PROJECT PIPELINE  ============")

    # ── Step 1: Load transactional + forecast data ────────────────────────
    print("Step 1: Loading data\n")
    forecast_data      = forecast_reader.get_forecast_data()
    print("Loaded forecast data\n")

    transactional_data = transactional_reader.get_transactional_data()
    reclass_notes      = transactional_reader.get_reclass_notes()
    hierarchy_map      = transactional_reader.get_hierarchy_map()
    intl_po_set        = transactional_reader.get_intl_po_set()
    print("Loaded transactional data\n")

    assert transactional_reader.data is not None, "Transactional data should be loaded"

    # ── Step 2: Determine P3 ID → tab mapping ────────────────────────────
    # New layout: config sheet gives explicit {p3_id: tab_name} pairs.
    # Legacy layout: fall back to single-sheet processing.
    p3_tab_map: dict[str, str] = template_reader.p3_tab_map

    if p3_tab_map:
        print(f"Step 2: Processing {len(p3_tab_map)} P3 ID(s) across separate tabs\n")
        all_exception_logs: list[tuple[str, ExceptionLog]] = []
        all_p3_ids: set[str] = set()

        for p3_id, tab_name in p3_tab_map.items():
            # Sanitise the tab name before use so all callers see the same name.
            tab_name = TemplateWriter._sanitise_tab_name(tab_name)
            print(f"  -- P3 ID: {p3_id}  ->  tab: '{tab_name}' --")

            # ── 2a: Clone a fresh template sheet for this P3 ID ──────────
            # clone_template_sheet deletes any existing tab with this name,
            # copies the master blank template, and switches the writer to it.
            template_writer.clone_template_sheet(tab_name)

            # Write the P3 ID into B2 on the cloned tab — this is the "corner"
            # cell that identifies the P3 ID for this sheet, replacing the old
            # workflow of typing it in manually.
            template_writer.sheet["B2"] = p3_id

            # ── 2b: Build hierarchy for this single P3 ID ────────────────
            exception_log = ExceptionLog()
            single_p3_map = {p3_id: []}   # p3_wbs_map for this P3 only

            hierarchy = build_project_hierarchy(
                projects=template_reader.projects,
                hierarchy_map=hierarchy_map,
                transactional_data=transactional_data,
                forecast_data=forecast_data,
                exception_log=exception_log,
                transactional_df=transactional_reader.data,
                p3_wbs_map=single_p3_map,
                reclass_notes=reclass_notes,
                template_pos={},   # cloned tab is always blank — no pre-existing POs
                intl_po_set=intl_po_set,
                p3_ids=[p3_id],
            )

            # ── 2c: Write hierarchy into the cloned tab ───────────────────
            pos = template_writer.insert_missing_po_rows(
                hierarchy,
                pos={},            # cloned tab is always blank — no pre-existing POs
                blank_po_rows=[],  # no blank placeholder rows on a fresh clone
                exception_log=exception_log,
            )
            template_writer.write_hierarchy(hierarchy, pos=pos)

            all_p3_ids.add(p3_id)
            all_exception_logs.append((p3_id, exception_log))

        # ── Step 3: Write combined exception reports ──────────────────────
        # Set _hierarchy_ids to ALL processed P3 IDs so the exception sheet
        # filters the transactional data correctly across every tab.
        template_writer._hierarchy_ids = all_p3_ids
        print("\nStep 3: Writing exception reports\n")
        # Merge all exception logs into one for the shared exception sheets
        merged_log = ExceptionLog()
        for _p3_id, el in all_exception_logs:
            for entry in el.entries:
                merged_log.entries.append(entry)
            merged_log._seen_keys.update(el._seen_keys)

        merged_log.summary()
        template_writer.write_exception_data_sheet(merged_log)
        template_writer.write_exception_sheet(merged_log, transactional_reader.data)

    else:
        # ── Legacy single-sheet path ──────────────────────────────────────
        print("Step 2: Building project hierarchy (legacy single-sheet mode)\n")
        exception_log = ExceptionLog()

        hierarchy = build_project_hierarchy(
            projects=template_reader.projects,
            hierarchy_map=hierarchy_map,
            transactional_data=transactional_data,
            forecast_data=forecast_data,
            exception_log=exception_log,
            transactional_df=transactional_reader.data,
            p3_wbs_map=template_reader.p3_wbs_map,
            reclass_notes=reclass_notes,
            template_pos=template_reader.pos,
            intl_po_set=intl_po_set,
            p3_ids=config.get('template', {}).get('p3_ids'),
        )

        print("Step 3: Writing template output\n")
        pos = template_writer.insert_missing_po_rows(
            hierarchy,
            pos=template_reader.pos,
            blank_po_rows=template_reader.blank_po_rows,
            exception_log=exception_log,
        )
        template_writer.write_hierarchy(hierarchy, pos=pos)

        _template_p3_ids = set(template_reader.p3_wbs_map.keys())
        if _template_p3_ids:
            template_writer._hierarchy_ids = _template_p3_ids

        print("Step 4: Writing exception reports\n")
        exception_log.summary()
        template_writer.write_exception_data_sheet(exception_log)
        template_writer.write_exception_sheet(exception_log, transactional_reader.data)

    template_writer.save()


if __name__ == "__main__":
    main()
