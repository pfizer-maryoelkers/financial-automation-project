"""project_utils.py
Build a project hierarchy from the transactional detail file.

The project pipeline is structurally identical to the OpEx pipeline but uses
the *WBS root code* (e.g. 'CE-BTS21076') as the top-level grouping key
instead of Cost Center.

A WBS root is derived from the full WBS element by keeping only the first two
dash-separated parts:
    CE-BTS21076            → CE-BTS21076  (already a root)
    CE-BTS21076-02-10      → CE-BTS21076
    CE-BTS21076-02-EX      → CE-BTS21076
    CE-BTS21076-02-EX-IE   → CE-BTS21076
"""

import re
import pandas as pd
from collections import defaultdict

from src.models import Project, WBSCode, PO, MonthlyMetrics, ExceptionLog, ExceptionType, P3ID
from src.project_template_reader import extract_project_root, wbs_charge_type

# Forecast month shift: maps a 3-letter month key one step back (same as OpEx).
_FORECAST_MONTH_SHIFT = {
    "Jan": "Dec", "Feb": "Jan", "Mar": "Feb", "Apr": "Mar",
    "May": "Apr", "Jun": "May", "Jul": "Jun", "Aug": "Jul",
    "Sep": "Aug", "Oct": "Sep", "Nov": "Oct", "Dec": "Nov",
}

_KNOWN_TYPES = {"Actual", "Accrual", "Reversal", "Reclass", "ER"}


def build_project_hierarchy(
    projects: list[str],
    hierarchy_map: dict,
    transactional_data: dict,
    forecast_data: dict,
    exception_log: ExceptionLog,
    transactional_df: pd.DataFrame,
    p3_wbs_map: dict[str, list[str]],
    reclass_data: dict | None = None,
    reclass_notes: dict | None = None,
    template_pos: dict | None = None,
    intl_po_set: set | None = None,
) -> dict[str, P3ID]:
    """Build the project hierarchy.

    Parameters
    ----------
    projects : list[str]
        Project root WBS codes read from the project template
        (e.g. ['CE-BTS21076', 'CE-BTS22001']).
    hierarchy_map : dict
        Row-indexed map from TransactionalDetailReader.get_hierarchy_map().
    transactional_data : dict
        Aggregated PO data from TransactionalDetailReader.get_transactional_data().
    forecast_data : dict
        PO forecast data from ForecastReader.get_forecast_data().
    exception_log : ExceptionLog
        Receives exception entries.
    transactional_df : pd.DataFrame
        Raw transactional DataFrame (used to read source_row_data per entry).
    p3_wbs_map : dict[str, list[str]]
        P3 ID to WBS mapping from the project template.
    template_pos : dict | None
        PO → row mapping from the project template (None = blank template).
    intl_po_set : set | None
        Set of international PO numbers (forecast shifted back one month).

    Returns
    -------
    dict[str, P3ID]  keyed by P3 ID.
    """
    er_pattern = re.compile(r'\bER\d+\b', re.IGNORECASE)

    def _find_er_in_row(row: dict) -> str | None:
        for key in ('gl_line_desc', 'gl_trans_desc', 'description'):
            text = row.get(key)
            if text:
                m = er_pattern.search(text)
                if m:
                    return m.group(0).upper()
        return None

    # Build a set of known P3 IDs for direct cost_center matching
    _known_p3_ids: set[str] = set(p3_wbs_map.keys())

    def _find_p3_id_for_wbs(wbs: str) -> str | None:
        if not wbs:
            return None
        root = extract_project_root(wbs)
        for p3_id, wbs_codes in p3_wbs_map.items():
            for template_wbs in wbs_codes:
                if extract_project_root(template_wbs) == root:
                    return p3_id
        return None

    # ── Step 1: Pre-group rows by P3 ID ─────────────────────────────────────
    # Primary match: WBS root → P3 ID via template mapping.
    # Fallback: when WBS is absent or unmatched, use the cost_center field
    # directly — this handles P3 IDs that have no WBS listed in the template.
    rows_by_p3: dict[str, list] = defaultdict(list)
    for row_idx, row in hierarchy_map.items():
        wbs         = row.get('wbs')
        cost_center = row.get('cost_center')
        p3_id = _find_p3_id_for_wbs(wbs) if wbs else None
        # Fallback: match cost_center directly when it is a known P3 ID
        if p3_id is None and cost_center and cost_center in _known_p3_ids:
            p3_id = cost_center
        if p3_id:
            rows_by_p3[p3_id].append((row_idx, row))

    # ── Step 2: Pre-scan for duplicate WBS codes ────────────────────────────
    wbs_to_p3s: dict[str, set] = defaultdict(set)
    er_real_wbs: dict[str, str] = {}
    for p3_id in p3_wbs_map.keys():
        for row_idx, row in rows_by_p3[p3_id]:
            wbs = row.get('wbs')
            po = row.get('po')
            if wbs:
                wbs_to_p3s[wbs].add(p3_id)
                if po and er_pattern.match(po) and po not in er_real_wbs:
                    er_real_wbs[po] = wbs
    duplicate_wbs_codes = {w for w, ps in wbs_to_p3s.items() if len(ps) > 1}

    # ── Step 3: Cross-P3 tracking (exception reporting only) ────────────────
    er_extracted_count = 0

    # ── Step 4: Build hierarchy ─────────────────────────────────────────────
    result: dict[str, P3ID] = {}
    for p3_id in p3_wbs_map.keys():
        p3_obj = P3ID(p3_id=p3_id)
        # Scope these per P3 ID so POs/WBS codes from different P3 IDs never
        # block each other — critical when rows are matched via cost_center
        # fallback (P3-ID-only template rows with no WBS listed).
        seen_pos: dict[str, tuple] = {}   # po → (p3_id, wbs) — scoped per P3 ID
        seen_wbs: dict[str, str]   = {}   # wbs → p3_id       — scoped per P3 ID
        for row_idx, row in rows_by_p3[p3_id]:
            po          = row.get('po')
            wbs         = row.get('wbs')
            legal_entity = row.get('legal_entity')
            country      = row.get('country')
            vendor_name  = row.get('vendor_name')
            gl_account   = row.get('gl_account')
            req_title    = row.get('req_title')

            source_row_data = transactional_df.loc[row_idx].to_dict() if row_idx in transactional_df.index else {}
            month      = source_row_data.get('Accounting Period')
            amount     = source_row_data.get('GL BER Corp Amount')
            trans_type = source_row_data.get('Type')

            # ── Exception: unrecognised transaction type ────────────────────
            if trans_type not in _KNOWN_TYPES:
                if template_pos and po and po in template_pos:
                    exception_log.log(
                        ExceptionType.UNMATCHED_TRANSACTION,
                        row_index=row_idx, po=po, wbs=wbs,
                        cost_center=p3_id, month=month, amount=amount,
                        transaction_type=trans_type, source_row_data=source_row_data,
                    )
                continue

            # ── Exception: Reclass ──────────────────────────────────────────
            if trans_type == "Reclass":
                exception_log.log(
                    ExceptionType.RECLASS,
                    row_index=row_idx, po=po, wbs=wbs,
                    cost_center=p3_id, month=month, amount=amount,
                    transaction_type=trans_type, source_row_data=source_row_data,
                )
                continue

            # ── ER extraction (both-missing and PO-only paths) ──────────────
            if not wbs and not po:
                er_found = _find_er_in_row(row)
                if er_found:
                    po, wbs = er_found, "ER"
                    er_extracted_count += 1
                else:
                    if template_pos:
                        exception_log.log(
                            ExceptionType.MISSING_WBS_AND_PO,
                            row_index=row_idx, cost_center=p3_id,
                            month=month, amount=amount,
                            transaction_type=trans_type, source_row_data=source_row_data,
                        )
                    continue

            if po and not wbs:
                er_found = _find_er_in_row(row)
                if not er_found:
                    m = er_pattern.match(po)
                    if m:
                        er_found = m.group(0).upper()
                if er_found:
                    po, wbs = er_found, "ER"
                    er_extracted_count += 1

            if not wbs:
                if template_pos:
                    exception_log.log(
                        ExceptionType.MISSING_WBS,
                        row_index=row_idx, po=po, cost_center=p3_id,
                        month=month, amount=amount,
                        transaction_type=trans_type, source_row_data=source_row_data,
                    )
                wbs = "NO_WBS"

            if not po:
                if template_pos:
                    exception_log.log(
                        ExceptionType.MISSING_PO,
                        row_index=row_idx, wbs=wbs, cost_center=p3_id,
                        month=month, amount=amount,
                        transaction_type=trans_type, source_row_data=source_row_data,
                    )
                continue

            # ── Duplicate WBS ───────────────────────────────────────────────
            if wbs in duplicate_wbs_codes:
                exception_log.log(
                    ExceptionType.DUPLICATE_WBS,
                    row_index=row_idx, wbs=wbs, po=po, cost_center=p3_id,
                    month=month, amount=amount,
                    transaction_type=trans_type, source_row_data=source_row_data,
                )
                if wbs in seen_wbs and seen_wbs[wbs] != p3_id:
                    continue
            seen_wbs[wbs] = p3_id

            # ── Duplicate PO ────────────────────────────────────────────────
            if po in seen_pos:
                prev_p3, prev_wbs = seen_pos[po]
                if prev_p3 != p3_id or prev_wbs != wbs:
                    exception_log.log(
                        ExceptionType.DUPLICATE_PO,
                        row_index=row_idx, po=po, wbs=wbs, cost_center=p3_id,
                        month=month, amount=amount,
                        transaction_type=trans_type, source_row_data=source_row_data,
                    )
                continue
            seen_pos[po] = (p3_id, wbs)

            # ── PO not on template ──────────────────────────────────────────
            # Log every transaction row so the exceptions tab shows per-month
            # amounts and vendor details for every missing PO.
            if template_pos and po not in template_pos and wbs != "ER":
                exception_log.log(
                    ExceptionType.PO_NOT_ON_TEMPLATE,
                    row_index=row_idx, po=po, wbs=wbs, cost_center=p3_id,
                    month=month, amount=amount,
                    transaction_type=trans_type, vendor_name=vendor_name,
                    source_row_data=source_row_data,
                )

            # ── Build Project → WBS → PO objects ───────────────────────────
            if wbs not in p3_obj.wbs_codes:
                p3_obj.wbs_codes[wbs] = WBSCode(
                    wbs_code=wbs,
                    cost_center=p3_id,
                    charge_type=wbs_charge_type(wbs),
                )
            if po not in p3_obj.wbs_codes[wbs].pos:
                p3_obj.wbs_codes[wbs].pos[po] = PO(
                    po_number=po,
                    legal_entity=legal_entity,
                    country=country,
                    vendor_name=vendor_name,
                    gl_account=gl_account,
                    req_title=req_title,
                    real_wbs=er_real_wbs.get(po) if wbs == "ER" else None,
                )
            po_obj = p3_obj.wbs_codes[wbs].pos[po]

            # Backfill fields that were None on the first row
            if po_obj.vendor_name is None and vendor_name is not None:
                po_obj.vendor_name = vendor_name
            if po_obj.legal_entity is None and legal_entity is not None:
                po_obj.legal_entity = legal_entity
            if po_obj.country is None and country is not None:
                po_obj.country = country
            if po_obj.gl_account is None and gl_account is not None:
                po_obj.gl_account = gl_account
            if po_obj.req_title is None and req_title is not None:
                po_obj.req_title = req_title

            # ── Monthly metrics from transactional data ─────────────────────
            po_lookup = str(po).strip().upper() if wbs == "ER" and po else po
            if po_lookup in transactional_data:
                po_data = transactional_data[po_lookup]
                if po_obj.gross_po_value is None:
                    po_obj.gross_po_value = po_data.get('gross_ber_total')
                for month_key, values in po_data.items():
                    if month_key in ('cost_center', 'wbs', 'gross_ber_total'):
                        continue
                    if month_key not in po_obj.monthly_data:
                        po_obj.monthly_data[month_key] = MonthlyMetrics()
                    m_obj = po_obj.monthly_data[month_key]
                    m_obj.actual           = values.get('Actual', 0.0)
                    m_obj.accrual          = values.get('Accrual', 0.0)
                    m_obj.accrual_reversal = values.get('Reversal', 0.0)

            # ── Forecast data ───────────────────────────────────────────────
            if po in forecast_data:
                is_intl = bool(intl_po_set and str(po).strip() in intl_po_set)
                for month_key, values in forecast_data[po].items():
                    fc_value = values.get('Forecast', 0.0)
                    if not fc_value:
                        continue
                    write_month = _FORECAST_MONTH_SHIFT.get(month_key, month_key) if is_intl else month_key
                    if write_month not in po_obj.monthly_data:
                        po_obj.monthly_data[write_month] = MonthlyMetrics()
                    po_obj.monthly_data[write_month].forecast = fc_value

        # ── Reclass notes ───────────────────────────────────────────────────
        if reclass_notes:
            for po_number, month_entries in reclass_notes.items():
                po_in_td = transactional_data.get(po_number, {})
                po_wbs   = po_in_td.get('wbs', '') or ''
                po_cc    = po_in_td.get('cost_center', '') or ''
                # Match this reclass to the current P3 ID via WBS root, or
                # via cost_center directly when no WBS is listed in the template.
                po_p3_id = _find_p3_id_for_wbs(po_wbs) if po_wbs else None
                if po_p3_id is None and po_cc and po_cc in _known_p3_ids:
                    po_p3_id = po_cc
                if po_p3_id != p3_id:
                    continue
                wbs_key = po_in_td.get('wbs', 'NO_WBS') or 'NO_WBS'
                if wbs_key not in p3_obj.wbs_codes:
                    p3_obj.wbs_codes[wbs_key] = WBSCode(
                        wbs_code=wbs_key,
                        cost_center=p3_id,
                        charge_type=wbs_charge_type(wbs_key),
                    )
                if po_number not in p3_obj.wbs_codes[wbs_key].pos:
                    p3_obj.wbs_codes[wbs_key].pos[po_number] = PO(po_number=po_number)
                    po_obj = p3_obj.wbs_codes[wbs_key].pos[po_number]
                    for mk, vals in po_in_td.items():
                        if mk in ('cost_center', 'wbs', 'gross_ber_total'):
                            continue
                        if mk not in po_obj.monthly_data:
                            po_obj.monthly_data[mk] = MonthlyMetrics()
                        po_obj.monthly_data[mk].actual          = vals.get('Actual', 0.0)
                        po_obj.monthly_data[mk].accrual         = vals.get('Accrual', 0.0)
                        po_obj.monthly_data[mk].accrual_reversal = vals.get('Reversal', 0.0)
                else:
                    po_obj = p3_obj.wbs_codes[wbs_key].pos[po_number]
                for month_label, entries in month_entries.items():
                    if month_label not in po_obj.reclass_adjustments:
                        po_obj.reclass_adjustments[month_label] = list(entries)
                    else:
                        existing_set = set(po_obj.reclass_adjustments[month_label])
                        for entry in entries:
                            if entry not in existing_set:
                                po_obj.reclass_adjustments[month_label].append(entry)
                                existing_set.add(entry)

        result[p3_id] = p3_obj

    if er_extracted_count > 0:
        print(f"  - ER numbers extracted from project rows: {er_extracted_count}")

    # ── Catch-all: log unclaimed rows that belong to a known P3 ID ──────────
    # Only rows whose cost_center matches a P3 ID that IS on the template are
    # relevant — rows belonging to entirely different projects are ignored.
    # Within that scope, log:
    #   • MISSING_PO  — no PO was found in Document num/PO# or CO Doc Line Item Txt
    #   • UNMATCHED_P3 — has a valid PO but was never routed into rows_by_p3
    #     (e.g. WBS root doesn't match any template entry for this P3 ID)
    claimed_indices: set = set()
    for p3_id in p3_wbs_map.keys():
        for row_idx, _ in rows_by_p3[p3_id]:
            claimed_indices.add(row_idx)

    n_missing_po = 0
    n_unmatched_p3 = 0
    for row_idx, row in hierarchy_map.items():
        if row_idx in claimed_indices:
            continue
        cost_center = row.get('cost_center')
        # Skip rows that don't belong to any P3 ID on the template
        if cost_center not in _known_p3_ids:
            continue

        source_row_data = transactional_df.loc[row_idx].to_dict() if row_idx in transactional_df.index else {}
        month      = source_row_data.get('Accounting Period')
        amount     = source_row_data.get('GL BER Corp Amount')
        trans_type = source_row_data.get('Type')
        po         = row.get('po')
        wbs        = row.get('wbs')

        if not po:
            exception_log.log(
                ExceptionType.MISSING_PO,
                row_index=row_idx, wbs=wbs, cost_center=cost_center,
                month=month, amount=amount,
                transaction_type=trans_type, source_row_data=source_row_data,
            )
            n_missing_po += 1
        else:
            exception_log.log(
                ExceptionType.UNMATCHED_P3,
                row_index=row_idx, po=po, wbs=wbs, cost_center=cost_center,
                month=month, amount=amount,
                transaction_type=trans_type, source_row_data=source_row_data,
            )
            n_unmatched_p3 += 1

    if n_missing_po:
        print(f"  - Rows with no PO (logged to exceptions): {n_missing_po}")
    if n_unmatched_p3:
        print(f"  - Rows with unmatched WBS for known P3 ID (logged to exceptions): {n_unmatched_p3}")

    return result
