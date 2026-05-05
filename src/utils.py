import yaml
import base64
import io
import pandas as pd
from collections import defaultdict
from src.models import CostCenter, WBSCode, PO, MonthlyMetrics, ExceptionLog, ExceptionType


def load_config(config_path='configs/config_base.yaml'):
    """Load and merge YAML configs.
    
    Defaults to values in 'config_base.yaml', overrides if there are any overrides needed
    """
    with open(config_path) as f:
        config = yaml.safe_load(f)
    
    return config

# Helper function to combine forecasts, actuals, and accrual data into one JSON formatted dictionary
def combine_data(forecast, transactional):
    combined = {}

    # All PO numbers that exist in either dataset
    all_pos = set(forecast.keys()) | set(transactional.keys())

    months_list = [
        "Jan","Feb","Mar","Apr","May","Jun",
        "Jul","Aug","Sep","Oct","Nov","Dec"
    ]

    for po in all_pos:
        combined[po] = {}

        for month in months_list:
            f = forecast.get(po, {}).get(month, {})
            t = transactional.get(po, {}).get(month, {})

            combined[po][month] = {
                "Forecast": f.get("Forecast", 0),
                "Actual": t.get("Actual", 0),
                "Accrual": t.get("Accrual", 0),
                "Accrual Reversal": t.get("Reversal", 0)
            }

    return combined

def convert_base64(bytes_string: str):
    # Converts string of bytes to Excel like object for pd/openpyxl to read
    decoded_bytes = base64.b64decode(bytes_string)
    excel_file_like_object = io.BytesIO(decoded_bytes)
    return excel_file_like_object


def build_hierarchy(
    cost_centers: list[str],
    hierarchy_map: dict,
    transactional_data: dict,
    forecast_data: dict,
    exception_log: ExceptionLog,
    transactional_df: pd.DataFrame
) -> dict[str, CostCenter]:
    
    # Step 1: Pre-group hierarchy_map rows by cost center
    rows_by_cost_center = defaultdict(list)
    for row_idx, row in hierarchy_map.items():
        if row['cost_center'] in cost_centers:
            rows_by_cost_center[row['cost_center']].append((row_idx, row))

    # Step 2: Pre-scan to identify duplicate WBS codes (WBS appearing under multiple cost centers)
    wbs_to_cost_centers = defaultdict(set)  # wbs -> set of cost_centers
    for cc_id in cost_centers:
        for row_idx, row in rows_by_cost_center[cc_id]:
            wbs = row['wbs']
            if wbs:  # Only track non-empty WBS
                wbs_to_cost_centers[wbs].add(cc_id)
    
    # Identify which WBS codes are duplicates (appear under multiple cost centers)
    duplicate_wbs_codes = {wbs for wbs, ccs in wbs_to_cost_centers.items() if len(ccs) > 1}

    # Step 3: Track seen POs and WBS codes for processing
    seen_pos = {}  # po -> (cost_center, wbs)
    seen_wbs = {}  # wbs -> cost_center (for tracking first occurrence)

    # Step 4: Build hierarchy
    result = {}
    for cc_id in cost_centers:
        cost_center = CostCenter(cost_center_id=cc_id)
        for row_idx, row in rows_by_cost_center[cc_id]:
            po = row['po']
            wbs = row['wbs']

            # Extract source row data from transactional DataFrame
            source_row_data = transactional_df.loc[row_idx].to_dict() if row_idx in transactional_df.index else {}
            month = source_row_data.get('Month')
            amount = source_row_data.get('GL BER Corp Amount')
            trans_type = source_row_data.get('Type')

            # Handling Exceptions (in priority order)

            # Check 1: Both WBS and PO missing (highest priority)
            if not wbs and not po:
                exception_log.log(
                    ExceptionType.MISSING_WBS_AND_PO,
                    row_index=row_idx,
                    cost_center=cc_id,
                    month=month,
                    amount=amount,
                    transaction_type=trans_type,
                    source_row_data=source_row_data
                )
                continue

            # Check 2: Individual missing checks
            if not wbs:
                exception_log.log(
                    ExceptionType.MISSING_WBS,
                    row_index=row_idx,
                    po=po,
                    cost_center=cc_id,
                    month=month,
                    amount=amount,
                    transaction_type=trans_type,
                    source_row_data=source_row_data
                )
                continue
            
            if not po:
                exception_log.log(
                    ExceptionType.MISSING_PO,
                    row_index=row_idx,
                    wbs=wbs,
                    cost_center=cc_id,
                    month=month,
                    amount=amount,
                    transaction_type=trans_type,
                    source_row_data=source_row_data
                )
                continue

            # Check 3: Duplicate WBS (WBS owned by multiple cost centers)
            # Log ALL occurrences of duplicate WBS codes as exceptions
            if wbs in duplicate_wbs_codes:
                exception_log.log(
                    ExceptionType.DUPLICATE_WBS,
                    row_index=row_idx,
                    wbs=wbs,
                    po=po,
                    cost_center=cc_id,
                    month=month,
                    amount=amount,
                    transaction_type=trans_type,
                    source_row_data=source_row_data
                )
                # Still process first occurrence in hierarchy, skip subsequent ones
                if wbs in seen_wbs and seen_wbs[wbs] != cc_id:
                    continue  # Skip this occurrence - first is canonical
            
            seen_wbs[wbs] = cc_id

            # Check 4: Duplicate PO
            if po in seen_pos:
                prev_cc, prev_wbs = seen_pos[po]
                if prev_cc != cc_id or prev_wbs != wbs:
                    exception_log.log(
                        ExceptionType.DUPLICATE_PO,
                        row_index=row_idx,
                        po=po,
                        wbs=wbs,
                        cost_center=cc_id,
                        month=month,
                        amount=amount,
                        transaction_type=trans_type,
                        source_row_data=source_row_data
                    )
                continue  # Skip regardless - first occurrence is canonical

            seen_pos[po] = (cc_id, wbs)

            # Build WBS and PO objects if not already seen
            if wbs not in cost_center.wbs_codes:
                cost_center.wbs_codes[wbs] = WBSCode(wbs_code=wbs, cost_center=cc_id)
            if po not in cost_center.wbs_codes[wbs].pos:
                cost_center.wbs_codes[wbs].pos[po] = PO(po_number=po)
            po_obj = cost_center.wbs_codes[wbs].pos[po]
            
            # Fill MonthlyMetrics from transactional data
            if po in transactional_data:
                po_data = transactional_data[po]
                for month, values in po_data.items():
                    if month in ('cost_center', 'wbs'):
                        continue
                    if month not in po_obj.monthly_data:
                        po_obj.monthly_data[month] = MonthlyMetrics()
                    metrics = po_obj.monthly_data[month]
                    metrics.actual = values.get('Actual', 0.0)
                    metrics.accrual = values.get('Accrual', 0.0)
                    metrics.accrual_reversal = values.get('Reversal', 0.0)
            
            # Fill forecast from forecast data
            if po in forecast_data:
                for month, values in forecast_data[po].items():
                    if month not in po_obj.monthly_data:
                        po_obj.monthly_data[month] = MonthlyMetrics()
                    po_obj.monthly_data[month].forecast = values.get('Forecast', 0.0)
            # Note: MISSING_FORECAST exception removed - no longer tracking
        result[cc_id] = cost_center
    
    return result