"""
pa_main.py
Entry point for running the ETL pipeline from Power Automate Desktop (PAD).

This script:
- Accepts file paths passed in by PAD
- Runs the ETL using the modular classes already implemented
- Writes the output Excel template to the specified location
- Returns clear console messages for PAD to interpret
"""

import argparse
import sys
import os
from src.forecast_reader import ForecastReader
from src.transactional_detail_reader import InvoiceActualReader, AccrualReader
from src.template_writer import FinancialTemplateV2Writer


def run_etl(forecast_path, trans_path, template_path, output_path):
    """
    Main orchestration function.
    Power Automate will call this and provide all required file paths.
    """

    # ----------- Validate inputs -----------
    for label, path in [
        ("Forecast file", forecast_path),
        ("C-TIES file", trans_path),
        ("Template file", template_path)
    ]:
        if not os.path.exists(path):
            raise FileNotFoundError(f"{label} not found at: {path}")

    print("Starting ETL process...")
    print(f"Forecast file: {forecast_path}")
    print(f"C-TIES file:  {trans_path}")
    print(f"Template file: {template_path}")
    print(f"Output path:   {output_path}")

    # ----------- Step 1: Load forecast data -----------
    forecast_reader = ForecastReader(forecast_path)
    forecast_reader.load_forecast()
    forecast_data = forecast_reader.get_forecast_data()

    # ----------- Step 2: Load actuals + accruals -----------
    actual_reader = InvoiceActualReader(trans_path)
    actual_reader.load_transactional_detail()
    actual_data = actual_reader.get_transactional_data()

    accrual_reader = AccrualReader(trans_path)
    accrual_reader.load_transactional_detail()
    accrual_data = accrual_reader.get_transactional_data()

    # ----------- Step 3: Load template -----------
    writer = FinancialTemplateV2Writer(template_path)
    writer.parse_template()
    selected_pos = writer.pos

    # ----------- Step 4: Combine data -----------
    def combine_data(forecast, actual, accrual):
        """(same implementation as main.py)"""
        combined = {}
        all_pos = set(forecast.keys()) | set(actual.keys()) | set(accrual.keys())

        for po in all_pos:
            combined[po] = {}
            months = set()
            for d in [forecast, actual, accrual]:
                if po in d:
                    months.update(d[po].keys())

            for month in months:
                combined[po][month] = {
                    'Forecast': forecast.get(po, {}).get(month, {}).get('Forecast', 0),
                    'Actual': actual.get(po, {}).get(month, {}).get('Actual', 0),
                    'Accrual': accrual.get(po, {}).get(month, {}).get('Accrual', 0),
                    'Accrual Reversal': accrual.get(po, {}).get(month, {}).get('Accrual Reversal', 0),
                    '2WM': accrual.get(po, {}).get(month, {}).get('2WM', False)
                }

        month_order = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']
        for po in combined:
            combined[po] = {m: combined[po][m] for m in month_order if m in combined[po]}
        return combined

    combined = combine_data(forecast_data, actual_data, accrual_data)

    # ----------- Step 5: Filter for POs in template -----------
    filtered = {po: v for po, v in combined.items() if po in selected_pos}

    # ----------- Step 6: Write template output -----------
    writer.write_data(filtered)
    writer.write_forecast_audit(forecast_reader.data, selected_pos)
    writer.write_transactions_audit(actual_reader.data, selected_pos)
    writer.save(output_path)

    print("ETL process completed successfully.")
    print(f"Output saved to: {output_path}")


def parse_args():
    parser = argparse.ArgumentParser(description="ETL Pipeline Runner for Power Automate Desktop")
    parser.add_argument("--forecast", required=True, help="Path to forecast Excel file")
    parser.add_argument("--transactions", required=True, help="Path to transactional C-TIES Excel file")
    parser.add_argument("--template", required=True, help="Path to the financial template")
    parser.add_argument("--output", required=True, help="Path to write the completed template output")
    return parser.parse_args()


def main():
    try:
        args = parse_args()
        run_etl(
            forecast_path=args.forecast,
            trans_path=args.transactions,
            template_path=args.template,
            output_path=args.output
        )
        sys.exit(0)

    except Exception as e:
        print("ETL FAILED:", str(e))
        sys.exit(1)


if __name__ == "__main__":
    main()
