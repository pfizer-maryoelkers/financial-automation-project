from src.forecast_reader import ForecastReader
from src.transactional_detail_reader import TransactionalDetailReader
from src.template_writer import TemplateWriter


forecast_paths = ["data/ibm_forecast.xlsx"]
tdf_path = "data/C-TIES AP09 2025.xlsx"
template_path = "data/templates/financial_template_v2_2026.xlsx"
output_path = "data/template_output.xlsx"

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


def main():
    ## Step 1: Initialize readers and load data
    print("Step 1: Initializing readers and loading data\n")

    fr = ForecastReader(forecast_paths)
    forecast_data = fr.get_forecast_data()
    print("Loaded forecast data\n")


    tdr = TransactionalDetailReader(tdf_path)
    transactional_data = tdr.get_transactional_data()
    print("Loaded transactional data\n")

    ## Step 2: Combining data and filtering POs
    print("Step 2: Combing data and filtering on selected POs\n")
    combined = combine_data(forecast_data, transactional_data)


    ## Step 3: Writing to template
    print("Step 3: Writing template output\n")
    tw = TemplateWriter(template_path)
    tw.write_data(combined)
    tw.write_forecast_source_sheet(fr.data)
    tw.write_transactional_source_sheet(tdr.data)
    print("Loaded template data\n")

    tw.wb.save(output_path)
    print(f"Template saved to {output_path}")


if __name__ == "__main__":
    main()
