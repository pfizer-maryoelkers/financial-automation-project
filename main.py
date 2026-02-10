from src.forecast_reader import ForecastReader
from src.transactional_detail_reader import InvoiceActualReader, AccrualReader
from src.template_writer import FinancialTemplateV2Writer


forecast_path = "data/ibm_forecast.xlsx"
tdf_path = "data/C-TIES AP09 2025.xlsx"
template_path = "data/Financial Spreadsheet Template v2.xlsx"
output_path = "data/template_output.xlsx"

# Helper function to combine forecasts, actuals, and accrual data into one JSON formatted dictionary
def combine_data(forecast, actual, accrual):
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
                '2WM': accrual.get(po, {}).get(month, {}).get('2WM', False),
                # Sources can be combined or kept separate if needed
            }

    # Reordering months for easier debugging
    month_order = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    
    for po in combined:
        # Sort the inner dict by month_order
        combined[po] = {month: combined[po][month] for month in month_order if month in combined[po]}
    return combined



def main():
    ## Step 1: Initialize readers and load data
    print("Step 1: Initializing readers and loading data\n")

    forecast_reader = ForecastReader(forecast_path)
    forecast_reader.load_forecast()
    forecast_data = forecast_reader.get_forecast_data()
    print("Loaded forecast data\n")

    actual_reader = InvoiceActualReader(tdf_path)
    actual_reader.load_transactional_detail()
    actual_data = actual_reader.get_transactional_data()
    print("Loaded actuals data\n")

    accrual_reader = AccrualReader(tdf_path)
    accrual_reader.load_transactional_detail()
    accrual_data = accrual_reader.get_transactional_data()
    print("Loaded accrual data\n")

    template_writer = FinancialTemplateV2Writer(template_path)
    template_writer.parse_template()
    selected_pos = template_writer.pos 
    print("Loaded template data\n")

    ## Step 2: Combining data and filtering POs
    print("Step 2: Combing data and filtering on selected POs\n")
    result = combine_data(forecast_data, actual_data, accrual_data)

    # Filter the result dictionary
    filtered_result = {po: data for po, data in result.items() if po in selected_pos}
    print(f"Filtered data to selcted POs: {selected_pos}\n")

    ## Step 3: Writing to template
    print("Step 3: Writing template output\n")
    template_writer.write_data(filtered_result)

    # Writing audit sheets
    template_writer.write_forecast_audit(forecast_reader.data, selected_pos)
    template_writer.write_transactions_audit(actual_reader.data, selected_pos)

    template_writer.save(output_path)
    print(f"Template saved to {output_path}")


if __name__ == "__main__":
    main()
