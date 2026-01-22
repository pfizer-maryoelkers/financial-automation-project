from src.forecast_reader import ForecastReader
from src.transactional_detail_reader import TransactionalDetailReader
from src.template_writer import TemplateWriter


# Function to combine forecasts, actuals, and accrual data into one dictionary
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
                # TODO: Add source field and decide how to handle sources
            }
    return combined



def main():
    # --- Step 1: Read the vendor forecast file 
    forecast_file = "data/ibm_forecast.xlsx" 
    fr = ForecastReader(forecast_file)
    fr.load_forecast()
    forecast_data = fr.get_forecast_data()  # { 'PO': {'Jan': {'Forecast': ...}, ...}, ... }

    # --- Step 2: Load and prepare the template 
    template_file = "data/Financial Spreadsheet Template v2.xlsx" 
    tw = TemplateWriter(template_file)
    tw.parse_headers()  # Builds header map and applies filter on header row

    # --- Step 3: Write only selected POs
    selected_pos = ["9500879389", "9500882917"]
    tw.write_forecast(forecast_data, po_filter=selected_pos)

    # --- Step 4: Save the updated template
    tw.save("data/completed_financial_template.xlsx")  # Save to a new file

if __name__ == "__main__":
    main()
