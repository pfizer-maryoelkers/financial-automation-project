from src.forecast_reader import ForecastReader
from src.transactional_detail_reader import TransactionalDetailReader
from src.template_writer import TemplateWriter


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
