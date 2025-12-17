from src.forecast_reader import ForecastReader
from src.transactional_detail_reader import TransactionalDetailReader
from src.template_writer import TemplateWriter

def main():
    # File paths (update as needed)
    forecast_file = "data/vendor_forecast.xlsx"
    transactional_detail_file = "data/transactional_detail_file.xlsx"
    template_file = "data/template.xlsx"
    output_file = "data/template_filled.xlsx"

    # Step 1: Initialize classes
    forecast_reader = ForecastReader(forecast_file)
    transactional_detail_reader = TransactionalDetailReader(transactional_detail_file)
    template_writer = TemplateWriter(template_file)

    # TODO: Add logic for transactional detail file in steps 2-4 (dependent on class implementation)

    # Step 2: Load forecast file
    forecast_reader.load_forecast()

    # Step 3: Extract forecast data
    forecast_data = forecast_reader.get_forecast_data()

    # Step 4: parse headers, and write forecast data into template
    template_writer.parse_headers()
    template_writer.write_forecast(forecast_data)

    # Step 5: Save updated template
    template_writer.save(output_file)

    print(f"Process completed. Output saved to: {output_file}")

if __name__ == "__main__":
    main()