import sys
from src.utils import load_config, combine_data, convert_base64
from src.forecast_reader import ForecastReader
from src.transactional_detail_reader import TransactionalDetailReader
from src.template_writer import TemplateWriter

# Loading config file
config_path = 'configs/config_base.yaml'
base64 = False
config = load_config(config_path)

if base64:
    #  if file paths in base64 format, convert bytes to Excel like objects
    for idx, _ in enumerate(config['forecast_reader']['file_paths']):
        config['forecast_reader']['file_paths'][idx] = convert_base64(config['forecast_reader']['file_paths'][idx])

    config['transactional_detail_reader']['file_path'] = convert_base64(config['transactional_detail_reader']['file_path'])
    config['template_writer']['file_path'] = convert_base64(config['template_writer']['file_path'])


## Initialize classes
forecast_reader = ForecastReader(
    file_paths=config['forecast_reader']['file_paths']
)

transactional_reader = TransactionalDetailReader(
    file_path=config['transactional_detail_reader']['file_path'],
    required_cols=config['transactional_detail_reader']['required_cols'],
    valid_types=config['transactional_detail_reader']['valid_types'],
    colmap=config['transactional_detail_reader']['colmap']
)

template_writer = TemplateWriter(
    file_path=config['template_writer']['file_path'],
    header_row=config['template_writer']['header_row'],
    po_column=config['template_writer']['po_column'],
    dec_acc_reversal_col=config['template_writer']['dec_acc_reversal_col']
)

# Output path
output_path = 'data/templates/output_test.xlsx'


def main():
    ## Step 1: Initialize readers and load data
    print("Step 1: Initializing readers and loading data\n")

    forecast_data = forecast_reader.get_forecast_data()
    print("Loaded forecast data\n")


    transactional_data = transactional_reader.get_transactional_data()
    print("Loaded transactional data\n")

    ## Step 2: Combining data and filtering POs
    print("Step 2: Combing data and filtering on selected POs\n")
    combined = combine_data(forecast_data, transactional_data)


    ## Step 3: Writing to template
    print("Step 3: Writing template output\n")
    template_writer.write_data(combined)
    template_writer.write_forecast_source_sheet(forecast_reader.data)
    template_writer.write_transactional_source_sheet(transactional_reader.data)
    print("Loaded template data\n")

    template_writer.wb.save(output_path)
    print(f"Template saved to {output_path}")


if __name__ == "__main__":
    main()
