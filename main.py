import sys
from src.utils import load_config, convert_base64, build_hierarchy
from src.forecast_reader import ForecastReader
from src.transactional_detail_reader import TransactionalDetailReader
from src.template_reader import TemplateReader
from src.template_writer import TemplateWriter
from src.models import ExceptionLog
# Loading config file
config_path = 'configs/config_base.yaml'
base64 = False
config = load_config(config_path)
if base64:
    for idx, _ in enumerate(config['forecast_reader']['file_paths']):
        config['forecast_reader']['file_paths'][idx] = convert_base64(config['forecast_reader']['file_paths'][idx])
    config['transactional_detail_reader']['file_path'] = convert_base64(config['transactional_detail_reader']['file_path'])
    config['template']['file_path'] = convert_base64(config['template']['file_path'])
## Initialize classes
forecast_reader = ForecastReader(
    file_paths=config['forecast_reader']['file_paths'],
    po_col=config['forecast_reader']['po_col']
)
transactional_reader = TransactionalDetailReader(
    file_path=config['transactional_detail_reader']['file_path'],
    required_cols=config['transactional_detail_reader']['required_cols'],
    valid_types=config['transactional_detail_reader']['valid_types'],
    colmap=config['transactional_detail_reader']['colmap']
)
template_reader = TemplateReader(
    file_path=config['template']['file_path'],
    header_row=config['template']['header_row'],
    po_col=config['template']['po_col'],
    po_stop_marker=config['template']['po_stop_marker'],
    cost_center_col=config['template']['cost_center_col'],
    cost_center_start_row=config['template']['cost_center_start_row']
)
template_writer = TemplateWriter(
    file_path=config['template']['file_path'],
    header_row=config['template']['header_row'],
    po_column=config['template']['po_col'],
    output_path=config['template_writer']['output_path'],
    overwrite=config['template_writer']['overwrite'],
    dec_acc_reversal_col=config['template_writer']['dec_acc_reversal_col'],
    forecast_source_cols=config['template_writer']['forecast_source_cols'],
    transactional_source_cols=config['template_writer']['transactional_source_cols']
)
def main():
    print("============")
    exception_log = ExceptionLog()
    ## Step 1: Load data
    print("Step 1: Loading data\n")
    forecast_data = forecast_reader.get_forecast_data()
    print("Loaded forecast data\n")
    transactional_data = transactional_reader.get_transactional_data()
    hierarchy_map = transactional_reader.get_hierarchy_map()
    print("Loaded transactional data\n")
    ## Step 2: Build hierarchy
    print("Step 2: Building hierarchy\n")
    hierarchy = build_hierarchy(
        cost_centers=template_reader.cost_centers,
        hierarchy_map=hierarchy_map,
        transactional_data=transactional_data,
        forecast_data=forecast_data,
        exception_log=exception_log
    )
    ## Step 3: Write to template
    print("Step 3: Writing template output\n")
    template_writer.write_hierarchy(hierarchy, pos=template_reader.pos)
    template_writer.write_forecast_source_sheet(forecast_reader.data)
    template_writer.write_transactional_source_sheet(transactional_reader.data)

    ## Step 4: Exception summary
    print("Step 4: Exception summary\n")
    exception_log.summary()
    template_writer.write_exception_sheet(exception_log)
    template_writer.save()


if __name__ == "__main__":
    main()