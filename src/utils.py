import yaml
import base64
import io

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

