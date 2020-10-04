from data_loader import DataLoader
from utils.config import get_config_from_json

config = get_config_from_json()

data_loader = DataLoader(config)
if config.data.force_write_csv:
  data_loader.create_inst_candle()
  data_loader.write_candles()
data_loader.get_candles_from_csv()
