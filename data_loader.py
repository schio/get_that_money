import win32com.client
import pandas as pd
from tqdm import tqdm
import time
import csv
import os
import numpy as np
import stockstats
from utils.supplementray_indicator import *


class DataLoader():
  def __init__(self, config):
    self.config = config
    self.candles_dir_path = os.path.join(os.getcwd(), f"candles_{self.config.data.market_kind}")

  def create_inst_candle(self):
    conn = win32com.client.Dispatch("CpUtil.CpCybos")
    if conn.IsConnect:
      print("Create inst_candle")
      self.inst_candle = win32com.client.Dispatch("CpSysDib.StockChart")
    else:
      print("Create inst_candle fail")
      exit()

  def get_market_info(self, market_kind, section_kinds):
    """
    GetStockMarketKind 
    typedef enum {
    [helpstring("구분없음")] CPC_MARKET_NULL  = 0,
    [helpstring("거래소")]    CPC_MARKET_KOSPI  = 1,
    [helpstring("코스닥")]    CPC_MARKET_KOSDAQ = 2,
    [helpstring("프리보드")] CPC_MARKET_FREEBOARD = 3,
    [helpstring("KRX")]    CPC_MARKET_KRX  = 4,
    }CPE_MARKET_KIND;

    GetStockSectionKind 
    typedef enum{
    [helpstring("구분없음")]   CPC_KSE_SECTION_KIND_NULL= 0,
    [helpstring("주권")]   CPC_KSE_SECTION_KIND_ST   = 1,
    [helpstring("투자회사")]   CPC_KSE_SECTION_KIND_MF    = 2,
    [helpstring("부동산투자회사"]   CPC_KSE_SECTION_KIND_RT    = 3,
    [helpstring("선박투자회사")]   CPC_KSE_SECTION_KIND_SC    = 4,
    [helpstring("사회간접자본투융자회사")] CPC_KSE_SECTION_KIND_IF = 5,
    [helpstring("주식예탁증서")]   CPC_KSE_SECTION_KIND_DR    = 6,
    [helpstring("신수인수권증권")]   CPC_KSE_SECTION_KIND_SW    = 7,
    [helpstring("신주인수권증서")]   CPC_KSE_SECTION_KIND_SR    = 8,
    [helpstring("주식워런트증권")]   CPC_KSE_SECTION_KIND_ELW = 9,
    [helpstring("상장지수펀드(ETF)")] CPC_KSE_SECTION_KIND_ETF = 10,
    [helpstring("수익증권")]    CPC_KSE_SECTION_KIND_BC    = 11,
    [helpstring("해외ETF")]      CPC_KSE_SECTION_KIND_FETF   = 12,
    [helpstring("외국주권")]    CPC_KSE_SECTION_KIND_FOREIGN = 13,
    [helpstring("선물")]      CPC_KSE_SECTION_KIND_FU    = 14,
    [helpstring("옵션")]      CPC_KSE_SECTION_KIND_OP    = 15,    
    } CPE_KSE_SECTION_KIND;
    """

    market_kind = market_kind.lower()
    cp_code_mgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
    market_kind_dict = {
        "null": 0,
        "kospi": 1,
        "kosdaq": 2,
        "freeboard": 3,
        "krx": 4
    }
    code_list = cp_code_mgr.GetStockListByMarket(market_kind_dict[market_kind])

    market = []
    for i, code in enumerate(code_list):
      second_code = cp_code_mgr.GetStockSectionKind(code)
      if second_code in section_kinds:
        name = cp_code_mgr.CodeToName(code)
        listing_date = cp_code_mgr.GetStockListedDate(code)
        market.append([code, name, second_code, listing_date])
    print(f"num of {market_kind} code: {len(market)}")
    return market

  def set_inst_candle(self, code, start_date, end_date, chart_type="D"):
    self.inst_candle.SetInputValue(0, code)  # 요청 종목 코드
    self.inst_candle.SetInputValue(1, ord("1"))  # 1 - 기간으로 요청, 2 - 갯수로 요청
    self.inst_candle.SetInputValue(2, end_date)  # 요청종료일, YYYYMMDD
    self.inst_candle.SetInputValue(3, start_date)  # 요청시작일, YYYYMMDD
    # 0: 날짜, 1: 시간, 2: 시가, 3: 고가, 4: 저가, 5: 종가, 6: 전일대비, 8: 거래량, 9: 거래대금, 10: 누적체결매도수량
    self.inst_candle.SetInputValue(5, (0, 2, 3, 4, 5, 6, 8, 9, 10))
    self.inst_candle.SetInputValue(6, ord(chart_type))  # ‘D’: 일, ‘W’: 주, ‘M’: 월, ‘m’: 분, ‘T’: 틱
    self.inst_candle.SetInputValue(9, ord("1"))  # 0 - 무수정주가, 1 - 수정주가
    self.inst_candle.BLockRequest()

  def create_dir(self, path):
    if not os.path.isdir(path):
      print(f"CREATE DIR: {path}")
      os.mkdir(path)

  def write_candle(self, candles_dir_path, code, start_date, end_date, chart_type="D"):
    self.set_inst_candle(code, start_date, end_date, chart_type)
    len_data = self.inst_candle.GetHeaderValue(3)
    len_field = self.inst_candle.GetHeaderValue(1)

    with open(os.path.join(candles_dir_path, f"{code}.csv"), "w", newline="") as f:
      candle = [None] * len_data
      writer = csv.writer(f)
      for i in range(len_data):
        row = [code]
        for j in range(len_field):
          row.append(self.inst_candle.GetDataValue(j, i))
        writer.writerow(row)

  def write_candles(self):
    from datetime import datetime

    self.create_dir(self.candles_dir_path)

    if self.config.data.candle_end_date is None:
      candle_end_date = datetime.now().strftime("%Y%m%d")
    else:
      candle_end_date = self.config.data.candle_end_date

    if self.config.data.candle_start_date is None:
      for code, name, _, listing_date in tqdm(get_market_info(self.config.data.market_kind, self.config.data.market_section_kinds)[:1]):
        self.write_candle(self.candles_dir_path, code, str(listing_date), candle_end_date)
        time.sleep(0.3)
    else:
      for code, name, _, _ in tqdm(get_market_info(self.config.data.market_kind, self.config.data.market_section_kinds)[:1]):
        self.write_candle(self.candles_dir_path, code, self.config.data.candle_start_date, candle_end_date)
        time.sleep(0.3)

  def _create_label(self, candle):
    close_y = candle['change_ratio_close_1'].apply(lambda x: 1 if x > 3 else 0)
    high_y = candle['change_ratio_high_1'].apply(lambda x: 1 if x > 5 else 0)
    candle['y'] = close_y * high_y

    return candle

  def get_candle_from_csv(self, path):
    cols = ("code", "date", "open", "close", "high", "low", "cr_close",
            "volume", "trading_value", "accumulate_sell_volume")
    # cr: 전일대비 change ratio
    # candle = np.loadtxt(path, delimiter=",", dtype={
    #     "names": cols,
    #     "formats": ("S8", "S8", "f", "f", "f", "f", "f", "f", "f", "f")
    # })
    candle = pd.read_csv(path, names=cols).iloc[::-1]
    if self.config.data.indicator.macd:
      for fastperiod in range(6, 30, 2):
        slowperiod = int(fastperiod * 2.3)
        signalperiod = slowperiod - fastperiod
        candle[f"macd_{fastperiod}_{slowperiod}_{signalperiod}"] = get_macd(
            candle, fastperiod, slowperiod, signalperiod)

    if self.config.data.indicator.willr:
      for timeperiod in range(6, 30, 2):
        candle[f"willr_{timeperiod}"] = get_willr(candle, timeperiod)

    if self.config.data.indicator.cci:
      for timeperiod in range(6, 28, 2):
        candle[f"cci_{timeperiod}"] = get_cci(candle, timeperiod)

    if self.config.data.indicator.ma:
      for timeperiod in range(6, 56, 2):
        candle[f"ma_{timeperiod}"] = get_ma(candle, timeperiod)

    if self.config.data.indicator.macd:
      for fastperiod in range(5, 30, 5):
        for slowperiod in range(3, 30, 3):
          candle[f"slow_stoch_k_{fastperiod}_{slowperiod}_{slowperiod}"], candle[f"slow_stoch_d_{fastperiod}_{slowperiod}_{slowperiod}"] = get_stoch(
              candle, fastperiod, slowperiod, slowperiod)

    if self.config.data.indicator.bollinger_bands:
      for timeperiod in range(3, 30, 2):
        candle[f"upperbb_{timeperiod}"], candle[f"middlebb_{timeperiod}"], candle[f"lowerbb_{timeperiod}"] = get_bbands(
            candle, timeperiod, int(timeperiod * 0.4), int(timeperiod * 0.4))

    if self.config.data.indicator.roc:
      for timeperiod in range(2, 30, 2):
        candle[f"roc_{timeperiod}"] = get_roc(candle, timeperiod)

    use_cols = ["open", "high", "low", "close", "volume"]
    stock = stockstats.StockDataFrame.retype(candle[use_cols])
    pd.set_option('mode.chained_assignment', None)
    for day in range(1, 30, 1):
      # candle[f"change_ratio_open_{day}"] = stock[f"open_-{day}_r"]
      candle[f"change_ratio_high_{day}"] = stock[f"high_-{day}_r"]
      # candle[f"change_ratio_low_{day}"] = stock[f"low_-{day}_r"]
      candle[f"change_ratio_close_{day}"] = stock[f"close_-{day}_r"]
      # candle[f"change_ratio_volume_{day}"] = stock[f"volume_-{day}_r"]
    # candle["positive_dmi"] = stock["pdi"]
    # candle["negative_dmi"] = stock["mdi"]
    return self._create_label(candle.dropna())

  def get_candles_from_csv(self):
    files = os.listdir(self.candles_dir_path)
    candles = []
    for file in tqdm(files[:]):
      path = os.path.join(self.candles_dir_path, file)
      candles.append(self.get_candle_from_csv(path))
    candles = pd.concat(candles)
    # candles = pd.DataFrame(np.concatenate(candles))

    pass
