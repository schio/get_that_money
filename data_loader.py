import win32com.client
import pandas as pd
from tqdm import tqdm
import time
import csv
import os

class DataLoader():
    def __init__(self):
        conn = win32com.client.Dispatch("CpUtil.CpCybos")
        if conn.IsConnect:
            print("connect success")
            self.inst_candle = win32com.client.Dispatch("CpSysDib.StockChart")
        else:
            print("connect fail")
            exit()
    
    def get_market_info(self, market_kind="kospi", section_kinds=[1]):
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
        """
        GetStockMarketKind 
        typedef enum {
        [helpstring("구분없음")] CPC_MARKET_NULL  = 0,
        [helpstring("거래소")]    CPC_MARKET_KOSPI  = 1,
        [helpstring("코스닥")]    CPC_MARKET_KOSDAQ = 2,
        [helpstring("프리보드")] CPC_MARKET_FREEBOARD = 3,
        [helpstring("KRX")]    CPC_MARKET_KRX  = 4,
        }CPE_MARKET_KIND;
        """
        
        """
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
        self.inst_candle.SetInputValue(0, code) # 요청 종목 코드
        self.inst_candle.SetInputValue(1, ord('1')) # 1 - 기간으로 요청, 2 - 갯수로 요청
        self.inst_candle.SetInputValue(2, end_date) # 요청종료일, YYYYMMDD
        self.inst_candle.SetInputValue(3, start_date) # 요청시작일, YYYYMMDD
        self.inst_candle.SetInputValue(5, (0,2,3,4,5,6,8,9,10)) # 0: 날짜, 1: 시간, 2: 시가, 3: 고가, 4: 저가, 5: 종가, 6: 전일대비, 8: 거래량, 9: 거래대금, 10: 누적체결매도수량
        self.inst_candle.SetInputValue(6, ord(chart_type)) # ‘D’: 일, ‘W’: 주, ‘M’: 월, ‘m’: 분, ‘T’: 틱
        self.inst_candle.SetInputValue(9, ord('1')) # 0 - 무수정주가, 1 - 수정주가
        self.inst_candle.BLockRequest()

    def get_candle(self, code, start_date, end_date, chart_type="D"):
        self.set_inst_candle(code, start_date, end_date, chart_type)
        len_data = self.inst_candle.GetHeaderValue(3)
        len_field = self.inst_candle.GetHeaderValue(1)
        
        candle = [None] * len_data
        for i in range(len_data):
            row = [code]
            for j in range(len_field):
                row.append(self.inst_candle.GetDataValue(j, i))
            candle[i] = row
        
        return candle
    
    def create_dir(self, path):
        if not os.path.isdir(path):
            print(f"CREATE DIR: {path}")
            os.mkdir(path)

    def write_candle(self, candles_dir_path, code, start_date, end_date, chart_type="D"):
        self.set_inst_candle(code, start_date, end_date, chart_type)
        len_data = self.inst_candle.GetHeaderValue(3)
        len_field = self.inst_candle.GetHeaderValue(1)
        
        with open(os.path.join(candles_dir_path, f"{code}.csv"), "w", newline='') as f:
            candle = [None] * len_data
            writer = csv.writer(f)
            for i in range(len_data):
                row = [code]
                for j in range(len_field):
                    row.append(self.inst_candle.GetDataValue(j, i))
                writer.writerow(row)
            
    
    def write_entire_period_candles(self, market_kind):
        from datetime import datetime
        candles_dir_path = os.path.join(os.getcwd(), f"candles_{market_kind}")
        self.create_dir(candles_dir_path)
        
        candles = []
        now = datetime.now().strftime("%Y%m%d")
        for code, name, _, listing_date in tqdm(self.get_market_info(market_kind)):
            self.write_candle(candles_dir_path, code, str(listing_date), now)
            time.sleep(0.5)
        return candles


data_loader = DataLoader()
print(len(data_loader.write_entire_period_candles("kospi")))


