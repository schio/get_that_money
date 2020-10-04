import talib
import pandas as pd


def get_macd(df, fastperiod=12, slowperiod=26, signalperiod=9):
  return talib.MACD(df.close, fastperiod=fastperiod, slowperiod=slowperiod, signalperiod=signalperiod)[0]


def get_willr(df, timeperiod=14):
  return talib.WILLR(df.high, df.low, df.close, timeperiod=timeperiod)


def get_cci(df, timeperiod=14):
  return talib.CCI(df.high, df.low, df.close, timeperiod=timeperiod)


def get_ma(df, timeperiod=30):
  return talib.MA(df.close, timeperiod=timeperiod, matype=0)


def get_stoch(df, fastk_period=5, slowk_period=3, slowd_period=3, slowk_matype=0, slowd_matype=0):
  return talib.STOCH(df.high, df.low, df.close, fastk_period=fastk_period, slowk_period=slowk_period, slowk_matype=slowk_matype, slowd_period=slowd_period, slowd_matype=slowd_matype)


def get_bbands(df, timeperiod=5, nbdevup=2, nbdevdn=2, matype=0):
  return talib.BBANDS(df.close, timeperiod=timeperiod, nbdevup=nbdevup, nbdevdn=nbdevdn, matype=matype)


def get_roc(df, timeperiod=10):
  return talib.ROC(df.close, timeperiod=timeperiod)
