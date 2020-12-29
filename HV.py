import datetime
from sys import path
path.append(r'C:\Users\F128537908\Desktop\WCFADOX_Python')
from WCFAdox import PCAX
import pandas as pd
from pandas.tseries.offsets import BDay
import numpy as np
PX=PCAX("192.168.65.11")
def vol(day):
    today=datetime.date.today()
    last=(today-BDay(day+10)).date()
    col='日期,股票代號,股票名稱,收盤價'
    df=PX.Pal_Data("日收盤還原表排行","D",
                   datetime.datetime.strftime(last,'%Y%m%d'),
                   datetime.datetime.strftime(today,'%Y%m%d'),
                   colist=col,isst='Y')
    df['收盤價']=pd.to_numeric(df['收盤價'])
    df['日期']=pd.to_datetime(df['日期'])
    df=df.set_index('日期')
    df=df.sort_index(ascending=True)
    df=df.dropna()
    tem=df.groupby('股票代號')['收盤價'].rolling(2)\
          .apply(lambda x: np.log(x[1]/x[0]))\
          .dropna()
    
    dict_vol={}
    for stock in tem.index.levels[0].values:
        if stock in tem.index.droplevel(-1):
            dict_vol[stock]=np.round(tem.loc[stock].iloc[-(day-1):].std()*(252**0.5)*100,2)
    return pd.Series(dict_vol)




