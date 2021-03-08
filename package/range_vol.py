import datetime
from sys import path
path.append(r'C:\Users\F128537908\Desktop\WCFADOX_Python')
from WCFAdox import PCAX
import pandas as pd
import numpy as np
from collections import defaultdict
from tool import last_x_workday

def run(today=datetime.date.today(),t0=0):
    '''
    input: today (dt)
    t0 0:今天 1:昨天
    '''
    
    PX=PCAX("192.168.65.11")
    #last_workday
    last_workday=last_x_workday(t0,today)
    last_workday_str=datetime.datetime.strftime(last_workday,'%Y%m%d')
    last_60_workday=last_x_workday(60,last_workday)
    
    #標的選擇
    data_index=PX.Mul_Data("日收盤表排行","D",last_workday_str,colist="股票代號",ps="<CM特殊,2012>").iloc[:,0].values
    stock_set=set(data_index)
    
    del last_workday,today
    #%%
    # =============================================================================
    # range_vol
    # =============================================================================
    col='日期,股票代號,股票名稱,最高價,最低價,收盤價'
    df=PX.Pal_Data("日收盤還原表排行","D",
                   datetime.datetime.strftime(last_60_workday,'%Y%m%d'),
                   last_workday_str,
                   colist=col,isst='Y')
    df=df.query('股票代號 in @stock_set')
    df['日期']=df['日期'].apply(pd.to_datetime)
    df[['最高價','最低價','收盤價']]=df[['最高價','最低價','收盤價']].apply(pd.to_numeric)
    
    table=df.set_index(['股票代號','日期'])
    table=table.sort_index()
    table['昨日收盤']=table.groupby('股票名稱')['收盤價'].transform(lambda s: s.shift(1))
    
    def rang_vol(data):
        if abs(data['最高價']-data['昨日收盤']) >= abs(data['昨日收盤']-data['最低價']):
            return np.log(data['最高價']/data['昨日收盤'])
        else:
            return np.log(data['最低價']/data['昨日收盤'])
    table['range_return']=table[['最高價','最低價','昨日收盤']].apply(rang_vol,axis=1)
    final={}
    for i in [20,40,60]:
        tem=table.groupby('股票代號')['range_return'].rolling(i-1).std()*(250**0.5)*100
        tem.index=tem.index.droplevel(0)
        final[i]=tem
    stock_dict=defaultdict(dict)
    for stock in stock_set:
        for i in [20,40,60]:
            stock_dict[stock][i]=(final[i].loc[stock].iloc[-1].round(1)).copy()
    range_col_dataframe=pd.DataFrame(stock_dict).T
    range_col_dataframe=range_col_dataframe.reindex(data_index)
    range_col_dataframe['avg']=range_col_dataframe.apply(np.mean,axis=1).round(1)
    return range_col_dataframe