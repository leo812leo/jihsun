import datetime
from sys import path
path.append(r'T:\WCFADOX_Python')
from WCFAdox import PCAX
import pandas as pd
from pandas.tseries.offsets import BDay
import matplotlib.pyplot as plot
import numpy as np
from tool import workingday,implied_volatility,last_x_workday,calculate
plot.rcParams['font.sans-serif'] = 'Arial Unicode MS'
PX=PCAX("192.168.65.11")
def last3_ask_vol(tt=60,filter_for_warrant="價內外<10 and 最後委買價>0.3 and flag=='c'",t_last_day=1):
    #%%
    # =============================================================================
    # input
    # 盤後時間用 today_str
    # 盤前時間用 last_workday_str
    # =============================================================================
    #today
    today=datetime.date.today()
    today_str=datetime.datetime.strftime(today,'%Y%m%d')
    #last_workday
    last_workday=last_x_workday(t_last_day,today)
    last_workday_str=datetime.datetime.strftime(last_workday,'%Y%m%d')
    today=last_workday
    today_str=last_workday_str
    
    #標的選擇
    data_index=PX.Mul_Data("日收盤表排行","D",today_str,colist="股票代號",ps="<CM特殊,2112>").iloc[:,0].values
    from collections import Counter
    data_index=[item for item, count in Counter(data_index).items()]
    stock_set=set(data_index)
    
    del last_workday,today
    #%%
    # =============================================================================
    # 計算權證最低vol
    # =============================================================================
    col="代號,名稱,標的代號,標的收盤價,到期日期,最後委買價,最後委賣價,最新履約價,最新執行比例"
    warrant_data=PX.Pal_Data("權證評估表","D",last_workday_str,last_workday_str,colist=col)
    warrant_data=warrant_data.query('標的代號 in @stock_set')
    
    #資料前處理
    warrant_data[['最後委賣價','最新履約價','最新執行比例','標的收盤價','最後委買價']]=\
    warrant_data[['最後委賣價','最新履約價','最新執行比例','標的收盤價','最後委買價']].apply(lambda s:pd.to_numeric(s))
    warrant_data['到期日期']=warrant_data['到期日期'].apply(lambda s:datetime.datetime.strptime(s,'%Y%m%d').date())
    warrant_data['t_modify']=warrant_data['到期日期'].apply(lambda s: workingday(datetime.date.today(),s)/250)
    warrant_data=warrant_data[warrant_data['t_modify']>tt/250]
    
    #filter
    warrant_data=warrant_data[warrant_data['到期日期']>datetime.date.today()]
    warrant_data=warrant_data.fillna(0)
    warrant_data=warrant_data[warrant_data['最後委賣價']!=0]
    #flag
    warrant_data=warrant_data[~((warrant_data['代號'].str[-1]=='B') | (warrant_data['代號'].str[-1]=='C')| (warrant_data['代號'].str[-1]=='X')| (warrant_data['代號'].str[-1]=='F'))]
    warrant_data['flag']=warrant_data['代號'].str[-1].apply(lambda  s: s.lower() if s=='P' else 'c' )
    mask=warrant_data['名稱'].str.contains("日盛")
    warrant_data=warrant_data[~mask]
    
    del mask,col,stock_set
    #%%
    # =============================================================================
    # 價內外計算
    # =============================================================================
    #價內外
    warrant_data['價內外']=warrant_data[['標的收盤價','最新履約價','flag']].apply(calculate ,axis=1)
    warrant_data=warrant_data.query(filter_for_warrant)
    #%%
    # =============================================================================
    # vol計算
    # =============================================================================
    generator=zip(warrant_data['最後委賣價'].values,
                  warrant_data['標的收盤價'].values,
                  warrant_data['最新履約價'].values,
                  warrant_data['t_modify'].values,
                  np.zeros(warrant_data.shape[0]), 
                  warrant_data['flag'].values,
                  warrant_data['最新執行比例'].values,
                  warrant_data.index.tolist(),
                   warrant_data['最後委買價'].values)  
    vol_for_each=[]
    for price, S, K, t, r, flag,ratio,index,s_check in generator:
        if (price-s_check)/s_check>=2: #filter
             vol_for_each.append(np.nan)
        else:
            vol_for_each.append(round(implied_volatility(price, S, K, t, r, flag,ratio)*100,1))
    warrant_data['askvol']=vol_for_each
    warrant_data=warrant_data.set_index('代號')
    warrant_data=warrant_data.groupby('標的代號').apply(lambda s: s.nsmallest(3,'askvol'))
    warrant_data=warrant_data.drop(['標的代號','t_modify','標的收盤價'],axis=1)
    return warrant_data
    