import datetime
from sys import path
path.append(r'T:\WCFADOX_Python')
from WCFAdox import PCAX
import pandas as pd
from tool import calculate,last_x_workday
from pandas.tseries.offsets import BDay
PX=PCAX("192.168.65.11")
def largest_underlying(t0,n_large):
    # =============================================================================
    # 日期
    # =============================================================================
    today=(datetime.date.today()-BDay(t0)).date()
    end_date   = last_x_workday(1,today)
    #%%
    
    col='日期,代號,標的代號,標的名稱,發行日期,到期日期,最新執行比例,[價內外程度(%)],[存續期間(月)],發行價格,標的收盤價,最新履約價'
    df1=PX.Mul_Data("權證評估表","D",datetime.datetime.strftime(end_date,"%Y%m%d")
                    ,colist=col)
    df1=df1.set_index('代號')
    col="代號,發行機構名稱"
    df2=PX.Mul_Data("權證基本資料表","Y",'2020',colist=col)
    df2=df2.set_index('代號')
    data1=pd.concat([df1,df2],axis=1,join='inner')
    data1=data1[~((data1.index.str[-1]=='B') | (data1.index.str[-1]=='X') | (data1.index.str[-1]=='C'))]
    to_num=data1.columns[5:11].to_list()
    data1[to_num]=data1[to_num].apply(pd.to_numeric)
    to_date=data1.columns[3:5].to_list()
    data1[to_date]=data1[to_date].applymap(pd.to_datetime)     
    data1['flag']=data1.index.str[-1].to_series().apply(lambda  s: s.lower() if s=='P' else 'c' ).to_list()
    data1['價內外程度(%)']=data1[['標的收盤價','最新履約價','flag']].apply(calculate ,axis=1)
    
    del df1,df2,to_date,to_num
    #%%發行分析
    start_date = datetime.datetime.strftime(last_x_workday(5,today),"%Y-%m-%d")
    end_date   = datetime.datetime.strftime(last_x_workday(1,today),"%Y-%m-%d")
    data_new=data1.query('發行日期>= @start_date and 發行日期 <= @end_date')
    
    #標的統計
    table4 = data_new.pivot_table(index=['發行機構名稱'],columns=['標的代號'],values=['發行價格'],fill_value=0,
                         aggfunc={'發行價格':'count'})
    table4.loc['總和',:]=table4.sum()
    table4=table4.droplevel(0,axis=1)
    largest=table4.loc['總和'].nlargest(n_large).index
    return largest
