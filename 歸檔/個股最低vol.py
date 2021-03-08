import datetime
from sys import path
path.append(r'C:\Users\F128537908\Desktop\WCFADOX_Python')
from WCFAdox import PCAX
import pandas as pd
import re
from itertools import chain
from pandas.tseries.offsets import BDay
import matplotlib.pyplot as plot
import time
import numpy as np
import scipy.stats as st
from scipy.optimize import brentq
plot.rcParams['font.sans-serif'] = 'Arial Unicode MS'
PX=PCAX("192.168.65.11")
#%%
# =============================================================================
# #假日
# =============================================================================
df=pd.read_csv('https://www.twse.com.tw/holidaySchedule/holidaySchedule?response=csv&queryYear='+\
               str(datetime.datetime.now().year-1911),header=1,encoding='big5')
tem=[re.findall(r'(\d+月\d+日)',i.values[0])for _,i in df[['日期']].iterrows()]
holiday=[(datetime.datetime.strptime(str(datetime.datetime.now().year)+i, "%Y%m月%d日")).date() for i in chain(*tem)]

df=pd.read_csv('https://www.twse.com.tw/holidaySchedule/holidaySchedule?response=csv&queryYear='+\
               str(datetime.datetime.now().year-1911-1),header=1,encoding='big5')
tem2=[re.findall(r'(\d+月\d+日)',i.values[0])for _,i in df[[' 日期']].iterrows()]
holiday2=[(datetime.datetime.strptime(str(datetime.datetime.now().year-1)+i, "%Y%m月%d日")).date() for i in chain(*tem2)]
holiday.extend(holiday2)
del df,holiday2,tem,tem2
#%%
#公式
def last_x_workday(date,x,holiday=holiday):                              #假日
    last_workday=(date- BDay(x)).date()                              #往前x天用工作天推(last_workday)
    num_holiday=sum([(i<date) & (i>=last_workday) for i in holiday]) #期間內假日總數
    if num_holiday!=0:
        return last_x_workday(last_workday,num_holiday)
    return last_workday
def workingday(start,end,holiday=holiday):
    work_day=np.busday_count( start, end )
    holiday_num=sum((np.array(holiday)>=start) & (np.array(holiday)<end))
    work_day_real=work_day-holiday_num
    return work_day_real
def Premium_cal(flag,s,k,r,sigma,t,ratio):
    def N(x):
        N=st.norm.cdf(x,scale=1)    
        return N
    
    d1=(np.log(s/k)+(r+0.5*sigma**2)*t)/(sigma*np.sqrt(t))
    d2=d1-sigma*np.sqrt(t)

    if flag=='c' or flag=='C':
        return (s*N(d1)-k*np.exp(-r*t)*N(d2))*ratio
    else:
        return (-s*N(-d1)+k*np.exp(-r*t)*N(-d2))*ratio
def implied_volatility(price, S, K, t, r, flag,ratio):
    """Calculate the Black-Scholes implied volatility.
    :param price: the Black-Scholes option price
    :type price: float
    :param S: underlying asset price
    :type S: float
    :param K: strike price
    :type K: float
    :param t: time to expiration in years
    :type t: float
    :param r: risk-free interest rate
    :type r: float
    :param flag: 'c' or 'p' for call or put.
    :type flag: str
  """
    f = lambda sigma: price - Premium_cal(flag, S, K, r,sigma, t, ratio)
    return brentq(
        f,
        a=-1e-5,
        b=10,
        xtol=1e-6,
        rtol=1e-6,
        maxiter=100000,
        full_output=False
        )
#%%
# =============================================================================
# input
# 盤後時間用 today_str
# 盤前時間用 last_workday_str
# =============================================================================
stock_set=set(['2330'])
#today
today=datetime.date.today()
today_str=datetime.datetime.strftime(today,'%Y%m%d')
#last_workday
last_workday=last_x_workday(today,1)
last_workday_str=datetime.datetime.strftime(last_workday,'%Y%m%d')
#last_60_workday
last_60_workday=last_x_workday(datetime.date.today(),60)

#篩選
filter_for_warrant="價內外<=5 and 最後委買價>0.3 and 價內外>=-5"
tt=90

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
for i in [10,20,60]:
    tem=table.groupby('股票代號')['range_return'].rolling(i-1).std()*(250**0.5)*100
    tem.index=tem.index.droplevel(0)
    final[i]=tem
    
from collections import defaultdict
stock_dict=defaultdict(dict)

for stock in stock_set:
    for i in [10,20,60]:
        stock_dict[stock][str(i)+'range_vol']=(final[i].loc[stock].iloc[-1].round(1)).copy()
        
range_col_dataframe=pd.DataFrame(stock_dict).T
#range_col_dataframe=range_col_dataframe.reindex(data_index)
del stock_dict,stock,i,tem,table,final,df,last_60_workday
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
warrant_data=warrant_data[~((warrant_data['代號'].str[-1]=='B') | (warrant_data['代號'].str[-1]=='C')| (warrant_data['代號'].str[-1]=='X'))]
warrant_data['flag']=warrant_data['代號'].str[-1].apply(lambda  s: s.lower() if s=='P' else 'c' )
mask=warrant_data['名稱'].str.contains("日盛")
warrant_data=warrant_data[~mask]

del mask,col,stock_set
#%%
# =============================================================================
# 價內外計算
# =============================================================================
#價內外
for_each=[]
for _,values in warrant_data[['標的收盤價','最新履約價','flag']].iterrows():
    if values['flag']=='c':
        for_each.append(round((values['標的收盤價']-values['最新履約價'])/values['最新履約價']*100,2))
    else:
        for_each.append(round(-(values['標的收盤價']-values['最新履約價'])/values['最新履約價']*100,2))
#篩選條件
warrant_data['價內外']=for_each
warrant_data=warrant_data.query(filter_for_warrant)
del values,for_each
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
               warrant_data['最後委賣價'].values)  
vol_for_each=[]
for price, S, K, t, r, flag,ratio,index,s_check in generator:
    if (price-s_check)/s_check>=2: #filter
         vol_for_each.append(np.nan)
    else:
        vol_for_each.append(round(implied_volatility(price, S, K, t, r, flag,ratio)*100,1))
warrant_data['askvol']=vol_for_each
warrant_data=warrant_data.set_index('代號')
temp=warrant_data.groupby('標的代號').apply(lambda s: s.nsmallest(3,'askvol'))
count={ index  : temp.loc[index,temp.columns[0]].count()       for index in temp.index.levels[0]}

warrant_data=warrant_data.groupby('標的代號').apply(lambda s: s.nsmallest(3,'askvol'))
warrant_data=warrant_data.drop(['標的代號','t_modify','標的收盤價'],axis=1)

final=range_col_dataframe.merge(warrant_data, left_index=True, right_on=['標的代號'])
final=final[np.append(final.columns[3:].values,final.columns[0:3].values)]

del price, S, K, t, r, flag,ratio,index,s_check,warrant_data,vol_for_each,temp
#%%excel
final.to_excel(r'D:\code\output專用\range_vol.xlsx')
count=final.index.levels[0].map(count).values
count=count[:-1]
csum=count.cumsum()
import xlwings as xw 
app = xw.App(visible=True, add_book=False)
wb2 = app.books.open(r'D:\code\output專用\range_vol.xlsx')
wb3 = app.books.open(r'D:\code\output專用\新發行權證分析.xlsx')

ws2=wb2.sheets[0]
ws2.api.Copy(Before=wb3.sheets[-1].api)
wb3.sheets[-1].delete()
wb3.sheets[-1].name='預備發行標的'


wb1 = app.books.open(r'D:\code\output專用\code_for_format.xlsm')
sht=wb3.sheets[-1]
sht.activate()
for i in csum:
    xw.Range((2+i,1),(4+i,14)).select()
    wb1.macro('format')()
wb1.macro('格式')()

wb2.save()
wb2.close()

wb3.save()
wb3.close()

wb1.close()
app.quit()