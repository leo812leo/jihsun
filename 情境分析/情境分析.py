import datetime
import scipy.stats as st
import numpy as np
from sys import path
path.append(r'C:\Users\F128537908\Desktop\WCFADOX_Python')
from WCFAdox import PCAX
import pandas as pd
import re
from itertools import chain
from pandas.tseries.offsets import BDay
import matplotlib.pyplot as plot
plot.rcParams['font.sans-serif'] = 'Arial Unicode MS'
def Premium_cal(flag,s,k,r,sigma,t,ratio)->float:
    d1=(np.log(s/k)+(r+0.5*sigma**2)*t)/(sigma*np.sqrt(t))
    d2=d1-sigma*np.sqrt(t)

    if flag=='c' or flag=='C':
        return (s*N(d1)-k*np.exp(-r*t)*N(d2))*ratio
    else:
        return (-s*N(-d1)+k*np.exp(-r*t)*N(-d2))*ratio
df=pd.read_csv('https://www.twse.com.tw/holidaySchedule/holidaySchedule?response=csv&queryYear='+\
               str(datetime.datetime.now().year-1911),header=1,encoding='big5')
tem=[re.findall(r'(\d+月\d+日)',i.values[0])for _,i in df[['日期']].iterrows()]
holiday=[(datetime.datetime.strptime(str(datetime.datetime.now().year)+i, "%Y%m月%d日")).date() for i in chain(*tem)]
def workingday(start,end,holiday=holiday)->int:
    work_day=np.busday_count( start, end )
    holiday_num=sum((np.array(holiday)>=start) & (np.array(holiday)<end))
    work_day_real=work_day-holiday_num
    return work_day_real
def gamma_cal(s,k,r,sigma,t):
    d1=(np.log(s/k)+(r+0.5*sigma**2)*t)/(sigma*np.sqrt(t))
    gamma=np.exp(-r*t)*dN(d1)/(s*sigma*np.sqrt(t)) 
    return gamma
def delta_cal(flag,s,k,r,sigma,t):
    d1=(np.log(s/k)+(r+0.5*sigma**2)*t)/(sigma*np.sqrt(t))
    if flag=='c' or flag=='C':
        return np.exp(-r*t)*N(d1)
    else:
        return np.exp(-r*t)*(N(d1)-1)
def N(x):
    N=st.norm.cdf(x,scale=1)    
    return N
def dN(X):
    dN=st.norm.pdf(X,scale = 1) 
    return dN
def pl(flag,s,k,r,sigma,t,ratio,oi,up_down,delta):
    option_pl=-(Premium_cal(flag,s*(1+up_down/100),k,r,sigma,(t-1/250),ratio)-Premium_cal(flag,s,k,r,sigma,t,ratio))*1000*oi
    stock_pl=delta*(up_down/100)           #避險部位為delta
    return option_pl+stock_pl
del df,tem
#%%
date_str='20200731'
stock='2330'
broker='富邦綜合證'
#%%
# table1
PX=PCAX("192.168.65.11")
col="代號,名稱,權證收盤價,標的收盤價,最新履約價,最後委買價,[隱含波動率(委買價)(%)]"
df1=PX.Mul_Data("權證評估表","D",date_str,colist=col)
df1=df1.set_index("代號")
col="代號,發行日期,到期日期,標的代號,最新執行比例,發行機構名稱"
year=str(datetime.date.today().year)
df2=PX.Mul_Data("權證基本資料表","Y",year,colist=col)
df2=df2.set_index("代號")
data1=pd.concat([df1,df2],axis=1,join='inner')
data1=data1[~((data1.index.str[-1]=='B') | (data1.index.str[-1]=='X') | (data1.index.str[-1]=='C'))]

to_num=data1.columns[1:6].to_list()+ [data1.columns[-2]]
data1[to_num]=data1[to_num].apply(pd.to_numeric)
data1[['發行日期','到期日期']]=data1[['發行日期','到期日期']].apply(pd.to_datetime)
for string in ['發行日期','到期日期']:
    data1[string]=data1[string].dt.date
# table2
col="股票代號,自營商庫存,自營商買賣超,[自營商買賣超金額(千)]"
df3=PX.Pal_Data("日自營商進出排行","D",date_str,date_str,colist=col)
df3=df3.set_index("股票代號")
df4=PX.Mul_Data("日收盤表排行","D",date_str,colist="股票代號,[成交金額(千)]",ps="<CM代號,權證>")
df4=df4.set_index("股票代號")
data2=pd.concat([df3,df4],axis=1,join='inner')
data2=data2.apply(pd.to_numeric)
table=pd.concat([data1,data2],axis=1,join='inner')
table['flag']=table.index.str[-1].to_series().apply(lambda  s: s.lower() if s=='P' else 'c' ).to_list()
table=table[table['標的代號']==stock]
table=table[table['發行機構名稱']==broker]
del df3,df4,data1,data2,year,df1,df2,col,to_num,string
#%%
for (i,(index,values)) in enumerate(table[['標的收盤價','最新履約價','隱含波動率(委買價)(%)','到期日期','flag','最新執行比例','自營商庫存']].iterrows()):
    s,k,sigma,enddate,flag,ratio,oi = values
    sigma=sigma/100;oi=-oi
    r=0
    t=workingday(datetime.datetime.strptime(date_str,"%Y%m%d").date(),enddate)/250
    delta=delta_cal(flag,s,k,r,sigma,t)*ratio*s*1000*oi    #正確delta(dollar)
    table.loc[index,'delta']=delta
    table.loc[index,'gamma']=-gamma_cal(s,k,r,sigma,t)*ratio*((s*1000)**2)*0.01*oi
    for i in range(-10,11):
        table.loc[index,str(i)]=pl(flag,s,k,r,sigma,t,ratio,oi,i,delta)
del s,k,r,sigma,t,ratio,oi
#%%
table.to_excel(date_str+'.xlsx')