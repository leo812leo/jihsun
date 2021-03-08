import datetime
import scipy.stats as st
import numpy as np
from scipy.optimize import brentq
import pandas as pd
import re
from itertools import chain
from pandas.tseries.offsets import BDay
from pandas.tseries.offsets import CustomBusinessDay

def Premium_cal(flag,s,k,r,sigma,t,ratio)->float:
    def N(x):
        N=st.norm.cdf(x,scale=1)    
        return N
    
    d1=(np.log(s/k)+(r+0.5*sigma**2)*t)/(sigma*np.sqrt(t))
    d2=d1-sigma*np.sqrt(t)

    if flag=='c' or flag=='C':
        return (s*N(d1)-k*np.exp(-r*t)*N(d2))*ratio
    else:
        return (-s*N(-d1)+k*np.exp(-r*t)*N(-d2))*ratio
# =============================================================================
#過去的假日
df=pd.read_csv('https://www.twse.com.tw/holidaySchedule/holidaySchedule?response=csv&queryYear='+\
               str(datetime.datetime.now().year-1911),header=1,encoding='big5')
df=df[~df.iloc[:,-1].str.contains('o').fillna(False)]
tem=[re.findall(r'(\d+月\d+日)',i.values[0])for _,i in df[['日期']].iterrows()]

df=pd.read_csv('https://www.twse.com.tw/holidaySchedule/holidaySchedule?response=csv&queryYear='+\
               str(datetime.datetime.now().year-1911-1),header=1,encoding='big5')
tem2=[re.findall(r'(\d+月\d+日)',i.values[0])for _,i in df[[' 日期']].iterrows()]

holiday=[(datetime.datetime.strptime(str(datetime.datetime.now().year)+i, "%Y%m月%d日")).date() for i in chain(*tem)]
holiday2=[(datetime.datetime.strptime(str(datetime.datetime.now().year-1)+i, "%Y%m月%d日")).date() for i in chain(*tem2)]
#未來的假日
today=datetime.date.today()    
day=pd.read_html('http://172.16.10.13/hedge/tradingday.php',encoding= 'big5')[-1].iloc[:,1].drop(0)
day=set(day.apply(lambda x : pd.to_datetime(x).date()).to_list())
date_range=pd.date_range(start=datetime.datetime.strftime(today,'%Y%m%d'),
              end=datetime.datetime.strftime(datetime.date(today.year+2,1,1),'%Y%m%d')).tolist()
date_range=set([day.date() for day in date_range])
temp=date_range-day
holiday.extend(temp)
holiday.extend(holiday2)
holiday=list(set(holiday))
holiday=[d  for d in holiday  if d.weekday() not in [5,6]]


# =============================================================================
def workingday(start,end,holiday=holiday)->int:
    work_day=np.busday_count( start, end )
    holiday_num=sum( (np.array(holiday)>=start) & (np.array(holiday)<end) )
    work_day_real=work_day-holiday_num
    return work_day_real

def last_x_workday(x,date=datetime.date.today(),holiday=holiday):                              #假日
    last_workday=(date- BDay(x)).date()                              #往前x天用工作天推(last_workday)
    num_holiday=sum([(i<date) & (i>=last_workday) for i in holiday]) #期間內假日總數
    if num_holiday!=0:
        return last_x_workday(num_holiday,last_workday)
    return last_workday
    
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
        a=-1e-4,
        b=100,
        xtol=1e-6,
        rtol=1e-6,
        maxiter=100000,
        full_output=False
        )

def calculate(df):
    if df['flag']=='c':
        return round(( (df['標的收盤價']-df['最新履約價']) /df['最新履約價'] )*100,2)
    elif df['flag']=='p':
        return round(( -(df['標的收盤價']-df['最新履約價']) /df['最新履約價']  )*100,2)
    else:
        raise ValueError("數值錯誤")
        

def TheoryPrice_BidVol_Theta(table,date_str):
    
    def theta(flag,s,k,r,sigma,t,ratio,oi):
        option_pl=-(Premium_cal(flag,s,k,r,sigma,(t-1/252),ratio)-Premium_cal(flag,s,k,r,sigma,t,ratio))*1000*oi
        return option_pl

    temp=[];temp2=[];temp3=[]
    for _,values in table[['標的收盤價','最新履約價','EWMA','到期日期','flag','最新執行比例','最後委買價','自營商庫存']].iterrows():
        s,k,sigma,enddate,flag,ratio,bid,在外張數 = values
        sigma=sigma/100
        r=0
        在外張數=-在外張數
        t=workingday(datetime.datetime.strptime(date_str,"%Y%m%d").date(),enddate)/250
        temp.append(Premium_cal(flag,s,k,r,sigma,t,ratio))
        bidvol=implied_volatility(bid, s, k, t, r, flag,ratio)
        temp2.append(implied_volatility(bid, s, k, t, r, flag,ratio))
        if bidvol!=0:
            temp3.append(theta(flag,s,k,r,bidvol,t,ratio,在外張數))
        else:
            temp3.append(theta(flag,s,k,r,sigma,t,ratio,在外張數))
    table['理論價'] = temp
    table['bid_vol']=np.where(np.array(bidvol)<0.0001,0,np.array(temp2))
    table['theta'] = temp3
    
    return table
    
    