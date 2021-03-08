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
plot.rcParams['font.sans-serif'] = 'Arial Unicode MS'

berfore=1
after=1

def is_weekday(date):
    if date.weekday()<5:
        return True
    else:
        return False
    
holiday_dict={}
for year in range(2008,2021):
    df=pd.read_csv('https://www.twse.com.tw/holidaySchedule/holidaySchedule?response=csv&queryYear='+\
                   str(year-1911),header=1,encoding='big5')
    df=df.dropna(subset=[' 日期'])
    tem=[re.findall(r'(\d+月\d+日)',i.values[0])for _,i in df[[' 日期']].iterrows()]
    holiday_dict[year]=[(datetime.datetime.strptime(str(year)+i, "%Y%m月%d日")).date()  
                        for i in chain(*tem)  if is_weekday(  (datetime.datetime.strptime(str(year)+i, "%Y%m月%d日")).date()  )]


def last_x_workday(year,date,x): 
    date=(datetime.datetime.strptime(date,"%Y%m%d")).date()#當天
    holiday=holiday_dict[year]                                       #假日
    last_workday=(date- BDay(x)).date()                              #往前x天用工作天推(last_workday)
    num_holiday=sum([(i<date) & (i>=last_workday) for i in holiday]) #期間內假日總數
    if num_holiday!=0:
        last_str=datetime.datetime.strftime(last_workday,"%Y%m%d")
        return last_x_workday(year,last_str,num_holiday)
    return last_workday


def data(date,stock,year):
    sqlcmd="SELECT 收盤價 FROM 日收盤還原表排行 WHERE 日期='"+ date +"'AND 股票代號='"+stock+"'"
    sqltables="日收盤還原表排行"
    res=PX.Sql_data(sqlcmd,sqltables)
    search_date=date
    count = 0
    while res.empty:
        count+=1
        if count>=5:
            return np.nan
        else:
            search_date=last_x_workday(year,search_date,1)
            search_date=datetime.datetime.strftime(search_date,"%Y%m%d")
            sqlcmd="SELECT 收盤價 FROM 日收盤還原表排行 WHERE 日期='"+ search_date +"'AND 股票代號='"+stock+"'"
            res=PX.Sql_data(sqlcmd,sqltables)        
    else:
        return float(res.iloc[0,0])

PX=PCAX("192.168.65.11")


col="年度,股票代號,股票名稱,除息日,除息前股價,[現金股利合計(元)],[現金股利殖利率(%)],配發次數"
df0=PX.Pal_Data("股利政策表","Y",'2008','2019',colist=col,isst='Y')
df0=df0[~(df0['股票代號'].apply(len)>4)]
df0=df0[df0['配發次數']=='1']
df0.iloc[:,-3:]=df0.iloc[:,-3:].apply(pd.to_numeric)
df0=df0.dropna(subset=['除息前股價','現金股利合計(元)'])
df0=df0[df0['除息日']!='']
df0['除息日_d']=pd.to_datetime(df0['除息日']).dt.date

tStart = time.time()
for i,values in df0.iterrows():
    year,stock,_,除息日,_,_,_,_,除息日_d=values
    search_date=last_x_workday(int(year)+1,除息日,berfore)
    df0.loc[i,'berfore']=data(datetime.datetime.strftime(search_date,"%Y%m%d"),stock,int(year)+1)


# def calculate(s):
#     global berfore
#     search_date=last_x_workday(int(s['年度'])+1,s['除息日'],berfore)
#     return data(datetime.datetime.strftime(search_date,"%Y%m%d"),s['股票代號'],int(s['年度'])+1)

# df0['berfore']=df0[['年度', '股票代號','除息日']].apply(calculate,axis=1)
tEnd = time.time()
print('程式執行時間:'+str(round(tEnd - tStart,4))+' sec')



#df0.pivot(index='股票代號',columns='年度',values=['現金股利殖利率(%)'])

