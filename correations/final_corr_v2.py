from sys import path
path.append(r'C:\Users\F128537908\Desktop\WCFADOX_Python')
import datetime 
import pandas as pd
from WCFAdox import PCAX
from pandas.tseries.offsets import BDay 
from sys import path
path.append(r'C:\Users\F128537908\Desktop\WCFADOX_Python')
import re
from itertools import chain



# =============================================================================
# #假日
# =============================================================================
df=pd.read_csv('https://www.twse.com.tw/holidaySchedule/holidaySchedule?response=csv&queryYear='+\
               str(datetime.datetime.now().year-1911),header=1,encoding='big5')
tem=[re.findall(r'(\d+月\d+日)',i.values[0])for _,i in df[[' 日期']].iterrows()]
holiday=[(datetime.datetime.strptime(str(datetime.datetime.now().year)+i, "%Y%m月%d日")).date() for i in chain(*tem)]

df=pd.read_csv('https://www.twse.com.tw/holidaySchedule/holidaySchedule?response=csv&queryYear='+\
               str(datetime.datetime.now().year-1911-1),header=1,encoding='big5')
tem2=[re.findall(r'(\d+月\d+日)',i.values[0])for _,i in df[[' 日期']].iterrows()]
holiday2=[(datetime.datetime.strptime(str(datetime.datetime.now().year-1)+i, "%Y%m月%d日")).date() for i in chain(*tem2)]
holiday.extend(holiday2)

#公式
def last_x_workday(date,x,holiday=holiday):                              #假日
    last_workday=(date- BDay(x)).date()                              #往前x天用工作天推(last_workday)
    num_holiday=sum([(i<date) & (i>=last_workday) for i in holiday]) #期間內假日總數
    if num_holiday!=0:
        return last_x_workday(last_workday,num_holiday)
    return last_workday

PX=PCAX("192.168.65.11")
today=datetime.date.today()
today_str=datetime.datetime.strftime(today,"%Y%m%d")



pairs=[['0050','00632R','TX'],['00637L','00655L','00638R']]

dataname = {}
data_dict={}
x=last_x_workday(today,130)
last_x_workday_str=datetime.date.strftime(x,"%Y%m%d")

data=PX.Pal_Data("日收盤表排行","D",last_x_workday_str,today_str,colist="日期,股票名稱,股票代號,收盤價,[漲幅(%)]",isst='Y',ps='<CM特殊,2011>')
TX=PX.Pal_Data("期貨交易行情表","D",last_x_workday_str,today_str,colist="日期,名稱,代號,收盤價,[漲幅(%)]");tx=TX[TX['代號']=='TX'];tx.columns=data.columns
etf=PX.Pal_Data("ETF淨值還原表","D",last_x_workday_str,today_str,colist="日期,股票名稱,股票代號,淨值,[報酬率(%)]")
etf['股票代號']=etf['股票代號']+'_淨值'
etf=etf.rename(columns={"淨值": "收盤價", "報酬率(%)": "漲幅(%)"})
data_final=pd.concat([data,tx,etf],axis=0,ignore_index=True)
from collections import defaultdict
final = defaultdict(dict)

for i,pair in enumerate(pairs):
        data_final_temp=data_final[data_final['股票代號'].isin(pair)]
        data_final_temp.loc[:,"漲幅(%)"]=data_final_temp.loc[:,"漲幅(%)"].astype(float).values
        temp=pd.pivot_table(data_final_temp,index="股票代號",columns="日期",values="漲幅(%)").T.dropna()
        temp=temp.iloc[-120:]
        temp=temp.rolling(60).corr().dropna()
        temp=[temp.iloc[i*3:i*3+3] for i in range(1,61)]
        mean=pd.concat(temp).groupby(level=-1).mean().round(3)
        std=pd.concat(temp).groupby(level=-1).std().round(6)
        D60=temp[-1].droplevel(0).round(3).stack()
        D60=D60[D60!=1]
        ff= (std/mean).stack()
        ff=ff[ff!=0]
        df=pd.concat([D60,ff],axis=1,join='inner')
        df.columns=['相關係數','變異係數']
        final[i]['final']=df
        
        
           
path = r'C:\Users\F128537908\Desktop\output2.xlsx'
writer = pd.ExcelWriter(path, engine = 'openpyxl',mode='w')
for i,(_,dct) in enumerate(final.items()):
    for key in ['final']:
        dct[key].to_excel(writer, sheet_name=str(i)+'_'+key)
writer.save()
