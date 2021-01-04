import datetime
from sys import path
path.append(r'C:\Users\F128537908\Desktop\WCFADOX_Python')
from WCFAdox import PCAX
import pandas as pd
from pandas.tseries.offsets import BDay

def  third_wen(dt):
    y=dt.year;m=dt.month
    day=21-(datetime.date(y,m,1).weekday()+4)%7 
    return dt.day==day
def  wen(dt):
    return dt.weekday()==2

PX=PCAX("192.168.65.11")
start='2017-09-08'
end='2020-09-23'
sqlcmd="SELECT 日期,代號,名稱,收盤價 FROM 日選擇權行情表 WHERE 日期 BETWEEN '"+start+"' and '" +end+ \
    "' and 代號 in ('TXPOB000','TXPOB100','TXPOS000','TXPOS100','TXXOB000','TXXOS000')"
sqltables="日選擇權行情表"
df1=PX.Sql_data(sqlcmd,sqltables)
df1['收盤價']=df1['收盤價'].apply(pd.to_numeric)
df1['日期']=df1['日期'].apply(lambda x:pd.to_datetime(x))
for string in ['近月','次月','週']:
    mask=df1['名稱'].str.contains(string)
    name_list=df1[mask].index
    df1.loc[name_list,'類別']=string
option_cp_value=df1.groupby(['類別','日期'])['收盤價'].sum()
option_cp_value=option_cp_value.reset_index()
option_cp_value.loc[(option_cp_value['類別']=='週') & (option_cp_value['日期'].apply(wen)),'結算日']=True
option_cp_value.loc[(option_cp_value['類別']!='週') & (option_cp_value['日期'].apply(third_wen)),'結算日']=True
option_cp_value.loc[:,'結算日']=option_cp_value.loc[:,'結算日'].fillna(False)

sqlcmd="SELECT 日期,代號,最高價,最低價,收盤價 FROM 期貨交易行情表 WHERE 日期 BETWEEN '"+start+"' and '" +end+ \
    "' and 代號 ='TX'"
sqltables="期貨交易行情表"
TX=PX.Sql_data(sqlcmd,sqltables)
TX[['最高價','最低價','收盤價']]=TX[['最高價','最低價','收盤價']].apply(pd.to_numeric)
TX['日期']=TX['日期'].apply(lambda x:pd.to_datetime(x))
TX['振幅']=TX['最高價']-TX['最低價']
dict_=TX.set_index('日期')['振幅']

週選=option_cp_value.query("類別=='週' and 結算日==True ")
週選['end']=週選['日期'].shift(-1).copy()
週選=週選.reset_index(drop=True)

月選=option_cp_value.query("類別=='近月'")
月選=月選.set_index('日期',drop=True)['收盤價']
from datetime import timedelta

df=pd.DataFrame(index=週選.index) 
for index,(start,end) in 週選[['日期','end']].dropna().iterrows():
    df.loc[index,'平均振福']=dict_.loc[start+timedelta(days=1):end].mean()
    df.loc[index,'振福標準差']=dict_.loc[start+timedelta(days=1):end].std()
    df.loc[index,'月選c+p']=月選[start]
週選=pd.concat([週選,df],axis=1)

import matplotlib. pyplot as plt
plt.rcParams['font.sans-serif'] = 'Arial Unicode MS'
fig2, ax = plt.subplots(figsize=(24,10))
ax.plot(週選['日期'],週選['收盤價'],label='c+p',marker = ".")
ax.plot(週選['日期'],週選['平均振福'],label='平均振福',marker = ".")
ax.plot(週選['日期'],週選['月選c+p'],label='月選',linestyle=':',marker = ".")
plt.xticks(rotation=90)
ax.legend()
plt.show()


temp=週選.set_index('日期').copy()
x=temp['收盤價']/temp['月選c+p']
x=x[x!=1].dropna()
fig, ax1 = plt.subplots()
plt.xlabel('date')
ax2 = ax1.twinx()
ax1.set_ylabel('比率', color='tab:blue')
ax1.plot(x, color='tab:blue',marker = "o", alpha=0.75)
ax1.tick_params(axis='y', labelcolor='tab:blue')
ax2.set_ylabel('點數', color='black')
ax2.plot(週選['日期'], 週選['平均振福'],marker = "o", color='black', alpha=1)
ax2.tick_params(axis='y', labelcolor='black')
ax1.axhline((x[x!=1].dropna()).mean(),color='r')
plt.show()