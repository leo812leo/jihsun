from sys import path
path.append(r'C:\Users\F128537908\Desktop\WCFADOX_Python')
from datetime import datetime
import pandas as pd
from WCFAdox import PCAX
from pandas.tseries.offsets import BDay 

PX=PCAX("192.168.65.11")
today=datetime.now()
today_str=datetime.strftime(today,"%Y%m%d")
str_list=[]
cluster={}

for day in [10,20,60]:
    str_list.append(datetime.strftime(datetime.now()- BDay(day) ,"%Y%m%d"))

pairs=[['0050','00632R','TX'],['0050_淨值','00632R_淨值','TX'],
       ['00637L','00655L','00638R'],['00637L_淨值','00655L_淨值','00638R_淨值']]
#pairs=[['0050','00632R','TX'],['00637L','00655L'],['1590','2049'],['1536','3665'],['2317','4938'],['2327','2492','2456'],['2408','2337','2344'],['2409','3481'],['2376','2337'],['2383','6274','6213'],['3105','2455','8086'],['3008','3406'],['5483','6488'],['3673','6456'],['6230','3324']]

dataname = {}
data_dict={}
for i,days in enumerate([10,20,60]):
  data=PX.Pal_Data("日收盤表排行","D",str_list[i],today_str,colist="日期,股票名稱,股票代號,收盤價,[漲幅(%)]",isst='Y',ps='<CM特殊,2011>')
  TX=PX.Pal_Data("期貨交易行情表","D",str_list[i],today_str,colist="日期,名稱,代號,收盤價,[漲幅(%)]");tx=TX[TX['代號']=='TX'];tx.columns=data.columns
  etf=PX.Pal_Data("ETF淨值還原表","D",str_list[i],today_str,colist="日期,股票名稱,股票代號,淨值,[報酬率(%)]")
  etf['股票代號']=etf['股票代號']+'_淨值'
  etf=etf.rename(columns={"淨值": "收盤價", "報酬率(%)": "漲幅(%)"})
  data_final=pd.concat([data,tx,etf],axis=0,ignore_index=True)
  dataname["D{}".format(days)]=data_final

for i,pair in enumerate(pairs):
    for key, value in dataname.items():
        value=value[value['股票代號'].isin(pair)]
        value.loc[:,"漲幅(%)"]=value.loc[:,"漲幅(%)"].astype(float).values
        a=pd.pivot_table(value,index="股票代號",columns="日期",values="漲幅(%)").T
        data_dict.setdefault(str(i),{}).update({key: a.corr()})
#        output.setdefault("pair"+str(i)+key,[]).append(a.corr())
final=[]
coff_dict={}
for i in range(len(pairs)):
    temp=list(data_dict[str(i)].values())
    final.append(pd.concat(temp).groupby(level=0).mean().round(3))
    coff_dict[str(i)]=pd.concat(temp).groupby(level=0).apply( lambda x:x.iloc[0]*x.iloc[2]/(x.iloc[1])**2 ).round(3)



path = r'C:\Users\F128537908\Desktop\output2.xlsx'
writer = pd.ExcelWriter(path, engine = 'openpyxl',mode='w')
for i,avg in enumerate(final):
    avg.to_excel(writer, sheet_name=str(i)+'-Avg')
    coff_dict[str(i)].to_excel(writer, sheet_name=str(i)+'-coff')
    data_dict[str(i)]['D60'].to_excel(writer, sheet_name=str(i)+'-60D')
writer.save()
