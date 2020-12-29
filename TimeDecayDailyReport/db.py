import xlwings as xw 
import datetime
import sqlite3
import pandas as pd
import os

app = xw.App(visible=True, add_book=False)
wb1 = app.books.open(r'T:\權證發行人評等\python版資料\broker_data\備份檔\broker.xlsx')
conn = sqlite3.connect(r'T:\權證發行人評等\python版資料\發行評等.db')

sht=wb1.sheets[-1]
df0=sht.range('A1').options(pd.DataFrame,expand='table').value[['價差大小分數','價格穩定分數','報價數量分數','總分']]
df0=df0.reset_index()
name=sht.name
df0['date']=name[:4]+'-'+name[4:6]+'-'+name[6:]
transforms={'中國信託綜合證':'中信','香港商麥格理':'麥格理','第一金證':'第一金'}
df0['券商名稱']=df0['券商名稱']\
                .apply(lambda s: transforms[s] if s in transforms else s[:2] )\
                .to_list()
df0=df0.set_index(['date','券商名稱'])
df0.to_sql('券商', conn, if_exists='append')
wb1.close()
del df0
#%%交易員
wb1 = app.books.open(r'T:\權證發行人評等\python版資料\trader_data\備份檔\trader.xlsx')
sht=wb1.sheets[-1]
df0=sht.range('A1').options(pd.DataFrame,expand='table').value[['價差大小分數','價格穩定分數','報價數量分數','平均分數']]
df0=df0.reset_index()
name=sht.name
df0['date']=name[:4]+'-'+name[4:6]+'-'+name[6:]
df0=df0.set_index(['date','群組代碼'])
df0.to_sql('交易員', conn, if_exists='append')
wb1.close()
del df0
#%%標的
path=r'T:\權證發行人評等\python版資料\stock_data' 
list_path=os.listdir(path)
listdiration=[path+'\\'+string for string in list_path]
listdiration.sort(key=lambda x: os.path.getmtime(x))
i=listdiration[-1]

wb1 = app.books.open(i)
sht=wb1.sheets[0]
df0=sht.range('A1').options(pd.DataFrame,expand='table').value[['價差大小分數','價格穩定分數','報價數量分數','平均分數','交易群組']]
df0=df0.reset_index()
name=i[-19:]
df0['date']=name[:4]+'-'+name[4:6]+'-'+name[6:8]
df0=df0.rename(columns={'index':'stock_id'})
df0=df0.set_index(['date','stock_id'])
df0.to_sql('標的', conn, if_exists='append')
wb1.close()
del df0
#%%權證
path=r'T:\權證發行人評等\python版資料\warrant_data' 
list_path=os.listdir(path)
listdiration=[path+'\\'+string for string in list_path]
listdiration.sort(key=lambda x: os.path.getmtime(x))
i=listdiration[-1]
wb1 = app.books.open(i)
sht=wb1.sheets[0]
df0=sht.range('A1').options(pd.DataFrame,expand='table').value[['標的代號','群組代碼','價差大小分數','價格穩定分數']]
df0=df0.reset_index()
name=i[-21:]
df0['date']=name[:4]+'-'+name[4:6]+'-'+name[6:8]
df0=df0.set_index(['date','權證代號','標的代號'])
df0.to_sql('權證', conn, if_exists='append')
wb1.close()
app.quit()
del df0

