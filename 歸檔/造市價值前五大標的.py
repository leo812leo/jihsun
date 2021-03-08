import datetime
from sys import path
path.append(r'C:\Users\F128537908\Desktop\WCFADOX_Python')
path.append("..")
from WCFAdox import PCAX
import pandas as pd
from pandas.tseries.offsets import BDay
from tool import Premium_cal,workingday,calculate,implied_volatility,TheoryPrice_BidVol_Theta
import numpy as np
#%%
# =============================================================================
# 一.資料下載
# 1.table1(權證評估表+權證基本資料表)
# 2.table2(日收盤表排行+日自營商進出排行)
# 3.類別資料
# =============================================================================
# table1
PX=PCAX("192.168.65.11");date_str=datetime.datetime.strftime(datetime.date.today()-BDay(1),'%Y%m%d')
df0=PX.Mul_Data("日常用技術指標表","D",date_str,colist='股票代號,[EWMA波動率(%)]')
df0=df0.set_index("股票代號")
df0=df0.apply(pd.to_numeric)

col="代號,名稱,權證收盤價,[發行數量(千)],價內金額,標的收盤價,最新履約價,發行價格,\
    [近一月歷史波動率(%)],[近三月歷史波動率(%)],[近六月歷史波動率(%)],權證成交量,最後委買價"
df1=PX.Mul_Data("權證評估表","D",date_str,colist=col)
df1=df1.set_index("代號")

col="代號,發行日期,到期日期,標的代號,最新執行比例,發行機構名稱"
year=str(datetime.date.today().year)
df2=PX.Mul_Data("權證基本資料表","Y",year,colist=col)
df2=df2.set_index("代號")
data1=pd.concat([df1,df2],axis=1,join='inner')
data1=data1[~((data1.index.str[-1]=='B') | (data1.index.str[-1]=='X') | (data1.index.str[-1]=='C') | (data1.index.str[-1]=='Q') | (data1.index.str[-1]=='F') )]
to_num=data1.columns[1:12].to_list()+ [data1.columns[-2]]
data1[to_num]=data1[to_num].apply(pd.to_numeric)
data1[['發行日期','到期日期']]=data1[['發行日期','到期日期']].apply(pd.to_datetime)
for string in ['發行日期','到期日期']:
    data1[string]=data1[string].dt.date

data1['平均歷史波動率']=data1[['近一月歷史波動率(%)','近三月歷史波動率(%)','近六月歷史波動率(%)']].mean(axis=1)
data1['EWMA']=data1['標的代號'].map(lambda x:df0.loc[x].values[0])
data1=data1.drop(columns=['近一月歷史波動率(%)','近三月歷史波動率(%)','近六月歷史波動率(%)'])
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
table['價內外']=table[['標的收盤價','最新履約價','flag']].apply(calculate ,axis=1)
stock=['1309',
       '1589',
       '1712',
       '2031',
       '2331',
       '2348',
       '2478',
       '2605',
       '3056',
       '3450',
       '4104',
       '4763',
       '5203',
       '6120',
       '6715']
table=table.query(" 標的代號 not in @stock")


del df3,df4,data1,data2,year,df1,df2,col,to_num,string
#%%
# 類別資料
台灣50=PX.Mul_Data("日收盤表排行","D",date_str,colist="股票代號",ps="<CM特殊,1>").iloc[:,0].values
上櫃=PX.Mul_Data("日收盤表排行","D",date_str,colist="股票代號",ps="<CM代號,2>").iloc[:,0].values
上市=PX.Mul_Data("日收盤表排行","D",date_str,colist="股票代號",ps="<CM代號,1>").iloc[:,0].values
#分類
上市台灣50=list(set(台灣50) & set(上市));上市非台灣50=list(set(上市) - set(台灣50));上櫃=list(上櫃)
dic={}
for stock in table['標的代號'].drop_duplicates().to_list():
    if stock in 上市台灣50:
        dic[stock]='上市台灣50'
    elif stock in 上市非台灣50:
        dic[stock]='上市非台灣50'
    elif stock in 上櫃:
        dic[stock]='上櫃'
    elif stock[0:2]=='00':
        dic[stock]='ETF'
    else:
        dic[stock]='指數'        
table['類別']=table['標的代號'].apply(lambda s: dic[s])
del 台灣50,上櫃,上市,dic,上市台灣50,上市非台灣50,stock
#%%
table=TheoryPrice_BidVol_Theta(table,date_str)
# =============================================================================
table['在外市值']=table['自營商庫存']*-1*table['權證收盤價']*1000
table['發行金額']=table['發行價格']*table['發行數量(千)']*1000
df=pd.DataFrame(index=table.index)
for index,values in table[['價內金額','權證收盤價','理論價','最後委買價','自營商庫存']].iterrows():
    價內金額,市價,理論價,委買價,在外張數=values
    在外張數=-在外張數
    if 價內金額>0:
        df.loc[index,'市價時間價值']=(市價-價內金額)*在外張數*1000
        df.loc[index,'理論價時間價值']=(理論價-價內金額)*在外張數*1000
        df.loc[index,'委買價時間價值']=(委買價-價內金額)*在外張數*1000
    elif 價內金額<=0:
        df.loc[index,'市價時間價值']=市價*在外張數*1000
        df.loc[index,'理論價時間價值']=理論價*在外張數*1000
        df.loc[index,'委買價時間價值']=委買價*在外張數*1000         
df['市價委買價偏離金額']=df['市價時間價值']-df['委買價時間價值']
df['市價理論價偏離金額']=df['市價時間價值']-df['理論價時間價值']
table=pd.concat([table,df],axis=1,join='inner')

#filter_for_warrant="最後委買價>0.3"
#table=table.query(filter_for_warrant)

del index,values,價內金額,市價,理論價,委買價,在外張數
#%%
dict_for_broker={}
xx=table.pivot_table(index=['發行機構名稱'],columns=['標的代號'],values=['theta'],aggfunc='sum')
for broker in xx.index.to_list(): 
    temp=xx.loc[broker,:].sort_values(ascending=False).dropna()
    temp=temp/temp.sum()*100
    stock=(temp[:temp.cumsum().searchsorted(70)+1].index.get_level_values(1).to_list()).copy()
    dict_for_broker[broker]=stock

df=pd.DataFrame()
for broker in xx.index.to_list():
    stock_list=dict_for_broker[broker]
    df=df.append(table.query("發行機構名稱 == @broker and 標的代號 in @stock_list"))
temp=df.groupby(['發行機構名稱','標的代號'])['theta'].sum().round(0)
final=df.groupby(['發行機構名稱','標的代號'])['市價理論價偏離金額','市價時間價值'].apply(lambda x: x['市價理論價偏離金額'].sum()/x['市價時間價值'].sum())                       
final.name='造市價值'
final=pd.DataFrame(final)
final['theta']=temp
final=final.reset_index()\
      .sort_values(['發行機構名稱','造市價值'], ascending=False).set_index(['發行機構名稱','標的代號'])
final.to_excel('造勢價值_明細.xlsx')

final=final.reset_index()
a=final.groupby('發行機構名稱')[['標的代號','造市價值','theta']].apply(lambda x: x.nlargest(5,'theta'))
a=a.droplevel(-1)
a=a.reset_index()
a['theta']=a.apply(lambda x: temp[(x['發行機構名稱'],x['標的代號'])], axis=1)
a=a.set_index(['發行機構名稱','標的代號'])
a.to_excel('造勢價值_fifth.xlsx')

