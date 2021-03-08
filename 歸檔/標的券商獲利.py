import datetime
import numpy as np
from sys import path
path.append(r'C:\Users\F128537908\Desktop\WCFADOX_Python')
path.append(r'D:\code')
from WCFAdox import PCAX
import pandas as pd
from pandas.tseries.offsets import BDay
import matplotlib.pyplot as plot
plot.rcParams['font.sans-serif'] = 'Arial Unicode MS'
from tool import Premium_cal,workingday,implied_volatility,calculate
def pl(flag,s,k,r,sigma,t,ratio,oi):
    option_pl=-(Premium_cal(flag,s,k,r,sigma,(t-1/252),ratio)-Premium_cal(flag,s,k,r,sigma,t,ratio))*1000*oi
    return option_pl
#%%
# =============================================================================
# 一.資料下載
# 1.table1(權證評估表+權證基本資料表)
# 2.table2(日收盤表排行+日自營商進出排行)
# 3.類別資料
# =============================================================================
# table1
PX=PCAX("192.168.65.11");date_str=datetime.datetime.strftime(datetime.date.today()-BDay(5),'%Y%m%d')
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
data1['EWMA']=data1['標的代號'].map(lambda x:  df0.loc[x].values[0])


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
#%% 產業名稱
#col='股票代號,產業代號,產業名稱,產業指數代號,產業指數名稱,指數彙編分類代號,指數彙編分類'
dic={}
col='股票代號,指數彙編分類'
stock_data=PX.Mul_Data("上市櫃公司基本資料","Y",str(datetime.date.today().year),colist=col)
stock_data=stock_data.set_index('股票代號')
dic=stock_data.to_dict()['指數彙編分類']

temp=table.query('類別 != "ETF" and 類別 !="指數" ' )
table['產業名稱']=temp['標的代號'].map(dic)
table['產業名稱']=table['產業名稱'].fillna("無")

del col,stock_data,temp,dic
#%%
# =============================================================================
# 二.
# 1.總表計算
'''計算
1.1. 理論價計算
1.2. 賣超金額、估計今日獲利、在外市值、發行金額
1.3. 市價時間價值、理論價時間價值、委買價時間價值
1.4. 市價委買價偏離金額、市價理論價偏離金額
''' 
# 2.交易員計算
'''
1.資料下載
2.建立Stock2Trader表
'''
#3.table輸出
'''
1. pivot_table(index=['發行機構名稱'])
2. 理論到期獲利率
3.格式
'''
#4.交易員輸出
'''
1.全市場(交易員標的)
2.日盛(交易員標的)
'''
# =============================================================================
#1.TABLE計算

#1.1理論價計算
temp=[];temp2=[];bidvol_list=[]
for index,values in table[['標的收盤價','最新履約價','EWMA','到期日期','flag','最新執行比例','自營商庫存','最後委買價']].iterrows():
    s,k,sigma,enddate,flag,ratio,在外張數,bid= values
    在外張數=-在外張數
    sigma=sigma/100
    r=0
    t=workingday(datetime.datetime.strptime(date_str,"%Y%m%d").date(),enddate)/252
    bidvol=implied_volatility(bid, s, k, t, r, flag,ratio)
    temp.append(Premium_cal(flag,s,k,r,sigma,t,ratio))
    if bidvol!=0:
        temp2.append(pl(flag,s,k,r,bidvol,t,ratio,在外張數))
        bidvol_list.append(bidvol)
    else:
        temp2.append(pl(flag,s,k,r,sigma,t,ratio,在外張數))
        bidvol_list.append(np.nan)
        print(index)
table['理論價'] = temp;table['theta'] = temp2;table['bid_vol']=bidvol_list

'''計算
1.賣超金額
2.估計今日獲利
3.在外市值
4.發行金額
''' 
#1.2
table['賣超金額']=table[['自營商買賣超','權證收盤價']].apply(lambda col:  col[0]*col[1] if col[0]<0 else 0,axis=1)
table['估計今日獲利']=(-1000*table['自營商買賣超金額(千)'])-(table['自營商買賣超']*-1*1000*table['理論價'])
table['在外市值']=table['自營商庫存']*-1*table['權證收盤價']*1000
table['發行金額']=table['發行價格']*table['發行數量(千)']*1000
df=pd.DataFrame(index=table.index)
#1.3
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
# =============================================================================
table['t_modify']=table['到期日期'].apply(lambda s: workingday(datetime.date.today(),s)/250)
table['價內外']=table[['標的收盤價','最新履約價','flag']].apply(calculate ,axis=1)
# =============================================================================
del index,values,s,k,sigma,enddate,flag,ratio,價內金額,市價,理論價,委買價,在外張數,r,t,temp
#%%  3.table輸出 
stock=['1907','2610','2618','2881','2882','1904']
temp=table.query("標的代號  in @stock")


final=table[table['產業名稱'].str.contains('傳產|金融', regex=True)]
list_fot_table=['市價時間價值','理論價時間價值','市價理論價偏離金額','成交金額(千)','自營商買賣超金額(千)','發行金額','theta']
dic={i:'sum' for i in list_fot_table};dic.update({'發行金額':['count']})
out=final.pivot_table(index=['產業名稱'],values=list_fot_table,aggfunc=dic)
out=out.droplevel(-1,axis=1)
out=out.rename(columns={'發行金額':'檔數'})
out.to_excel("金融傳產.xlsx")

out2=table.pivot_table(index='發行機構名稱',columns='標的代號',values='市價理論價偏離金額',aggfunc='sum')
out2.to_excel("券商標的明細(市價-理論).xlsx")

out3=temp.pivot_table(index='標的代號',values=list_fot_table,aggfunc=dic)
out3=out3.droplevel(-1,axis=1)
out3.to_excel("特殊標的.xlsx")


