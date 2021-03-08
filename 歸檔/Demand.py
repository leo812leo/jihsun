import datetime
import numpy as np
from sys import path
path.append(r'C:\Users\F128537908\Desktop\WCFADOX_Python')
from WCFAdox import PCAX
import pandas as pd
import re
from itertools import chain
from pandas.tseries.offsets import BDay
from tool import Premium_cal,workingday,calculate,implied_volatility

df=pd.read_csv('https://www.twse.com.tw/holidaySchedule/holidaySchedule?response=csv&queryYear='+\
               str(datetime.datetime.now().year-1911),header=1,encoding='big5')
tem=[re.findall(r'(\d+月\d+日)',i.values[0])for _,i in df[['日期']].iterrows()]
holiday=[(datetime.datetime.strptime(str(datetime.datetime.now().year)+i, "%Y%m月%d日")).date() for i in chain(*tem)]

def fun1(df):
    x=df.sum().nlargest(20)
    x=x.reset_index();x=x['標的名稱'].to_list()
    df_temp=df.copy()
    df_temp.columns=df_temp.columns.levels[-1]
    df_temp=df_temp.loc[:,x]
    return df_temp
del df,tem
#%%
# =============================================================================
# 一.資料下載
# 1.table1(權證評估表+權證基本資料表)
# 2.table2(日收盤表排行+日自營商進出排行)
# 3.類別資料
# =============================================================================
# table1
for i in range(1,7):
    exec('demand_'+str(i)+'={}')

PX=PCAX("192.168.65.11")
for i,date_str in enumerate([datetime.datetime.strftime(datetime.date.today()-BDay(1),'%Y%m%d'),datetime.datetime.strftime(datetime.date.today()-BDay(3),'%Y%m%d')]):  
    df0=PX.Mul_Data("日常用技術指標表","D",date_str,colist='股票代號,[EWMA波動率(%)]')
    df0=df0.set_index("股票代號")
    df0=df0.apply(pd.to_numeric)
    
    col="代號,名稱,權證收盤價,[發行數量(千)],價內金額,標的收盤價,最新履約價,發行價格,\
        [近一月歷史波動率(%)],[近三月歷史波動率(%)],[近六月歷史波動率(%)],權證成交量,最後委買價,標的名稱"
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
    data1['EWMA']=data1['標的代號'].map(lambda x: df0.loc[x].values[0])
    
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
# =============================================================================
    table['t_modify']=table['到期日期'].apply(lambda s: workingday(datetime.date.today(),s)/250)
    filter_for_warrant="價內外<10 and 最後委買價>0.3"
    table=table[table['t_modify']>90/250]
    table=table.query(filter_for_warrant)
# =============================================================================

    
    
    del df3,df4,data1,data2,year,df1,df2,col,to_num,string
    #%%類別資料
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
    col='股票代號,產業名稱'
    stock_data=PX.Mul_Data("上市櫃公司基本資料","Y",str((datetime.date.today()-BDay(1)).year),colist=col)
    stock_data=stock_data.set_index('股票代號')
    dic=stock_data.to_dict()['產業名稱']
    table['產業名稱']=table['標的代號'].map(dic)
    #%%
    #1.TABLE計算
    
    #1.1
    temp=[]
    temp2=[]
    def pl(flag,s,k,r,sigma,t,ratio,oi):
        option_pl=-(Premium_cal(flag,s,k,r,sigma,(t-1/252),ratio)-Premium_cal(flag,s,k,r,sigma,t,ratio))*1000*oi
        return option_pl
    
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
        else:
            temp2.append(pl(flag,s,k,r,sigma,t,ratio,在外張數))
            print(index)
    
    table['理論價'] = temp;table['theta'] = temp2
    
    '''計算
    1.賣超金額
    2.估計今日獲利
    3.在外市值
    4.發行金額
    ''' 
    #1.2
    table['估計今日獲利']=(-1000*table['自營商買賣超金額(千)'])-(table['自營商買賣超']*-1*1000*table['理論價'])
    df=pd.DataFrame(index=table.index)
    
    #1.3
    for index,values in table[['價內金額','權證收盤價','理論價','最後委買價','自營商庫存','flag']].iterrows():
        價內金額,市價,理論價,委買價,在外張數,flag=values
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
    
    del index,values,s,k,sigma,enddate,flag,ratio,價內金額,市價,理論價,委買價,在外張數,r,t,temp
    #%%  3.table輸出 
    #TABLE     
    list_fot_table=['市價時間價值','理論價時間價值','委買價時間價值','市價理論價偏離金額','市價委買價偏離金額'
                    ,'成交金額(千)','自營商買賣超金額(千)','估計今日獲利']
    
    final=table.pivot_table(index=['發行機構名稱'],values=list_fot_table,aggfunc='sum')
    columns_name=['自營商買賣超金額(千)'];name_list=['淨買賣超金額']
    final=final.rename(columns=dict(zip(columns_name,name_list)))
    
    final['理論到期獲利率']=final['市價理論價偏離金額']/final['市價時間價值']
    #格式
    transforms={'中國信託綜合證':'中信','香港商麥格理':'麥格理','第一金證':'第一金'}
    final.index=final.index.to_series()\
                .apply(lambda s: transforms[s] if s in transforms else s[:2] )\
                .to_list()
    
    index_range=['~-15','-15 ~ -10','-10 ~ -5','-5 ~ 0','0 ~ 5','5 ~ 10','10 ~ 15','15 ~']
    table['價內外區間']=pd.cut(table['價內外'],[-np.inf,-15,-10,-5,0,5,10,15,np.inf],labels=index_range)
    index_range_2=['~0.6','0.6 ~ 0.9','0.9 ~ 1.2','1.2 ~ 1.5','1.5 ~ 2','2 ~']
    table['價格區間']=pd.cut(table['權證收盤價'],[-np.inf,0.6,0.9,1.2,1.5,2,np.inf],labels=index_range_2)
    
    demand_1[i]=table.pivot_table(index=['發行機構名稱'],columns=['類別'],values=['市價時間價值','theta'],aggfunc=np.mean).fillna(0)
    demand_2[i]=table.pivot_table(index=['發行機構名稱'],columns=['價內外區間'],values=['市價時間價值','theta'],aggfunc=np.mean).fillna(0)
    demand_3[i]=table.pivot_table(index=['發行機構名稱'],columns=['產業名稱'],values=['市價時間價值','theta'],aggfunc=np.mean).fillna(0)
    demand_4[i]=table.pivot_table(index=['發行機構名稱'],columns=['標的名稱'],values=['theta'],aggfunc=np.mean).fillna(0)
    demand_5[i]=table.pivot_table(index=['發行機構名稱'],columns=['標的名稱'],values=['市價時間價值'],aggfunc=np.mean).fillna(0)
    demand_6[i]=table.pivot_table(index=['發行機構名稱'],columns=['價格區間'],values=['市價時間價值','theta'],aggfunc=np.mean).fillna(0)    
    
for i in range(1,7):
    #exec('demand_'+str(i)+'=demand_'+str(i)+'[1]'+'-demand_'+str(i)+'[0]') 
    exec('demand_'+str(i)+'=demand_'+str(i)+'[0]') 
    
    exec('demand_'+str(i)+'.index=demand_'+str(i)+'.index.to_series()\
         .apply(lambda s: transforms[s] if s in transforms else s[:2] ).to_list()')
demand_4=fun1(demand_4)
demand_5=fun1(demand_5)

demand_1.to_excel('test.xlsx')
with pd.ExcelWriter('test.xlsx', engine="openpyxl",mode='a') as writer:  
    for i in range(2,7):
        temp=eval('demand_'+str(i))
        temp.to_excel(writer, sheet_name=str(i))
writer.save()




del list_fot_table,columns_name,name_list,dic,transforms,df,holiday
#%%
xx=table.pivot_table(index=['標的名稱'],values=['市價時間價值','theta','市價理論價偏離金額'],aggfunc=['sum',np.mean]).fillna(0)
xx.to_excel('標的.xlsx')
x2=table.pivot_table(index=['發行機構名稱'],columns=['類別'],values=['市價時間價值'],aggfunc='count').fillna(0)
x2.to_excel('類別需求.xlsx')
x3=table.pivot_table(index=['類別'],values=['市價時間價值','theta'],aggfunc=['sum',np.mean]).fillna(0)
x3.to_excel('類別.xlsx')