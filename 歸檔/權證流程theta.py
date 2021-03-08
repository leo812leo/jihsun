import datetime
import numpy as np
from sys import path
path.append(r'T:\WCFADOX_Python')
from WCFAdox import PCAX
import pandas as pd
from pandas.tseries.offsets import BDay
from tool import Premium_cal,workingday,implied_volatility,calculate
from datetime import timedelta
PX=PCAX("192.168.65.11")
#%%公式
def pl(flag,s,k,r,sigma,t,ratio,oi):
    option_pl=-(Premium_cal(flag,s,k,r,sigma,(t-1/252),ratio)-Premium_cal(flag,s,k,r,sigma,t,ratio))*1000*oi
    return option_pl
#%%
# =============================================================================
# =============================================================================
# input
t0=1 #0今日 1昨日
range_vol_ratio=0.9
exception_ratio=0.9
n_large=15 #總theta前幾大
##篩選條建1(算theta用)
filter_for_warrant="價內外<10 and 最後委買價>0.3 and 發行機構名稱 != '元大證'"
t_filter=60  #存續期間>60日 

# =============================================================================
# =============================================================================

today=datetime.date.today()-BDay(t0)
#%%標的篩選
#base1 熱門標的
from new_warrant_package import largest_underlying
base1=set(largest_underlying(t0,20))
#base2 技術面
df_stock=PX.Mul_Data("日收盤速選","D",datetime.datetime.strftime(today,'%Y%m%d'),colist='股票代號,[較20日均量縮放N%]',ps='<CM特殊,2011>')
df_stock['較20日均量縮放N%']=df_stock['較20日均量縮放N%'].apply(pd.to_numeric)
df_temp=PX.Mul_Data("日收盤表排行","D",datetime.datetime.strftime(today,'%Y%m%d'),colist='股票代號,[成交金額(千)]',ps='<CM特殊,2011>')
df_temp['成交金額(千)']=df_temp['成交金額(千)'].apply(pd.to_numeric)
df_temp=df_temp.dropna()
mask2=set(df_temp.loc[df_temp['成交金額(千)']>= df_temp['成交金額(千)'].quantile(0.33),'股票代號'].to_list())
base2=set(df_stock.loc[df_stock["較20日均量縮放N%"]  >= 20,'股票代號'].to_list()) & mask2

#base3 基本面
df_Financial_report=PX.Mul_Data("月營收速選","M",str(today.year)+"{0:02d}".format(today.month-1)
                                ,colist="股票代號,年月,[單月合併營收年成長(%)連N個月大於零]",ps='<CM特殊,2011>')
df_Financial_report['單月合併營收年成長(%)連N個月大於零']=\
    df_Financial_report['單月合併營收年成長(%)連N個月大於零'].apply(pd.to_numeric)
base3=set(df_Financial_report.loc[df_Financial_report['單月合併營收年成長(%)連N個月大於零']>=3,'股票代號'].to_list())
base = base1 | base2 | base3
del df_stock,df_temp,df_Financial_report,mask2
del base1,base2,base3
#%%
demand_1={}

for i,date_str in enumerate([datetime.datetime.strftime(today,'%Y%m%d'),
                             datetime.datetime.strftime(datetime.date.today()-BDay(1),'%Y%m%d')]):
    ## table1
    #日常用技術指標表
    df0=PX.Mul_Data("日常用技術指標表","D",date_str,colist='股票代號,[EWMA波動率(%)]')
    df0=df0.set_index("股票代號")
    df0=df0.apply(pd.to_numeric)
    #權證評估表
    col="代號,名稱,權證收盤價,價內金額,標的收盤價,最新履約價,發行價格,權證成交量,最後委買價,標的名稱"
    df1=PX.Mul_Data("權證評估表","D",date_str,colist=col)
    df1=df1.set_index("代號")
    #權證基本資料表
    col="代號,發行日期,到期日期,標的代號,最新執行比例,發行機構名稱"
    year=str(today.year)
    df2=PX.Mul_Data("權證基本資料表","Y",year,colist=col)
    df2=df2.set_index("代號")
    #concat
    data1=pd.concat([df1,df2],axis=1,join='inner')
    #前處理
    data1=data1[~((data1.index.str[-1]=='B') | (data1.index.str[-1]=='X') | (data1.index.str[-1]=='C') | (data1.index.str[-1]=='Q') | (data1.index.str[-1]=='F') )]
    to_num=data1.columns[1:8].to_list()+ [data1.columns[-2]]
    data1[to_num]=data1[to_num].apply(pd.to_numeric)
    data1[['發行日期','到期日期']]=data1[['發行日期','到期日期']].apply(pd.to_datetime)
    for string in ['發行日期','到期日期']:
        data1[string]=data1[string].dt.date
    data1['EWMA']=data1['標的代號'].map(df0['EWMA波動率(%)'])
    ## table2
    #日自營商進出排行
    col="股票代號,自營商庫存,自營商買賣超,[自營商買賣超金額(千)]"
    df3=PX.Pal_Data("日自營商進出排行","D",date_str,date_str,colist=col)
    df3=df3.set_index("股票代號")
    #日收盤表排行
    df4=PX.Mul_Data("日收盤表排行","D",date_str,colist="股票代號,[成交金額(千)]",ps="<CM代號,權證>")
    df4=df4.set_index("股票代號")
    data2=pd.concat([df3,df4],axis=1,join='inner')
    data2=data2.apply(pd.to_numeric)
    
    ##table1+table2
    table=pd.concat([data1,data2],axis=1,join='inner')
    #資料處理(flag、價內外、t_modify...)
    table['flag']=table.index.str[-1].to_series().apply(lambda  s: s.lower() if s=='P' else 'c' ).to_list()

    table['價內外']=table[['標的收盤價','最新履約價','flag']].apply(calculate ,axis=1)
# =============================================================================
    table=table.query(filter_for_warrant)
    table['t_modify']=table['到期日期'].apply(lambda s: workingday(datetime.datetime.strptime(date_str,'%Y%m%d').date(),s)/250)
    table['上市天數']=table['發行日期'].apply(lambda s: workingday(s,datetime.datetime.strptime(date_str,'%Y%m%d').date())-1)
    table=table[table['t_modify']>t_filter/250]
# =============================================================================
    del df3,df4,data1,data2,year,df1,df2,col,to_num,string
    ##類別資料
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
    ## TABLE計算
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
    #1.2
    df=pd.DataFrame(index=table.index)    
    for index,values in table[['價內金額','權證收盤價','理論價','最後委買價','自營商庫存','flag']].iterrows():
        價內金額,市價,理論價,委買價,在外張數,flag=values
        在外張數=-在外張數
        if 價內金額>0:
            df.loc[index,'市價時間價值']=(市價-價內金額)*在外張數*1000
            df.loc[index,'理論價時間價值']=(理論價-價內金額)*在外張數*1000
        elif 價內金額<=0:
            df.loc[index,'市價時間價值']=市價*在外張數*1000
            df.loc[index,'理論價時間價值']=理論價*在外張數*1000              
    df['市價理論價偏離金額']=df['市價時間價值']-df['理論價時間價值']
    table=pd.concat([table,df],axis=1,join='inner')
    
    del index,values,s,k,sigma,enddate,flag,ratio,價內金額,市價,理論價,委買價,在外張數,r,t,temp
    del temp2,bidvol_list,bidvol,bid,df,df0
    ##demand_1
    demand_1[i]=table.pivot_table(index=['標的名稱'],columns=['flag'],values=['theta'],aggfunc=np.mean).fillna(0)
    if i ==1:    
        demand_2=table.pivot_table(index=['標的名稱'],columns=['flag'],values=['theta'],aggfunc=np.sum).fillna(0) 
del t_filter,filter_for_warrant
for i in range(1,2):
    exec('demand_'+str(i)+'=demand_'+str(i)+'[1]'+'-demand_'+str(i)+'[0]')
del i
#%%
filter1=demand_1[demand_1>=demand_1.quantile(q=0.8)].iloc[:,0].dropna()
filter2=demand_2.iloc[:,0].nlargest(n_large)
underlying=(set(filter1.index) | set(filter2.index))

#his_vol
df0=PX.Mul_Data("日常用技術指標表","D",date_str,colist='股票名稱,股票代號,[近三月歷史波動率(%)]')
dict_for_underlying=df0.set_index('股票名稱')['股票代號'].to_dict()
#underlying
underlying=pd.DataFrame(list(underlying))
underlying['代號']=underlying[0].map(dict_for_underlying)
underlying=underlying.set_index('代號')
underlying=underlying[underlying.index.isin(base)]

df0=df0.set_index('股票代號')
stock_list=list(set(df0.index) & set(underlying.index))
his_vol=df0.reindex(stock_list)['近三月歷史波動率(%)']
his_vol=his_vol.apply(pd.to_numeric)
del df0
#%%
#range_vol
from  range_vol import run
range_df=run(t0=1)
range_df=range_df[60]*range_vol_ratio
stock_list=list(set(range_df.index) & set(underlying.index))
range_vol=range_df.reindex(stock_list)
del range_df
#%%
#last3_ask_vol
from 最低vol_package import last3_ask_vol
ask_vol=last3_ask_vol()
ask_vol=ask_vol.droplevel(level=-1)
stock_list=list(set(ask_vol.index) & set(underlying.index))
ask_vol=ask_vol.loc[stock_list]
ask_vol=pd.pivot_table(ask_vol,index='標的代號',values='askvol',aggfunc='mean')
#%%
#條件
median=table[table['上市天數']<=30].groupby('標的代號')['bid_vol'].apply(np.median)
median=median.reindex(underlying.index)*100

final=pd.concat([his_vol,range_vol,ask_vol],axis=1,sort=False)
final['max']=final.apply(max,axis=1)
final=pd.concat([final,median],axis=1,sort=False)
final.columns=['HV','range','ask_vol','max','median']
final_output=final[(final['median']>final['max']) | (final['median'].apply(np.isnan))]
#%%
#例外
stock_list=underlying.index.to_list()
x=str(stock_list)
x='('+x[1:-1]+')' 
start=datetime.datetime.strftime(datetime.date(today.year-10,today.month,today.day),'%Y-%m-%d')
end=datetime.datetime.strftime(today,'%Y-%m-%d')
sqlcmd="SELECT 股票代號,日期,[近三月歷史波動率(%)] FROM 日常用技術指標表 WHERE 日期 BETWEEN '"+start+"' and '" +end+"' and 股票代號 in "+x
sqltables="日常用技術指標表"
df_long_term=PX.Sql_data(sqlcmd,sqltables)
df_long_term['日期']=df_long_term['日期'].apply(pd.to_datetime)
df_long_term['近三月歷史波動率(%)']=df_long_term['近三月歷史波動率(%)'].apply(pd.to_numeric)
df_long_term=df_long_term.sort_values(['股票代號','日期'])
df=df_long_term.groupby(['股票代號']).apply(lambda x : x['近三月歷史波動率(%)'].rank(pct=True).iloc[-1])
exception=df[df>exception_ratio].index
del x,sqltables,sqlcmd,start,end,df_long_term
del ask_vol,range_vol,his_vol
del underlying,stock_list,range_vol_ratio,exception_ratio
#%%
from best_parameter_package  import Best_sale_warrant
StockList= list(set(final_output.index) | set(exception))
parameter=Best_sale_warrant(StockList,t0)
parameter['最大時間價值增量'].to_excel('test3.xlsx')