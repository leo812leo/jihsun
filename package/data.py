import datetime
from sys import path
path.append(r'T:\WCFADOX_Python')
from WCFAdox import PCAX
import pandas as pd
from tool import calculate,workingday
def 分類(x,類別,table_for_divide):
    if 類別=='波動率':
        name='隱含波動率(委買價)(%)'
    elif 類別=='價格':
        name='權證收盤價'        
    stock=x['標的代號']
    divide_for_stock=table_for_divide[(類別,stock)]
    if x[name]>=divide_for_stock[0.7]:
        output=3
    elif x[name]<divide_for_stock[0.7] and x[name]>=divide_for_stock[0.3]:
        output=2        
    elif x[name]<divide_for_stock[0.3]:
        output=1  
    return  output 
def data_request(today_str,類別=False,百分比=False):
    """
    Parameters
    ----------
    today_str : TYPE
        DESCRIPTION.
    類別 : boolean, optional
        category data. The default is False.
    百分比 : boolean, optional
        bid_vol and warrant_price data require. The default is False.

    Returns
    -------
    table : object
        table of warrant include id,name,warrant_close,bid_vol...etc.

    """
    PX=PCAX("192.168.65.11")
    df0=PX.Mul_Data("日常用技術指標表","D",today_str,colist='股票代號,[EWMA波動率(%)]')
    df0=df0.set_index("股票代號")
    df0=df0.apply(pd.to_numeric)
    
    col="代號,名稱,權證收盤價,[發行數量(千)],價內金額,標的收盤價,最新履約價,發行價格,\
        [近一月歷史波動率(%)],[近三月歷史波動率(%)],[近六月歷史波動率(%)],權證成交量,最後委買價,標的名稱,[隱含波動率(委買價)(%)]"
    df1=PX.Mul_Data("權證評估表","D",today_str,colist=col)
    df1=df1.set_index("代號")
    
    col="代號,發行日期,到期日期,標的代號,最新執行比例,發行機構名稱"
    year=str(datetime.date.today().year)
    df2=PX.Mul_Data("權證基本資料表","Y",year,colist=col)
    df2=df2.set_index("代號")
    data1=pd.concat([df1,df2],axis=1,join='inner')
    data1=data1[~((data1.index.str[-1]=='B') | (data1.index.str[-1]=='X') | (data1.index.str[-1]=='C') | (data1.index.str[-1]=='Q') | (data1.index.str[-1]=='F') )]
    to_num=data1.columns[1:12].to_list()+  [data1.columns[13]]+[data1.columns[-2]]
    data1[to_num]=data1[to_num].apply(pd.to_numeric)
    data1[['發行日期','到期日期']]=data1[['發行日期','到期日期']].apply(pd.to_datetime)
    for string in ['發行日期','到期日期']:
        data1[string]=data1[string].dt.date
    data1=data1[~(data1['標的代號']=='IX0001')]
    data1['平均歷史波動率']=data1[['近一月歷史波動率(%)','近三月歷史波動率(%)','近六月歷史波動率(%)']].mean(axis=1)
    data1['EWMA']=data1['標的代號'].map(lambda x: df0.loc[x].values[0])
    
    
    data1=data1.drop(columns=['近一月歷史波動率(%)','近三月歷史波動率(%)','近六月歷史波動率(%)'])
    # table2
    col="股票代號,自營商庫存,自營商買賣超,[自營商買賣超金額(千)]"
    df3=PX.Pal_Data("日自營商進出排行","D",today_str,today_str,colist=col)
    df3=df3.set_index("股票代號")
    df4=PX.Mul_Data("日收盤表排行","D",today_str,colist="股票代號,[成交金額(千)]",ps="<CM代號,權證>")
    df4=df4.set_index("股票代號")
    data2=pd.concat([df3,df4],axis=1,join='inner')
    data2=data2.apply(pd.to_numeric)
    table=pd.concat([data1,data2],axis=1,join='inner')
    table['flag']=table.index.str[-1].to_series().apply(lambda  s: s.lower() if s=='P' else 'c' ).to_list()
    table['價內外']=table[['標的收盤價','最新履約價','flag']].apply(calculate ,axis=1)
    # =============================================================================
    table['t_modify']=table['到期日期'].apply(lambda s: workingday(datetime.datetime.strptime(today_str,"%Y%m%d").date(),s)/250)
    #%%類別資料
    if 類別==True:
        台灣50=PX.Mul_Data("日收盤表排行","D",today_str,colist="股票代號",ps="<CM特殊,1>").iloc[:,0].values
        上櫃=PX.Mul_Data("日收盤表排行","D",today_str,colist="股票代號",ps="<CM代號,2>").iloc[:,0].values
        上市=PX.Mul_Data("日收盤表排行","D",today_str,colist="股票代號",ps="<CM代號,1>").iloc[:,0].values
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
    #%%
    if 百分比==True:
        divide=table.groupby(['標的代號'])['隱含波動率(委買價)(%)','權證收盤價'].quantile([0.3,0.7])
        divide.columns=['波動率','價格']
        temp=divide.unstack(0).to_dict()
        table['vol_分類']=table[['標的代號','隱含波動率(委買價)(%)']].apply(lambda x: 分類(x,'波動率',temp),axis=1)
        table['價格_分類']=table[['標的代號','權證收盤價']].apply(lambda x: 分類(x,'價格',temp),axis=1)
    return table