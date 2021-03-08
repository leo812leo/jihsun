import datetime
from sys import path
path.append(r'C:\Users\F128537908\Desktop\WCFADOX_Python')
path.append("..")
from WCFAdox import PCAX
import pandas as pd
from pandas.tseries.offsets import BDay
from tool import workingday,calculate,TheoryPrice_BidVol_Theta
def fun1(df):
    x=df.sum().nlargest(20)
    x=x.reset_index();x=x['標的名稱'].to_list()
    df_temp=df.copy()
    df_temp.columns=df_temp.columns.levels[-1]
    df_temp=df_temp.loc[:,x]
    return df_temp
#%%
# =============================================================================
# 一.資料下載
# 1.table1(權證評估表+權證基本資料表)
# 2.table2(日收盤表排行+日自營商進出排行)
# 3.類別資料
# =============================================================================
# table1

PX=PCAX("192.168.65.11")
date_str = datetime.datetime.strftime(datetime.date.today()-BDay(1),'%Y%m%d')

df0=PX.Mul_Data("日常用技術指標表","D",date_str,colist='股票代號,[EWMA波動率(%)]')
df0=df0.set_index("股票代號")
df0=df0.apply(pd.to_numeric)

col="代號,名稱,權證收盤價,[發行數量(千)],價內金額,標的收盤價,最新履約價,發行價格,\
    [近一月歷史波動率(%)],[近三月歷史波動率(%)],[近六月歷史波動率(%)],權證成交量,最後委買價,標的名稱,[隱含波動率(委買價)(%)]"
df1=PX.Mul_Data("權證評估表","D",date_str,colist=col)
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
table=table[table['t_modify']>20/250]

# =============================================================================
table['波動率pr']=(table.groupby(['標的代號','flag'])['隱含波動率(委買價)(%)'].rank(pct=True,method='max')*100).round(2)
table=table[table['flag']=='c']
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
table=TheoryPrice_BidVol_Theta(table,date_str)
#獲利面
table['估計今日獲利']=(-1000*table['自營商買賣超金額(千)'])-(table['自營商買賣超']*-1*1000*table['理論價'])
df=pd.DataFrame(index=table.index)
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

del index,values,價內金額,市價,理論價,委買價,在外張數
#%%  3.table輸出    

temp2=table.groupby(['標的代號'])['權證收盤價'].count()
stock=temp2[temp2>20].index.to_list()
table=table[table['標的代號'].isin(stock)]


divide=table.groupby(['標的代號'])['隱含波動率(委買價)(%)','權證收盤價'].quantile([0.3,0.7])
divide.columns=['波動率','價格']
temp3=divide.unstack(0).to_dict()


filter_for_warrant="發行機構名稱 not in ('元大證','凱基證')"
table2=table.query(filter_for_warrant)


#warrant=table2.groupby(['標的代號'])['theta'].nlargest(30).index.levels[-1]

def 分類(x,類別,table_for_divide=temp3):
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


table2['vol_分類']=table2[['標的代號','隱含波動率(委買價)(%)']].apply(lambda x: 分類(x,'波動率'),axis=1)
table2['價格_分類']=table2[['標的代號','權證收盤價']].apply(lambda x: 分類(x,'價格'),axis=1)
final=table2.pivot_table(index=['標的代號','vol_分類'],columns=['價格_分類'],values=['theta'],aggfunc='mean').fillna(0)
broker=table2.pivot_table(index=['發行機構名稱','vol_分類','價格_分類'],columns=['類別'],values=['theta'],aggfunc='sum').fillna(0)



#%%   
ETF=set(table2.query("類別=='ETF'")['標的代號'])
T50=set(table2.query("類別=='上市台灣50'")['標的代號'])
SE=set(table2.query("類別=='上市非台灣50'")['標的代號'])
OTC=set(table2.query("類別=='上櫃'")['標的代號'])

for i,stock_type in enumerate(['ETF','T50','SE','OTC']):
    temp=final.loc[list(eval(stock_type))]
    if i ==0:
        temp.to_excel('分類結果.xlsx',sheet_name='ETF')
    else:
        with pd.ExcelWriter('分類結果.xlsx', engine='openpyxl',mode='a') as writer:    
            temp.to_excel(writer,stock_type)   
#%%
final.sort_index(level=0, ascending=False, inplace=True)
import seaborn as sns
import matplotlib.pyplot as plot
plot.rcParams['font.sans-serif'] = 'Arial Unicode MS'

for i in stock:
    sns.heatmap(final.loc[i],square=True,cmap='YlOrRd', linewidths=1.3, linecolor='black', center=0.00,
                annot=True, fmt ='2.0f',xticklabels=['低','中','高'],yticklabels=['高','中','低'])
    plot.xlabel('價格')
    plot.ylabel('vol')
    plot.title(i)
    if i in ETF:
        fig_dir=".\\圖片"+'\\'+'ETF'+'\\'+i+'.jpg'
    elif i in T50:
        fig_dir=".\\圖片"+'\\'+'T50'+'\\'+i+'.jpg'
    elif i in SE:
        fig_dir=".\\圖片"+'\\'+'上市'+'\\'+i+'.jpg'
    elif i in OTC:
        fig_dir=".\\圖片"+'\\'+'上櫃'+'\\'+i+'.jpg'        
    plot.savefig(fig_dir)
    plot.clf()


