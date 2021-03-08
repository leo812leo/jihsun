import datetime
from sys import path
path.append(r'T:\WCFADOX_Python')
path.append('.\package')
from WCFAdox import PCAX
import pandas as pd
from pandas.tseries.offsets import BDay
import numpy as np
import re
from itertools import chain
PX=PCAX("192.168.65.11")
from tool import Premium_cal,workingday,calculate,last_x_workday

#%%日期
# =============================================================================
#過去的假日
df=pd.read_csv('https://www.twse.com.tw/holidaySchedule/holidaySchedule?response=csv&queryYear='+\
               str(datetime.datetime.now().year-1911),header=1,encoding='big5')
tem=[re.findall(r'(\d+月\d+日)',i.values[0])for _,i in df[['日期']].iterrows()]

df=pd.read_csv('https://www.twse.com.tw/holidaySchedule/holidaySchedule?response=csv&queryYear='+\
               str(datetime.datetime.now().year-1911-1),header=1,encoding='big5')
tem2=[re.findall(r'(\d+月\d+日)',i.values[0])for _,i in df[[' 日期']].iterrows()]

holiday=[(datetime.datetime.strptime(str(datetime.datetime.now().year)+i, "%Y%m月%d日")).date() for i in chain(*tem)]
holiday2=[(datetime.datetime.strptime(str(datetime.datetime.now().year-1)+i, "%Y%m月%d日")).date() for i in chain(*tem2)]
#未來的假日
today=datetime.date.today()    
day=pd.read_html('http://172.16.10.13/hedge/tradingday.php',encoding= 'big5')[-1].iloc[:,1].drop(0)
day=set(day.apply(lambda x : pd.to_datetime(x).date()).to_list())
date_range=pd.date_range(start=datetime.datetime.strftime(today,'%Y%m%d'),
              end=datetime.datetime.strftime(datetime.date(today.year+2,1,1),'%Y%m%d')).tolist()
date_range=set([day.date() for day in date_range])
temp=date_range-day
holiday.extend(temp)
holiday.extend(holiday2)
holiday=list(set(holiday))
holiday=[d  for d in holiday  if d.weekday() not in [5,6]]



# =============================================================================
# 日期
# =============================================================================
today=(datetime.date.today()-BDay(0)).date()
#%%資料獲取

#取資料(25天市場上權證資料)
col='日期,代號,名稱,標的代號,上市日期,到期日期,最新執行比例,[存續期間(月)],發行價格,標的收盤價,最新履約價,價內金額,權證收盤價,最後委買價,[隱含波動率(委買價)(%)]'
df1=PX.Pal_Data("權證評估表","D",datetime.datetime.strftime(last_x_workday(26,today),"%Y%m%d")
                                ,datetime.datetime.strftime(last_x_workday(1,today),"%Y%m%d"),colist=col) #25日全市場權證資料
#資料整理
df1=df1[~((df1.代號.str[-1]=='B') | (df1.代號.str[-1]=='X') | (df1.代號.str[-1]=='C'))]
df1['券商']=df1['名稱'].str.extract(r'(..)..[售購]')
df1['flag']=df1.代號.str[-1].apply(lambda  s: s.lower() if s=='P' else 'c' ).to_list()
#日盛標的篩選(上市日期20天)
temp=df1.query('券商=="日盛"')
temp['上市日期']=temp['上市日期'].apply(lambda x :pd.to_datetime(x))
start_date = datetime.datetime.strftime(last_x_workday(21,today),"%Y-%m-%d")
end_date   = datetime.datetime.strftime(last_x_workday(1,today),"%Y-%m-%d")
StockList=list(set(temp.query('上市日期>= @start_date and 上市日期 <= @end_date ')['標的代號'])) #觀察近20日日盛權證 發行標的
#其他加篩選(同標的、日期25日)
df1=df1.query('標的代號 in @StockList')
to_num=['最新執行比例','存續期間(月)','標的收盤價','最新履約價','價內金額','發行價格','權證收盤價','最後委買價','隱含波動率(委買價)(%)']
df1[to_num]=df1[to_num].apply(pd.to_numeric)
to_date=['日期','上市日期','到期日期']
df1[to_date]=df1[to_date].applymap(pd.to_datetime)
df1['價內外程度(%)']=df1[['標的收盤價','最新履約價','flag']].apply(calculate ,axis=1)
start_date = datetime.datetime.strftime(last_x_workday(26,today),"%Y-%m-%d")
end_date   = datetime.datetime.strftime(last_x_workday(1,today),"%Y-%m-%d") #25日 新發行權證資料 
df1=df1.query('上市日期>= @start_date and 上市日期 <= @end_date ')   

del to_num,temp,to_date,start_date,end_date,col
#%%發行分析
'''
觀察近20日日盛 新發權證
比較權證為近25日同業 同標的權證

最大時間價值 日期,參數
目前最大時間價值 參數
自家權證狀況
'''
#庫存
warrant=list(set(df1.代號))
col="日期,股票代號,自營商庫存,自營商買賣超"
df3=PX.Pal_Data("日自營商進出排行","D",datetime.datetime.strftime(last_x_workday(26,today),"%Y%m%d"),
                                      datetime.datetime.strftime(last_x_workday(1,today),"%Y%m%d"),colist=col)
df3=df3.query('股票代號 in @warrant')
df3[['自營商庫存','自營商買賣超']]=df3[['自營商庫存','自營商買賣超']].apply(pd.to_numeric)
df3=df3.rename(columns={'股票代號':'代號'})
to_date=['日期']
df3[to_date]=df3[to_date].applymap(pd.to_datetime)


#table(計算市價時間價值)
table=pd.merge(df1,df3, on=['代號','日期'],how='inner')
table[['自營商庫存','自營商買賣超']]=-table[['自營商庫存','自營商買賣超']]
table=table[table['隱含波動率(委買價)(%)']!=0]
del col,df3,to_date,df1
#%%col計算

for index,values in table[['價內金額','權證收盤價','最後委買價','自營商庫存','自營商買賣超']].iterrows():
    價內金額,市價,委買價,在外張數,買賣超=values
    if 價內金額>0:
        table.loc[index,'市價時間價值']=(市價-價內金額)*在外張數*1000
        table.loc[index,'市價時間價值增量']=(市價-價內金額)*買賣超*1000        
    elif 價內金額<=0:
        table.loc[index,'市價時間價值']=市價*在外張數*1000 
        table.loc[index,'市價時間價值增量']=(市價-價內金額)*買賣超*1000
        
for index,values in table[['日期','上市日期','到期日期']].iterrows(): 
    日期,上市日期,到期日期=values
    table.loc[index,'交易天數']=workingday(上市日期.date(),日期.date())+1
    table.loc[index,'距到期日']=workingday(日期.date(),到期日期.date())
    table.loc[index,'距到期日(日曆)']= (到期日期.date()-日期.date()).days


def pl(flag,s,k,r,sigma,t,ratio,oi,up_down):
    option_pl=-(Premium_cal(flag,s*(1+up_down/100),k,r,sigma,(t-1/365),ratio)-Premium_cal(flag,s,k,r,sigma,t,ratio))*1000*oi
    return option_pl

for index,values in table[['標的收盤價','最新履約價','隱含波動率(委買價)(%)','距到期日(日曆)','自營商庫存','flag','最新執行比例']].iterrows():
    s,k,sigma,t,oi,flag,ratio=values
    sigma=sigma/100;r=0;t=t/365
    table.loc[index,'theta']= int(pl(flag,s,k,r,sigma,t,ratio,oi,0))
    
#%%存量觀察

to_date=['日期','上市日期','到期日期']
table[to_date]=table[to_date].applymap(lambda x : x.date())
finnal_output={}
plot_out={}
table=table.drop(['價內金額','權證收盤價'],axis=1)
today2=last_x_workday(1,today)
日盛權證=table.query("券商=='日盛'")
日盛權證=日盛權證[日盛權證['日期']==today2]
# =============================================================================
# 1.25天內最大時間價值
# 2.目前最大時間價值
# 3.25天內最大OI
# 4.目前最大OI
# =============================================================================
def sort_jihsun_last(df):
    sort_list=list(set(df["券商"])-{'日盛'})
    sort_list.append("日盛")
    df['券商'] = pd.Categorical(df['券商'],sort_list)
    return df.sort_values('券商')
rename_dict={'最新執行比例':'行使比例','隱含波動率(委買價)(%)':'委買vol',
             '價內外程度(%)':'價內外','最新履約價':'履約價','最後委買價':'委買價','存續期間(月)':'存續(月)',
             '自營商庫存':'OI','自營商買賣超':'OI增量'}

sort_col=['日期','上市日期','到期日期',
          '券商','flag','行使比例','存續(月)',
          '發行價格','委買價','標的收盤價','履約價','價內外','委買vol','波動率pr',
         'OI','市價時間價值','theta','OI增量','市價時間價值增量',
         '交易天數',	'距到期日']
def make_table(col,target=table ,whole=table ,table_jisun=日盛權證):
    temp=target.loc[target.groupby("代號")[col].idxmax()]            #25天內時間價值最大時 權證參數
    temp=temp.groupby('標的代號').apply(lambda s: s.nlargest(3,col))    #25天內 同標的 時間價值最大權證參數
    temp['波動率pr']=temp['代號'].map(vol_pct_dict)
    #最大oi增量
    warrant=list(set(temp.代號))
    temp_x=whole.loc[whole.groupby("代號")['自營商買賣超'].idxmax()]
    temp_x=temp_x[temp_x['代號'].isin(warrant)]
    temp_x=temp_x[['標的代號','代號','日期','自營商買賣超']]
    temp_x=temp_x.set_index(["標的代號","代號"])
    
    temp=temp.reset_index(drop=True).set_index(["代號","名稱"])
    table_jisun['波動率pr']=table_jisun['代號'].map(vol_pct_dict)
    table_jisun=table_jisun.reset_index(drop=True).set_index(["代號","名稱"])
    temp=temp.append(table_jisun)
    temp=temp.groupby('標的代號').apply(sort_jihsun_last) #把日盛排在最後
    temp=temp.drop('標的代號',axis=1)
    temp=temp.rename(columns=rename_dict)
    temp=temp[sort_col]
    temp=temp[~temp.index.duplicated(keep='first')]
    return temp,temp_x
#資料處理_今日
now=table[table['日期']==today2]
now['波動率pr']=(now.groupby(['標的代號','flag'])['隱含波動率(委買價)(%)'].rank(pct=True)*100).round(2)
temp=now.set_index('代號')
vol_pct_dict=temp['波動率pr'].to_dict()
#最大時間價值 日期,參數
#finnal_output['最大時間價值']=make_table("市價時間價值")

#最大庫存 日期,參數
finnal_output['最大庫存'],plot_out['最大庫存']=make_table("自營商庫存")

#####目前最大時間價值 參數#####

#目前最大庫存 參數
finnal_output['目前最大庫存'],plot_out['目前最大庫存']=make_table("自營商庫存",target=now)   #2目前 同標的 時間價值最大權證參數


#%% 增量觀察
# =============================================================================
# 1.25天內最大時間價值增量
# 3.25天內最大OI增量
# =============================================================================
#最大時間價值增量 日期,參數
finnal_output['最大時間價值增量'],plot_out['最大時間價值增量']=make_table("市價時間價值增量")


#最大庫存增量 日期,參數
finnal_output['最大庫存增量'],plot_out['最大庫存增量']=make_table("自營商買賣超")   #25天內 同標的 時間價值最大權證參數

#另外算
count_out=[]
for sheet in ['最大庫存','目前最大庫存','最大時間價值增量','最大庫存增量']:
    x=finnal_output[sheet].copy(deep=True)
    count=[len(x.loc[index].index.get_level_values(0).drop_duplicates())for index in x.index.levels[0].values ]
    count_out.append(np.cumsum(count))

#%%excel
i=0
for index,values in finnal_output.items():
    if i==0:
        values.to_excel(r'.\output專用\新發行權證分析\新發行權證追蹤.xlsx',sheet_name=index)
    else:
        writer = pd.ExcelWriter(r'.\output專用\新發行權證分析\新發行權證追蹤.xlsx', engine = 'openpyxl',mode='a')
        values.to_excel(writer, sheet_name=index)
        writer.save()
    i+=1


#%%圖
# =============================================================================
#公式
def last_x_workday(date,x,holiday=holiday):                              #假日
    last_workday=(date- BDay(x)).date()                              #往前x天用工作天推(last_workday)
    num_holiday=sum([(i<date) & (i>=last_workday) for i in holiday]) #期間內假日總數
    if num_holiday!=0:
        return last_x_workday(last_workday,num_holiday)
    return datetime.datetime.strftime(last_workday,"%Y%m%d"),datetime.datetime.strftime(last_workday,"%Y-%m-%d")
# =============================================================================
import os
for sheet in ['最大庫存','目前最大庫存','最大時間價值增量','最大庫存增量']:
    dirction= os.path.join(os.getcwd(),'output專用','新發行權證分析','fig',sheet)
    if not os.path.exists(dirction):
        os.makedirs(dirction)
    else:
        print(dirction + '已經建立！')
from plot_for_warrant import plot_candles_for_broker
import matplotlib.pyplot as plt


start=last_x_workday(today,100)[0]
import gc
for sheet in ['最大庫存','目前最大庫存','最大時間價值增量','最大庫存增量']:
    for stock in set(plot_out[sheet].index.levels[0]):
        df_warrant=plot_out[sheet].loc[stock]
        plot_candles_for_broker(start,today2.strftime("%Y%m%d"),stock,df_warrant)
        plt.savefig(os.path.join(os.getcwd(),'output專用','新發行權證分析','fig',sheet,stock+'.jpg'))
        plt.close('all')
        gc.collect()
#%%excel
path=r'.\output專用\新發行權證分析\新發行權證追蹤.xlsx'
import xlwings as xw 

app = xw.App(visible=True, add_book=False)
wb1 = app.books.open(r'T:\權證發行動態分析\權證發行動態分析_python\格式用\code_for_format.xlsm')
wb2 = app.books.open(path)
wb2.activate()
wb2.sheets[1].select()


for sheet_i in range(len(wb2.sheets)):
    wb2.sheets[sheet_i].select()    
    for time,i in enumerate(count_out[sheet_i]):
        if time==0:
            xw.Range((2,1),(i+1,3)).select()
            wb1.macro('format')()  
            
            xw.Range((2,4),(i+1,6)).select()
            wb1.macro('format')()
            
            xw.Range((2,7),(i+1,16)).select()
            wb1.macro('format')()
            
            xw.Range((2,16),(i+1,18)).select()
            wb1.macro('format')()
            
            xw.Range((2,18),(i+1,21)).select()
            wb1.macro('format')()
    
            xw.Range((2,21),(i+1,23)).select()
            wb1.macro('format')()

            xw.Range((2,23),(i+1,24)).select()
            wb1.macro('format')()
        else:
            xw.Range((x,1),(i+1,3)).select()
            wb1.macro('format')()  
            
            xw.Range((x,4),(i+1,6)).select()
            wb1.macro('format')()
            
            xw.Range((x,7),(i+1,16)).select()
            wb1.macro('format')()
            
            xw.Range((x,16),(i+1,18)).select()
            wb1.macro('format')()
            
            xw.Range((x,18),(i+1,21)).select()
            wb1.macro('format')()
    
            xw.Range((x,21),(i+1,23)).select()
            wb1.macro('format')()
            
            xw.Range((x,23),(i+1,24)).select()
            wb1.macro('format')()
        x=i+2
    rng=xw.Range((1,1))
    rng.current_region.autofit()
    wb1.macro('超連接')()
wb1.close()
wb2.save()
wb2.close()       
app.quit()
print('完成')