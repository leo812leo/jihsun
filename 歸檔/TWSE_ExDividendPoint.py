from sys import path
path.append(r'C:\Users\F128537908\Desktop\WCFADOX_Python')
from datetime import datetime,timedelta
import pandas as pd
import numpy as np
from WCFAdox import PCAX
from workalendar.asia import Taiwan
from pandas.tseries.offsets import BDay 
import re
from itertools import chain
#%%
df=pd.read_csv('https://www.twse.com.tw/holidaySchedule/holidaySchedule?response=csv&queryYear='+\
               str(datetime.now().year-1911),header=1,encoding='big5')
tem=[re.findall(r'(\d+月\d+日)',i.values[0])for _,i in df[['日期']].iterrows()]
holiday=[(datetime.strptime(str(datetime.now().year)+i, "%Y%m月%d日")).date() for i in chain(*tem)]

def format_for_cmoney(year,Q):
    return str(year) +"{:0>2d}".format(Q)

def table(today):
    Y=today.year
    if (datetime(Y,1,1)<=today) & (today<=datetime(Y,3,31)):
        season=1
    elif (datetime(Y,3,31)<today) & (today<=datetime(Y,6,30)):
        season=2
    elif (datetime(Y,6,30)<today) & (today<=datetime(Y,9,30)):
        season=3
    else:
        season=4
    return season
def get_near_Q(today_X):
    if isinstance(today_X, datetime):
        Y=today_X.year
        Q=table(today_X)
        if Q==1:
            YY=Y-2;QQ=4
        else :
            YY=Y-1;QQ=Q-1
    return format_for_cmoney(YY,QQ)
def estimate(day_str,延後天數):
    return datetime.strptime(day_str, "%Y%m%d")+timedelta(days=延後天數)
def dataframe(data,cols):
    ans=[]
    for col in cols:
        deviation=data[col].apply(pd.Series)
        ans.append(deviation.merge(data['代號'], right_index = True, left_index = True)\
                            .melt(id_vars = ['代號'], value_name = col)\
                            .drop("variable", axis = 1).dropna())
    final=pd.merge(ans[0],ans[1], right_index = True, left_index = True, on=['代號'])
    return final
#設定連線主機IP並產生物件
PX=PCAX("192.168.65.11")
today=datetime.now()

if today.weekday()==0:
    yesterday=datetime.now()- BDay(2)
else:
    yesterday=datetime.now()- BDay(1)
yesterday_str=datetime.strftime(yesterday,"%Y%m%d")

#%%權重計算
加權=PX.Pal_Data("日收盤表排行","D",yesterday_str,yesterday_str,ps='<CM代號,1>')
加權=加權.set_index("股票代號")
加權=加權.loc[:,['收盤價','總市值(億)']]
加權=加權.apply(lambda s:pd.to_numeric(s, errors='coerce'))
加權['權重']=加權['總市值(億)']/加權['總市值(億)'].sum()

MSCI=PX.Pal_Data("日收盤表排行","D",yesterday_str,yesterday_str,ps='<CM特殊,2>')
MSCI=MSCI.set_index("股票代號")
MSCI=MSCI.loc[:,['收盤價','總市值(億)']]
MSCI=MSCI.apply(lambda s:pd.to_numeric(s, errors='coerce'))
MSCI['權重']=MSCI['總市值(億)']/MSCI['總市值(億)'].sum()


total=PX.Pal_Data("日收盤表排行","D",yesterday_str,yesterday_str,colist='股票代號,收盤價')
加權指數=float(total[total['股票代號']=='TWA00'].values[0][1])
MSCI指數=float(total[total['股票代號']=='RN099'].values[0][1])

# =============================================================================
# 股利政策-日期
# =============================================================================
#%%發一次
預測今年=True
if 預測今年:
    #天數差計算(只算發一次的)
    舊股東會預告=PX.Pal_Data("股東會預告","y",str(today.year-1),str(today.year-1));舊股東會預告=舊股東會預告.set_index("股票代號")
    舊股東會預告=舊股東會預告[舊股東會預告['臨時(常)會']=='常會']
    舊股東會=舊股東會預告['股東會日期']
    舊股東會=舊股東會.apply(lambda s: pd.to_datetime(s))
    
    天數預估=PX.Pal_Data("股利政策表","y",str(today.year-2), str(today.year-2),colist='股票代號,配發次數,除息日',ps='<CM代號,1>');index=天數預估['股票代號'].values
    天數預估=天數預估.loc[天數預估['配發次數']=='1']
    天數預估['除息日']=天數預估['除息日'].apply(lambda s: pd.to_datetime(s));天數預估=天數預估.set_index("股票代號")
    天數=(天數預估['除息日']-舊股東會).mean().days
    del 舊股東會預告,天數預估,舊股東會
    #除息日以去年推(只算發一次的)
    #年資料(可能部分有誤)
    股利政策表=PX.Pal_Data("股利政策表","y",str(today.year-2), str(today.year),ps='<CM代號,1>');股利政策表=股利政策表.set_index("股票代號")
    舊股利政策表=股利政策表.loc[股利政策表['年度']==str(today.year-2),['股票名稱','配發次數','除息日','現金股利合計(元)','現金股利發放率(%)']]
    舊股利政策表=舊股利政策表.loc[舊股利政策表['配發次數']=='1']
    以去年除息估計=舊股利政策表['除息日'].map(lambda v: '' if v=='' else datetime.strptime(v,"%Y%m%d")+timedelta(days=365))
    以去年除息估計=以去年除息估計.map(lambda d: d+timedelta(days=2) if d.weekday()==5 else  
                                            d+timedelta(days=1) if d.weekday()==6 
                                            else d)
    #由股東會推估
    股東會預告=PX.Pal_Data("股東會預告","y",str(today.year),str(today.year))
    股東會預告=股東會預告.set_index("股票代號")
    股東會預告=股東會預告[股東會預告['臨時(常)會']=='常會']
    股東會預告=股東會預告['股東會日期']
    #日期對到工作日
    常會估計=股東會預告.map(lambda d: estimate(d,天數+2) if estimate(d,天數).weekday()==5 else  
                   estimate(d,天數+1) if estimate(d,天數).weekday()==6 
                   else estimate(d,天數))
    del 股東會預告,天數
    #新公布日期(年資料來看)
    新股利政策公布_年=股利政策表.loc[股利政策表['年度']==str(today.year),['股票名稱','配發次數','除息日','現金股利合計(元)','現金股利發放率(%)']]
    新日期公布_年=新股利政策公布_年[新股利政策公布_年['除息日']!=""]
    新日期公布_年=新日期公布_年[新日期公布_年['配發次數']=='1']
else:
    #天數差計算(只算發一次的)
    舊股東會預告=PX.Pal_Data("股東會預告","y",str(today.year),str(today.year));舊股東會預告=舊股東會預告.set_index("股票代號")
    舊股東會預告=舊股東會預告[舊股東會預告['臨時(常)會']=='常會']
    舊股東會=舊股東會預告['股東會日期']
    舊股東會=舊股東會.apply(lambda s: pd.to_datetime(s))
    
    天數預估=PX.Pal_Data("股利政策表","y",str(today.year-1), str(today.year-1),colist='股票代號,配發次數,除息日',ps='<CM代號,1>');index=天數預估['股票代號'].values
    天數預估=天數預估.loc[天數預估['配發次數']=='1']
    天數預估['除息日']=天數預估['除息日'].apply(lambda s: pd.to_datetime(s));天數預估=天數預估.set_index("股票代號")
    天數=(天數預估['除息日']-舊股東會).mean().days
    del 舊股東會預告,天數預估,舊股東會
    #除息日以去年推(只算發一次的)
    #年資料(可能部分有誤)
    股利政策表=PX.Pal_Data("股利政策表","y",str(today.year-1), str(today.year),ps='<CM代號,1>');股利政策表=股利政策表.set_index("股票代號")
    舊股利政策表=股利政策表.loc[股利政策表['年度']==str(today.year-1),['股票名稱','配發次數','除息日','現金股利合計(元)','現金股利發放率(%)']]
    舊股利政策表=舊股利政策表.loc[舊股利政策表['配發次數']=='1']
    以去年除息估計=舊股利政策表['除息日'].map(lambda v: '' if v=='' else datetime.strptime(v,"%Y%m%d")+timedelta(days=365))
    以去年除息估計=以去年除息估計.map(lambda d: d+timedelta(days=2) if d.weekday()==5 else  
                                            d+timedelta(days=1) if d.weekday()==6 
                                            else d)
    #由股東會推估
    股東會預告=PX.Pal_Data("股東會預告","y",str(today.year+1),str(today.year+1))
    股東會預告=股東會預告.set_index("股票代號")
    股東會預告=股東會預告[股東會預告['臨時(常)會']=='常會']
    股東會預告=股東會預告['股東會日期']
    #日期對到工作日
    常會估計=股東會預告.map(lambda d: estimate(d,天數+2) if estimate(d,天數).weekday()==5 else  
                   estimate(d,天數+1) if estimate(d,天數).weekday()==6 
                   else estimate(d,天數))
    del 股東會預告,天數
    #新公布日期(年資料來看)
    新股利政策公布_年=股利政策表.loc[股利政策表['年度']==str(today.year),['股票名稱','配發次數','除息日','現金股利合計(元)','現金股利發放率(%)']]
    新日期公布_年=新股利政策公布_年[新股利政策公布_年['除息日']!=""]
    新日期公布_年=新日期公布_年[新日期公布_年['配發次數']=='1']    


#%%不只發一次資料處理
'''不只發一次股利優先順序  1.已公布 2.去年除息日推估'''

季股利政策表=PX.Pal_Data("季股利政策表","Q",get_near_Q(today),format_for_cmoney(today.year,table(today)),ps='<CM代號,1>')
季股利政策表['現金股利合計(元)']=季股利政策表['現金股利合計(元)'].apply(lambda s:pd.to_numeric(s, errors='coerce') )
季股利政策表=季股利政策表[季股利政策表['現金股利合計(元)']!=0]
季股利政策表['除息日']=季股利政策表['除息日'].apply(lambda s:pd.to_datetime(s))
季股利政策表=季股利政策表.set_index(["股票代號","年季"])
股利超過一次=[]
股利超過一次_日期={}
#找股利超過一次
for stock in 季股利政策表.index.get_level_values(0).drop_duplicates():
    temp=季股利政策表.loc[(stock,slice(None)),'除息日']
    condiction=temp>datetime(datetime.now().year,1,1)
    if temp[condiction].count()>1:
        股利超過一次.append(stock)
#已公布
for stock in 股利超過一次:
    condiction=季股利政策表.loc[(stock,slice(None)),'除息日']>datetime(datetime.now().year+1,1,1)
    temp=季股利政策表.loc[(stock,slice(None)),'除息日'][condiction].map(lambda d: d+timedelta(days=2) if d.weekday()==5 else  
               d+timedelta(days=1) if d.weekday()==6 
               else d)
    股利超過一次_日期.setdefault(stock,list(temp.values))
#去年推估(先處理)
for stock in 股利超過一次:
    condiction=季股利政策表.loc[(stock,slice(None)),'除息日'].map(lambda c:c.year)==datetime.now().year      #取去年除息日資料
    去年推估=季股利政策表.loc[(stock,slice(None)),'除息日'][condiction].map(lambda c: c+timedelta(days=365))#取去年除息日+一年
    去年推估=去年推估.map(lambda d: d+timedelta(days=2) if d.weekday()==5 else  
                  d+timedelta(days=1) if d.weekday()==6 
                  else d)
    for i,stock_day in enumerate(去年推估.values):
        if not (not 股利超過一次_日期[stock]):                                                                #有公布 跟去年+一年日期比
          if min(abs(股利超過一次_日期[stock]-stock_day))  /  np.timedelta64(1, 'D') >=80:
             股利超過一次_日期[stock].append(stock_day)
        else:                                                                                               #無公布 直接用去年
            股利超過一次_日期[stock]=list(去年推估.values)
#%%輸出
"""如果公佈則已公布為主，如未公布先以去年除權日加一年(第2順位)，股東會+天數(第3順位)"""
output={}
for i,stock in enumerate(index):
    if (stock in 常會估計.index) and (stock not in 股利超過一次):
        output[stock]={'日期':常會估計[stock]}
    if (stock in 以去年除息估計.index) and (stock not in 股利超過一次):
        output[stock]={'日期':以去年除息估計[stock]}
    if (stock in 新日期公布_年.index) and (stock not in 股利超過一次):
        output[stock]={'日期':新日期公布_年.loc[stock,'除息日']}
    if stock in 股利超過一次:
        if not (股利超過一次_日期[stock]==[]):
            output[stock]={'日期':股利超過一次_日期[stock]}
        else:
            output[stock]={'日期':[]}
del 股利超過一次_日期,以去年除息估計,去年推估,常會估計,新日期公布_年


# =============================================================================
#點數
# =============================================================================
"""1.公布 2.eps yoy*上次股利 3.eps推估(財報公布)
    發超過一次另外處理
"""
#%%發一次
#新配息公布
新配息公布_年=新股利政策公布_年[新股利政策公布_年['現金股利合計(元)']!=0]
新配息公布_年=新配息公布_年.loc[:,['現金股利合計(元)','現金股利發放率(%)']]
filter1=季股利政策表.index.get_level_values(0).isin(股利超過一次)#股利超過一次
filter2=季股利政策表.index.get_level_values(0).isin(新配息公布_年.index.to_list())
filter_all= filter1 | filter2
#去年股利
舊配息推估=季股利政策表.loc[~filter_all,'現金股利合計(元)']
舊配息推估=舊配息推估.reset_index().set_index('股票代號')['現金股利合計(元)']
舊配息推估=舊配息推估.dropna()
#%%以eps_yoy估計
每股盈餘成長率=PX.Pal_Data("季IFRS財報(總表)","Q",'201903','201904',
              colist='年季,代號,[公告累計基本每股盈餘成長率(%)]',ps='<CM代號,1>')
每股盈餘成長率['公告累計基本每股盈餘成長率(%)']=每股盈餘成長率['公告累計基本每股盈餘成長率(%)'].apply(lambda s:pd.to_numeric(s, errors='coerce') )
每股盈餘成長率=每股盈餘成長率.set_index(['代號','年季']).dropna()
股利成長={}
for stock in index:
    if stock in 每股盈餘成長率.index.get_level_values(0):
        if 每股盈餘成長率.loc[stock].count()[0]==1:
            股利成長[stock]=每股盈餘成長率.loc[stock,'公告累計基本每股盈餘成長率(%)'].values[0]
        else:
            temp=每股盈餘成長率.loc[stock]
            condiction=temp.index=='201904'
            股利成長[stock]=temp.loc[condiction,'公告累計基本每股盈餘成長率(%)'].values[0]
股利成長=pd.Series(股利成長)
EPS預估股利=((1+股利成長/100)*舊配息推估).dropna()

#%%股利超過一次   
#已公布
股利超過一次_配息_公布={}
for stock in 股利超過一次:
    condiction=季股利政策表.loc[(stock,slice(None)),'除息日']>datetime.now()
    股利超過一次_配息_公布.setdefault(stock,list(季股利政策表.loc[(stock,slice(None)),'現金股利合計(元)'][condiction].values))
    if 股利超過一次_配息_公布[stock]==[]:
        股利超過一次_配息_公布[stock]=[季股利政策表.loc[(stock,slice(None)),'現金股利合計(元)'].values[-1]]


# ######debug######
股利政策改變=['2707','3557','5283','9188']

#error
################################
for st in 股利政策改變:
    output[st]={}
    
################################
for i,stock in enumerate(index):
    if (stock in 舊配息推估.index) and (stock not in 股利超過一次):
        if type(舊配息推估[stock])== pd.Series:
            股利政策改變.append(stock)                                          #去年股利發一次，但20194Q，卻有一個資料以上
        output[stock]['配息']=舊配息推估[stock]
    if (stock in 新配息公布_年.index) and (stock not in 股利超過一次):
        output[stock]['配息']=float(新配息公布_年.loc[stock,'現金股利合計(元)'])
    if stock in 股利超過一次:
        if not (股利超過一次_配息_公布[stock]==[]):                              #不適空白
            output[stock]['配息']=股利超過一次_配息_公布[stock]                  #用去年的
            
# for stock in 股利政策改變:
#     temp=季股利政策表.loc[(stock,slice(None)),:]
#     temp=temp[temp['除息日']>datetime.now()]
#     output[stock]['日期']=temp['除息日'].values[0]
#     output[stock]['配息']=temp['現金股利合計(元)'].values[0]


兩邊不相等=[]
for  i,[key,value] in enumerate(output.items()):
    a=list(value.values())        
    if len(value)!=2:
        print('缺項',key)
    else:
        if type(a[0])!=list:
            a[0]=[a[0]]
        if type(a[1])!=list:
            a[1]=[a[1]]
        if len(a[0])!=len(a[1]):
            bais=len(a[0])-len(a[1])
            兩邊不相等.append([key,bais])
            print('兩邊不相等',key)
for stock,bais in 兩邊不相等:
    base=[output[stock]['配息'][-1]]
    output[stock]['配息'].extend(base*bais)

# =============================================================================
# 算點數
# =============================================================================
data=pd.DataFrame.from_dict(output).T.dropna().reset_index()
data=data.rename(columns={'index':'代號'})

final=dataframe(data,['日期','配息'])
final['日期']=pd.to_datetime(final['日期'],unit='ns')
final=final.sort_values(by=['日期']).reset_index(drop=True)
final['點數影響_台指']=final['配息']/加權.loc[final.代號.values,'收盤價'].values*加權.loc[final.代號.values,'權重'].values*加權指數
final['點數影響_msci']=(final['配息']/MSCI.loc[final.代號.values,'收盤價'].values*MSCI.loc[final.代號.values,'權重'].values*MSCI指數).fillna(0)
#日期序列
weekmask = ' Mon Tue Wed Thu Fri'
dat_range=pd.bdate_range(start='2021/01/01', end='2021/12/31', freq='C', weekmask=weekmask, holidays=holiday)
#以日期序列重新編排整理
out_final={}
for day in dat_range:
    dday=datetime.strftime(day,"%Y-%m-%d")
    out_final.setdefault(dday,{})
    out_final[dday]['點數影響_台指']=sum(final.loc[final['日期']==dday,'點數影響_台指'])
    out_final[dday]['點數影響_msci']=sum(final.loc[final['日期']==dday,'點數影響_msci'])
    out_final[dday]['股票代號']=final.loc[final['日期']==dday,'代號'].values
     
#輸出
path = r'C:\Users\F128537908\Desktop\除息點數.xlsx'
writer = pd.ExcelWriter(path, engine = 'openpyxl',mode='w')
out_final=pd.DataFrame.from_dict(out_final, orient='index')
out_final.to_excel(writer,sheet_name='日期')
final.loc[:,'日期']=final.loc[:,'日期'].apply(lambda v : v.date())
final.set_index('代號').to_excel(writer,sheet_name='股票')
writer.save()