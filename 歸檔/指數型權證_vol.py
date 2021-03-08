import datetime
import numpy as np
from sys import path
path.append(r'C:\Users\F128537908\Desktop\WCFADOX_Python')
from WCFAdox import PCAX
import pandas as pd
from collections import defaultdict
from tool import workingday,implied_volatility
from pandas.tseries.offsets import BDay
from BlackScholes import delta_cal,gamma_cal
#%%公式
def closing_date(date):
    try :
        type(date)==datetime.datetime
        first_day=datetime.datetime(date.year,date.month,1)
        index=first_day.weekday()
        dic={0:17,1:16,2:15,3:21,4:20,5:19,6:18}
        weekday_for_month=dic[index]
    except :
        raise ValueError('input is not datetime')
    return weekday_for_month
def nearest(items, pivot):
    return min(items, key=lambda x: abs(x - pivot))
def point(future_day,warrant_day,df):
    if warrant_day>future_day:
        condiction1=df.index>future_day ; condiction2=df.index<=warrant_day
        result=df[condiction1 & condiction2]
        return -result['台指'].sum()
    elif warrant_day<future_day:
        condiction1=df.index<=future_day ; condiction2=df.index>warrant_day
        result=df[condiction1 & condiction2]
        return result['台指'].sum()
    elif warrant_day==future_day:
        return 0
def calculate(df):
    if df['flag']=='c':
        return round(( (df['期貨價位']-df['最新履約價']) /df['最新履約價'] )*100,2)
    elif df['flag']=='p':
        return round(( -(df['期貨價位']-df['最新履約價']) /df['最新履約價']  )*100,2)
    else:
        raise ValueError("數值錯誤")
#%%
# =============================================================================
# #期貨資料
# =============================================================================
PX=PCAX("192.168.65.11");date_str=datetime.datetime.strftime(datetime.date.today()-BDay(1),'%Y%m%d')
Future_data=PX.Pal_Data("期貨交易行情表","D",date_str,date_str)
year=str(datetime.datetime.now().year)
mask1=Future_data['代號'].apply(lambda x:x.startswith('TX'+year))
mask2=Future_data['代號'].str.find('／')==-1

Future_data=Future_data[mask1 & mask2]
Future_data=Future_data.set_index('代號')
Future_data['交割日']=Future_data['交割月份'].apply(lambda str_datr: datetime.datetime.strptime(
                        str_datr+str(
                            closing_date(datetime.datetime.strptime(str_datr+'01','%Y%m%d')))
                        ,'%Y%m%d').date())
items=Future_data['交割日'].values

#%%
#dict warrant to feature
temp=zip(Future_data.index.values,Future_data['收盤價'].values,Future_data['交割日'].values)
dic=defaultdict(dict)
for index,S,expdate in temp:
    dic[expdate]['spot']=S
    dic[expdate]['underlying']=index

# =============================================================================
# #除權息資料
# =============================================================================
exdividend_data=pd.read_html('http://172.16.10.11/hedge/indexyield.php',encoding='big5',index_col=0)[0]
exdividend_data.index=pd.to_datetime(exdividend_data.index).date
today=datetime.date.today().strftime('%Y%m%d')


# =============================================================================
# #權證資料
# =============================================================================
#preprocessing
col="代號,名稱,標的代號,到期日期,最後委買價,最後委賣價,最新履約價,最新執行比例"
warrant_data=PX.Pal_Data("權證評估表","D",date_str,date_str,colist=col)
warrant_data=warrant_data[warrant_data['標的代號']=='TWA00']
warrant_data[['最後委買價','最後委賣價','最新履約價','最新執行比例']]=\
warrant_data[['最後委買價','最後委賣價','最新履約價','最新執行比例']].apply(lambda s:pd.to_numeric(s))
warrant_data['到期日期']=warrant_data['到期日期'].apply(lambda s:datetime.datetime.strptime(s,'%Y%m%d').date())
warrant_data['t_modify']=warrant_data['到期日期'].apply(lambda s: workingday(datetime.date.today(),s)/252)
#filter
warrant_data=warrant_data[warrant_data['到期日期']>datetime.date.today()]
warrant_data=warrant_data.fillna(0)
warrant_data=warrant_data[warrant_data['最後委賣價']!=0]

#flag
warrant_data=warrant_data[~((warrant_data['代號'].str[-1]=='B') | (warrant_data['代號'].str[-1]=='C'))]
warrant_data['flag']=warrant_data['代號'].str[-1].apply(lambda  s: s.lower() if s=='P' else 'c' )
#期貨價位
expdate_future=warrant_data['到期日期'].apply(lambda day: nearest(items,day))
warrant_data['期貨代號']=expdate_future.map(lambda expdate: dic[expdate]['underlying'])
warrant_data['期貨價位']=expdate_future.map(lambda expdate: int(dic[expdate]['spot']))
warrant_data=warrant_data[np.logical_or(warrant_data['flag']=='c',warrant_data['flag']=='p')]

#point_for_each
startdates=expdate_future
enddates=warrant_data['到期日期']
point_for_each=[]
for startdate,enddate in zip(startdates,enddates):
    point_for_each.append(point(startdate,enddate,exdividend_data))
#合併
warrant_data['差異點數']=point_for_each
warrant_data['spot_modify']=warrant_data['差異點數']+warrant_data['期貨價位']
# =============================================================================
# vol計算
# =============================================================================
generator=zip(warrant_data['最後委賣價'].values,
              warrant_data['spot_modify'].values,
              warrant_data['最新履約價'].values,
              warrant_data['t_modify'].values,
              np.zeros(warrant_data.shape[0]), 
              warrant_data['flag'].values,
              warrant_data['最新執行比例'].values,
              warrant_data['最後委買價'].values,
              warrant_data.index.tolist())  
vol_for_each=[]
delta=[]
gamma=[]
for price, S, K, t, r, flag,ratio,bid,index in generator:
    vol=implied_volatility(price, S, K, t, r, flag,ratio)
    vol_for_each.append(vol)
    delta.append(delta_cal(S, K, r,vol,t)*S*1000*ratio)
    gamma.append(gamma_cal(S, K, r,vol,t)*0.01*(S**2)*1000)
    
warrant_data['askvol']=vol_for_each
warrant_data['期貨交割日']=startdates
warrant_data['delta']=delta
warrant_data['gamma']=gamma
warrant_data=warrant_data.set_index('代號')
warrant_data['價內外']=warrant_data[['期貨價位','最新履約價','flag']].apply(calculate ,axis=1)
#output
output=warrant_data[[ '名稱','flag','期貨價位','差異點數', 'spot_modify', '最新履約價','價內外','到期日期','期貨交割日','t_modify',
                     '最後委買價', '最後委賣價','askvol','delta','gamma']]
output.to_excel('test.xlsx')







