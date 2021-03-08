# -*- coding: utf-8 -*-
"""
Created on Mon Dec 21 12:06:42 2020

@author: F128537908
"""
from sys import path
path.append(r'..')
path.append(r'C:\Users\F128537908\Desktop\WCFADOX_Python')
import requests
import pandas as pd
from data import  data_request
from HV import vol
from tool import TheoryPrice_BidVol_Theta
from datetime import datetime
from pandas.tseries.offsets import BDay
from range_vol import run


t=1
date=(datetime.today()-BDay(t)).date()
date_str_1=date.strftime("%Y%m%d")
date_str_2=date.strftime("%Y/%m/%d")
del date,t
#%% 日盛全市場佔比
data=data_request(date_str_1)
table=TheoryPrice_BidVol_Theta(data,date_str_1)
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

table=table.query("發行機構名稱 in ['國泰綜合證','第一金證','富邦綜合證']")

out=table.groupby('發行機構名稱').apply(lambda x :x.nlargest(10,'市價理論價偏離金額'))
out.to_excel('out.xlsx')