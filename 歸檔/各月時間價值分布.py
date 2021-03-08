import datetime
from sys import path
path.append(r'C:\Users\F128537908\Desktop\WCFADOX_Python')
from WCFAdox import PCAX
import pandas as pd

PX=PCAX("192.168.65.11")
col="代號,發行日期"
df1=PX.Pal_Data("權證基本資料表","Y",'2019','2020',colist=col)
df1=df1[(df1.代號.str[-1]!='B') & (df1.代號.str[-1]!='X') & (df1.代號.str[-1]!='C')]
df1['發行日期']=pd.to_datetime(df1['發行日期'])
df1=df1.set_index('發行日期')
dict1={}
for i in range(1,7):
    dict1[i]=df1.loc['2020-'+str(i)].count().iloc[0]

    
start = datetime.date(2020,1,1)
end = datetime.date.today()
bussiness_days_rng =pd.date_range(start, end, freq='BM')
expiration=[datetime.date.strftime(i,"%Y%m%d") for i in bussiness_days_rng]
expiration[1]='20200227'
expiration_index=[i[4:] for i in expiration]
def calculate_time_values():
    global data
    for index,values in data[['價內金額','權證收盤價','自營商庫存']].iterrows():
        價內金額,市價,在外張數=values
        在外張數=-在外張數
        if 價內金額>0:
            data.loc[index,'市價時間價值']=(市價-價內金額)*在外張數*1000
        elif 價內金額<=0:
            data.loc[index,'市價時間價值']=市價*在外張數*1000
dict2={}            
for  i,date_str in  enumerate(expiration):         
    col="代號,權證收盤價,價內金額"
    df2=PX.Mul_Data("權證評估表","D",date_str,colist=col)
    df2=df2[(df2.代號.str[-1]!='B') & (df2.代號.str[-1]!='X') & (df2.代號.str[-1]!='C')]
    df2=df2.set_index('代號')        
    
    col="股票代號,自營商庫存"
    df3=PX.Mul_Data("日自營商進出排行","D",date_str,colist=col)
    df3=df3.set_index("股票代號")
    data=pd.concat([df2,df3],axis=1,join='inner')
    data=data.apply(pd.to_numeric)
    calculate_time_values()
    dict2[i+1]=[data['市價時間價值'].sum(),data.count()[0]]



final=pd.concat([pd.Series(dict1,name='本月發行檔數'),
                 pd.DataFrame(dict2,index=['市價時間價值','檔數']).T],
                axis=1)
final.index=expiration_index
final['平均每檔']=final['市價時間價值']/final['檔數']
# final.to_excel('final.xlsx')




mean=final.iloc[-1,-1]
sigma=data['市價時間價值'].std()
output=pd.value_counts(pd.cut(data['市價時間價值'],[mean-sigma,mean,mean+sigma,mean+2*sigma,data['市價時間價值'].max()]))


import seaborn as sns
import matplotlib.pyplot as plt
plt.rcParams['font.sans-serif'] = 'Arial Unicode MS'

fig, ax = plt.subplots()
sns.violinplot(data['市價時間價值']/10000,jitter=1)
ax.set_xlabel('市價時間價值(萬)',fontsize=16)
plt.show()







