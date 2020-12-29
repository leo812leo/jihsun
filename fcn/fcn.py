# =============================================================================
# """Base Parameters"""
# =============================================================================
day=126
today_str='20200921'
stock=['AAPL','GOOGL','PYPL']


import numpy as np
import datetime
import time
from sys import path
path.append(r'T:\WCFADOX_Python')
import pandas as pd
from WCFAdox import PCAX
from pandas.tseries.offsets import BDay
import matplotlib.pyplot as plt

PX=PCAX("192.168.65.11")
today=datetime.datetime.strptime(today_str,"%Y%m%d").date()
last_x_workday_str=datetime.date.strftime(today- BDay(day+10),"%Y%m%d")

future_x_workday_str=datetime.date.strftime(today+ BDay(day+10),"%Y%m%d")
data=PX.Pal_Data("美股日收盤還原表排行","D",last_x_workday_str,today_str,colist="日期,代號,收盤價,[漲幅(%)]",isst='Y')

#data=PX.Pal_Data("美股日收盤還原表排行","D",today_str,future_x_workday_str,colist="日期,代號,收盤價,[漲幅(%)]",isst='Y')


data=data.query("代號 in @stock")
data[['收盤價','漲幅(%)']]=data[['收盤價','漲幅(%)']].apply(pd.to_numeric)
temp1=pd.pivot_table(data,index="日期",columns="代號",values="收盤價").dropna()

temp1=temp1.iloc[-day:]
temp1=temp1.rolling(2).apply(lambda x: np.log(x[1]/x[0])).dropna()
Corr=temp1.corr().round(3)
Hv=(temp1.std()*(252**0.5)).round(4)
del temp1,stock,last_x_workday_str,today_str,data,today
#%%
start_time = time.time()
# =============================================================================
# """Base Parameters"""
# =============================================================================
"Parameters Price Process"
d     = 3      # no. of underlyings
S_0   = 100     # Initial Value of Assets at first of february
K     = 62.4     # Strike Price
K_in  = 50
K_out = 90
r     = 0.12/100
mu    = (r* np.ones([1,d]))
sigma=Hv.to_numpy()
coupon=0.1
面額=100
"Parameters Monte Carlo"
N     = 10**5  # Number of Monte Carlo samples
"""Construct Time Parameter"""
dT  = 1/252
T=1/dT
num_of_dt=63

"Construct Covariance Matrix and Decomposition"
corr = Corr
L = np.linalg.cholesky(corr)
"Construct Brownian Motion step"
Delta_W =[ np.matmul(np.sqrt(dT) * L, np.random.normal(0, 1, (d,num_of_dt ))) for i in range(N)]
Delta_W=np.array(Delta_W)
def run(num_of_mn):
    out=[]
    out.append(np.array([100,100,100]))
    KO=False
    KI=False
    S_0=100
    values=0
    a=set()
    for num_dt in range(num_of_dt):
        S = S_0*np.exp((mu - 0.5 * sigma**2) * dT + sigma *Delta_W[num_of_mn,:,num_dt] )  
        if (num_dt+1) %21==0: #配息21的倍數
            values += (面額*coupon*1/12)/ ((1+r/T)**(num_dt+1)) #利息現值
            
        if (num_dt+1)>21 and (KO==False or KI==False):   #開始記憶 coupon_time>=1
            a.update( set((S[0]>=K_out).nonzero()[0].tolist()) )
            if len(a)==d:
                KO=True   
            if (S[0]<=K_out).any():
                KI=True      
        if (num_dt+1)>21 and KO: 
            values=  values + (面額+ 面額*coupon/12 * ((num_dt+1)%21)/21  ) / ((1+r/T)**(num_dt+1))
            return  values
        
        elif (num_dt+1) == num_of_dt:    #沒被KO且到期
           if (S.min()<K and KI):
               values=   values + ((面額/K)*S[0][S.argmin()])/ ((1+r/T)**(num_dt+1))
               return values
           
           else:
               values = values+ (面額)  / ((1+r/T)**(num_dt+1))
               return values
        out.append(S[0])
        S_0 = S.copy()
def mn(N):
    final=[]
    for i in range(N):
         final.append(run(i))
    return final
ans=mn(N)
price=np.mean(ans)
plt.hist(ans,bins=100)

print("--- %s seconds ---" % np.round((time.time() - start_time), 2))
