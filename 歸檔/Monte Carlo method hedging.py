from random import gauss
from math import exp, sqrt
import matplotlib.pyplot as plot
import numpy as np
import scipy.stats as st
import pandas as pd
plot.rcParams['font.sans-serif'] = 'Arial Unicode MS'

def N(x):
    N=st.norm.cdf(x,scale=1)    
    return N
def dN(X):
    dN=st.norm.pdf(X,scale = 1) 
    return dN
def option_cal(s,k,r,sigma,t):
    d1=(np.log(s/k)+(r+0.5*sigma**2)*t)/(sigma*np.sqrt(t))
    d2=d1-sigma*np.sqrt(t)
    c=s*N(d1)-k*np.exp(-r*t)*N(d2)
    return c 
def delta_cal(s,k,r,sigma,t):
    d1=(np.log(s/k)+(r+0.5*sigma**2)*t)/(sigma*np.sqrt(t))
    delta=np.exp(-r*t)*N(d1) 
    return delta
def generate_asset_price(S,sigma,r,T):
    return S * exp((r - 0.5 * sigma**2) * T + sigma * sqrt(T) * gauss(0,1.0))
#------------input------------------#
S = 90 
sigma = 0.99#option
r = 0.008
T = 1/ 252
K = 100

sigma_real=0.1
sigma_hedge=0.99
simulations=180
#----------------------------------#
for i in range(30):
    ss=[]
    for i in range(simulations):
        if i==0:
            S_T = S
        else:
            S_T = generate_asset_price(ss[-1],sigma_real,r,T)
        ss.append(S_T)
    
    # plot.figure(figsize=(7.5,5))  
    # plot.plot(ss)
    # plot.plot([0,simulations],[K,K])
    # plot.ylabel('stock price',fontsize=16)
    # plot.xlabel('days',fontsize=16)
    # plot.show()




    delta_list=[]
    c_list=[]
    for s,t in zip(ss,np.arange(simulations,0,-1)/252):
        option=option_cal(s,K,r,sigma,t)
        delta=delta_cal(s,K,r,sigma_hedge,t)
        c_list.append(option)
        delta_list.append(delta)
        
        
    plot.figure(1,figsize=(7.5,5))    
    plot.plot(delta_list,label=str(sigma_hedge))
    plot.ylabel('delta',fontsize=16)
    plot.xlabel('days',fontsize=16)
    # plot.legend(loc='upper left') 
    # plot.figure(figsize=(7.5,5))    
    # plot.plot(c_list)
    # plot.ylabel('warrant price',fontsize=16)
    # plot.xlabel('days',fontsize=16)
    
    
    #動態避險 每天盤後補滿delta oi穩定
    c_list=pd.Series(c_list)
    delta_list=pd.Series(delta_list)
    ss=pd.Series(ss)
    ss_change=(ss-ss.shift()).fillna(0)
    
    權證損益=(c_list.shift()-c_list).fillna(0) #第一天無損益 
    避險部位=delta_list
    避險部位損益=避險部位*ss_change             #最添梧避險部位損益
    每日損益=權證損益+避險部位損益
    每日累積損益=每日損益.cumsum()
    
    plot.figure(2,figsize=(7.5,5))
    #plot.plot(權證損益,label='權證損益')
    #plot.plot(避險部位損益,label='避險部位損益')
    plot.plot(每日累積損益,label='每日累積損益'+str(sigma_hedge))
plot.ylabel('p&l',fontsize=16)
plot.xlabel('days',fontsize=16)
# plot.legend(loc='upper left') 
# plot.show()



