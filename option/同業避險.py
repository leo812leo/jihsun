import matplotlib.pyplot as plot
from scipy.stats import norm
import math
import numpy as np

def d2_cal(s,k,r,sigma,t):
    d1=(np.log(s/k)+(r+0.5*sigma**2)*t)/(sigma*np.sqrt(t))
    d2=d1-sigma*np.sqrt(t)
    return d2
def N(x):
    N=norm.cdf(x,scale=1)    
    return N
def dN(X):
    dN=norm.pdf(X,scale = 1)
    return dN

def delta_cal(s,k,r,sigma,t):
    d1=(np.log(s/k)+(r+0.5*sigma**2)*t)/(sigma*np.sqrt(t))
    delta=np.exp(-r*t)*N(d1) 
    return delta
def gamma_cal(s,k,r,sigma,t):
    d1=(np.log(s/k)+(r+0.5*sigma**2)*t)/(sigma*np.sqrt(t))
    gamma=np.exp(-r*t)*dN(d1)/(s*sigma*np.sqrt(t)) 
    return gamma

def bscall_cal(s,k,r,sigma,t):
    d1=(np.log(s/k)+(r+0.5*sigma**2)*t)/(sigma*np.sqrt(t))
    d2=d2_cal(s,k,r,sigma,t)
    call=s*N(d1)-k*np.exp(-r*t)*N(d2)
    return call
 
def bsput_cal(s,k,r,sigma,t):
    call=bscall_cal(s,k,r,sigma,t)
    put=k*np.exp(-r*t)+call-s
    return put



"""example 1  我方權證:隱含波動率:60%、價外10% ；同業隱含波動率20~100%
              情況:漲10%      劃出P/L"""
        
t=1;s=100;k=s*(1-0.1);r=0#皆相同
sigma_self=0.6
sigma_competitors=np.arange(0.2,1,0.01)


#intial
my_warrant_initial         =bscall_cal(s,k,r,sigma_self,t)
competitors_warrant_initial=bscall_cal(s,k,r,sigma_competitors,t)
my_warrant_delta           =delta_cal(s,k,r,sigma_self,t)
competitors_warrant_delta  =delta_cal(s,k,r,sigma_competitors,t)
num_of_competitors=my_warrant_delta/competitors_warrant_delta

#final
my_warrant_final         =bscall_cal(s*1.1,k,r,sigma_self,t)
competitors_warrant_final=bscall_cal(s*1.1,k,r,sigma_competitors,t)
#P/L
income_my =-(           my_warrant_final          -    my_warrant_initial)*1
income_competitors =  (competitors_warrant_final -    competitors_warrant_initial)*num_of_competitors
total_income=  income_my +  income_competitors






plot.figure(1) 
#plot.subplot(1,2,1)
plot.plot(sigma_competitors*100,total_income)
plot.ylabel('profit & loss',fontsize=16)
plot.xlabel('implied vol',fontsize=16)
plot.title('P/L of different vol',fontsize=30) 

plot.figure(2) 
#plot.subplot(1,2,2)
gamma_my=gamma_cal(s,k,r,sigma_self,t)
gamma_competitors=gamma_cal(s,k,r,sigma_competitors,t)
g1=plot.plot(sigma_competitors*100,[gamma_my]*len(gamma_competitors), label='my warrant')
g2=plot.plot(sigma_competitors*100,gamma_competitors, label='competitor warrant')
plot.ylabel('intial gamma',fontsize=16)
plot.xlabel('implied vol',fontsize=16)
plot.title('gamma different vol',fontsize=30)
plot.legend(loc='upper right')        

#%%
        
"""example 2  我方權證:隱含波動率:60%、價外10% ；同業隱含波動率60%      
              情況:漲1~10%      劃出P/L"""
t=1;s=100;r=0;sigma=0.6
k=s*(1-0.1);k_competitors=s*(1-np.arange(0,0.1,0.001))

#intial
my_warrant_initial         =bscall_cal(s,k,r,sigma,t)
competitors_warrant_initial=bscall_cal(s,k_competitors,r,sigma,t)

my_warrant_delta           =delta_cal(s,k,r,sigma,t)
competitors_warrant_delta  =delta_cal(s,k_competitors,r,sigma,t)

num_of_competitors=my_warrant_delta/competitors_warrant_delta

#final
my_warrant_final          =bscall_cal(s*1.1,k,r,sigma,t)
competitors_warrant_final =bscall_cal(s*1.1,k_competitors,r,sigma,t)
#P/L
income_my =-(           my_warrant_final          -    my_warrant_initial)*1
income_competitors =(  competitors_warrant_final -    competitors_warrant_initial)*num_of_competitors
total_income=  income_my +  income_competitors


plot.figure() 
#c1=plot.plot(np.arange(0,0.1,0.001),[income_my]*len(k_competitors), label='P/L my warrant')
#c2=plot.plot(np.arange(0,0.1,0.001),income_competitors, label='P/L competitor''s warrant')
c3=plot.plot(np.arange(0,0.1,0.001)*100,total_income, label='total')
plot.ylabel('profit & loss',fontsize=16)
plot.xlabel('OTM%',fontsize=16)
plot.title('P/L of different OTM%',fontsize=30) 
plot.legend(loc='upper right')
#plot.subplot(121)

plot.figure() 
gamma_my=gamma_cal(s,k,r,sigma,t)
gamma_competitors=gamma_cal(s,k_competitors,r,sigma,t)

g1=plot.plot(np.arange(0,0.1,0.001)*100,[gamma_my]*len(gamma_competitors), label='my warrant')
g2=plot.plot(np.arange(0,0.1,0.001)*100,gamma_competitors, label='competitor warrant')
plot.ylabel('gamma',fontsize=16)
plot.xlabel('OTM%',fontsize=16)
plot.title('gamma different OTM%',fontsize=30)
plot.legend(loc='upper right')  
#plot.subplot(122)
        
        
        