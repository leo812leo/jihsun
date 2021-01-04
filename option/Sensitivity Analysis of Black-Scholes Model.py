import numpy as np
import scipy.stats as st
import matplotlib.pyplot as plot

def d1_cal(s,k,r,sigma,t):
    d1=(np.log(s/k)+(r+0.5*sigma**2)*t)/(sigma*np.sqrt(t))
    return d1
def d2_cal(d1,sigma,t):
    d2=d1-sigma*np.sqrt(t)
    return d2
def N(x):
    N=st.norm.cdf(x,scale=1)    
    return N
def dN(X):
    dN=st.norm.pdf(X,scale = 1) 
    return dN
def c_cal(s,d1,d2,k,r,t):
    c=s*N(d1)-k*np.exp(-r*t)*N(d2)
    return c 
def delta_cal(d1,r,t):
    delta=np.exp(-r*t)*N(d1) 
    return delta
def gamma_cal(d1,r,t,s,sigma):
    gamma=np.exp(-r*t)*dN(d1)/(s*sigma*np.sqrt(t)) 
    return gamma
def theta_cal(s,d1,d2,sigma,t,k,r):
    tem=-s*dN(d1)*sigma*np.exp(-r*t)/(2*np.sqrt(t))+r*s*N(d1)*np.exp(-r*t)-r*k*np.exp(-r*t)*N(d2)
    theta=tem/250
    return theta
def vega_cal(s,t,d1,r):
    vega=s*np.sqrt(t)*dN(d1)*np.exp(-r*t)/100
    return vega
def rho_cal(k,t,r,d2):
    rho=k*t*np.exp(-r*t)*N(d2)/100
    return rho
def profit(s,k,pc):
    if pc=='c':
        profit=np.zeros(len(s))        
        for n in range(0,len(s)):
            if s[n]>k:
                profit[n]=(s[n]-k)
            else:
                profit[n]=0
            
    if pc=='p':
            if s[n]<k:
                profit[n]=(k-s[n])
            else:
                profit[n]=0
    return profit
##c=  #call price
d1=d1_cal(s,k,r,sigma,t)
d2=d2_cal(d1,sigma,t)
Nd1=N(d1)
Nd2=N(d2)
c=c_cal(s,d1,d2,k,r,t)

#greek
delta=delta_cal(d1,r,t)
gamma=gamma_cal(d1,r,t,s,sigma)
theta=theta_cal(s,d1,d2,sigma,t,k,r)
vega=vega_cal(s,t,d1,r)
rho=rho_cal(k,t,r,d2)
profit=profit(s,k,pc)

plot.figure(1)  
#first_day
s=50;k=50;t=180/252;sigma=0.8;r=0.009;pc='c'
d1=d1_cal(s,k,r,sigma,t)
d2=d2_cal(d1,sigma,t)
Nd1=N(d1)
Nd2=N(d2)
c=c_cal(s,d1,d2,k,r,t)
delta=delta_cal(d1,r,t)
#second_day
s_2=np.arange(10,101,1);k=50;t=180/252;sigma=0.8;r=0.009;pc='c'
d1=d1_cal(s_2,k,r,sigma,t)
d2=d2_cal(d1,sigma,t)
Nd1=N(d1)
Nd2=N(d2)
c_2=c_cal(s_2,d1,d2,k,r,t)
delta_2=delta_cal(d1,r,t)

profit=-c+c_2 - (delta*(-s+s_2))
plot.plot(s_2,profit,color='r',label='hedge')








#
#
#plot.figure(2)  
#plot.plot(s,delta)
#plot.title('delta vs spot',fontsize=30) 
#plot.ylabel('delta',fontsize=16)
#plot.xlabel('spot',fontsize=16)
#
#
plot.figure(5)  
plot.plot(s_2,gamma)
plot.title('gamma vs spot',fontsize=30) 
plot.ylabel('gamma',fontsize=16)
plot.xlabel('spot',fontsize=16)
#plot.figure(6)  
#plot.plot(s,theta)
#plot.title('theta vs spot',fontsize=30) 
#plot.ylabel('theta',fontsize=16)
#plot.xlabel('spot',fontsize=16)

