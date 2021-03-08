import numpy as np
import scipy.stats as st

def N(x):
    N=st.norm.cdf(x,scale=1)    
    return N
def dN(X):
    dN=st.norm.pdf(X,scale = 1) 
    return dN
def Premium_cal(s,k,r,sigma,t):
    d1=(np.log(s/k)+(r+0.5*sigma**2)*t)/(sigma*np.sqrt(t))
    d2=d1-sigma*np.sqrt(t)
    c=s*N(d1)-k*np.exp(-r*t)*N(d2)
    return c 
def delta_cal(s,k,r,sigma,t):
    d1=(np.log(s/k)+(r+0.5*sigma**2)*t)/(sigma*np.sqrt(t))
    delta=np.exp(-r*t)*N(d1) 
    return delta
def gamma_cal(s,k,r,sigma,t):
    d1=(np.log(s/k)+(r+0.5*sigma**2)*t)/(sigma*np.sqrt(t))
    gamma=np.exp(-r*t)*dN(d1)/(s*sigma*np.sqrt(t)) 
    return gamma
def theta_cal(s,k,r,sigma,t):
    d1=(np.log(s/k)+(r+0.5*sigma**2)*t)/(sigma*np.sqrt(t))
    d2=d1-sigma*np.sqrt(t)
    tem=-s*dN(d1)*sigma*np.exp(-r*t)/(2*np.sqrt(t))+r*s*N(d1)*np.exp(-r*t)-r*k*np.exp(-r*t)*N(d2)
    theta=tem/250
    return theta
def vega_cal(s,k,r,sigma,t):
    d1=(np.log(s/k)+(r+0.5*sigma**2)*t)/(sigma*np.sqrt(t))
    vega=s*np.sqrt(t)*dN(d1)*np.exp(-r*t)/100
    return vega
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