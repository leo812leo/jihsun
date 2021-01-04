import matplotlib.pyplot as plot
from scipy.stats import norm
import math
import numpy as np



#greek數值解

def delta_cal(s,series):
    ds=s[1]-s[0]
    delta=-(series[0:len(series)-1]-series[1:len(series)]) / ds
    return delta

def gamma_cal(s,series):
    ds=s[1]-s[0]
    delta=delta_cal(s,series)
    gamma=-(delta[0:len(delta)-1]-delta[1:len(delta)]) / ds
    return gamma

def d2_cal(s,k,r,sigma,t):
    d1=(np.log(s/k)+(r+0.5*sigma**2)*t)/(sigma*np.sqrt(t))
    d2=d1-sigma*np.sqrt(t)
    return d2
def N(x):
    N=norm.cdf(x,scale=1)    
    return N


def bscall_cal(s,k,r,sigma,t):
    d1=(np.log(s/k)+(r+0.5*sigma**2)*t)/(sigma*np.sqrt(t))
    d2=d2_cal(s,k,r,sigma,t)
    call=s*N(d1)-k*np.exp(-r*t)*N(d2)
    return call
 
def bsput_cal(s,k,r,sigma,t):
    call=bscall_cal(s,k,r,sigma,t)
    put=k*np.exp(-r*t)+call-s
    return put




def mu_cal(r,sigma):
    mu=(r-0.5*sigma**2)/sigma**2
    return mu 
def lamb_cal(mu,r,sigma):
    lamb=math.sqrt(mu**2+(2*r)/sigma**2)
    return lamb

def x(S,K,sigma,T,mu):
    x=(math.log(S/K))/(sigma*math.sqrt(T))+(mu+1)*sigma*math.sqrt(T)
    return x

def y_1(H,S,K,sigma,T,mu):
    y_1=(math.log(H**2/(S*K)))/(sigma*math.sqrt(T))+(mu+1)*sigma*math.sqrt(T)
    return y_1

def y_2(H,S,sigma,T,mu):
    y_2=(math.log(H/S))/(sigma*math.sqrt(T))+(mu+1)*sigma*math.sqrt(T)
    return y_2
   
def Z_cal(H,S,sigma,T,lamb):
    Z=(math.log(H/S))/(sigma*math.sqrt(T))+lamb*sigma*math.sqrt(T)
    return Z

####
def A_cal(phi,S,T,K,r,sigma):    
    mu=mu_cal(r,sigma)
    x_1=x(S,K,sigma,T,mu)
    A=phi*S*norm.cdf(phi*x_1)-phi*K*math.exp(-r*T)*norm.cdf(phi*x_1-phi*sigma*math.sqrt(T))
    return A

def B_cal(phi,S,T,K,r,sigma,H):
    mu=mu_cal(r,sigma)
    x_2=x(S,H,sigma,T,mu)
    B=phi*S*norm.cdf(phi*x_2)-phi*K*math.exp(-r*T)*norm.cdf(phi*x_2-phi*sigma*math.sqrt(T))
    return B

def C_cal(phi,S,T,K,r,sigma,H,eta):
    mu=mu_cal(r,sigma)
    y1=y_1(H,S,K,sigma,T,mu)
    C=phi*S*(H/S)**(2*(mu+1))*norm.cdf(eta*y1)-phi*K*math.exp(-r*T)*(H/S)**(2*mu)*norm.cdf(eta*y1-eta*sigma*math.sqrt(T))
    return C

def D_cal(phi,S,T,K,r,sigma,H,eta):
    mu=mu_cal(r,sigma)   
    y2=y_2(H,S,sigma,T,mu)
    D=phi*S*(H/S)**(2*(mu+1))*norm.cdf(eta*y2)-phi*K*math.exp(-r*T)*(H/S)**(2*mu)*norm.cdf(eta*y2-eta*sigma*math.sqrt(T))
    return D





def IC(S,T,K,r,sigma,H,xx1,option_type):
    if (option_type=='UIC'):#UIC
        eta=-1;phi=1
        A=A_cal(phi,S,T,K,r,sigma)
        B=B_cal(phi,S,T,K,r,sigma,H)
        C=C_cal(phi,S,T,K,r,sigma,H,eta)  
        D=D_cal(phi,S,T,K,r,sigma,H,eta)
        if(S<H):    
            if (K>H):
                IC=A
            elif (K<H):
                IC=B-C+D
        else:
            IC=bscall_cal(S,K,r,sigma,T)         
    elif (option_type=='DIC'):#DIC
        eta=1;phi=1
        A=A_cal(phi,S,T,K,r,sigma)
        B=B_cal(phi,S,T,K,r,sigma,H)
        C=C_cal(phi,S,T,K,r,sigma,H,eta)  
        D=D_cal(phi,S,T,K,r,sigma,H,eta)
        if(S>H):
            if (K>H):
                IC=C
            elif (K<H):
                IC=A-B+D
        else:
            IC=bscall_cal(S,K,r,sigma,T)             
    return IC


"""up/down-in put"""
def IP(S,T,K,r,sigma,H,xx1,option_type):
    if (option_type=='UIP'):
        eta=-1;phi=-1
        A=A_cal(phi,S,T,K,r,sigma)
        B=B_cal(phi,S,T,K,r,sigma,H)
        C=C_cal(phi,S,T,K,r,sigma,H,eta)  
        D=D_cal(phi,S,T,K,r,sigma,H,eta)
        if(S<=H):#UIP
            if (K>H):
                IP=A-B+D
            elif (K<H):
                IP=C
        else:
            IP=bsput_cal(S,K,r,sigma,T)            
    elif (option_type=='DIP'): 
        eta=1;phi=-1
        A=A_cal(phi,S,T,K,r,sigma)
        B=B_cal(phi,S,T,K,r,sigma,H)
        C=C_cal(phi,S,T,K,r,sigma,H,eta)  
        D=D_cal(phi,S,T,K,r,sigma,H,eta)          
        if (S>H):#DIP
            if (K>H):
                IP=B-C+D
            elif (K<H):
                IP=A
        else:
            IP=bsput_cal(S,K,r,sigma,T) 
    return IP

"""up/down-out call"""
def OC(S,T,K,r,sigma,H,xx1,option_type):
    if (option_type=='UOC'):
        eta=-1;phi=1
        A=A_cal(phi,S,T,K,r,sigma)
        B=B_cal(phi,S,T,K,r,sigma,H)
        C=C_cal(phi,S,T,K,r,sigma,H,eta)  
        D=D_cal(phi,S,T,K,r,sigma,H,eta)
        if (S<=H):#UOC
            if (K>H):
                OC=0
            elif (K<H):
                OC=A-B+C-D
        else:
            OC=0
    elif (option_type=='DOC'):            
        eta=1;phi=1
        A=A_cal(phi,S,T,K,r,sigma)
        B=B_cal(phi,S,T,K,r,sigma,H)
        C=C_cal(phi,S,T,K,r,sigma,H,eta) 
        D=D_cal(phi,S,T,K,r,sigma,H,eta)
        if (S>=H):#DOC   
            if (K>H):
                OC=A-C
            elif (K<H):
                OC=B-D
        else:
            OC=0 
    return OC

"""up/down-out put"""
def OP(S,T,K,r,sigma,H,xx1,option_type):
    if (option_type=='UOP'):  
        eta=-1;phi=-1
        A=A_cal(phi,S,T,K,r,sigma)
        B=B_cal(phi,S,T,K,r,sigma,H)
        C=C_cal(phi,S,T,K,r,sigma,H,eta)  
        D=D_cal(phi,S,T,K,r,sigma,H,eta)
        if (S<=H):#UOP
            if (K>H):
                OP=B-D
            elif (K<H):
                OP=A-C
        else:
            OP=0
            
    elif (option_type=='DOP'):                        
        eta=1;phi=-1
        A=A_cal(phi,S,T,K,r,sigma)
        B=B_cal(phi,S,T,K,r,sigma,H)
        C=C_cal(phi,S,T,K,r,sigma,H,eta)  
        D=D_cal(phi,S,T,K,r,sigma,H,eta)
        if (S>H):#DOP
            if (K>H):
                OP=A-B+C-D
            elif(K<H):
                OP=0        
        else:
            OP=0                       
    return OP

Spot=np.arange(60,140,0.01)   #spot
K=100;r=0.1;sigma=0.25;T=1;xx1=0;
U=120;D=80
#-------------------------------------------------------------#
    
UIC=[IC(S,T,K,r,sigma,U,xx1,'UIC') for S in Spot] 
UOC=[OC(S,T,K,r,sigma,U,xx1,'UOC') for S in Spot] 
UIP=[IP(S,T,K,r,sigma,U,xx1,'UIP') for S in Spot] 
UOP=[OP(S,T,K,r,sigma,U,xx1,'UOP') for S in Spot]

DIC=[IC(S,T,K,r,sigma,D,xx1,'DIC') for S in Spot] 
DOC=[OC(S,T,K,r,sigma,D,xx1,'DOC') for S in Spot] 
DIP=[IP(S,T,K,r,sigma,D,xx1,'DIP') for S in Spot] 
DOP=[OP(S,T,K,r,sigma,D,xx1,'DOP') for S in Spot]

delta_total=[]
gamma_total=[]
for name in [UIC,UOC,UIP,UOP,DIC,DOC,DIP,DOP]:
    delta_total.append(delta_cal(Spot,np.array(name)))

for name in [UIC,UOC,UIP,UOP,DIC,DOC,DIP,DOP]:
    gamma_total.append(gamma_cal(Spot,np.array(name)) )
    
for n in range(4):
    gamma_total[n][5999]=0 
for n in range(4,8,1):
    gamma_total[n][1999]=0 
    
#-------------------------------------------------------------#
for name,str_name in zip([UIC,UOC,UIP,UOP,DIC,DOC,DIP,DOP],['UIC','UOC','UIP','UOP','DIC','DOC','DIP','DOP']):
    plot.figure() 
    plot.plot(Spot,name, label=str_name)
    plot.ylabel('value',fontsize=16)
    plot.xlabel('spot',fontsize=16)
    plot.title(str_name,fontsize=30) 
    
for num,str_name in zip(range(8),['UIC','UOC','UIP','UOP','DIC','DOC','DIP','DOP']):
    plot.figure() 
    plot.plot(Spot[1:],delta_total[num])
    plot.ylabel('value',fontsize=16)
    plot.xlabel('spot',fontsize=16)
    plot.title(str_name+'_delta',fontsize=30) 
    
for num,str_name in zip(range(8),['UIC','UOC','UIP','UOP','DIC','DOC','DIP','DOP']):
    plot.figure() 
    plot.plot(Spot[2:],gamma_total[num])
    plot.ylabel('value',fontsize=16)
    plot.xlabel('spot',fontsize=16)
    plot.title(str_name+'_gamma',fontsize=30) 








#GREEK_CALCULATE
delta_UIC=delta_cal(Spot,np.array(UIC));delta_DIC=delta_cal(Spot,np.array(DIC))
delta_UOC=delta_cal(Spot,np.array(UOC));delta_UIC=delta_cal(Spot,np.array(DOC))

gamma_UIP=gamma_cal(Spot,np.array(UIP));gamma_DIP=gamma_cal(Spot,np.array(DIP))
gamma_UOP=gamma_cal(Spot,np.array(UOP));gamma_DOP=gamma_cal(Spot,np.array(DOP))

def figure(Spot,option_type):
    plot.plot(Spot,UIC, label=str(option_type))
    plot.ylabel('value',fontsize=16)
    plot.xlabel('spot',fontsize=16)
    plot.title('up barrier call',fontsize=30) 
    plot.legend(loc='upper left')
    
#-------------------------------------------------------------#
plot.figure(1) 
#c1=plot.plot(s,bscall, label='call')
c2=plot.plot(Spot,UIC, label='up & in call')
c3=plot.plot(Spot,UOC, label='up & out call')
plot.ylabel('value',fontsize=16)
plot.xlabel('spot',fontsize=16)
plot.title('up barrier call',fontsize=30) 
plot.legend(loc='upper left')
