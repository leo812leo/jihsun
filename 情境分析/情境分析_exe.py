from sys import path
path.append(r'T:\WCFADOX_Python')
import datetime
import scipy.stats as st
import numpy as np
from WCFAdox import PCAX
import pandas as pd
import re
from itertools import chain
from pandas.tseries.offsets import BDay
import matplotlib.pyplot as plot
plot.rcParams['font.sans-serif'] = 'Arial Unicode MS'
#%%
import requests
from json import loads
from pandas import DataFrame
class PCAX():
    def __init__(self,sip):
        self._ip=sip

    def _check_perid(self,ptype,perid):
        #日期>D,年月>M,年季>Q,年度>Y
        chk_ptype=str.lower(ptype)
        if((chk_ptype!="d")&(chk_ptype!="m")&(chk_ptype!="q")&(chk_ptype!="y")):
            return "資料表頻率輸入錯誤:日與週資料請輸入D,月資料請輸入M,季資料請輸入Q,年資料請輸入Y,"
        else:
            if((chk_ptype=="y")&(len(perid)==4)):
                return "年度"
            elif((chk_ptype=="q")&(len(perid)==6)):
                return "年季"
            elif((chk_ptype=="m")&(len(perid)==6)):
                return "年月"
            elif((chk_ptype=="d")&(len(perid)==8)):
                return "日期"
            else:
                return "資料日期格式不正確:日與週資料請輸入8碼,月與季資料請輸入6碼,年資料請輸入4碼"
    
    def _check_isst(self,isst): #Y>股票,其他>非股票
        if(str.lower(isst)=="y"):
            return "股票代號"
        else:
            return "代號"

    def Mul_Data(self,tbname,ptype,date,colist = "",isst="Y",ps=""):
        chk_result=self._check_perid(ptype,date)    
        if(len(chk_result)!=2):
            print(chk_result)
        else:
            if(colist==""):
                col="*"
            else:
                col=colist
            sql="select "+col+" from "+tbname+" where "+chk_result+" = '"+date+"'"
            if(ps!=""):
                sql=sql+" and "+self._check_isst(isst)+" in "+ps
            return self._return_qry(sql,tbname)                        
    def _return_qry(self,cmds,tables):
        srv="http://"+self._ip+"/CMoneyAdox/AdoxcService.svc/QueryTableLight"
        params ="{\"FormatSQL\": \""+cmds+"\",\"TableNames\": [\""+tables+"\"]}"
        params=params.encode('utf-8')
        headers = {'Content-Type': 'application/json'}
        response = requests.request("POST", srv, data=params, headers=headers,timeout=600)
        qryresult=loads(response.text)
        errcode=qryresult.get("Error")        
        if (errcode==""):
            data=qryresult.get('ResultValue')
            data=data.split("\r\n")
            data=data[0:len(data)-1]
            for i in range(0,len(data)):
                data[i]=data[i].split("^")
            return DataFrame(data[1:],columns=data[0])
 
        else:
            print("查詢產生錯誤，錯誤訊息為:")
            print(errcode)
            return ""
    def Pal_Data(self,tbname,ptype,begin,end,colist = "",isst="Y",ps=""):
        chk_result1=self._check_perid(ptype,begin)    
        chk_result2=self._check_perid(ptype,end)    
        if(len(chk_result1)!=2):
            print(chk_result1)
        elif(len(chk_result2)!=2):
            print(chk_result2)
        else:
            if(colist==""):
                col="*"
            else:
                col=colist
            sql="select "+col+" from "+tbname+" where "+chk_result1+" between '"+begin+"' and '"+end+"'"
            if(ps!=""):
                sql=sql+" and "+self._check_isst(isst)+" in "+ps
            return self._return_qry(sql,tbname)                
#%%
def Premium_cal(flag,s,k,r,sigma,t,ratio)->float:
    d1=(np.log(s/k)+(r+0.5*sigma**2)*t)/(sigma*np.sqrt(t))
    d2=d1-sigma*np.sqrt(t)

    if flag=='c' or flag=='C':
        return (s*N(d1)-k*np.exp(-r*t)*N(d2))*ratio
    else:
        return (-s*N(-d1)+k*np.exp(-r*t)*N(-d2))*ratio
df=pd.read_csv('https://www.twse.com.tw/holidaySchedule/holidaySchedule?response=csv&queryYear='+\
               str(datetime.datetime.now().year-1911),header=1,encoding='big5')
tem=[re.findall(r'(\d+月\d+日)',i.values[0])for _,i in df[['日期']].iterrows()]
holiday=[(datetime.datetime.strptime(str(datetime.datetime.now().year)+i, "%Y%m月%d日")).date() for i in chain(*tem)]
def workingday(start,end,holiday=holiday)->int:
    work_day=np.busday_count( start, end )
    holiday_num=sum((np.array(holiday)>=start) & (np.array(holiday)<end))
    work_day_real=work_day-holiday_num
    return work_day_real
def gamma_cal(s,k,r,sigma,t):
    d1=(np.log(s/k)+(r+0.5*sigma**2)*t)/(sigma*np.sqrt(t))
    gamma=np.exp(-r*t)*dN(d1)/(s*sigma*np.sqrt(t)) 
    return gamma
def delta_cal(flag,s,k,r,sigma,t):
    d1=(np.log(s/k)+(r+0.5*sigma**2)*t)/(sigma*np.sqrt(t))
    if flag=='c' or flag=='C':
        return np.exp(-r*t)*N(d1)
    else:
        return np.exp(-r*t)*(N(d1)-1)
def N(x):
    N=st.norm.cdf(x,scale=1)    
    return N
def dN(X):
    dN=st.norm.pdf(X,scale = 1) 
    return dN
def pl(flag,s,k,r,sigma,t,ratio,oi,up_down,delta):
    option_pl=-(Premium_cal(flag,s*(1+up_down/100),k,r,sigma,(t-1/250),ratio)-Premium_cal(flag,s,k,r,sigma,t,ratio))*1000*oi
    stock_pl=delta*(up_down/100)           #避險部位為delta
    return option_pl+stock_pl
del df,tem


#%%
def run():
    # table1
    PX=PCAX("192.168.65.11")
    col="代號,名稱,權證收盤價,標的收盤價,最新履約價,最後委買價,[隱含波動率(委買價)(%)]"
    df1=PX.Mul_Data("權證評估表","D",date_str,colist=col)
    df1=df1.set_index("代號")
    col="代號,發行日期,到期日期,標的代號,最新執行比例,發行機構名稱"
    year=str(datetime.date.today().year)
    df2=PX.Mul_Data("權證基本資料表","Y",year,colist=col)
    df2=df2.set_index("代號")
    data1=pd.concat([df1,df2],axis=1,join='inner')
    data1=data1[~((data1.index.str[-1]=='B') | (data1.index.str[-1]=='X') | (data1.index.str[-1]=='C'))]
    
    to_num=data1.columns[1:6].to_list()+ [data1.columns[-2]]
    data1[to_num]=data1[to_num].apply(pd.to_numeric)
    data1[['發行日期','到期日期']]=data1[['發行日期','到期日期']].apply(pd.to_datetime)
    for string in ['發行日期','到期日期']:
        data1[string]=data1[string].dt.date
    # table2
    col="股票代號,自營商庫存,自營商買賣超,[自營商買賣超金額(千)]"
    df3=PX.Pal_Data("日自營商進出排行","D",date_str,date_str,colist=col)
    df3=df3.set_index("股票代號")
    df4=PX.Mul_Data("日收盤表排行","D",date_str,colist="股票代號,[成交金額(千)]",ps="<CM代號,權證>")
    df4=df4.set_index("股票代號")
    data2=pd.concat([df3,df4],axis=1,join='inner')
    data2=data2.apply(pd.to_numeric)
    table=pd.concat([data1,data2],axis=1,join='inner')
    table['flag']=table.index.str[-1].to_series().apply(lambda  s: s.lower() if s=='P' else 'c' ).to_list()
    table=table[table['標的代號']==stock]
    table=table[table['發行機構名稱']==broker]
    #%%
    for (i,(index,values)) in enumerate(table[['標的收盤價','最新履約價','隱含波動率(委買價)(%)','到期日期','flag','最新執行比例','自營商庫存']].iterrows()):
        s,k,sigma,enddate,flag,ratio,oi = values
        sigma=sigma/100;oi=-oi
        r=0
        t=workingday(datetime.datetime.strptime(date_str,"%Y%m%d").date(),enddate)/250
        delta=delta_cal(flag,s,k,r,sigma,t)*ratio*s*1000*oi    #正確delta(dollar)
        table.loc[index,'delta']=delta
        table.loc[index,'gamma']=-gamma_cal(s,k,r,sigma,t)*ratio*((s)**2)*1000*0.01*oi
        for i in range(-10,11):
            table.loc[index,str(i)]=pl(flag,s,k,r,sigma,t,ratio,oi,i,delta)
    del s,k,r,sigma,t,ratio,oi
    #%%
    savefile = asksaveasfilename(filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*") ))
    table.to_excel(savefile+ ".xlsx")
    ms.showinfo(title='完成', message='完成')
#%%

list_b=['中國信託綜合證', '亞東證', '元大證', '元富證', '兆豐證', '凱基證', '台新證', '國泰綜合證', '國票綜合證', 
        '富邦綜合證', '康和綜合證', '日盛證', '永豐金證', '第一金證', '統一綜合證', '群益證', '華南永昌綜合證', '香港商麥格理']
import tkinter as tk
from tkinter.filedialog import asksaveasfilename
import tkinter.messagebox as ms
from tkinter import ttk
date_str='';tock='';broker=''
def input_common():
    global date_str,open_delta,stock,broker
    date_str = e.get()
    stock = e3.get()
    broker=comboExample.get()
    ms.showinfo(title='資料輸入', message='輸入完成請按開始計算!!!')
window = tk.Tk()
window.title('情境分析')
window.geometry('300x350')
l = tk.Label(window, text='日期 範例:20080803',font=('Arial', 12))
l.pack()
e = tk.Entry(window)
e.pack()
l = tk.Label(window, text='股票代號',font=('Arial', 12))
l.pack()
e3 = tk.Entry(window)
e3.pack()

labelTop = tk.Label(window,
                    text = "Choose broker",font=('Arial', 12))
labelTop.pack()
comboExample = ttk.Combobox(window, 
                            values=list_b)
comboExample.pack()

#Button1
b1 = tk.Button(window, text='輸入完成', width=15,height=2, command=input_common)
b1.pack(side=tk.LEFT)
b2 = tk.Button(window, text='開始計算', width=15,height=2, command=run)
b2.pack(side=tk.RIGHT)
window.mainloop()




