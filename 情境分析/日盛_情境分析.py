import datetime
import scipy.stats as st
import numpy as np
import pandas as pd
import re
from itertools import chain
import matplotlib.pyplot as plot
from pandas.tseries.offsets import CustomBusinessDay
import requests
from io import StringIO


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
cb=CustomBusinessDay(1,holidays=holiday)
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
def pl(flag,s,k,r,sigma,t,ratio,oi,up_down):
    option_pl=-(Premium_cal(flag,s*(1+up_down/100),k,r,sigma,(t-1/252),ratio)-Premium_cal(flag,s,k,r,sigma,t,ratio))*1000*oi
    return option_pl
del df,tem
#%%
def run():
    ### data1_權證部位
    PX=PCAX("192.168.65.11") 
    if after_hours==1:
        date_str=datetime.datetime.strftime(datetime.date.today(),"%Y%m%d") #今天收盤 
        date_str2=datetime.datetime.strftime(datetime.date.today(),"%Y/%m/%d") 
    else:
        date_str=datetime.datetime.strftime(datetime.date.today()-cb,"%Y%m%d")  #昨天收盤     
        date_str2=datetime.datetime.strftime(datetime.date.today()-cb,"%Y/%m/%d") 
        
    #網頁資料
    df1=pd.read_html('http://172.16.10.13/hedge/wmaker/warmktmakercfg_query.php',header=0)[0]
    df1=df1[['權證代號','權證名稱','標的代號', '標的名稱','權證市價','在外流通','BVol%','權證理論價', '價內外程度', '距到期日','調整Delta','行使比例']]
    df1=df1.set_index("權證代號") 
    
    #cmoney資料    標的收盤價+含波動率(委買價)+最新履約價
    year=str(datetime.date.today().year)
    col="代號,最新履約價"
    df2=PX.Mul_Data("權證基本資料表","Y",year,colist=col)
    df2=df2.set_index("代號")
    col="代號,標的收盤價,[隱含波動率(委買價)(%)],到期日期"
    df3=PX.Mul_Data("權證評估表","D",date_str,colist=col)
    df3=df3.set_index("代號")
    df3['到期日期']=df3['到期日期'].apply(pd.to_datetime)    
    #合併
    data1=pd.concat([df1,df2,df3],axis=1,join='inner')
    to_num=data1.columns[-4:-1]
    data1[to_num]=data1[to_num].apply(pd.to_numeric)

    data1['flag']=data1.index.str[-1].to_series().apply(lambda  s: s.lower() if s=='P' else 'c' ).to_list()
    
    #filter
    data1=data1[data1['標的代號']==stock]
    
    del df1,df2,to_num,col,year
    #%%
    #權證整理
    權證整理={}
    for index,values in data1[['標的收盤價','最新履約價','BVol%','距到期日','flag','行使比例','在外流通']].iterrows():
        s,k,sigma,enddate,flag,ratio,oi = values
        sigma=sigma/100
        r=0.008;t=enddate/252
        data1.loc[index,'delta']=delta_cal(flag,s,k,r,sigma,t)*ratio*(s)*1000*oi
        data1.loc[index,'gamma']=-gamma_cal(s,k,r,sigma,t)*ratio*((s)**2)*1000*0.01*oi
        for i in range(-10,11,2):
            權證整理.setdefault('損益', pd.DataFrame()).loc[index,str(i)+'%']=pl(flag,s,k,r,sigma,t,ratio,oi,i)
            s_temp=s*(1+i/100)
            權證整理.setdefault('gamma', pd.DataFrame()).loc[index,str(i)+'%']=-gamma_cal(s_temp,k,r,sigma,t)*ratio*((s_temp)**2)*1000*0.01*oi
            權證整理.setdefault('delta', pd.DataFrame()).loc[index,str(i)+'%']=-delta_cal(flag,s_temp,k,r,sigma,t)*ratio*(s_temp)*1000*oi   
    del s,k,r,sigma,t,ratio,oi,flag,enddate,i,index,s_temp,values
    #%%
    ### data2_同業權證部位+期貨
    df1=pd.read_html('http://172.16.10.13/hedge/warrealmkt2.php',header=0)[0]  
    df1.columns=df1.iloc[-1,:]
    df1=df1.drop(df1.shape[0]-1)
    #filter
    df1=df1[df1['標的代碼']==stock]
    
    同業避險=df1[df1['商品名稱'].str.contains('[購售]', regex=True)]
    同業避險['flag']=同業避險['商品代號'].str[-1].apply(lambda  s: s.lower() if s=='P' else 'c' )
    同業避險=同業避險.set_index('商品代號')
    同業避險=pd.concat([同業避險,df3],axis=1,join='inner')
    to_num=同業避險.columns[-3:-1]
    同業避險[to_num]=同業避險[to_num].apply(pd.to_numeric)
    同業避險['隱含波動率(委買價)(%)']=同業避險['隱含波動率(委買價)(%)'].fillna(同業避險['隱含波動率(委買價)(%)'].mean())
    del to_num,df3
    同業整理={}
    if 同業避險.empty:
        同業整理.setdefault('損益', pd.Series([0]))
        同業整理.setdefault('delta', pd.Series([0]))
        同業整理.setdefault('gamma', pd.Series([0]))
    else:
        同業避險[['標的收盤價','履約價','bidvol','行使比率','即時口(張)數']] = 同業避險[['標的收盤價','履約價','bidvol','行使比率','即時口(張)數']].astype(float)
        for index,values in 同業避險[['標的收盤價','履約價','bidvol','到期日期','flag','行使比率','即時口(張)數']].iterrows():
            s,k,sigma,enddate,flag,ratio,oi = values
            r=0;t= workingday(datetime.date.today(),enddate.date())/252
            #sigma=sigma/100
            同業避險.loc[index,'delta']=delta_cal(flag,s,k,r,sigma,t)*ratio*(s)*1000*oi
            同業避險.loc[index,'gamma']=+gamma_cal(s,k,r,sigma,t)*ratio*((s)**2)*1000*0.01*oi
            for i in range(-10,11,2):
                同業整理.setdefault('損益', pd.DataFrame()).loc[index,str(i)+'%']=-pl(flag,s,k,r,sigma,t,ratio,oi,i)
                s_temp=s*(1+i/100)
                同業整理.setdefault('gamma', pd.DataFrame()).loc[index,str(i)+'%']=+gamma_cal(s_temp,k,r,sigma,t)*ratio*((s_temp)**2)*1000*0.01*oi  
                同業整理.setdefault('delta', pd.DataFrame()).loc[index,str(i)+'%']=+delta_cal(flag,s_temp,k,r,sigma,t)*ratio*(s_temp)*1000*oi  
        del s,k,r,sigma,t,ratio,oi,flag,enddate,i,index,s_temp,values     
    ### 現貨+期貨
    期貨避險=df1[~df1['商品名稱'].str.contains('[購售]', regex=True)]    
    期貨避險=期貨避險[~期貨避險['商品名稱'].str.contains('TXO', regex=True)]
    期貨避險=期貨避險.set_index('商品名稱')
    期貨避險.index = 期貨避險.index + ('_'+期貨避險.groupby(level=0).cumcount().astype(str)).replace('_0','')
    期貨整理={}
    if 期貨避險.empty:
        期貨整理.setdefault('損益', pd.Series([0]))
        期貨整理.setdefault('delta', pd.Series([0]))
        期貨整理.setdefault('gamma', pd.Series([0]))  
    else:
        期貨避險[['標的價格','即時口(張)數']]=期貨避險[['標的價格','即時口(張)數']].astype(float)
        for index,values in 期貨避險[['標的價格','即時口(張)數','標的代碼']].iterrows():
            s,oi,code = values
            if code[:2]=='00':
                xx=10
            elif index[0]=='M':
                xx=0.1
            else:
                xx=2
            delta=s*1000*oi*xx    #正確delta(dollar)
            期貨避險.loc[index,'delta']=delta
            for i in range(-10,11,2):
                期貨整理.setdefault('損益', pd.DataFrame()).loc[index,str(i)+'%']=delta*(i/100)
                s_temp=s*(1+i/100)
                期貨整理.setdefault('gamma', pd.DataFrame()).loc[index,str(i)+'%']=0
                期貨整理.setdefault('delta', pd.DataFrame()).loc[index,str(i)+'%']=s_temp*1000*oi*xx 
        del i,delta,code,s,oi,index,values,s_temp,xx,df1
    #%%
    現貨整理={}

    url='http://172.16.10.13/hedge/Stockposdy_war.php'
    form_data = {
        'qrytrdday':date_str2}
    r = requests.post(url, data=form_data)
    r.encoding='ut-f8'
    df4=pd.read_html(StringIO(r.text),index_col=0)[0]
    col=['股票名稱','庫存股數','收盤價','市值']
    df4.columns=col
    df4.index.name='股票代號'
    df4.index=df4.index.astype(str)
    df4=df4.rename(index={'50':'0050'})
    if stock in df4.index:
        現貨市值=float(df4.loc[stock,'市值'])
        for i in range(-10,11,2):
            現貨整理.setdefault('損益', pd.DataFrame()).loc[stock,str(i)+'%']=現貨市值*(i/100)
            s_temp=現貨市值*(1+i/100)
            現貨整理.setdefault('gamma', pd.DataFrame()).loc[stock,str(i)+'%']=0
            現貨整理.setdefault('delta', pd.DataFrame()).loc[stock,str(i)+'%']=s_temp
    else:
        現貨整理['損益']=pd.Series([0]);現貨整理['gamma']=pd.Series([0]);現貨整理['delta']=pd.Series([0])
    #%% 總結
    name=['權證整理','同業整理','期貨整理','現貨整理']
    total={}    
    for state in ['損益','gamma','delta']:
        exec('total["'+state+'"] = '+
             name[0]+'["'+state+'"].sum() +' +  
             name[1]+'["'+state+'"].sum() +' +
             name[2]+'["'+state+'"].sum() +' +
             name[3]+'["'+state+'"].sum() ')
    table=pd.DataFrame(total)

    #%%
    savefile = asksaveasfilename(filetypes=(("Excel files", "*.xlsx"), ("All files", "*.*") ))
    table.to_excel(savefile+ ".xlsx")
    def click():
        wb2.macro('num')()
        wb2.macro('format')()
    
    import xlwings as xw 
    app = xw.App(visible=False, add_book=False)
    wb2 = app.books.open(r'D:\code\output專用\code_for_format.xlsm')
    wb1 = app.books.open(savefile+ ".xlsx")
    for name in ['權證整理','同業整理','期貨整理','現貨整理']:
        wb1.sheets.add(name)
        xw.sheets.active
        row=eval(name+"['損益'].shape[0]")
        i1=(1,1)
        exec("xw.Range("+str(i1)+").value="+name+"['損益']")
        xw.Range(i1).value='p/l'
        xw.Range(tuple(np.array(i1)+1),tuple([row+1,12])).select()
        click()
        
        x=1+row+2
        i2=(x,1)
        exec("xw.Range("+str(i2)+").value="+name+"['delta']")
        xw.Range(i2).value='delta'
        xw.Range(tuple(np.array(i2)+1),tuple([x+row,12])).select()
        click()    
        
        x=1+2*(row+2)
        i3=(x,1)
        exec("xw.Range("+str(i3)+").value="+name+"['gamma']")
        xw.Range(i3).value='gamma'
        xw.Range(tuple(np.array(i3)+1),tuple([x+row,12])).select()
        click()
        
        #autofit
        rng=xw.Range('A1:L50')
        rng.columns.autofit()
        
    wb1.sheets[-1].name='總整理'
    wb1.sheets[-1].select()
    xw.Range('A1:D12').select()
    click()
    
    wb1.save()
    wb1.close()
    app.quit()
    ms.showinfo(title='完成', message='完成')    
    
    
#%%
import tkinter as tk
from tkinter.filedialog import asksaveasfilename
import tkinter.messagebox as ms
from tkinter import IntVar,Checkbutton
date_str='';tock='';broker=''
def input_common():
    global date_str,open_delta,stock,broker,after_hours
    stock = e3.get()
    after_hours=Var1.get()
    ms.showinfo(title='資料輸入', message='輸入完成請按開始計算!!!')
window = tk.Tk()
window.title('情境分析[盤後使用]')
window.geometry('500x300')
l = tk.Label(window, text='股票代號',font=('Arial', 13))
l.pack()
e3 = tk.Entry(window)
e3.pack()
Var1 = IntVar()
C1 = Checkbutton(window, text = "盤後1500~2100(無隔夜)", variable = Var1,font=('Arial', 12))
C1.pack()
l = tk.Label(window, text='收盤當天請打勾\n不得有個股選擇權,融券及TXO ',font=('Arial', 12))
l.pack()

#Button1
b1 = tk.Button(window, text='輸入完成', width=15,height=2, command=input_common)
b1.pack(side=tk.LEFT)
b2 = tk.Button(window, text='開始計算', width=15,height=2, command=run)
b2.pack(side=tk.RIGHT)
window.mainloop()




