import tkinter as tk
from sys import path
path.append(r'T:\WCFADOX_Python')
import datetime 
import pandas as pd
from WCFAdox import PCAX
from pandas.tseries.offsets import BDay 
import re
from itertools import chain
import tkinter.messagebox as ms

def run():
    global pairs,e2
    string=e2.get()
    if string=='':
        days_before=0
    else:
        days_before=int(string)
    del string        
    ms.showinfo(title='開始計算', message='資料取至' + str(days_before) +'工作天前\n\n開始計算!!')
    # =============================================================================
    # #假日
    # =============================================================================
    df=pd.read_csv('https://www.twse.com.tw/holidaySchedule/holidaySchedule?response=csv&queryYear='+\
                    str(datetime.datetime.now().year-1911),header=1,encoding='big5')
    tem=[re.findall(r'(\d+月\d+日)',i.values[0])for _,i in df[[' 日期']].iterrows()]
    holiday=[(datetime.datetime.strptime(str(datetime.datetime.now().year)+i, "%Y%m月%d日")).date() for i in chain(*tem)]
    
    df=pd.read_csv('https://www.twse.com.tw/holidaySchedule/holidaySchedule?response=csv&queryYear='+\
                    str(datetime.datetime.now().year-1911-1),header=1,encoding='big5')
    tem2=[re.findall(r'(\d+月\d+日)',i.values[0])for _,i in df[[' 日期']].iterrows()]
    holiday2=[(datetime.datetime.strptime(str(datetime.datetime.now().year-1)+i, "%Y%m月%d日")).date() for i in chain(*tem2)]
    holiday.extend(holiday2)
    
    #公式
    def last_x_workday(date,x,holiday=holiday):                              #假日
        last_workday=(date- BDay(x)).date()                              #往前x天用工作天推(last_workday)
        num_holiday=sum([(i<date) & (i>=last_workday) for i in holiday]) #期間內假日總數
        if num_holiday!=0:
            return last_x_workday(last_workday,num_holiday)
        return last_workday
    
    PX=PCAX("192.168.65.11")
    today=(datetime.date.today()- BDay(days_before)).date()
    today_str=datetime.datetime.strftime(today,"%Y%m%d")
    
    
    x=last_x_workday(today,121)
    last_x_workday_str=datetime.date.strftime(x,"%Y%m%d")
    
    data=PX.Pal_Data("日收盤表排行","D",last_x_workday_str,today_str,colist="日期,股票名稱,股票代號,收盤價,[漲幅(%)]",isst='Y',ps='<CM特殊,2011>')
    TX=PX.Pal_Data("期貨交易行情表","D",last_x_workday_str,today_str,colist="日期,名稱,代號,結算價");tx=TX[TX['代號']=='TX']
    tx["漲幅(%)"]=tx.sort_values("日期")["結算價"].rolling(2).apply(lambda x: (x[1]-x[0]) / x[0]*100).round(2)
    tx.columns=data.columns
    etf=PX.Pal_Data("ETF淨值還原表","D",last_x_workday_str,today_str,colist="日期,股票名稱,股票代號,淨值,[報酬率(%)]")
    etf['股票代號']=etf['股票代號']+'_淨值'
    etf=etf.rename(columns={"淨值": "收盤價", "報酬率(%)": "漲幅(%)"})
    data_final=pd.concat([data,tx,etf],axis=0,ignore_index=True,sort=True)
    final ={}
    
    for i,pair in enumerate(pairs):
            length=len(pair)
            data_final_temp=data_final[data_final['股票代號'].isin(pair)]
            data_final_temp.loc[:,"漲幅(%)"]=data_final_temp.loc[:,"漲幅(%)"].astype(float).values
            temp=pd.pivot_table(data_final_temp,index="股票代號",columns="日期",values="漲幅(%)").T.dropna()
            temp=temp.iloc[-120:]
            temp=temp.rolling(60).corr().dropna()
            temp=[temp.iloc[i*length:i*length+length] for i in range(1,61)]
            mean=pd.concat(temp).groupby(level=-1).mean().round(3)
            std=pd.concat(temp).groupby(level=-1).std().round(6)
            D60=temp[-1].droplevel(0).round(3).stack()
            D60=D60[D60!=1]
            ff= (std/mean).stack()
            ff=ff[ff!=0]
            df=pd.concat([D60,ff],axis=1,join='inner')
            df.index.names=['申請標的', '相關有價證券']
            df.columns=['60日相關係數','穩定性分析：60日相關係數之變異係數']
            df['相關性是否符合規定']=df['60日相關係數'].apply(lambda x : '是' if abs(x)>=0.9 else '否')
            df['穩定性是否符合規定']=df['穩定性分析：60日相關係數之變異係數'].apply(lambda x : '是' if abs(x)<=0.05 else '否') 
            df=df[[df.columns[i] for i in [0,2,1,3]]]
            final[i]=df
    for i,_ in enumerate(pairs):
            if i==0:
                final[0].to_excel('corr.xlsx',sheet_name=str(i+1))
            else:
                with pd.ExcelWriter('corr.xlsx',mode='a', engine="openpyxl") as writer: 
                    final[i].to_excel(writer,sheet_name=str(i+1))
    ms.showinfo(title='完成', message='完成!!! 請打開excel檔')
# =============================================================================
# tkinter
# =============================================================================
pairs=[['0050','00632R','TX'],['00637L','00655L','00638R']]
window = tk.Tk()
window.title('correation cofficient')
window.geometry('500x300')
l = tk.Label(window, text='請輸入標的代碼 \n範例:2330,2454(逗號隔開)',font=('Arial', 12))
l.pack()
e = tk.Entry(window)
e.pack()
n=3
def insert_point():
    global n,pairs
    var = e.get()
    e.delete(0, 'end')
    pairs.append([str(stock) for stock in var.split(',')])
    t.insert('insert',str(n)+':' +var+'\n')
    n+=1
#Button1
b1 = tk.Button(window, text='輸入標的', width=15,height=2, command=insert_point)
b1.pack()

#label
l2 = tk.Label(window, text='工作天前',font=('Arial', 12))
l2.pack(side='right')


#TEXT1
t = tk.Text(window,  width=30,height=10)
t.insert("insert","1:0050,00632R,TX \n2:00637L,00655L,00638R\n")
t.pack(side='left')

e2 = tk.Entry(window)
e2.pack(side='right')

#Button2
b2 = tk.Button(window, text='開始計算', width=15,height=2, command=run)
b2.pack(side='bottom')
window.mainloop()