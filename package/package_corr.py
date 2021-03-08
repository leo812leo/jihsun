
def run(path_file,day):
    from sys import path
    path.append(r'C:\Users\F128537908\Desktop\WCFADOX_Python')
    import xlwings as xw 
    import datetime 
    import pandas as pd
    from WCFAdox import PCAX
    from pandas.tseries.offsets import BDay 
    from sys import path
    path.append(r'C:\Users\F128537908\Desktop\WCFADOX_Python')
    import re
    from itertools import chain
    from xlwings.constants import DeleteShiftDirection
    # =============================================================================
    # #假日
    # =============================================================================
    df=pd.read_csv('https://www.twse.com.tw/holidaySchedule/holidaySchedule?response=csv&queryYear='+\
                   str(datetime.datetime.now().year-1911),header=1,encoding='big5')
    tem=[re.findall(r'(\d+月\d+日)',i.values[0])for _,i in df[['日期']].iterrows()]
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
    today=day.date()
    today_str=datetime.datetime.strftime(today,"%Y%m%d")
    pairs=[['0050','00632R','TX']]
    x=last_x_workday(today,130)
    last_x_workday_str=datetime.date.strftime(x,"%Y%m%d")
    
    data=PX.Pal_Data("日收盤表排行","D",last_x_workday_str,today_str,colist="日期,股票名稱,股票代號,收盤價,[漲幅(%)]",isst='Y',ps='<CM特殊,2011>')
    TX=PX.Pal_Data("期貨交易行情表","D",last_x_workday_str,today_str,colist="日期,名稱,代號,結算價");tx=TX[TX['代號']=='TX']
    tx["漲幅(%)"]=tx.sort_values("日期")["結算價"].rolling(2).apply(lambda x: (x[1]-x[0]) / x[0]*100,raw=True).round(2)
    tx.columns=data.columns
    etf=PX.Pal_Data("ETF淨值還原表","D",last_x_workday_str,today_str,colist="日期,股票名稱,股票代號,淨值,[報酬率(%)]")
    etf['股票代號']=etf['股票代號']+'_淨值'
    etf=etf.rename(columns={"淨值": "收盤價", "報酬率(%)": "漲幅(%)"})
    data_final=pd.concat([data,tx,etf],axis=0,ignore_index=True)
    final ={}
    
    for i,pair in enumerate(pairs):
            data_final_temp=data_final[data_final['股票代號'].isin(pair)]
            data_final_temp.loc[:,"漲幅(%)"]=data_final_temp.loc[:,"漲幅(%)"].astype(float).values
            temp=pd.pivot_table(data_final_temp,index="股票代號",columns="日期",values="漲幅(%)").T.dropna()
            temp=temp.iloc[-120:]
            temp=temp.rolling(60).corr().dropna()
            temp=[temp.iloc[i*3:i*3+3] for i in range(1,61)]
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
            
    app = xw.App(visible=True, add_book=False)    
    wb1 = app.books.open(path_file)  #資料來源
    ws=wb1.sheets['corr']
    ws.select()
    for i,_ in enumerate(pairs):
        if i==0:
            xw.Range((1,1)).value=final[0]
            row=(final[0]).shape[0]+2
        else:
            xw.Range((row,1)).value=final[1]
            xw.Range(str(row)+':'+str(row)).api.Delete(DeleteShiftDirection.xlShiftUp) 
            row=(final[i]).shape[0]+row
        ws.range('A1').expand().columns.autofit()
    wb1.save()
    wb1.close()
    app.quit()