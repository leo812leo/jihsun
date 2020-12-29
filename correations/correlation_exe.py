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

from datetime import datetime
import tkinter as tk
from pandas import concat, pivot_table, ExcelWriter
from pandas.tseries.offsets import BDay 
from tkinter.filedialog import askdirectory
import tkinter.messagebox as ms
def run():
    PX=PCAX("192.168.65.11")
    today=datetime.now()
    today_str=datetime.strftime(today,"%Y%m%d")
    str_list=[]
    for day in [60]:
        str_list.append(datetime.strftime(datetime.now()- BDay(day) ,"%Y%m%d"))    
    dataname = {}
    data_dict={}
    for i,days in enumerate([60]):
      data=PX.Pal_Data("日收盤表排行","D",str_list[i],today_str,colist="日期,股票名稱,股票代號,收盤價,[漲幅(%)]",isst='Y',ps='<CM特殊,2011>')
      TX=PX.Pal_Data("期貨交易行情表","D",str_list[i],today_str,colist="日期,名稱,代號,收盤價,[漲幅(%)]");tx=TX[TX['代號']=='TX'];tx.columns=data.columns
      data_final=concat([data,tx],axis=0,ignore_index=True)
      dataname["D{}".format(days)]=data_final
    
    for i,pair in enumerate(pairs):
        for key, value in dataname.items():
            value=value[value['股票代號'].isin(pair)]
            value.loc[:,"漲幅(%)"]=value.loc[:,"漲幅(%)"].astype(float).values
            a=pivot_table(value,index="股票代號",columns="日期",values="漲幅(%)").T
            data_dict.setdefault(str(i),{}).update({key: a.corr()})
    file_path = askdirectory()    
    writer = ExcelWriter(file_path+'//output.xlsx', engine = 'openpyxl',mode='w')
    for i in range(len(pairs)):
        data_dict[str(i)]['D60'].to_excel(writer, sheet_name=str(i)+'-60D')
    writer.save()
    ms.showinfo(title='完成', message='完成!!! 請打開excel檔')
    
# =============================================================================
# tkinter
# =============================================================================
pairs=[]
window = tk.Tk()
window.title('correation cofficient')
window.geometry('400x200')
l = tk.Label(window, text='請輸入標的代碼 \n範例:2330,2317,2454(逗號隔開)',font=('Arial', 12))
l.pack()
e = tk.Entry(window)
e.pack()
n=0
def insert_point():
    global n
    var = e.get()
    e.delete(0, 'end')
    pairs.append([str(stock) for stock in var.split(',')])
    t.insert('insert',str(n)+':' +var+'\n')
    n+=1
#Button1
b1 = tk.Button(window, text='輸入標的', width=15,height=2, command=insert_point)
b1.pack()
#TEXT1
t = tk.Text(window, height=5)
t.pack()
#Button2
b2 = tk.Button(window, text='開始計算', width=15,height=2, command=run)
b2.pack()
window.mainloop()



