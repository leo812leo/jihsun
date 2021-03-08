import datetime
from sys import path
path.append(r'T:\WCFADOX_Python')
from WCFAdox import PCAX
import pandas as pd
import re
from itertools import chain
from pandas.tseries.offsets import BDay
import numpy as np
PX=PCAX("192.168.65.11")
#%%日期
# =============================================================================
# 假日
# =============================================================================
df=pd.read_csv('https://www.twse.com.tw/holidaySchedule/holidaySchedule?response=csv&queryYear='+\
               str(datetime.datetime.now().year-1911),header=1,encoding='big5')
tem=[re.findall(r'(\d+月\d+日)',i.values[0])for _,i in df[['日期']].iterrows()]
holiday=[(datetime.datetime.strptime(str(datetime.datetime.now().year)+i, "%Y%m月%d日")).date() for i in chain(*tem)]
#公式
def last_x_workday(date,x,holiday=holiday):                              #假日
    last_workday=(date- BDay(x)).date()                              #往前x天用工作天推(last_workday)
    num_holiday=sum([(i<date) & (i>=last_workday) for i in holiday]) #期間內假日總數
    if num_holiday!=0:
        return last_x_workday(last_workday,num_holiday)
    return datetime.datetime.strftime(last_workday,"%Y%m%d"),datetime.datetime.strftime(last_workday,"%Y-%m-%d")
del tem,df
# =============================================================================
# 日期
# =============================================================================
today=datetime.date.today()
date_str=datetime.datetime.strftime(today,"%Y%m%d")
start_date = last_x_workday(today,5)
end_date   = last_x_workday(today,1)
#%%

col='日期,代號,標的代號,標的名稱,發行日期,到期日期,最新執行比例,[發行波動率(%)],[存續期間(月)],發行價格,標的收盤價,最新履約價'
df1=PX.Mul_Data("權證評估表","D",end_date[0],colist=col)
df1=df1.set_index('代號')
col="代號,發行機構名稱,[認購／認售]"
df2=PX.Mul_Data("權證基本資料表","Y",str(today.year),colist=col)
df2=df2.set_index('代號')
data1=pd.concat([df1,df2],axis=1,join='inner')
data1=data1[~((data1.index.str[-1]=='B') | (data1.index.str[-1]=='X') | (data1.index.str[-1]=='C'))]
to_num=data1.columns[5:11].to_list()
data1[to_num]=data1[to_num].apply(pd.to_numeric)
to_date=data1.columns[3:5].to_list()
data1[to_date]=data1[to_date].applymap(pd.to_datetime)
def calculate(df):
    if df['認購／認售']=='認購':
        return round(( (df['標的收盤價']-df['最新履約價']) /df['最新履約價'] )*100,2)
    elif df['認購／認售']=='認售':
        return round(( -(df['標的收盤價']-df['最新履約價']) /df['最新履約價']  )*100,2)
    else:
        raise ValueError("數值錯誤")
        
data1['價內外程度(%)']=data1[['標的收盤價','最新履約價','認購／認售']].apply(calculate ,axis=1)

del df1,df2,to_date,to_num
#%%
# 類別資料
台灣50=PX.Mul_Data("日收盤表排行","D",date_str,colist="股票代號",ps="<CM特殊,1>").iloc[:,0].values
台灣100=PX.Mul_Data("日收盤表排行","D",date_str,colist="股票代號",ps="<CM特殊,3>").iloc[:,0].values
上櫃=PX.Mul_Data("日收盤表排行","D",date_str,colist="股票代號",ps="<CM代號,2>").iloc[:,0].values
上市=PX.Mul_Data("日收盤表排行","D",date_str,colist="股票代號",ps="<CM代號,1>").iloc[:,0].values
#分類
上市台灣50=list(set(台灣50) & set(上市))
上市中型100=list(set(上市) & set(台灣100))
上市 =list(set(上市) - set(上市) - set(台灣100))
上櫃=list(上櫃)
dic={}
for stock in data1['標的代號'].drop_duplicates().to_list():
    if stock in 上市台灣50:
        dic[stock]='上市台灣50'
    elif stock in 上市中型100:
        dic[stock]='上市中型100'
    elif stock in 上市:
        dic[stock]='其他上市'
    elif stock in 上櫃:
        dic[stock]='上櫃'
    elif stock[0:2]=='00':
        dic[stock]='ETF'
    else:
        dic[stock]='指數'    

data1['類別']=(data1['標的代號'].map(dic)).copy()
del 上市台灣50,上市中型100,上市,上櫃,stock,台灣50,台灣100
#%% 產業名稱
#col='股票代號,產業代號,產業名稱,產業指數代號,產業指數名稱,指數彙編分類代號,指數彙編分類'
col='股票代號,產業名稱'
stock_data=PX.Mul_Data("上市櫃公司基本資料","Y",str(today.year),colist=col)
stock_data=stock_data.set_index('股票代號')
dic=stock_data.to_dict()['產業名稱']

temp=data1.query('類別 != "ETF" and 類別 !="指數" ' )
data1['產業名稱']=temp['標的代號'].map(dic)
del col,stock_data,temp,dic
#%%發行分析
start_date = last_x_workday(today,5)[1];end_date   = last_x_workday(today,1)[1]
data_new=data1.query('發行日期>= @start_date and 發行日期 <= @end_date')


#新發行資料整理(增量)
#  存續期間 + 發行價格
table1=data_new.pivot_table(index=['發行機構名稱'],values=['存續期間(月)','發行價格'],
                     aggfunc={'存續期間(月)':[np.mean,'count'],
                              '發行價格':[np.mean]})
#  類別
table2 = data_new.pivot_table(index=['發行機構名稱'],columns=['類別'],values=['發行價格'],fill_value=0,
                     aggfunc={'發行價格':'count'})
table2.columns=table2.columns.droplevel(0)
#  產業名稱
table3 = data_new.pivot_table(index=['發行機構名稱'],columns=['產業名稱'],values=['發行價格'],fill_value=0,
                     aggfunc={'發行價格':'count'})
table3.columns=table3.columns.droplevel(0)
#純量
table0 = data1.pivot_table(index=['發行機構名稱'],values=['發行價格'],fill_value=0,
                     aggfunc={'發行價格':'count'})
#標的統計


table4 = data_new.pivot_table(index=['發行機構名稱'],columns=['標的名稱'],values=['發行價格'],fill_value=0,
                     aggfunc={'發行價格':'count'})
table4.loc['總和',:]=table4.sum()
table4=table4.droplevel(0,axis=1)
largest=table4.loc['總和'].nlargest(20).index
table4=table4.loc[:,largest]

del start_date,end_date,data1,today

#%%
#券商發行資料
table1=table1.droplevel(-1,axis=1)
table1.columns=['新發行檔數','存續期間(月)','平均價格']
table0.columns=['市場總檔數']
final_1=pd.concat([table0,table1],axis=1,sort=True).fillna(0)
final_1['新增比例(%)']=final_1.新發行檔數/(final_1.市場總檔數-final_1.新發行檔數)*100
final_1[['市場總檔數','新發行檔數']]=final_1[[ '市場總檔數','新發行檔數']].astype(int)
final_1[['存續期間(月)','平均價格','新增比例(%)']]=final_1[[ '存續期間(月)','平均價格','新增比例(%)']].round(2)
final_1=final_1.sort_values('市場總檔數',ascending=0)
sort_sample=final_1.index

del table1,table0
#%%
#券商發行 類別
final_2=table2.reindex(sort_sample).dropna().astype(int)
del table2
#券商發行 產業別
final_3=table3.reindex(sort_sample).dropna().astype(int)
final_3.columns=final_3.columns.str.replace('電子–','')
final_3.loc['sum']=final_3.sum()
final_3=final_3.sort_values(by='sum',axis=1,ascending=False)
final_3=final_3.drop('sum')
del table3
final_4=table4.reindex(sort_sample).dropna().astype(int)
del table4
detial=data_new.query("標的名稱 in @largest")
detial['標的名稱']=pd.Categorical(detial['標的名稱'],largest.to_list())
detial=detial.sort_values(["標的名稱","發行機構名稱"])
#%%
#新增 價內外程度
def cutter(series):
    return pd.cut(series,[-np.inf,-15,-10,-5,0,5,10,15,np.inf]).value_counts()
index_range=['~-15','-15 ~ -10','-10 ~ -5','-5 ~ 0','0 ~ 5','5 ~ 10','10 ~ 15','15 ~']
final_5=data_new.groupby(["發行機構名稱"])['價內外程度(%)'].apply(cutter)
final_5=final_5.unstack(level=-1)
final_5.columns=index_range
final_5=final_5.reindex(sort_sample).dropna().astype(int)
del index_range
#%%
# with pd.ExcelWriter('output.xlsx', engine="openpyxl",mode='w') as writer:  
#     for i,name in zip(range(1,6),['券商發行','發行類別','發行產業','發行標的','發行價內外程度']):
#         temp=eval('final_'+str(i))
#         temp.to_excel(writer, sheet_name=name)
# writer.save()
transforms={'中國信託綜合證':'中信','香港商麥格理':'麥格理','第一金證':'第一金'}
for i in range(1,6):
    exec('final_'+str(i)+'=final_'+str(i)+'.rename(index=transforms)')
    exec('final_'+str(i)+'.index='+'final_'+str(i) +'.index.str[:2]')

    
# =============================================================================
#  excel   
# =============================================================================
shape_dict={}
for i in range(1,6):
    temp=eval('final_'+str(i))
    shape_dict[str(i)]=temp.shape
#%%
import xlwings as xw 
app = xw.App(visible=True, add_book=False)
wb1 = app.books.open(r'.\格式用\code_for_format.xlsm')
wb2 = app.books.open(r'.\格式用\新發行權證分析.xlsx')
sht=wb2.sheets[0]
sht.select()
sht.clear()
def t2t(t0,t):
    row,col=tuple(np.array(t)+np.array(t0))
    return [(row,col),(row,t0[1]),(t0[0],col)]
def color(t0,t):
     row,col=tuple(np.array(t)+np.array(t0))
     xw.Range(tuple(np.array(t0)+1),(row,col)).select()
     wb1.macro('color')()
#table1
intial=(1,1)
xw.Range(intial).value=final_1
for i in range(0,3):
    xw.Range(intial,t2t(intial,shape_dict['1'])[i]).select()
    wb1.macro('format')()
 

#table2
intial=(1,shape_dict['1'][1]+3)
xw.Range(intial).value=final_2

for i in range(0,3):
    xw.Range(intial,t2t(intial,shape_dict['2'])[i]).select()
    wb1.macro('format')()
color(intial,shape_dict['2']) 
    
#table5
intial=(max(shape_dict['2'][0],shape_dict['1'][0])+3,1)
xw.Range(intial).value=final_5
for i in range(0,3):
    xw.Range(intial,t2t(intial,shape_dict['5'])[i]).select()
    wb1.macro('format')()
color(intial,shape_dict['5']) 
wb1.macro('total_format')()

#%%
sht=wb2.sheets[1]
sht.select()
sht.clear()
#table3
intial=(1,1)
xw.Range(intial).value=final_3
for i in range(0,3):
    xw.Range(intial,t2t(intial,shape_dict['3'])[i]).select()
    wb1.macro('format')()
color(intial,shape_dict['3']) 
#table4
intial=(shape_dict['3'][0]+3,1)
xw.Range(intial).value=final_4
for i in range(0,3):
    xw.Range(intial,t2t(intial,shape_dict['4'])[i]).select()
    wb1.macro('format')()
color(intial,shape_dict['4']) 
wb1.macro('total_format_largdata')()
#%%
#detial
sht=wb2.sheets['明細']
sht.select()
sht.clear()

xw.Range((1,1)).value=detial
rng=xw.Range('A1')
rng=rng.current_region
rng.autofit()
wb1.macro('format')()
#%%
wb1.close()
wb2.save('.\\備份\\權證發行分析'+date_str+'.xlsx')
file='權證發行分析'+date_str+'.xlsx'
wb2.close()
app.quit()
import range_vol


#%%
# =============================================================================
# Email
# =============================================================================
from email.mime.text import MIMEText               #專門傳送正文
from email.mime.multipart import MIMEMultipart     #傳送多個部分
from email.mime.application import MIMEApplication #傳送附件
from smtplib  import SMTP   #傳送郵件

df_mail=pd.read_excel(r'Email.xls')
Email_list=df_mail['Email'].values.tolist()
Email_str=','.join(Email_list)
send_user = '新金融商品處<ewarrant@jsun.com>'   #發件人
receive_users = Email_str   #收件人，可為list
subject = date_str+'週發行動態分析'  #郵件主題

email_text = '\
資料路徑:\r\n\
    備份檔案  : T:\權證發行動態分析\權證發行動態分析_python\備份\{file_name}\r\n\
    寄件者設定: T:\權證發行動態分析\權證發行動態分析_python\Email_Future.xls\r\n\r\n\
目前波動率檢核(同業權證選條件):\r\n\
    價內外<10 and 最後委買價>0.3\r\n\
    距到期日>=90,檔數=3' #郵件正文

server_address = 'mx00.jsun.com'   #伺服器地址

#構造一個郵件體：正文 附件
msg = MIMEMultipart()
msg['Subject']=subject    #主題
msg['From']=send_user      #發件人
msg['To']=receive_users           #收件人

#構建正文
email_text=email_text.format(file_name=file)
part_text=MIMEText(email_text)
msg.attach(part_text)             #把正文加到郵件體裡面去
path=r'T:\權證發行動態分析\權證發行動態分析_python\備份\權證發行分析'+date_str+'.xlsx'
#構建郵件附件
part_attach1 = MIMEApplication(open(path,'rb').read())   #開啟附件
part_attach1.add_header('Content-Disposition','attachment',filename=path[-19::]) #為附件命名
msg.attach(part_attach1)   #新增附件
# 傳送郵件 SMTP
smtp= SMTP(server_address,25)  # 連線伺服器，SMTP_SSL是安全傳輸

smtp.sendmail(send_user, Email_list, msg.as_string())  # 傳送郵件
print('郵件傳送成功！')

