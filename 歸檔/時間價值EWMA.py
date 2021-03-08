import datetime
import scipy.stats as st
import numpy as np
from sys import path
path.append(r'C:\Users\F128537908\Desktop\WCFADOX_Python')
from WCFAdox import PCAX
import pandas as pd
import re
from itertools import chain
from pandas.tseries.offsets import BDay
import matplotlib.pyplot as plot
from email.mime.text import MIMEText               #專門傳送正文
from email.mime.multipart import MIMEMultipart     #傳送多個部分
from smtplib  import SMTP   #傳送郵件
import os
plot.rcParams['font.sans-serif'] = 'Arial Unicode MS'
def error(errortype,date_str):
    df_mail=pd.read_excel(r'Email_發行評等.xls')
    Email_list=df_mail['Email'].dropna().values.tolist()
    Email_str=','.join(Email_list)
    send_user = '新金融商品處<ewarrant@jsun.com>'   #發件人
    receive_users = Email_str   #收件人，可為list
    subject = "(Error)"+date_str+'發行人評等'  #郵件主題
    text="錯誤訊息:"+errortype
    #郵件正文
    server_address = 'mx00.jsun.com'   #伺服器地址
    #構造一個郵件體：正文 附件
    #構建正文
    msg = MIMEText(text,'plain','utf-8')
    #把正文加到郵件體裡面去
    msg['Subject']=subject    #主題
    msg['From']=send_user      #發件人
    msg['To']=receive_users           #收件人
    try:
        # 傳送郵件 SMTP
        smtp= SMTP(server_address,25)  # 連線伺服器，SMTP_SSL是安全傳輸
        # #smtp.login(send_user, password)
        smtp.sendmail(send_user, Email_list, msg.as_string())  # 傳送郵件
    except smtplib.SMTPException:
        print( "Error: 無法傳送郵件")
from tool import workingday,Premium_cal,TheoryPrice_BidVol_Theta
#%%
# =============================================================================
# 一.資料下載
# 1.table1(權證評估表+權證基本資料表)
# 2.table2(日收盤表排行+日自營商進出排行)
# 3.類別資料
# =============================================================================
# table1
from data import data_request
PX=PCAX("192.168.65.11");date_str=datetime.datetime.strftime(datetime.date.today()-BDay(1),'%Y%m%d')

table=data_request(date_str,類別=True)

# df0=PX.Mul_Data("日常用技術指標表","D",date_str,colist='股票代號,[EWMA波動率(%)]')
# df0=df0.set_index("股票代號")
# df0=df0.apply(pd.to_numeric)

# col="代號,名稱,權證收盤價,[發行數量(千)],價內金額,標的收盤價,最新履約價,發行價格,\
#     [近一月歷史波動率(%)],[近三月歷史波動率(%)],[近六月歷史波動率(%)],權證成交量,最後委買價"
# df1=PX.Mul_Data("權證評估表","D",date_str,colist=col)
# df1=df1.set_index("代號")

# col="代號,發行日期,到期日期,標的代號,最新執行比例,發行機構名稱"
# year=str(datetime.date.today().year)
# df2=PX.Mul_Data("權證基本資料表","Y",year,colist=col)
# df2=df2.set_index("代號")
# data1=pd.concat([df1,df2],axis=1,join='inner')
# data1=data1[~((data1.index.str[-1]=='B') | (data1.index.str[-1]=='X') | (data1.index.str[-1]=='C'))]
# to_num=data1.columns[1:12].to_list()+ [data1.columns[-2]]
# data1[to_num]=data1[to_num].apply(pd.to_numeric)
# data1[['發行日期','到期日期']]=data1[['發行日期','到期日期']].apply(pd.to_datetime)
# for string in ['發行日期','到期日期']:
#     data1[string]=data1[string].dt.date

# data1['平均歷史波動率']=data1[['近一月歷史波動率(%)','近三月歷史波動率(%)','近六月歷史波動率(%)']].mean(axis=1)
# data1['EWMA']=data1['標的代號'].map(lambda x:df0.loc[x].values[0])
# data1=data1.drop(columns=['近一月歷史波動率(%)','近三月歷史波動率(%)','近六月歷史波動率(%)'])
# # table2
# col="股票代號,自營商庫存,自營商買賣超,[自營商買賣超金額(千)]"
# df3=PX.Pal_Data("日自營商進出排行","D",date_str,date_str,colist=col)
# df3=df3.set_index("股票代號")
# df4=PX.Mul_Data("日收盤表排行","D",date_str,colist="股票代號,[成交金額(千)]",ps="<CM代號,權證>")
# df4=df4.set_index("股票代號")
# data2=pd.concat([df3,df4],axis=1,join='inner')
# data2=data2.apply(pd.to_numeric)


# table=pd.concat([data1,data2],axis=1,join='inner')
# table['flag']=table.index.str[-1].to_series().apply(lambda  s: s.lower() if s=='P' else 'c' ).to_list()
#%%
# =============================================================================
# 二.
# 1.總表計算
'''計算
1.1. 理論價計算
1.2. 賣超金額、估計今日獲利、在外市值、發行金額
1.3. 市價時間價值、理論價時間價值、委買價時間價值
1.4. 市價委買價偏離金額、市價理論價偏離金額
''' 
# 2.交易員計算
'''
1.資料下載
2.建立Stock2Trader表
'''
#3.table輸出
'''
1. pivot_table(index=['發行機構名稱'])
2. 理論到期獲利率
3.格式
'''
#4.交易員輸出
'''
1.全市場(交易員標的)
2.日盛(交易員標的)
'''
# =============================================================================
#1.TABLE計算
table=TheoryPrice_BidVol_Theta(table,date_str)
'''計算
1.賣超金額
2.估計今日獲利
3.在外市值
4.發行金額
''' 
#1.2
table['賣超金額']=table[['自營商買賣超','權證收盤價']].apply(lambda col:  col[0]*col[1] if col[0]<0 else 0,axis=1)
table['估計今日獲利']=(-1000*table['自營商買賣超金額(千)'])-(table['自營商買賣超']*-1*1000*table['理論價'])
table['在外市值']=table['自營商庫存']*-1*table['權證收盤價']*1000
table['發行金額']=table['發行價格']*table['發行數量(千)']*1000
df=pd.DataFrame(index=table.index)
#1.3
for index,values in table[['價內金額','權證收盤價','理論價','最後委買價','自營商庫存']].iterrows():
    價內金額,市價,理論價,委買價,在外張數=values
    在外張數=-在外張數
    if 價內金額>0:
        df.loc[index,'市價時間價值']=(市價-價內金額)*在外張數*1000
        df.loc[index,'理論價時間價值']=(理論價-價內金額)*在外張數*1000
        df.loc[index,'委買價時間價值']=(委買價-價內金額)*在外張數*1000
    elif 價內金額<=0:
        df.loc[index,'市價時間價值']=市價*在外張數*1000
        df.loc[index,'理論價時間價值']=理論價*在外張數*1000
        df.loc[index,'委買價時間價值']=委買價*在外張數*1000         
        
df['市價委買價偏離金額']=df['市價時間價值']-df['委買價時間價值']
df['市價理論價偏離金額']=df['市價時間價值']-df['理論價時間價值']
table=pd.concat([table,df],axis=1,join='inner')
del 價內金額,市價,理論價,委買價,在外張數
#%%  2.交易員計算
#資料下載
Trader_Table=pd.read_html('http://172.16.10.13/hedge/makerauthlist.php',encoding='big5')[0]
Trader_Table['群組代碼']=Trader_Table['群組代碼'].astype(str)
Trader_Table.set_index("股票代號",inplace=True)
Trader_Table=Trader_Table.rename(index={'FITX':'TWA00'})

#建立Stock2Trader表
Stock2Trader={}
for stock in set(Trader_Table.index.values):
    if type(Trader_Table.loc[stock,'群組代碼'])==str:
        Stock2Trader[stock]=Trader_Table.loc[stock,'群組代碼']                    
    else:                        
        Stock2Trader[stock]=Trader_Table.loc[stock,'群組代碼'].values[0]

table['交易員']=table['標的代號'].map( lambda s: Stock2Trader.setdefault(s, np.nan))

del Trader_Table,stock,Stock2Trader
#%%  3.table輸出 
#TABLE     
list_fot_table=['在外市值','市價時間價值','理論價時間價值','委買價時間價值','市價理論價偏離金額','市價委買價偏離金額'
                ,'成交金額(千)','發行金額','賣超金額','自營商買賣超金額(千)','估計今日獲利']
dic={i:'sum' for i in list_fot_table};dic.update({'發行金額':['sum','count']})
final=table.pivot_table(index=['發行機構名稱'],values=list_fot_table,aggfunc=dic)
final.columns = final.columns.map(lambda v:'_'.join(v) if v[1]=='count' else v[0])

columns_name=['發行金額_count','自營商買賣超金額(千)'];name_list=['掛牌檔數市佔率','淨買賣超金額']
final=final.rename(columns=dict(zip(columns_name,name_list)))

final['理論到期獲利率']=final['市價理論價偏離金額']/final['市價時間價值']
#格式
transforms={'中國信託綜合證':'中信','香港商麥格理':'麥格理','第一金證':'第一金'}
final.index=final.index.to_series()\
            .apply(lambda s: transforms[s] if s in transforms else s[:2] )\
            .to_list()

del list_fot_table,columns_name,name_list,dic,transforms,df,holiday
#%% 4.交易員輸出

#全市場(交易員標的)
list_fot_trader=['市價時間價值','理論價時間價值','委買價時間價值','在外市值','市價理論價偏離金額']
trader_table=table.pivot_table(index=['交易員'],values=list_fot_trader,aggfunc='sum')
trader_table['理論到期獲利率']=trader_table['市價理論價偏離金額']/trader_table['市價時間價值']
trader_table=trader_table.drop('市價理論價偏離金額',axis=1)
#交易員
temp=table[table['發行機構名稱']=='日盛證']
trader_table2=temp.pivot_table(index=['交易員'],values=list_fot_trader,aggfunc='sum')
trader_table2['理論到期獲利率']=trader_table2['市價理論價偏離金額']/trader_table2['市價時間價值']
trader_table2=trader_table2.drop('市價理論價偏離金額',axis=1)

del list_fot_trader,temp
#%%
# =============================================================================
# 輸出格式整理
# 1.數值
# 2.市佔率(百分比)
# 3.副總觀察
# 4.交易員
# 5.作圖
# 6.格式
# =============================================================================
#1.數值
final_1=final[['在外市值','市價時間價值','理論價時間價值','委買價時間價值','市價委買價偏離金額','市價理論價偏離金額','估計今日獲利','理論到期獲利率']]
final_1.loc[:,['理論到期獲利率']]=(final_1.loc[:,['理論到期獲利率']]*100).applymap(lambda v:round(v,2)).copy()
final_1=final_1.sort_values('市價時間價值',ascending=0)

#2.百分比
final_2=final[['市價時間價值','理論價時間價值','委買價時間價值','成交金額(千)','發行金額','掛牌檔數市佔率','賣超金額','淨買賣超金額']]
final_2=final_2/final_2.sum()*100
final_2=final_2.applymap(lambda v:round(v,2)).copy()
final_2=final_2.sort_values('市價時間價值',ascending=0)

#3.副總
final_1['發行效率']=round(final_1['市價時間價值']/ 1000000/final_2['掛牌檔數市佔率'],1)
final_1['成交額市價時間價值比']=round(final_2['成交金額(千)']/final_2['市價時間價值'],2)

#作圖
'時間價值市佔'
'理論到期獲利率'
'成交金額市佔/市價時間價值市佔'
x=final_1['發行效率']
y=final_1['理論到期獲利率']
fig1, ax =plot.subplots(figsize=(10,10))
ax.scatter(x, y,s=final_2['市價時間價值']*100,c=range(1,x.shape[0]+1),alpha=0.7)
for  i,txt in enumerate(final_1.index.values):
     ax.annotate(txt, (x[i], y[i]),size=18)
ax.set_xlabel('發行效率', fontsize=20)
ax.set_ylabel('造市價值(%)', fontsize=20)
fig1.savefig('temp.jpg')

#交易員
sign1=trader_table2['理論到期獲利率'].map(lambda s: 1 if s>0 else -1)
sign2=trader_table['理論到期獲利率'].map(lambda s: 1 if s>0 else -1)
mask=(sign1*sign2).map(lambda s: True if s==1 else False)

trader=round(trader_table2/trader_table*100,2)
trader['理論到期獲利率']=round(trader['理論到期獲利率']/100,1)
trader['理論到期獲利率'][~mask]=np.nan
trader['理論到期獲利率']=trader['理論到期獲利率']*sign1
import sqlite3
conn = sqlite3.connect(r'T:\每日產出報表\sql\券商時間價值.db')
def table_exist(conn, sheet_name,day):
    return list(conn.execute(
        "select count(*) from "+sheet_name+" WHERE DATE(date) =  '"+day+"'"))[0][0] == 0


d=datetime.datetime.strptime(date_str,"%Y%m%d").date()
lsd=d.strftime("%Y-%m-%d")
for i,(df_name,sheet_name) in enumerate(zip(['final_1','final_2'],['數值','百分比'])):
    #上傳格式
    to_sql=df_name+'_to_sql'
    exec(to_sql+'='+df_name+'.copy()')
    exec(to_sql+"['date']=datetime.datetime.strptime(date_str,'%Y%m%d')")
    exec(to_sql+".reset_index(inplace=True)")
    exec(to_sql+'='+to_sql+".rename(columns={'index':'券商'})")    
    exec(to_sql+'='+to_sql+".set_index(['date','券商'])")
    if table_exist(conn,sheet_name,lsd):
        eval(to_sql+".to_sql('"+sheet_name+"', conn, if_exists='append')")
    else:
        print("表格已存在")
del  final_1_to_sql,final_2_to_sql
#%%
成交金額=pd.Series([])
date=datetime.datetime.strptime(date_str,"%Y%m%d")
成交金額[date]=round(final['成交金額(千)'].sum()*1000/100000000,2)
成交金額.name='成交金額'
成交金額.to_sql('成交金額(億)', conn, if_exists='append')
        
        
#%%

# sql="select * from "+sheet_name
# ret=pd.read_sql(sql,conn,index_col=['date','券商'])
# ret.reset_index(inplace=True)
# ret = ret.drop_duplicates(['券商', 'date'], keep='last')
# ret = ret.sort_values(['date','券商']).set_index(['date','券商'])    
# ret.to_sql(sheet_name, conn, if_exists='replace')   
#圖(時間相關)
import matplotlib.dates as mdates
sheet_name='百分比'
start=(d-BDay(14)).strftime("%Y-%m-%d")
sql="select date,券商,市價時間價值 from "+sheet_name +" WHERE DATE(date) BETWEEN '"+start+"' AND '"+lsd +"'"+"AND 券商 in ('日盛','台新','兆豐','國泰','富邦','中信','元富')"
df=pd.read_sql(sql,conn)
df['date']=pd.to_datetime(df['date'])

fig2, ax = plot.subplots(figsize=(8,6))
for broker in  set(df['券商']):
    temp=df[df['券商']==broker]
    if broker=='日盛':
        ax.plot(temp['date'], temp['市價時間價值'], label=broker,marker = "*", markersize=12)
    else:
        ax.plot(temp['date'], temp['市價時間價值'], label=broker,marker = "o",linestyle=':')
ax.set_xticks(temp['date'])
ax.set_ylabel('時間價值市佔率(%)', fontsize=16)
ax.set_title('時間價值佔比', fontsize=24)
ax.set_ylim([0, 8])
lgd=ax.legend(bbox_to_anchor=(1.05, 0.5),fontsize=15)
ax.xaxis.set_major_formatter(mdates.DateFormatter("%m-%d"))
ax.grid(1)
_=plot.xticks(rotation=90, fontsize=16) 
fig2.savefig('temp2.jpg', bbox_extra_artists=(lgd,), bbox_inches='tight')



#%%
#格式
#final_1
final_1=final_1.apply(lambda s:round(s,2) if s.name  in ['理論到期獲利率','發行效率','成交額市價時間價值比'] else  s.astype(int))
final_1[['理論到期獲利率']]=final_1[['理論到期獲利率']].applymap(lambda x : "{:-3.2f}%".format(x))
final_1[['發行效率','成交額市價時間價值比']]=final_1[['發行效率','成交額市價時間價值比']].applymap(lambda x : "{:-3.2f}".format(x))
col_comma=[col  for col in final_1.columns.values if col not in['理論到期獲利率','發行效率','成交額市價時間價值比'] ]
final_1[col_comma]=final_1[col_comma].applymap(lambda x : "{:,}".format(x))


from itertools import product
col1=['在外市值','市價','理論','委買','市價委買價','市價理論價','理論今日獲利','造市價值','發行效率','成交市價時間價值比']
name=chain([('','在外市值')] , product(['時間價值'],col1[1:4]) , product(['偏離金額'],col1[4:6]) , product(['獲利'],col1[6:8]), product(['指標'],col1[8:]) )
final_1.columns=pd.MultiIndex.from_tuples(list(name))



#final_2
final_2=final_2.applymap(lambda x : "{:-3.2f}%".format(x))
column=['市價','理論價','委買價','成交金額','發行金額','掛牌檔數','賣超金額','淨賣超金額']
name=chain(product(['時間價值市佔率'],column[0:3]) ,product(['市佔率'],column[3:]))
final_2.columns=pd.MultiIndex.from_tuples(list(name))


#交易員
trader.iloc[:,:4]=trader.iloc[:,:4].astype(str)+'%'
col=trader.columns.to_list()
name=chain(product(['市佔率'],col[:4]) ,product(['比值'],col[4:]))
trader.columns=pd.MultiIndex.from_tuples(list(name))

#%%
# =============================================================================
# Email
'''
'''
# =============================================================================
#df_mail=pd.read_excel(r'T:\權證發行人評等\python版資料\Email_發行評等.xls')
# Email_list=df_mail['Email'].dropna().values.tolist()
# Email_str=','.join(Email_list)
Email_str='leo812leo@jsun.com,Damonchen@jsun.com,teru7233@jsun.com,wen996@jsun.com,jacklin@jsun.com,ykpang@jsun.com'
Email_list=Email_str.split(',')

send_user = '新金融商品處<ewarrant@jsun.com>'   #發件人
receive_users = Email_str   #收件人，可為list
subject = date_str+'券商時間價值EWMA(合計)'  #郵件主題

html = """
<html><body>
<p><font size="6"><b>表格:</b></font></p>
<strong>券商時間價值(數值)</strong>
{table1}
造市價值=理論到期獲利<br>
<br>
<strong>券商時間價值(百分比)</strong>
{table2}
<br>
<strong>交易員時間價值</strong>
日盛全市場比
{table3}
<hr/>
<p><font size="6"><b>作圖</b></font></p>
形狀大小為市值佔比<br>

"""
#郵件正文

server_address = 'mx00.jsun.com'   #伺服器地址
mail_type = '1'    #郵件型別

#構造一個郵件體：正文 附件
##-------------一日------------------##
table_1=final_1.to_html()
table_1=re.sub(r'halign="left">','halign="left" bgcolor="#9D9D9D">',table_1,flags=re.S)
table_1=re.sub(r'<th>.*?','<th bgcolor="#9D9D9D">',table_1,flags=re.S)
table_1=re.sub(r'<tr>\n      <th bgcolor="#9D9D9D">日盛</th>'\
                  ,'<tr bgcolor="#FFAD86">\n      <th bgcolor="#9D9D9D">日盛</th>',table_1,flags=re.S)
table_1=re.sub(r'<td>',r'<td align="right">',table_1,flags=re.S)   

    
table_2=final_2.to_html()
table_2=re.sub(r'halign="left">','halign="left" bgcolor="#9D9D9D">',table_2,flags=re.S)
table_2=re.sub(r'<th>.*?','<th bgcolor="#9D9D9D">',table_2,flags=re.S)
table_2=re.sub(r'<tr>\n      <th bgcolor="#9D9D9D">日盛</th>'\
                  ,'<tr bgcolor="#FFAD86">\n      <th bgcolor="#9D9D9D">日盛</th>',table_2,flags=re.S)    
table_2=re.sub(r'<td>',r'<td align="right">',table_2,flags=re.S)   

table_Trader=trader.to_html()
table_Trader=re.sub(r'<tr>\s+ <th>交易員</th>.*? </tr>','',table_Trader,flags=re.I|re.M|re.S)
table_Trader=re.sub(r'<th ','<th bgcolor="#9D9D9D" ',table_Trader,flags=re.S)
table_Trader=re.sub(r'<th>','<th bgcolor="#9D9D9D">',table_Trader,flags=re.S)
table_Trader=re.sub(r'<td>',r'<td align="right">',table_Trader,flags=re.S)  

##-------------15日------------------##

##-------------輸出------------------##
html = html.format(table1=table_1,table2=table_2,table3=table_Trader)
    
#構建正文

#把正文加到郵件體裡面去
msg = MIMEMultipart()
msg['Subject']=subject    #主題
msg['From']=send_user      #發件人
msg['To']=receive_users           #收件人

from email import encoders
from email.mime.base import MIMEBase
for index in range(len(['temp.jpg','temp2.jpg'])):
     html += '<img src="cid:' + str(index) + '" align="center" width=auto>'
     
msg.attach(MIMEText(html+'<br>有任何建議再通知我 昱霖</body></html>', 'html'))
index=0
for file in ['temp.jpg','temp2.jpg']:
    with open(file, 'rb') as f:
        # 設定附件的MIME和檔名，這裡是png型別:
        mime = MIMEBase('image', 'png', filename=file)
        # 把附件的內容讀進來:
        mime.set_payload(f.read())
        
        # 加上必要的頭資訊:
        mime.add_header('Content-Disposition', 'attachment', filename='tp.png')
        mime.add_header('Content-ID', '<' + str(index) + '>')
        mime.add_header('X-Attachment-Id', '0')
        # 用Base64編碼:
        encoders.encode_base64(mime)
        # 新增到MIMEMultipart:
        index += 1
        msg.attach(mime)
        f.close()

# 傳送郵件 SMTP
smtp= SMTP(server_address,25)  # 連線伺服器，SMTP_SSL是安全傳輸

# #smtp.login(send_user, password)
smtp.sendmail(send_user, Email_list, msg.as_string())  # 傳送郵件
print('------------------') 
print('郵件傳送成功！')
print('------------------') 
try:
    os.remove("temp.jpg")
    os.remove("temp2.jpg")
except OSError as e:
    print(e)
else:
    print("File is deleted successfully")
