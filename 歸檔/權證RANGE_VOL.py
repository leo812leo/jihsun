import datetime
from sys import path
path.append(r'C:\Users\F128537908\Desktop\WCFADOX_Python')
import pandas as pd
from pandas.tseries.offsets import BDay
import matplotlib.pyplot as plot
from email.mime.text import MIMEText               #專門傳送正文
from smtplib  import SMTP   #傳送郵件
from tool import workingday,Premium_cal
from data import data_request
from range_vol import run
import re
from itertools import chain
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


#%%
# table
date_str=datetime.datetime.strftime(datetime.date.today()-BDay(1),'%Y%m%d')
table=data_request(date_str,類別=True)
range_vol=run(t0=1)[60]
#%% 理論價計算
table['range_vol']=table['標的代號'].map(range_vol)
temp=[]
for _,values in table[['標的收盤價','最新履約價','range_vol','到期日期','flag','最新執行比例']].iterrows():
    s,k,sigma,enddate,flag,ratio = values
    sigma=sigma/100
    r=0
    t=workingday(datetime.datetime.strptime(date_str,"%Y%m%d").date(),enddate)/250
    temp.append(Premium_cal(flag,s,k,r,sigma,t,ratio))
table['理論價'] = temp
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

del list_fot_table,columns_name,name_list,dic,transforms,df

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
# =============================================================================
# #格式
# =============================================================================
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






# =============================================================================
# Email
# =============================================================================
#df_mail=pd.read_excel(r'T:\權證發行人評等\python版資料\Email_發行評等.xls')
# Email_list=df_mail['Email'].dropna().values.tolist()
# Email_str=','.join(Email_list)
Email_str='leo812leo@jsun.com,Damonchen@jsun.com,teru7233@jsun.com,wen996@jsun.com,jacklin@jsun.com,ykpang@jsun.com'
Email_list=Email_str.split(',')

send_user = '新金融商品處<ewarrant@jsun.com>'   #發件人
receive_users = Email_str   #收件人，可為list
subject = date_str+'券商時間價值Range(合計)'  #郵件主題

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

##-------------輸出------------------##
html = html.format(table1=table_1,table2=table_2)
    
#構建正文
from email.mime.multipart import MIMEMultipart 
#把正文加到郵件體裡面去
msg = MIMEMultipart()
msg['Subject']=subject    #主題
msg['From']=send_user      #發件人
msg['To']=receive_users           #收件人
msg.attach(MIMEText(html+'<br>有任何建議再通知我 昱霖</body></html>', 'html'))

# 傳送郵件 SMTP
smtp= SMTP(server_address,25)  # 連線伺服器，SMTP_SSL是安全傳輸

# #smtp.login(send_user, password)
smtp.sendmail(send_user, Email_list, msg.as_string())  # 傳送郵件
print('------------------') 
print('郵件傳送成功！')
print('------------------') 




