import datetime
from sys import path
path.append(r'T:\WCFADOX_Python')
from WCFAdox import PCAX
import pandas as pd
import re
from itertools import chain,product
import numpy as np
from pandas.tseries.offsets import CustomBusinessDay
from email.mime.text import MIMEText               #專門傳送正文
from email.mime.multipart import MIMEMultipart     #傳送多個部分
from smtplib  import SMTP   #傳送郵件
PX=PCAX("192.168.65.11")
#%%CustomBusinessDay
df=pd.read_csv('https://www.twse.com.tw/holidaySchedule/holidaySchedule?response=csv&queryYear='+\
               str(datetime.datetime.now().year-1911),header=1,encoding='big5')
tem=[re.findall(r'(\d+月\d+日)',i.values[0])for _,i in df[[' 日期']].iterrows()]
holiday=[(datetime.datetime.strptime(str(datetime.datetime.now().year)+i, "%Y%m月%d日")).date() for i in chain(*tem)]
cb=CustomBusinessDay(1,holidays=holiday)
last_10_day=datetime.datetime.strftime(datetime.date.today()-cb*10,"%Y%m%d") 
today=datetime.datetime.strftime(datetime.date.today(),"%Y%m%d") 
del tem,df
#%%
df1=PX.Pal_Data("日券商分點排行Top15(權證)","D",last_10_day,today) #10日全市場權證資料
df1=df1[df1.代號.str[-1].str.contains('[0-9P]')]
df1=df1.replace('',np.nan)
df1=df1.set_index(['日期','代號'])
df2=df1[df1.名稱.str.contains('日盛')]

dic={}
cols=[("{0}張{1}(券商),{0}張{1}".format(state,i)).split(',') for state,i in product('賣買',range(1,6,1))]
for col in cols:
    for index,values in df1[col].dropna().iterrows():
        日期,代號=index
        分點,買賣張=values
        if col[0][0]=='買':
            dic.setdefault(分點,{}).setdefault(代號,{}).update({日期:+int(買賣張)})
        else:
            dic.setdefault(分點,{}).setdefault(代號,{}).update({日期:-int(買賣張)})            
def check(series,ratio,volumn):
    if  series[0] > volumn and series[1]/series[0] <= -ratio and series[0]>0:
        return -series[1]/series[0] 
    else:
        return np.nan
output={}
for 分點 in set(dic):   
    output[分點]=round(pd.DataFrame(dic[分點]).rolling(2,min_periods=2)\
                     .apply(lambda x: check(x,0.8,100),raw=True).sum()
                     .sum().sum())     
final=pd.Series(output)   
ninety=set(final[final>=final.quantile(q=0.9)].index)
del df1,output,col,dic,index,values,日期,代號,分點
#%%
df2=df2.loc[today]
dic={}
for col in cols:
    for index,values in df2[col].dropna().iterrows():
        代號=index
        分點,買賣超=values
        if col[0][0]=='買':
            dic.setdefault(分點,{}).setdefault(代號,{}).update({'買':+int(買賣超)})
        else:
            dic.setdefault(分點,{}).setdefault(代號,{}).update({'賣':-int(買賣超)})
new_keys=set(dic) & ninety
dict_you_want = { new_key: dic[new_key] for new_key in new_keys }
del dict_you_want['日盛']

dic2={}
for 分點,dicc in dict_you_want.items():
    for warrant,values  in dicc.items():
        for _,value  in values.items():        
            if abs(value) >=200:
                
                if value>0:   
                    dic2.setdefault(warrant,{}).setdefault('倒回',[]).append(分點+"："+str(value))
                if value<0:   
                    dic2.setdefault(warrant,{}).setdefault('攻擊',[]).append(分點+"："+str(value))
                
warrant_list=list(dic2.keys())
Trader_Table=pd.read_html('http://172.16.10.13/hedge/makerauthlist.php',encoding='big5')[0]
Trader_Table['交易員']=Trader_Table['交易員'].apply(lambda x: x.strip())
Trader_Table=Trader_Table.rename(index={'FITX':'TWA00'})   
         
Trader_Table=Trader_Table.query("權證代碼 in @warrant_list")
Trader_Table['被倒回']=Trader_Table['權證代碼'].map(lambda n: " ; ".join( dic2[n].setdefault('倒回','') ))
Trader_Table['被攻擊']=Trader_Table['權證代碼'].map(lambda n: " ; ".join( dic2[n].setdefault('攻擊','') ))

Trader_Table=Trader_Table.drop(['真實群組代碼','群組代碼','授權代碼'],axis=1)
Trader_Table=Trader_Table.set_index(['交易員','股票代號'])
Trader_Table=Trader_Table[['股票名稱','權證代碼', '權證名稱',  '被倒回', '被攻擊']]
Trader_Table=Trader_Table.sort_index()

#%%
# =============================================================================
# Email
'''
'''
# =============================================================================
#df_mail=pd.read_excel(r'T:\權證發行人評等\python版資料\Email_發行評等.xls')
# Email_list=df_mail['Email'].dropna().values.tolist()
# Email_str=','.join(Email_list)
Email_str='leo812leo@jsun.com,jacklin@jsun.com,Damonchen@jsun.com,teru7233@jsun.com,dingyichen@jsun.com'
#Email_str='leo812leo@jsun.com'
Email_list=Email_str.split(',')

send_user = '新金融商品處<ewarrant@jsun.com>'   #發件人
receive_users = Email_str   #收件人，可為list
subject = today+'權證隔日衝'  #郵件主題

html = """
<html><body>
<p><font size="6"><b>隔日衝券商統計:</b></font></p>
{table1}
<br>
<br>
"""
#郵件正文

server_address = 'mx00.jsun.com'   #伺服器地址
mail_type = '1'    #郵件型別

#構造一個郵件體：正文 附件
##-------------輸出------------------##
Table=Trader_Table.to_html()
Table=re.sub(r'<tr style="text-align: right;">','<tr>',Table,flags=re.S)
Table=re.sub('<tr>\s+<th>交易員</th>\s+<th>股票代號</th>[\s+<th></th>]*?</tr>','',Table)
Table=re.sub(r'<tr>','<tr bgcolor="#BABABA">',Table,flags=re.S,count=1)
Table=re.sub(r'<th','<th bgcolor="#BABABA"',Table,flags=re.S)
Table=re.sub(r'<th bgcolor="#BABABA"ead>','<thead>',Table,flags=re.S)

html = html.format(table1=Table)
#構建正文

#把正文加到郵件體裡面去
msg = MIMEMultipart()
msg['Subject']=subject    #主題
msg['From']=send_user      #發件人
msg['To']=receive_users           #收件人
msg.attach(MIMEText(html+'<br>有任何建議再通知我 昱霖</body></html>', 'html'))
#構建正文



# 傳送郵件 SMTP
smtp= SMTP(server_address,25)  # 連線伺服器，SMTP_SSL是安全傳輸
# #smtp.login(send_user, password)
smtp.sendmail(send_user, Email_list, msg.as_string())  # 傳送郵件
print('------------------') 
print('郵件傳送成功！')
print('------------------') 
