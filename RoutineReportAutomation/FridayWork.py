from datetime import datetime
import xlwings as xw 
import os
import shutil
from tkinter import filedialog
from pandas.tseries.offsets import BDay
import pandas as pd

# =============================================================================
date=datetime.now()-BDay(1)
# =============================================================================



temp={}
#找最新excel檔
for file_type,folder in zip(['excel','word'],['風險管理週報(EXCEL)','週報backup(WORD)']):
    path="T:\\權證風險管理週報\\"+folder+"\\"+str(date.year)
    listdiration=os.listdir(path)
    listdiration=[path+'\\'+string for string in listdiration]
    listdiration.sort(key=lambda x: os.path.getmtime(x))
    temp[file_type]=listdiration[-1]
#命名
str_name_e=datetime.strftime(date,'%m%d')
str_name_w=str(datetime.now().year-1911) + datetime.strftime(date,'%m%d')
new_name_e=temp['excel'][:-9]+str_name_e+'.xlsm'
new_name_w=temp['word'][:-11]+str_name_w+'.doc'

#新增檔案
for file_type,new_name in zip(['excel','word'],[new_name_e,new_name_w]):
    if os.path.isfile(new_name):
      print(file_type+"檔案存在!!!!!")
    else:
      shutil.copy(temp[file_type],new_name)
      print(file_type+"產生完成")
#選擇excel檔案
import tkinter as tk
root = tk.Tk()
root.withdraw()
file_path = filedialog.askopenfilename()    
app = xw.App(visible=True, add_book=False)
# =============================================================================
# 複製貼上，並執行巨集
# =============================================================================
wb1 = app.books.open(file_path)  #資料來源
wb2 = app.books.open(new_name_e) #資料修改

#刪除'new-1','new-2'
wb2.sheets['new-1'].delete();wb2.sheets['new-2'].delete()

for sheet_num in ['-1','-2']:
    sheetname=str_name_e+sheet_num
    ws1=wb1.sheets[sheetname]
    ws1.api.Copy(Before=wb2.sheets['均量'].api)
    wb2.sheets[sheetname].name='new'+sheet_num
#修改風險上限
ws2=xw.sheets['new-2']
ws2.select()
df = xw.Range('A3').options(pd.DataFrame, 
                             header=1,
                             expand='table').value
tem=df[df['Time Decay(合)']=='風險上限']  
tem=tem.iloc[:,-4:-1]    

xw.sheets['市場風險'].select()
xw.Range('C3').value =tem.T.values

#執行巨集
def click():
    wb2.macro('CM全部更新')()
    wb2.macro('查詢網頁')()
    wb2.macro('更新彙總表格')()
    wb2.macro('市場風險')()
ws2=wb2.sheets['日盛淨值']
ws2.select()
#抓網頁上的資料並做運算
import requests
from io import StringIO
import pandas as pd
url='https://brks.twse.com.tw/server-java/t17sc13_1'
form_data = {
    'step':1,
    'CK':1,
    'CN':1160,
    'YY':'2005',
    'SS':1}
#抓網站日盛淨值
r = requests.post(url, data=form_data)
r.encoding='big5'
df=pd.read_html(StringIO(r.text))[-1]
ws2.range('B2').value =df.iloc[0,5]
click()
#檢核表
ws2=wb2.sheets['檢核表']
ws2.select()
form_data = {'qrytrdday': datetime.strftime(date,'%Y/%m/%d')}
r = requests.post('http://172.16.10.13/hedge/Stockposdy_war.php', data=form_data)
r.encoding='utf-8'
df=pd.read_html(StringIO(r.text))[0]
ws2.range('A2:E100').clear()
ws2.range('A2').value =df.values

#end
wb2.save()
wb2.close()
wb1.close()
app.quit()
from sys import path
path.append(r'D:\code')

#corr
from package_corr import run
run(new_name_e,date)




