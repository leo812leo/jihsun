from sys import path
path.append(r'T:\WCFADOX_Python')
import pandas as pd
from WCFAdox import PCAX
PX=PCAX("192.168.65.11")
from talib import abstract     
import matplotlib.pyplot as plt
from plot_candles import plot_candles
plt.rcParams['font.sans-serif'] = 'Arial Unicode MS'

def plot_candles_for_broker(start,end,stock,df_warrant):
    
    col='日期,股票代號,開盤價,最高價,最低價,收盤價,成交量'
    df=PX.Pal_Data("日收盤表排行","D",
                   start,
                   end,colist=col,isst='Y'
                   ,ps='<CM特殊,2012>')
    df.columns=['date','stock_id','open','high','low','close','volume']
    df['date']=df['date'].apply(pd.to_datetime)
    df=df.set_index('date')
    df=df.sort_index(ascending=True)
    
    num=['open','high','low','close','volume']
    df[num]=df[num].apply(pd.to_numeric)
    df_stock=df[df['stock_id']==stock]
    
    # 創建各種指標
    SMA20 = abstract.SMA(df_stock,timeperiod=20)

    plot_candles(
                 # 起始時間、結束時間
                 start_time=start,
                 end_time=end,
             
                 # 股票的資料
                 pricing=df_stock, 
                 title=stock, 
                 # 是否畫出成交量？
                 volume_bars=False, 
                 # 將某些指標（如SMA）跟 K 線圖畫在一起
                 overlays=[SMA20]
                )
    loc_num=[df_stock.index.get_loc(df_warrant.日期.values[i])for i in range(len(df_warrant.日期.values))]
    minus=df_stock.close.std()
    minus_list=[minus*(i+1)*0.8 for i in range(len(loc_num))]
    for i,war,pp,oi in zip(loc_num,df_warrant.index.values,minus_list,df_warrant.自營商買賣超.values):
        plt.axvline(x=i,c='k')
        plt.annotate(war+'\n'+str(oi)+'張', (i,df_stock['close'][i]-pp),size=14)
    