import datetime
from sys import path
path.append(r'T:\WCFADOX_Python')
from WCFAdox import PCAX
import pandas as pd
from pandas.tseries.offsets import BDay
from tool import Premium_cal,workingday,calculate,last_x_workday
PX=PCAX("192.168.65.11")
def Best_sale_warrant(StockList,t0):
    # =============================================================================
    # 日期
    # =============================================================================
    today=(datetime.date.today()-BDay(t0)).date()
    #%%資料獲取  
    #取資料(25天市場上權證資料)
    col='日期,代號,名稱,標的代號,上市日期,到期日期,最新執行比例,[存續期間(月)],發行價格,標的收盤價,最新履約價,價內金額,權證收盤價,最後委買價,[隱含波動率(委買價)(%)]'
    df1=PX.Pal_Data("權證評估表","D",
                    datetime.datetime.strftime(last_x_workday(26,today),"%Y%m%d"),
                    datetime.datetime.strftime(last_x_workday(1,today),"%Y%m%d"),colist=col) #25日全市場權證資料
    #資料整理
    df1=df1[~((df1.代號.str[-1]=='B') | (df1.代號.str[-1]=='X') | (df1.代號.str[-1]=='C'))]
    df1['券商']=df1['名稱'].str.extract(r'(..)..[售購]')
    df1['flag']=df1.代號.str[-1].apply(lambda  s: s.lower() if s=='P' else 'c' ).to_list()
    
    #其他加篩選(同標的、日期25日)
    df1=df1.query('標的代號 in @StockList')
    to_num=['最新執行比例','存續期間(月)','標的收盤價','最新履約價','價內金額','發行價格','權證收盤價','最後委買價','隱含波動率(委買價)(%)']
    df1[to_num]=df1[to_num].apply(pd.to_numeric)
    to_date=['日期','上市日期','到期日期']
    df1[to_date]=df1[to_date].applymap(pd.to_datetime)
    df1['價內外程度(%)']=df1[['標的收盤價','最新履約價','flag']].apply(calculate ,axis=1)
    start_date = datetime.datetime.strftime(last_x_workday(26,today),"%Y-%m-%d")
    end_date   = datetime.datetime.strftime(last_x_workday(1,today),"%Y-%m-%d") #25日 新發行權證資料 
    df1=df1.query('上市日期>= @start_date and 上市日期 <= @end_date ')   
    
    del to_num,to_date,start_date,end_date,col
    #%%發行分析
    '''
    觀察近20日日盛 新發權證
    比較權證為近25日同業 同標的權證
    
    最大時間價值 日期,參數
    目前最大時間價值 參數
    自家權證狀況
    '''
    #庫存
    warrant=list(set(df1.代號))
    col="日期,股票代號,自營商庫存,自營商買賣超"
    df3=PX.Pal_Data("日自營商進出排行","D",datetime.datetime.strftime(last_x_workday(26,today),"%Y%m%d"),
                                          datetime.datetime.strftime(last_x_workday(1,today),"%Y%m%d"),colist=col)
    df3=df3.query('股票代號 in @warrant')
    df3[['自營商庫存','自營商買賣超']]=df3[['自營商庫存','自營商買賣超']].apply(pd.to_numeric)
    df3=df3.rename(columns={'股票代號':'代號'})
    to_date=['日期']
    df3[to_date]=df3[to_date].applymap(pd.to_datetime)
    
    
    #table(計算市價時間價值)
    table=pd.merge(df1,df3, on=['代號','日期'],how='inner')
    table=table[(table['券商']!='元大') & (table['flag']=='c')]
    table[['自營商庫存','自營商買賣超']]=-table[['自營商庫存','自營商買賣超']]
    
    del col,df3,to_date,df1
    #%%col計算
    
    for index,values in table[['價內金額','權證收盤價','最後委買價','自營商庫存','自營商買賣超']].iterrows():
        價內金額,市價,委買價,在外張數,買賣超=values
        if 價內金額>0:
            table.loc[index,'市價時間價值']=(市價-價內金額)*在外張數*1000
            table.loc[index,'市價時間價值增量']=(市價-價內金額)*買賣超*1000        
        elif 價內金額<=0:
            table.loc[index,'市價時間價值']=市價*在外張數*1000 
            table.loc[index,'市價時間價值增量']=(市價-價內金額)*買賣超*1000
            
    for index,values in table[['日期','上市日期','到期日期']].iterrows(): 
        日期,上市日期,到期日期=values
        table.loc[index,'交易天數']=workingday(上市日期.date(),日期.date())+1
        table.loc[index,'距到期日']=workingday(日期.date(),到期日期.date())
        table.loc[index,'距到期日(日曆)']= (到期日期.date()-日期.date()).days
    def pl(flag,s,k,r,sigma,t,ratio,oi,up_down):
        option_pl=-(Premium_cal(flag,s*(1+up_down/100),k,r,sigma,(t-1/365),ratio)-Premium_cal(flag,s,k,r,sigma,t,ratio))*1000*oi
        return option_pl
    
    for index,values in table[['標的收盤價','最新履約價','隱含波動率(委買價)(%)','距到期日(日曆)','自營商庫存','flag','最新執行比例']].iterrows():
        s,k,sigma,t,oi,flag,ratio=values
        sigma=sigma/100;r=0;t=t/365
        table.loc[index,'theta']= int(pl(flag,s,k,r,sigma,t,ratio,oi,0))
        
    #%%存量觀察
    to_date=['日期','上市日期','到期日期']
    table[to_date]=table[to_date].applymap(lambda x : x.date())
    finnal_output={}
    table=table.drop(['價內金額','權證收盤價'],axis=1)
    today2=datetime.datetime.strftime(last_x_workday(1,today),"%Y%m%d")
    日盛權證=table.query("券商=='日盛'")
    日盛權證=日盛權證[日盛權證['日期']==datetime.datetime.strptime(today2,"%Y%m%d").date()]
    # =============================================================================
    # 1.25天內最大時間價值
    # 2.目前最大時間價值
    # 3.25天內最大OI
    # 4.目前最大OI
    # =============================================================================
    def sort_jihsun_last(df):
        sort_list=list(set(df["券商"])-{'日盛'})
        sort_list.append("日盛")
        df['券商'] = pd.Categorical(df['券商'],sort_list)
        return df.sort_values('券商')
    rename_dict={'最新執行比例':'行使比例','隱含波動率(委買價)(%)':'委買vol',
                 '價內外程度(%)':'價內外','最新履約價':'履約價','最後委買價':'委買價','存續期間(月)':'存續(月)',
                 '自營商庫存':'OI','自營商買賣超':'OI增量'}
    
    sort_col=['日期','上市日期','到期日期',
              '券商','flag','行使比例','存續(月)',
              '發行價格','委買價','標的收盤價','履約價','價內外','委買vol','波動率pr',
             'OI','市價時間價值','theta','OI增量','市價時間價值增量',
             '交易天數',	'距到期日']
    def make_table(col,target=table ,whole=table ,table_jisun=日盛權證):
        temp=target.loc[target.groupby("代號")[col].idxmax()]            #25天內時間價值最大時 權證參數
        temp=temp.groupby('標的代號').apply(lambda s: s.nlargest(3,col))    #25天內 同標的 時間價值最大權證參數
        temp['波動率pr']=temp['代號'].map(vol_pct_dict)
        #最大oi增量
        warrant=list(set(temp.代號))
        temp_x=whole.loc[whole.groupby("代號")['自營商買賣超'].idxmax()]
        temp_x=temp_x[temp_x['代號'].isin(warrant)]
        temp_x=temp_x[['標的代號','代號','日期','自營商買賣超']]
        temp_x=temp_x.set_index(["標的代號","代號"])
        
        temp=temp.reset_index(drop=True).set_index(["代號","名稱"])
        table_jisun['波動率pr']=table_jisun['代號'].map(vol_pct_dict)
        table_jisun=table_jisun.reset_index(drop=True).set_index(["代號","名稱"])
        temp=temp.append(table_jisun)
        temp=temp.groupby('標的代號').apply(sort_jihsun_last) #把日盛排在最後
        temp=temp.drop('標的代號',axis=1)
        temp=temp.rename(columns=rename_dict)
        temp=temp[sort_col]
        temp=temp[~temp.index.duplicated(keep='first')]
        return temp,temp_x
    #資料處理_今日
    now=table[table['日期']==datetime.datetime.strptime(today2,"%Y%m%d").date()]
    now['波動率pr']=(now.groupby(['標的代號','flag'])['隱含波動率(委買價)(%)'].rank(pct=True)*100).round(2)
    temp=now.set_index('代號')
    vol_pct_dict=temp['波動率pr'].to_dict()
    
    # #最大庫存 日期,參數
    # finnal_output['最大OI'],_=make_table("自營商庫存")
    
    # #####目前最大時間價值 參數#####
    # #目前最大庫存 參數
    # finnal_output['目前最大OI'],_=make_table("自營商庫存",target=now)   #2目前 同標的 時間價值最大權證參數
    #%% 增量觀察
    # =============================================================================
    # 1.25天內最大時間價值增量
    # 3.25天內最大OI增量
    # =============================================================================
    #最大時間價值增量 日期,參數
    finnal_output['最大時間價值增量'],_=make_table("市價時間價值增量")
    #最大庫存增量 日期,參數
    finnal_output['最大庫存增量'],_=make_table("自營商買賣超")   #25天內 同標的 時間價值最大權證參數
    
    return finnal_output
    # # #%%excel
    # i=0
    # for index,values in finnal_output.items():
    #     if i==0:
    #         values.to_excel('test.xlsx',sheet_name=index)
    #     else:
    #         writer = pd.ExcelWriter('test.xlsx', engine = 'openpyxl',mode='a')
    #         values.to_excel(writer, sheet_name=index)
    #         writer.save()
    #     i+=1