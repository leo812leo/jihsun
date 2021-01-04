# -*- coding: utf-8 -*-

import requests
import json
import pandas as pd

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
        qryresult=json.loads(response.text)
        errcode=qryresult.get("Error")        
        if (errcode==""):
            data=qryresult.get('ResultValue')
            data=data.split("\r\n")
            data=data[0:len(data)-1]
            for i in range(0,len(data)):
                data[i]=data[i].split("^")
            return pd.DataFrame(data[1:],columns=data[0])
 
        else:
            print("查詢產生錯誤，錯誤訊息為:")
            print(errcode)
            return ""

    
    def Mul_Data(self,tbname,ptype,date,colist = "",isst="Y",ps=""):
        chk_result=self._check_perid(ptype,date)    
        if(len(chk_result)!=2):
            print(chk_result)
        else:
            if(colist==""):
                col="*"
            else:
                col=colist
            sql="select "+col+" from "+tbname+" where "+chk_result+" = '"+date+"'"
            if(ps!=""):
                sql=sql+" and "+self._check_isst(isst)+" in "+ps
            return self._return_qry(sql,tbname)                

    def Sil_Data(self,tbname,ptype,sid,begin,end,colist = "",isst="Y"):
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
            sql="select "+col+" from "+tbname+" where ("+chk_result1+" between '"+begin+"' and '"+end+"') and "+self._check_isst(isst)+"='"+sid+"'" 
            return self._return_qry(sql,tbname)                

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

    def Sql_data(self,sqlcmds,sqltables):
        return self._return_qry(sqlcmds,sqltables)                
        
        
        
        
