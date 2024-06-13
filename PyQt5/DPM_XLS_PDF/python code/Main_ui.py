import typing
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QWidget ,QLineEdit
from PyQt5.QtCore import QThread, pyqtSignal
import time
import pandas as pd
from sqlalchemy import create_engine
import pyodbc 
import sys
import pdfkit
import pymssql
import os
import codecs
import math
import win32com.client as win32
from win32com.client import gencache
import win32timezone

class Ui_Form(object):
    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(480, 513)
        self.textEdit = QtWidgets.QTextEdit(Form)
        self.textEdit.setGeometry(QtCore.QRect(20, 20, 241, 400))
        self.textEdit.setObjectName("textEdit")
        self.Select_File = QtWidgets.QPushButton(Form)
        self.Select_File.setGeometry(QtCore.QRect(280, 350, 91, 31))
        self.Select_File.setContextMenuPolicy(QtCore.Qt.DefaultContextMenu)
        self.Select_File.setObjectName("Select_File")
        self.Go = QtWidgets.QPushButton(Form)
        self.Go.setGeometry(QtCore.QRect(280, 390, 91, 31))
        self.Go.setObjectName("Go")
        self.TSB_Select = QtWidgets.QRadioButton(Form)
        self.TSB_Select.setGeometry(QtCore.QRect(280, 30, 180, 31))
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.TSB_Select.setFont(font)
        self.TSB_Select.setObjectName("TSB_Select")
        self.FBIB_Select = QtWidgets.QRadioButton(Form)
        self.FBIB_Select.setGeometry(QtCore.QRect(280, 60, 180, 31))
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.FBIB_Select.setFont(font)
        self.FBIB_Select.setObjectName("FBIB_Select")
        self.DBS_Select = QtWidgets.QRadioButton(Form)
        self.DBS_Select.setGeometry(QtCore.QRect(280, 90, 180, 31))
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.DBS_Select.setFont(font)
        self.DBS_Select.setObjectName("DBS_Select")

        self.DBS_Select1 = QtWidgets.QRadioButton(Form)
        self.DBS_Select1.setGeometry(QtCore.QRect(280, 120, 180, 31))
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.DBS_Select1.setFont(font)
        self.DBS_Select1.setObjectName("DBS_Select1")

        self.DBS_Select2 = QtWidgets.QRadioButton(Form)
        self.DBS_Select2.setGeometry(QtCore.QRect(280, 150, 180, 31))
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.DBS_Select2.setFont(font)
        self.DBS_Select2.setObjectName("DBS_Select2")

        #self.DBS_Select3 = QtWidgets.QRadioButton(Form)
        #self.DBS_Select3.setGeometry(QtCore.QRect(280, 180, 180, 31))
        #font = QtGui.QFont()
        #font.setPointSize(10)
        #font.setBold(True)
        #font.setWeight(75)
        #self.DBS_Select3.setFont(font)
        #self.DBS_Select3.setObjectName("DBS_Select3")

        
        self.PWD_Select = QtWidgets.QRadioButton(Form)
        self.PWD_Select.setGeometry(QtCore.QRect(280, 300, 105, 31))
        font = QtGui.QFont()
        font.setPointSize(10)
        font.setBold(True)
        font.setWeight(75)
        self.PWD_Select.setFont(font)
        self.PWD_Select.setObjectName("PWD_Select")

        self.PWDEdit = QtWidgets.QLineEdit(Form)
        self.PWDEdit.setGeometry(QtCore.QRect(360, 300, 80, 31))
        self.PWDEdit.setObjectName("PWDEdit")
        #self.PWDEdit.textChanged.connect(self.PWD_C)
        
        self.label = QtWidgets.QLabel(Form)
        self.label.setGeometry(QtCore.QRect(445, 480, 51, 20))
        self.label.setObjectName("label")
        self.progressBar = QtWidgets.QProgressBar(Form)
        self.progressBar.setGeometry(QtCore.QRect(20, 425, 251, 23))
        self.progressBar.setObjectName("progressBar")
        self.progressBar_2 = QtWidgets.QProgressBar(Form)
        self.progressBar_2.setGeometry(QtCore.QRect(20, 455, 251, 23))
        self.progressBar_2.setObjectName("progressBar_2")
        self.label_2 = QtWidgets.QLabel(Form)
        self.label_2.setGeometry(QtCore.QRect(290, 425, 111, 20))
        self.label_2.setObjectName("label_2")
        self.label_3 = QtWidgets.QLabel(Form)
        self.label_3.setGeometry(QtCore.QRect(290, 455, 111, 21))
        self.label_3.setObjectName("label_3")

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "消費明細(轉檔小工具 xlsx to pdf)"))
        self.Select_File.setText(_translate("Form", "選取檔案"))
        self.Go.setText(_translate("Form", "開始解析"))
        self.TSB_Select.setText(_translate("Form", "台 新 銀 行"))
        self.FBIB_Select.setText(_translate("Form", "遠 東 銀 行"))
        self.DBS_Select.setText(_translate("Form", "星 展 銀 行 (xlsx 直轉 pdf)"))
        self.DBS_Select1.setText(_translate("Form", "星 展 銀 行 (呆帳備查卡)"))
        self.DBS_Select2.setText(_translate("Form", "星 展 銀 行 (繳款明細)"))
        #self.DBS_Select3.setText(_translate("Form", "星 展 銀 行 (債金計算)"))
        self.PWD_Select.setText(_translate("Form", "密碼解密"))
        self.label.setText(_translate("Form", "V1.0.4"))
        self.label_2.setText(_translate("Form", "資料上傳進度"))
        self.label_3.setText(_translate("Form", "PDF產製進度"))
# 继承QThread
class Worker(QtCore.QThread):
    progress = QtCore.pyqtSignal(int)
    progress_2 = QtCore.pyqtSignal(int)
    signal = QtCore.pyqtSignal(str)
    finish = QtCore.pyqtSignal(str)

    def __init__( self, server , database , username , password ,  Flag , path 
                 , Table1 , Table2 , Table3 , Table4 , Table5 , Result1 , Result2 , Result3 , Result4 , Result5 , Data_dt, PWD_C):
        super(Worker, self).__init__()
        self.server     = server
        self.database   = database
        self.username   = username
        self.password   = password
        self.Path       = path
        self.Type       = Flag
        self.Table1     = Table1
        self.Table2     = Table2
        self.Table3     = Table3
        self.Table4     = Table4
        self.Table5     = Table5
        self.Result1    = Result1
        self.Result2    = Result2
        self.Result3    = Result3
        self.Result4    = Result4
        self.Result5    = Result5
        self.Data_dt    = Data_dt
        self.PWD_C      = PWD_C
        #print(self.PWD_C)

    def cursor_execute1(self):
        conn_str = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={self.server};DATABASE={self.database};UID={self.username};PWD={self.password}'
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
        datadt = time.strftime("%Y-%m-%d", time.localtime())
        param = ''
        cursor.execute(f"""EXEC USP_FEIB_01_INS_Bill_FA '{datadt}', '{param}'""")
        conn.commit()
        cursor.close()
        conn.close()

    def cursor_execute2(self):
        conn_str = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={self.server};DATABASE={self.database};UID={self.username};PWD={self.password}'
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
        datadt = time.strftime("%Y-%m-%d", time.localtime())
        param = ''
        cursor.execute(f"""EXEC USP_FEIB_02_INS_Bill_AIG '{datadt}', '{param}'""")
        conn.commit()
        cursor.close()
        conn.close()
        
    def cursor_execute3(self):
        conn_str = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={self.server};DATABASE={self.database};UID={self.username};PWD={self.password}'
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
        datadt = time.strftime("%Y-%m-%d", time.localtime())
        param = ''
        cursor.execute(f"""EXEC USP_TSB_01_INS_Bill '{datadt}', '{param}'""")
        conn.commit()
        cursor.close()
        conn.close()
    
    def cursor_execute4(self):
        conn_str = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={self.server};DATABASE={self.database};UID={self.username};PWD={self.password}'
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
        datadt = time.strftime("%Y-%m-%d", time.localtime())
        param = ''
        cursor.execute(f"""EXEC USP_DBS_01_INS_Bill '{datadt}', '{param}'""")
        conn.commit()
        cursor.close()
        conn.close()

    def cursor_execute5(self):
        conn_str = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={self.server};DATABASE={self.database};UID={self.username};PWD={self.password}'
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
        datadt = time.strftime("%Y-%m-%d", time.localtime())
        param = ''
        cursor.execute(f"""EXEC USP_DBS_02_APP_Review '{datadt}', '{param}'""")
        conn.commit()
        cursor.close()
        conn.close()

    def Truncate_Table1(self):
        conn_str = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={self.server};DATABASE={self.database};UID={self.username};PWD={self.password}'
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
        cursor.execute(f"""Truncate Table STG_FEIB_FA """)
        conn.commit()
        cursor.close()
        conn.close()

    def Truncate_Table2(self):
        conn_str = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={self.server};DATABASE={self.database};UID={self.username};PWD={self.password}'
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
        cursor.execute(f"""Truncate Table STG_FEIB_AIG""")
        conn.commit()
        cursor.close()
        conn.close()
        
    def Truncate_Table3(self):
        conn_str = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={self.server};DATABASE={self.database};UID={self.username};PWD={self.password}'
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
        cursor.execute(f"""Truncate Table STG_TSB_Bill""")
        conn.commit()
        cursor.close()
        conn.close()

    def Truncate_Table4(self):
        conn_str = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={self.server};DATABASE={self.database};UID={self.username};PWD={self.password}'
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
        cursor.execute(f"""Truncate Table STG_DBS_01_INS_Bill""")
        conn.commit()
        cursor.close()
        conn.close()

    def Truncate_Table5(self):
        conn_str = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={self.server};DATABASE={self.database};UID={self.username};PWD={self.password}'
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
        cursor.execute(f"""Truncate Table STG_DBS_02_APP_Review""")
        conn.commit()
        cursor.close()
        conn.close()

    def data1(self):
            
        conn = pymssql.connect(server=self.server, user=self.username, password=self.password, database = self.database)
        cursor = conn.cursor()
        script = f"""
            SELECT 
            DISTINCT [CaseID]
            FROM [UCS_ETL].[dbo].{self.Result1}	
            WHERE DataDt = '{self.Data_dt}'
            """
        cursor.execute(script)
        c_src = cursor.fetchall()
        cursor.close()
        conn.close()
        return c_src
    
    def data2(self):
            
        conn = pymssql.connect(server=self.server, user=self.username, password=self.password, database = self.database)
        cursor = conn.cursor()
        script = f"""
            SELECT 
            DISTINCT [CaseID]
            FROM [UCS_ETL].[dbo].{self.Result2}	
            WHERE DataDt = '{self.Data_dt}'
            """
        cursor.execute(script)
        c_src = cursor.fetchall()
        cursor.close()
        conn.close()
        return c_src
    
    def data3(self):
            
        conn = pymssql.connect(server=self.server, user=self.username, password=self.password, database = self.database)
        cursor = conn.cursor()
        script = f"""
            SELECT 
            DISTINCT [CaseID]
            FROM [UCS_ETL].[dbo].{self.Result3}	
            WHERE DataDt = '{self.Data_dt}'
            """
        cursor.execute(script)
        c_src = cursor.fetchall()
        cursor.close()
        conn.close()
        return c_src
    
    def data4(self):
            
        conn = pymssql.connect(server=self.server, user=self.username, password=self.password, database = self.database)
        cursor = conn.cursor()
        script = f"""
            --SELECT 
            --DISTINCT [Creditor_Reference]
            --FROM [UCS_ETL].[dbo].{self.Result4}	
            --WHERE DataDt = '{self.Data_dt}'
            SELECT 
            DISTINCT s2.casei
            FROM [UCS_ETL].[dbo].{self.Result4}	s1
			INNER JOIN skiptrace.dbo.casetb	s2 ON s1.[Creditor_Reference] = s2.AcctI_1
            INNER JOIN skiptrace.dbo.clienttb s3 on s2.clienti = s3.clienti
            WHERE DataDt = '{self.Data_dt}' and s2.statusi not in (2,3,4) and s3.masterclienti = 1000126
            """
        cursor.execute(script)
        c_src = cursor.fetchall()
        cursor.close()
        conn.close()
        return c_src
    
    def data5(self):
            
        conn = pymssql.connect(server=self.server, user=self.username, password=self.password, database = self.database)
        cursor = conn.cursor()
        script = f"""
            SELECT 
            DISTINCT [ID]
            FROM [UCS_ETL].[dbo].{self.Result5}	
            WHERE DataDt = '{self.Data_dt}'
            """
        cursor.execute(script)
        c_src = cursor.fetchall()
        cursor.close()
        conn.close()
        return c_src

    def OutPut1(self):
        conn = pymssql.connect(server=self.server, user=self.username, password=self.password, database = self.database)
        script = f"""
            SELECT 
            [歸戶帳號]          = CaseId
            ,[歸戶卡人中文姓名]  = CaseCName
            ,[帳務月份]         = Bill_Mon
            ,[繳消明細序號]     = Seq
            ,[消費卡號]         = Cos_Acct
            ,[消費日期]         = Cos_Date
            ,[清算日期]         = Exp_Date
            ,[入帳日期]         = Tran_Date
            ,[特店中文名稱]     = CN_Name
            ,[特店英文名稱]     = EN_Name
            ,[目的地金額]       = Bill
            FROM [UCS_ETL].[dbo].{self.Result1}	
            WHERE DataDt = '{self.Data_dt}' and CaseID = '{self.ID}'
            ORDER BY CaseId , Bill_Mon , Seq
            """
        df = pd.read_sql(script,conn)
        f = codecs.open('exp.html','w')
        title = "FE 繳消報表"
        a = df.to_html(index=False)
        x = str(a)
        try:
            f.write(x.replace('class="dataframe">','class="dataframe"><caption><font size="6">{}</font></caption>'.format(title)).replace('class', 'cellspacing=\"0\" class'))
        except UnicodeEncodeError as e :
            # Extracting information from the error
            encoding_used = e.encoding
            problematic_text = e.object[e.start:e.end]
            error_reason = e.reason

            f.write(x.replace(problematic_text,'?').replace('class="dataframe">','class="dataframe"><caption><font size="6">{}</font></caption>'.format(title)).replace('class', 'cellspacing=\"0\" class'))

        f.close()
    
    def OutPut2(self):
        conn = pymssql.connect(server=self.server, user=self.username, password=self.password, database = self.database)
        script = f"""
            SELECT 
            [歸戶帳號]          = CaseId
            ,[歸戶卡人中文姓名]  = Case_CName
            --,['']               =Reserved
            ,[帳務年月]         = Bill_Mon
            ,[消費卡號]         = Cos_Acct
            ,[消費日期]         = Cos_Date
            ,[入帳日期]         = Tran_Date
            ,[特店中文名稱]     = CN_Name
            ,[目的地金額]       = Bill
            ,[原始卡號]         = Org_Acct
            FROM [UCS_ETL].[dbo].{self.Result2}	
            WHERE DataDt = '{self.Data_dt}' and CaseID = '{self.ID}'
            ORDER BY CaseId , Bill_Mon 
            """
        df = pd.read_sql(script,conn)
        f = codecs.open('exp.html','w')
        title = "AIG 繳消報表"
        a = df.to_html(index=False)
        x = str(a)
        try:
            f.write(x.replace('class="dataframe">','class="dataframe"><caption><font size="6">{}</font></caption>'.format(title)).replace('class', 'cellspacing=\"0\" class'))
        except UnicodeEncodeError as e :
            # Extracting information from the error
            encoding_used = e.encoding
            problematic_text = e.object[e.start:e.end]
            error_reason = e.reason

            f.write(x.replace(problematic_text,'?').replace('class="dataframe">','class="dataframe"><caption><font size="6">{}</font></caption>'.format(title)).replace('class', 'cellspacing=\"0\" class'))

        f.close()
    
    def OutPut3(self):
        conn = pymssql.connect(server=self.server, user=self.username, password=self.password, database = self.database)
        script = f"""
            SELECT 
            [COLNAME]           = COLNAME
            ,[CUST_ID]          = CaseID
            ,[CHN_NAME]         = CHN_NAME
            ,[對內金額]          = Bill_IN
            ,[對外金額]          = Bill_OUT
            ,[MQ送查日]          = Send_Date
            ,[MQ資料日]          = Data_Date
            ,[CUST_ID1]         = CUST_ID1
            ,[保人]             = Guarator
            ,[BNK_CD]           = BNK_CD
            ,[BANK_NAME]        = BANK_NAME
            ,[查詢理由]         = Reason
            ,[備註]             = Remark
            FROM [UCS_ETL].[dbo].{self.Result3}	
            WHERE DataDt = '{self.Data_dt}' and CaseID = '{self.ID}'
            ORDER BY CaseID , Send_Date
            """
        df = pd.read_sql(script,conn)
        f = codecs.open('exp.html','w', 'big5')
        title = "台新銀行_繳消報表"
        a = df.to_html(index=False)
        x = str(a)
        try:
            f.write(x.replace('class="dataframe">','class="dataframe"><caption><font size="6">{}</font></caption>'.format(title)).replace('class', 'cellspacing=\"0\" class'))
        except UnicodeEncodeError as e :
            # Extracting information from the error
            encoding_used = e.encoding
            problematic_text = e.object[e.start:e.end]
            error_reason = e.reason

            f.write(x.replace(problematic_text,'?').replace('class="dataframe">','class="dataframe"><caption><font size="6">{}</font></caption>'.format(title)).replace('class', 'cellspacing=\"0\" class'))

        f.close()
    
    def OutPut4(self):
        conn = pymssql.connect(server=self.server, user=self.username, password=self.password, database = self.database)
        script = f"""
            SELECT 
               [Legacy_id]
            ,  [ID]
            ,  [Type]
            ,  [Created]
            ,  [Entered]
            ,  [Tendered]
            ,  [Posted]
            ,  [Amount]
            ,  [Location]
            ,  [Ext_Ref_Id]
            ,  [Addl_Ext_Ref_Id]
            ,  [Account_Number]
            ,  [Creditor_Reference]
			--,DataDt
            FROM [UCS_ETL].[dbo].{self.Result4} S1
			INNER JOIN skiptrace.dbo.casetb	s2 ON s1.[Creditor_Reference] = s2.AcctI_1
            WHERE DataDt = '{self.Data_dt}' and s2.casei = '{self.ID}'
			--WHERE DataDt = '2024-05-15' and Creditor_Reference = 'Y120050391CD'
            ORDER BY Creditor_Reference , Entered
            """
        df = pd.read_sql(script,conn)
        f = codecs.open('exp.html','w', 'big5')
        title = "星展銀行_繳款明細"
        a = df.to_html(index=False)
        x = str(a)
        try:
            f.write(x.replace('class="dataframe">','class="dataframe"><caption><font size="6">{}</font></caption>'.format(title)).replace('class', 'cellspacing=\"0\" class'))
        except UnicodeEncodeError as e :
            # Extracting information from the error
            encoding_used = e.encoding
            problematic_text = e.object[e.start:e.end]
            error_reason = e.reason

            f.write(x.replace(problematic_text,'?').replace('class="dataframe">','class="dataframe"><caption><font size="6">{}</font></caption>'.format(title)).replace('class', 'cellspacing=\"0\" class'))

        f.close()
    
    def OutPut5(self):
        conn = pymssql.connect(server=self.server, user=self.username, password=self.password, database = self.database)
        script = f"""
            SELECT 
             [申調日期] = [APP_Dt]
            ,[催員/法務] = [Legal_P]
            ,[委案日] = [Submitdate]
            ,[委案後已回收] = [Payment]
            ,[轉呆後總回收] = [Payment_TTL]
            ,[調閱呆帳備查卡] = [APP_Card]
            ,[調閱繳款明細] = [Payment_List]
            ,[ID] = [ID]
            ,[姓名] = [Name]
            ,[Creditor_Nbr] = [Creditor_Nbr]
            ,[請求金額] = [Bill]
            ,[請求本金] = [Original_Amount]
            ,[利息起算日] = [Acc_Date]
            ,[利率] = [Rate]
            ,[程序/裁判費] = [Ref_Fee]
            ,[執行費] = [Fee]
            FROM [UCS_ETL].[dbo].{self.Result4}	
            WHERE DataDt = '{self.Data_dt}' and ID = '{self.ID}'
            ORDER BY ID , Entered
            """
        df = pd.read_sql(script,conn)
        f = codecs.open('exp.html','w', 'big5')
        title = "星展銀行_債金計算"
        a = df.to_html(index=False)
        x = str(a)
        try:
            f.write(x.replace('class="dataframe">','class="dataframe"><caption><font size="6">{}</font></caption>'.format(title)).replace('class', 'cellspacing=\"0\" class'))
        except UnicodeEncodeError as e :
            # Extracting information from the error
            encoding_used = e.encoding
            problematic_text = e.object[e.start:e.end]
            error_reason = e.reason

            f.write(x.replace(problematic_text,'?').replace('class="dataframe">','class="dataframe"><caption><font size="6">{}</font></caption>'.format(title)).replace('class', 'cellspacing=\"0\" class'))

        f.close()
    
    def run(self):
        #print(self.Type)
        #try:
        connection_str = (f"""mssql+pyodbc://{self.username}:{self.password}@{self.server}/{self.database}?driver=ODBC+Driver+17+for+SQL+Server""")
        engine = create_engine(connection_str)
        Cols1,Cols2 = '',''
        self.progress.emit(0)
        self.progress_2.emit(0)  
        self.Truncate_Table1()
        self.Truncate_Table2()
        self.Truncate_Table3()
        self.Truncate_Table4()
        self.Truncate_Table5()
        print(self.PWD_C)
        if self.Type == '2' and len(self.Path) > 0:
            for C in range(len(self.Path)):
                self.progress_2.emit(0)
                xl = pd.ExcelFile(self.Path[C])
                if len(xl.sheet_names) == 1 :
                    df1 = pd.read_excel(xl, xl.sheet_names[0],header=0,dtype=str)
                    df1.fillna('', inplace=True)
                    for i in range(len(df1.columns)):
                        x = "'"+'Col'+str(i+1)+"'"
                        Cols_Name1="'"+df1.columns[i]+"'"
                        if i == 0:
                            Cols1 = Cols_Name1 +':' + x

                        else:
                            Cols1 = Cols1 +',' + Cols_Name1 +':' + x
                    Cols1 = '{'+Cols1+'}'
                    df1.rename(columns=eval(Cols1),inplace=True)
                    df1.to_sql(self.Table1, con=engine, if_exists='append', index=False)
                else:
                    df1 = pd.read_excel(xl, xl.sheet_names[0],header=0,dtype=str)
                    df1.fillna('', inplace=True)

                    for i in range(len(df1.columns)):
                        x = "'"+'Col'+str(i+1)+"'"
                        Cols_Name1="'"+df1.columns[i]+"'"
                        if i == 0:
                            Cols1 = Cols_Name1 +':' + x

                        else:
                            Cols1 = Cols1 +',' + Cols_Name1 +':' + x
                    Cols1 = '{'+Cols1+'}'
                    df1.rename(columns=eval(Cols1),inplace=True)
                    df1.to_sql(self.Table1, con=engine, if_exists='append', index=False)
                    df2 = pd.read_excel(xl, xl.sheet_names[1],header=1,dtype=str)
                    df2.fillna('', inplace=True)

                    for i in range(len(df2.columns)):
                        x = "'"+'Col'+str(i+1)+"'"
                        Cols_Name2="'"+df2.columns[i]+"'"
                        if i == 0:
                            Cols2 = Cols_Name2 +':' + x

                        else:
                            Cols2 = Cols2 +',' + Cols_Name2 +':' + x
                    Cols2 = '{'+Cols2+'}'
                    df2.rename(columns=eval(Cols2),inplace=True)
                    df2.to_sql(self.Table2, con=engine, if_exists='append', index=False)
                #self.ui.label_2.setText('已上載完成!')
                Status = int(math.ceil(100/(len(range(len(self.Path))))*(C+1)))
                self.progress.emit(Status)
                
            self.cursor_execute1()
            self.cursor_execute2()

                

            src1 = self.data1()
            for X1 in range(len(src1)):
                self.ID = src1[X1][0]
                self.OutPut1()
                
                config = pdfkit.configuration(wkhtmltopdf="C:\Program Files\wkhtmltopdf\wkhtmltopdf.exe")
                pdfkit_options = {'encoding':"big5",'page-height': '300','page-width': '250'}
                OutFile = self.Path[0].rsplit('/', maxsplit=1)[0]+'/FE'
                try:
                    os.makedirs(OutFile)
                except:
                    pass
                time.sleep(0.5)
                pdfkit.from_file('exp.html', OutFile+f'\{self.ID}.pdf',configuration=config,options= pdfkit_options)
                Status2 = int(math.ceil(100/(len(range(len(src1))))*(X1+1)))
                self.progress_2.emit(Status2)
            src2 = self.data2()
            self.progress_2.emit(0)
            for X2 in range(len(src2)):
                self.ID = src2[X2][0]
                self.OutPut2()
                
                
                config = pdfkit.configuration(wkhtmltopdf="C:\Program Files\wkhtmltopdf\wkhtmltopdf.exe")
                pdfkit_options = {'encoding':"big5",'page-height': '300','page-width': '250'}
                OutFile = self.Path[0].rsplit('/', maxsplit=1)[0]+'/AIG'
                try:
                    os.makedirs(OutFile)
                except:
                    pass
                time.sleep(0.5)
                pdfkit.from_file('exp.html', OutFile+f'\{self.ID}.pdf',configuration=config,options= pdfkit_options)
                Status2 = int(math.ceil(100/(len(range(len(src2))))*(X2+1)))
                self.progress_2.emit(Status2)

        elif self.Type == '3' and len(self.Path) > 0:
            for C in range(len(self.Path)):
                self.progress_2.emit(0)
                # in_file為要處理的excel檔案名稱加路徑
                in_file = os.path.join(self.Path[C])
                in_file = in_file.replace('/','\\')
                # 處理完的檔案，附檔名改成pdf
                out_file = os.path.join(os.path.splitext(self.Path[C])[0] + ".pdf")
                out_file = out_file.replace('/','\\')
                print(out_file)
                # excel的檔案先查看列印的頁面，是否是自己要的分頁方式，免得頁面被分割
                excel = win32.DispatchEx("Excel.Application")
                excel.Quit()
                excel.Interactive = False
                excel.Visible = False
                workbook = excel.Workbooks.Open(in_file, None, True)
                workbook.ActiveSheet.ExportAsFixedFormat(0, out_file) 
                # 移除excel檔案（非必要）
                #os.remove(in_file)
                #
                
                Status = int(math.ceil(100/(len(range(len(self.Path))))*(C+1)))
                
                self.progress.emit(Status)
                self.progress_2.emit(Status)
        elif self.Type == '4' and len(self.Path) > 0:
            for C in range(len(self.Path)):
                self.progress_2.emit(0)
                print('A')
                # in_file為要處理的excel檔案名稱加路徑
                in_file = os.path.join(self.Path[C])
                in_file = in_file.replace('/','\\')
                OutFile = self.Path[C].replace(self.Path[C].split('/')[-2],'').replace(self.Path[C].split('/')[-1],'')+'\PDF'
                try:
                    os.makedirs(OutFile)
                except:
                    pass
                # 處理完的檔案，附檔名改成pdf
                print('B')
                Filename = self.Path[C].split('/')[-1].split('_')[1] + ".pdf"
                print('B1')
                Filename = OutFile+'\\'+ Filename
                print('B2')
                print(Filename)
                # excel的檔案先查看列印的頁面，是否是自己要的分頁方式，免得頁面被分割
                #gencache.EnsureDispatch('Excel.Application')
                excel = win32.DispatchEx("Excel.Application")
                #excel.Quit()
                excel.Interactive = False
                excel.Visible = False
                try:
                    workbook = excel.Workbooks.Open(in_file, None, True)
                    sheet = workbook.Worksheets[0]  # 选择第一个工作表

                    # 导出为 PDF
                    sheet.Columns("G:M").AutoFit()
                    sheet.Columns("E").ColumnWidth = 13
                    sheet.Columns("F").ColumnWidth = 13
                    sheet.Rows(6).RowHeight = 90
                    # The text you're searching for
                    search_text = "前欠利息不計息"

                    # Loop through each cell in the specified column
                    for i in range(1, sheet.UsedRange.Rows.Count + 1):
                        for j in range(1, sheet.UsedRange.Columns.Count + 1):
                            if sheet.Cells(i, j).Value == search_text:
                                Row,Column = i,j
                    print(Row,Column)
                    # 页面设置
                    sheet.PageSetup.PaperSize = 9
                    if Row > 30 :
                        sheet.PageSetup.Zoom = 45
                    else:
                        sheet.PageSetup.Zoom = 60


                    sheet.PageSetup.PrintArea = rf"A1:Q{Row}"
                    sheet.ExportAsFixedFormat(Type=0,  # xlTypePDF
                                Filename=Filename,
                                Quality=0,  # xlQualityStandard
                                IncludeDocProperties=True,
                                IgnorePrintAreas=False,
                                OpenAfterPublish=False) 
                finally:
                    workbook.Close(SaveChanges=False)
                    excel.Quit()
                
                Status = int(math.ceil(100/(len(range(len(self.Path))))*(C+1)))
                
                self.progress.emit(Status)
                self.progress_2.emit(Status)
            
        elif self.Type == '1' and len(self.Path) > 0:
            for C in range(len(self.Path)):              
                xl = pd.ExcelFile(self.Path[C])
                if len(xl.sheet_names) == 1 :
                    df1 = pd.read_excel(xl, xl.sheet_names[0],header=0,dtype=str)
                    df1.fillna('', inplace=True)
                    for i in range(len(df1.columns)):
                        x = "'"+'Col'+str(i+1)+"'"
                        Cols_Name1="'"+df1.columns[i]+"'"
                        if i == 0:
                            Cols1 = Cols_Name1 +':' + x

                        else:
                            Cols1 = Cols1 +',' + Cols_Name1 +':' + x
                    Cols1 = '{'+Cols1+'}'
                    df1.rename(columns=eval(Cols1),inplace=True)
                    df1.to_sql(self.Table3, con=engine, if_exists='append', index=False)
                Status = int(math.ceil(100/(len(range(len(self.Path))))*(C+1)))
                self.progress.emit(Status)
            self.cursor_execute3()
            src3 = self.data3()
            for X3 in range(len(src3)):
                self.ID = src3[X3][0]
        
                try:
                    self.OutPut3()
                except UnicodeEncodeError:
                    self.OutPut3_rename()
                
                config = pdfkit.configuration(wkhtmltopdf="C:\Program Files\wkhtmltopdf\wkhtmltopdf.exe")
                pdfkit_options = {'encoding':"big5",'page-height': '300','page-width': '250'}
                OutFile = self.Path[0].rsplit('/', maxsplit=1)[0]+'/PDF'

                try:
                    os.makedirs(OutFile)
                except:
                    pass
                time.sleep(0.5)
                pdfkit.from_file('exp.html', OutFile+f'\{self.ID}.pdf',configuration=config,options= pdfkit_options)
                Status3 = int(math.ceil(100/(len(range(len(src3))))*(X3+1)))
                self.progress_2.emit(Status3)

        elif self.Type == '5' and len(self.Path) > 0:
            for C in range(len(self.Path)):              
                xl = pd.ExcelFile(self.Path[C])
                df1 = pd.read_excel(xl, xl.sheet_names[0],header=0,dtype=str)
                df1.fillna('', inplace=True)
                for i in range(len(df1.columns)):
                    x = "'"+'Col'+str(i+1)+"'"
                    Cols_Name1="'"+df1.columns[i]+"'"
                    if i == 0:
                        Cols1 = Cols_Name1 +':' + x

                    else:
                        Cols1 = Cols1 +',' + Cols_Name1 +':' + x
                Cols1 = '{'+Cols1+'}'
                
                df1.rename(columns=eval(Cols1),inplace=True)
                df1.to_sql(self.Table4, con=engine, if_exists='append', index=False)
                Status = int(math.ceil(100/(len(range(len(self.Path))))*(C+1)))
                self.progress.emit(Status)
            self.cursor_execute4()
            src4 = self.data4()
            for X3 in range(len(src4)):
                self.ID = src4[X3][0]
                
                try:
                    self.OutPut4()
                except UnicodeEncodeError:
                    self.OutPut4_rename()
                
                config = pdfkit.configuration(wkhtmltopdf="C:\Program Files\wkhtmltopdf\wkhtmltopdf.exe")
                pdfkit_options = {'encoding':"big5",'page-height': '300','page-width': '250'}
                OutFile = self.Path[0].rsplit('/', maxsplit=1)[0]+'/PDF'

                try:
                    os.makedirs(OutFile)
                except:
                    pass
                time.sleep(0.5)
                pdfkit.from_file('exp.html', OutFile+f'\{self.ID}.pdf',configuration=config,options= pdfkit_options)
                Status3 = int(math.ceil(100/(len(range(len(src4))))*(X3+1)))
                self.progress_2.emit(Status3)

        elif self.Type == '6' and len(self.Path) > 0:
            for C in range(len(self.Path)):              
                xl = pd.ExcelFile(self.Path[C])
                if len(xl.sheet_names) == 1 :
                    df1 = pd.read_excel(xl, xl.sheet_names[0],header=0,dtype=str)
                    df1.fillna('', inplace=True)
                    for i in range(len(df1.columns)):
                        x = "'"+'Col'+str(i+1)+"'"
                        Cols_Name1="'"+df1.columns[i]+"'"
                        if i == 0:
                            Cols1 = Cols_Name1 +':' + x

                        else:
                            Cols1 = Cols1 +',' + Cols_Name1 +':' + x
                    Cols1 = '{'+Cols1+'}'
                    df1.rename(columns=eval(Cols1),inplace=True)
                    df1.to_sql(self.Table5, con=engine, if_exists='append', index=False)
                Status = int(math.ceil(100/(len(range(len(self.Path))))*(C+1)))
                self.progress.emit(Status)
            self.cursor_execute5()
            src5 = self.data5()
            for X3 in range(len(src5)):
                self.ID = src5[X3][0]
        
                try:
                    self.OutPut5()
                except UnicodeEncodeError:
                    self.OutPut5_rename()
                
                config = pdfkit.configuration(wkhtmltopdf="C:\Program Files\wkhtmltopdf\wkhtmltopdf.exe")
                pdfkit_options = {'encoding':"big5",'page-height': '300','page-width': '250'}
                OutFile = self.Path[0].rsplit('/', maxsplit=1)[0]+'/PDF'

                try:
                    os.makedirs(OutFile)
                except:
                    pass
                time.sleep(0.5)
                pdfkit.from_file('exp.html', OutFile+f'\{self.ID}.pdf',configuration=config,options= pdfkit_options)
                Status3 = int(math.ceil(100/(len(range(len(src3))))*(X3+1)))
                self.progress_2.emit(Status3)
        elif self.Type == '7' and len(self.Path) > 0:
            
            for C in range(len(self.Path)):
                # pw_str为打开密码, 若无访问密码, 则设为 ''。如果要编辑的文件是需要密码的话，则修改上行代码
                xcl = win32.Dispatch("Excel.Application")
                Path_Var = str(self.Path[C].replace('/','\\'))
                wb = xcl.Workbooks.Open(Path_Var, True, True, None, self.PWD_C)
                xcl.DisplayAlerts = False
                OutFile = self.Path[C].replace(self.Path[C].split('/')[-2],'').replace(self.Path[C].split('/')[-1],'')+'解密'
            
                try:
                    os.makedirs(OutFile)
                except:
                    pass
                new_filename = OutFile+'/'+self.Path[C].split('/')[-1]

                # 保存时可设置访问密码，或者不设置密码。
                wb.SaveAs(new_filename.replace('/','\\'), None, '', '')
                xcl.Quit()
                
                Status = int(math.ceil(100/(len(range(len(self.Path))))*(C+1)))
                
                self.progress.emit(Status)
                self.progress_2.emit(Status)
        
        self.finish.emit('0')
        #except AttributeError:
        #    self.finish.emit('1')
        #except IndexError:
        #    self.finish.emit('2')
        #except :
        #    self.finish.emit('3')

            
            

class MainWindows_Controller(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_Form()
        self.ui.setupUi(self)
        self.setup_control()
        self.ETL_Control()
        self.Start_Control()
        #self.PWD()
        self.server = '10.10.0.94'
        self.database = 'UCS_ETL'
        self.username = 'CLUSER'
        self.password = 'Ucredit7607'
        self.Table1 = 'STG_FEIB_FA'
        self.Table2 = 'STG_FEIB_AIG'
        self.Table3 = 'STG_TSB_Bill'
        self.Table4 = 'STG_DBS_01_INS_Bill'
        self.Table5 = 'STG_DBS_02_APP_Review'
        self.Result1 = 'ODS_FEIB_FA'
        self.Result2 = 'ODS_FEIB_AIG'
        self.Result3 = 'ODS_TSB_Bill'
        self.Result4 = 'ODS_DBS_01_INS_Bill'
        self.Result5 = 'ODS_DBS_02_APP_Review'
        self.Data_dt = time.strftime("%Y-%m-%d")
    
    def setup_control(self):
        self.ui.Select_File.clicked.connect(self.openFiles)
        #self.ui.PWD_Select.clicked.connect(self.PWD)

    def ETL_Control(self):
        self.ui.TSB_Select.toggled.connect(self.onClicked)
        self.ui.FBIB_Select.toggled.connect(self.onClicked)
        self.ui.DBS_Select.toggled.connect(self.onClicked)
        self.ui.DBS_Select1.toggled.connect(self.onClicked)
        self.ui.DBS_Select2.toggled.connect(self.onClicked)
        #self.ui.DBS_Select3.toggled.connect(self.onClicked)
        self.ui.PWD_Select.toggled.connect(self.onClicked)
        self.ui.TSB_Select.setChecked(True)
    
    def openFiles(self):
        filePath, filterType = QtWidgets.QFileDialog.getOpenFileNames(self,'選取多個檔案','./')  # 選取多個檔案
        #print(filePath, filterType )
        self.path = filePath
        for i in self.path:
            self.ui.textEdit.append(str(i))

        return self.path
    
    #def PWD(self):
    #    self.PWD = self.ui.PWDEdit
    #    return self.PWD
    
    def Start_Control(self):
        self.ui.Go.clicked.connect(self.Main)

    def onClicked(self):
        radioBtn = self.sender()
        self.Flag = ''
        if radioBtn.isChecked():
            if radioBtn.text() == '遠 東 銀 行':
                self.Flag = '2'
            elif radioBtn.text() == '台 新 銀 行':
                self.Flag = '1'
            elif radioBtn.text() == '星 展 銀 行 (xlsx 直轉 pdf)':
                self.Flag = '3'
            elif radioBtn.text() == '星 展 銀 行 (呆帳備查卡)':
                self.Flag = '4'
            elif radioBtn.text() == '星 展 銀 行 (繳款明細)':
                self.Flag = '5'
            elif radioBtn.text() == '星 展 銀 行 (債金計算)':
                self.Flag = '6'
            elif radioBtn.text() == '密碼解密':
                self.Flag = '7'
            else:
                self.Flag = '0'
        return self.Flag
    
    def cursor_execute1(self):
        conn_str = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={self.server};DATABASE={self.database};UID={self.username};PWD={self.password}'
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
        datadt = time.strftime("%Y-%m-%d", time.localtime())
        param = ''
        cursor.execute(f"""EXEC USP_FEIB_01_INS_Bill_FA '{datadt}', '{param}'""")
        conn.commit()
        cursor.close()
        conn.close()

    def cursor_execute2(self):
        conn_str = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={self.server};DATABASE={self.database};UID={self.username};PWD={self.password}'
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
        datadt = time.strftime("%Y-%m-%d", time.localtime())
        param = ''
        cursor.execute(f"""EXEC USP_FEIB_02_INS_Bill_AIG '{datadt}', '{param}'""")
        conn.commit()
        cursor.close()
        conn.close()
    
    def cursor_execute3(self):
        conn_str = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={self.server};DATABASE={self.database};UID={self.username};PWD={self.password}'
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
        datadt = time.strftime("%Y-%m-%d", time.localtime())
        param = ''
        cursor.execute(f"""EXEC USP_TSB_01_INS_Bill '{datadt}', '{param}'""")
        conn.commit()
        cursor.close()
        conn.close()
    
    def cursor_execute4(self):
        conn_str = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={self.server};DATABASE={self.database};UID={self.username};PWD={self.password}'
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
        datadt = time.strftime("%Y-%m-%d", time.localtime())
        param = ''
        cursor.execute(f"""EXEC USP_DBS_01_INS_Bill '{datadt}', '{param}'""")
        conn.commit()
        cursor.close()
        conn.close()

    def cursor_execute5(self):
        conn_str = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={self.server};DATABASE={self.database};UID={self.username};PWD={self.password}'
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
        datadt = time.strftime("%Y-%m-%d", time.localtime())
        param = ''
        cursor.execute(f"""EXEC USP_DBS_02_APP_Review '{datadt}', '{param}'""")
        conn.commit()
        cursor.close()
        conn.close()

    def AttributeError_message(self):
        QtWidgets.QMessageBox.warning(self, "警告", "請重新選擇檔案", QtWidgets.QMessageBox.Retry)
    
    def IndexError_message(self):
        QtWidgets.QMessageBox.warning(self, "警告", "請確認檔案是否正常", QtWidgets.QMessageBox.Retry)
    
    def TypeError_message(self):
        QtWidgets.QMessageBox.warning(self, "警告", "請將檔案寄給DI人員進行排查", QtWidgets.QMessageBox.Retry)
    
    def Success_message(self):
        QtWidgets.QMessageBox.information(self, "完成", "此次作業已完成!", QtWidgets.QMessageBox.Ok)

    def message(self,info):
        if info == '0':
            QtWidgets.QMessageBox.information(self, "完成", "此次作業已完成!", QtWidgets.QMessageBox.Ok)
        elif info == '1':
            QtWidgets.QMessageBox.warning(self, "警告", "請重新選擇檔案", QtWidgets.QMessageBox.Retry)
        elif info == '2':
            QtWidgets.QMessageBox.warning(self, "警告", "請確認檔案是否正常", QtWidgets.QMessageBox.Retry)
        elif info == '3':
            QtWidgets.QMessageBox.warning(self, "警告", "請將檔案寄給DI人員進行排查", QtWidgets.QMessageBox.Retry)

    def Main(self):
        Type = self.Flag
        Path = self.path
        PWD_C = self.ui.PWDEdit.text()
 
        self.worker = Worker( self.server , self.database , self.username , self.password , Type , Path 
                             , self.Table1 , self.Table2 , self.Table3 , self.Table4 , self.Table5 ,self.Result1 , self.Result2, self.Result3 
                             , self.Result4 , self.Result5 , self.Data_dt , PWD_C)
        self.worker.progress.connect(self.ui.progressBar.setValue)
        self.worker.progress_2.connect(self.ui.progressBar_2.setValue)
        self.worker.finish.connect(self.message)
        self.worker.start()
        
        
        

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    window = MainWindows_Controller()
    window.show()
    sys.exit(app.exec_())
