#test211
# import pkgutil
# Share1 = pkgutil.extend_path(r"C:\Users\ITProg02\AppData\Local\anaconda3\Lib\site-packages, "Share")
# import sys
# sys.path.insert(0, r"C:\Users\ITProg02\AppData\Local\anaconda3\Lib\site-packages\Share\Honour_Share")
# sys.path.append(r"C:\Users\ITProg02\AppData\Local\anaconda3\Lib\site-packages\Share\Honour_Share")
# from Share.Honour_Share import update_ver
import numpy as np
from Share import Honour_Share
#################实现自动先把ui 文件转为python 文件#################开始
import qt_ui_to_py
import ui_main_inputbs

qt_ui_to_py.runMain()
#################实现自动先把ui 文件转为python 文件#################结束
# from PyQt5.QtGui import *
from PyQt5.QtGui import QStandardItemModel,QStandardItem
from PyQt5.QtWidgets import QMainWindow, QApplication, QWidget, QFileDialog,QAbstractItemView,QMessageBox
import os
import math

import psutil

import shutil
import datetime

import win32com.client
import configparser
import pandas as pd
# import win32api
# import win32con
# import sqlite3
from sqlite3 import connect
from collections import Counter

from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Table,PageBreak,Paragraph,Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

def df_to_pdf(df, PDFile):

    # Create a PDF document
    pdf_filename = PDFile
    doc = SimpleDocTemplate(pdf_filename)

    # Load the "simhei" font
    pdfmetrics.registerFont(TTFont('simhei', r'C:\Windows\Fonts\SimSun.ttc'))

    # Convert the DataFrame to a list of lists (table data)
    elements = []
    styles = getSampleStyleSheet()
    styleN = ParagraphStyle(
        name='CellStyle',
        parent=styles['Normal'],
        fontName='simhei',
        fontSize=12,  # Set the desired font size here for the cells
        leading=40,
        leftIndent=20
    )
    # table_data = [df.columns] + df.values.tolist()
    bw1 = df.groupby(["工单号","货号"])
    for group in bw1.groups:
        gp = bw1.get_group(group)

        table_data = [gp.columns] + gp.values.tolist()
        # Create a Table object and set its properties
        table = Table(table_data)
        table.setStyle([
            ('TEXTCOLOR', (0, 0), (-1, 0), (0, 0, 1)),  # Header row text color
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),  # Center alignment
            ('FONTNAME', (0, 0), (-1, -1), 'simhei'),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),  # Header padding
            ('BACKGROUND', (0, 0), (-1, 0), (0.9, 0.9, 0.9)),  # Header background color
            ('GRID', (0, 0), (-1, -1), 1, colors.black),  # 設置網格線
        ])
        elements.append(table)
        elements.append(Spacer(1, 12))
        total_mark=Paragraph(group[0]+" "+group[1]+"補數合計:"+str(int(gp['數量'].sum()))+" PCS.",style=styleN)
        elements.append(total_mark)
        elements.append(PageBreak())


    # Build the PDF document
    doc.build(elements)

    print(f"PDF file '{pdf_filename}' has been created with the DataFrame.")

def Email_lotus(file,dfpdffile,config):
    global key,df_all,joblist
    s = win32com.client.Dispatch('Notes.NotesSession')
    db_ddata=s.GetDatabase(config[1], r"PublicNSF\ddata.nsf")  #server , NSF path
    doc_ddata=db_ddata.GetDocumentByUNID("823BE41DAED99F2A48258810002FC9B5")
    st=config[5]+'To'
    ct = config[5] + 'CC'
    SendTo=doc_ddata.getitemvalue(st)
    CopyTo=doc_ddata.getitemvalue(ct)
    print("sendto ",SendTo)
    print("CopyTo ", CopyTo)
    db = s.GetDatabase(config[1], config[6])  #server , NSF path
    doc = db.CreateDocument
    doc.form = "Memo"

    # body1=doc.CreateRichTextItem("body")

    # tabs=np.array([["1","2"],["3","4"]])
    # tabs=[["1","2"],["3","4"]]
    # style=s.CreateRichTextParagraphStyle
    # style.LeftMargin =0.1
    # style.FirstLineLeftMargin =0.1
    # style.RightMargin = 10

    # style[2].LeftMargin = 1
    # style[2].FirstLineLeftMargin = 1
    # style[2].RightMargin = 2
    body1 = doc.CreateRichTextItem("body")
    richStyle = s.CreateRichTextStyle
    richStyle.NotesFont = 1
    richStyle.FontSize = 12
    # richStyle.NotesFont = body1.GetNotesFont("Courier",True)
    body1.AppendStyle(richStyle)

    attachment = doc.CreateRichTextItem("Attachment")
    attachment.EmbedObject(1454, "", file, "補數檔案")
    if config[3] == "Yes":
        attachment.EmbedObject(1454, "", dfpdffile, "補數檔案")
    doc.SendTo=SendTo
    doc.CopyTo = CopyTo


    datetime1=datetime.datetime.strftime(datetime.datetime.now(), "%Y-%m-%d %H:%M")
    jl=""
    for j in joblist:
        jl=jl+j+";"
    subject=config[5] + ' '+datetime1+" 補數數據. 工單:"+jl
    doc.Subject=subject

    body=""


    for k in key:
        body=body+k+"\n"

    body1.Appendtext(body)
    # head1="工单版本".ljust(8) + "貨號".ljust(10) + "識別碼".ljust(7) + "補數原因".ljust(18) + "編號".ljust(13)
    # head2 = "JobVer".ljust(12) + "Item".ljust(12) + "Code".ljust(10) + "Reason".ljust(22) + "seq".ljust(20)
    head1="工单版本\t貨號\t\t\t識別碼\t補數原因\t\t\t編號"
    head2 = "JobVer\tItem\t\t\tCode\t\tReason\t\t\tseq"
    body1.Appendtext(head1)
    body1.AddNewLine(1)
    body1.Appendtext(head2)
    body1.AddNewLine(1)

    for index, row in df_all.iterrows():
        if  pd.isna(row.iloc[5]):
            if type(row.iloc[4])==str:
                seq=row.iloc[4]
            else:
                seq=str(int(row.iloc[4]))
        else:
            seq=str(int(row.iloc[4]))+"-"+str(int(row.iloc[5]))
        # rows=str(row.iloc[0]).ljust(12)+str(row.iloc[1]).ljust(12)+str(row.iloc[3]).ljust(10)+str(row.iloc[9]).ljust(20)+seq
        rows = str(row.iloc[0])+"\t" + str(row.iloc[1])+"\t\t" + str(row.iloc[3])+"\t\t" + str(row.iloc[9])+"\t\t\t\t" + seq
        body1.Appendtext(rows)
        body1.AddNewLine(1)

    # doc.body=body

    # doc.save(True,False)
    doc.Send(False)

def Email_lotus_bak20250110(file,config):
    global key

    s = win32com.client.Dispatch('Notes.NotesSession')
    db_ddata=s.GetDatabase(config[1], r"PublicNSF\ddata.nsf")  #server , NSF path
    doc_ddata=db_ddata.GetDocumentByUNID("823BE41DAED99F2A48258810002FC9B5")
    st=config[5]+'To'
    ct = config[5] + 'CC'
    SendTo=doc_ddata.getitemvalue(st)
    CopyTo=doc_ddata.getitemvalue(ct)
    print("sendto ",SendTo)
    print("CopyTo ", CopyTo)
    db = s.GetDatabase(config[1], config[6])  #server , NSF path
    doc = db.CreateDocument
    doc.form = "Memo"

    attachment = doc.CreateRichTextItem("Attachment")

    attachment.EmbedObject(1454, "", file, "補數檔案")
    doc.SendTo=SendTo
    doc.CopyTo = CopyTo


    datetime1=datetime.datetime.strftime(datetime.datetime.now(), "%Y-%m-%d %H:%M")
    subject=config[5] + ' '+datetime1+" 補數數據."
    doc.Subject=subject
    body=""
    for k in key:
        body=body+k+"\n"
    doc.body=body

    # doc.save(True,False)
    doc.Send(False)

def Write_lotus(config,strdatetime,cnt,layout):
    def create_doc():
        doc = db.CreateDocument
        doc.form = "BsSeq"
        BsKey = dfv[0] + '#' + dfv[1] + '#' + config[5] + "-" + strdatetime
        ofname = config[5] + "-" + strdatetime
        doc.BsKey = BsKey
        doc.ofname = ofname
        doc.dp1 = config[5]
        doc.name1 = dfv[8]
        doc.Jobver = dfv[0]
        doc.item_code = dfv[1]
        doc.id_code = dfv[3]
        doc.remark = dfv[12]
        doc.groupseq = groupseq
        doc.num1 = seq_list
        doc.qty = len(seq_list)
        doc.bsreason = dfv[9]

        # doc.ComputeWithForm(True,True)
        doc.save(True, False)
        key.add(BsKey)  # 集合去重復

    global key,df_all
    print('all',df_all)
    s = win32com.client.Dispatch('Notes.NotesSession')
    db = s.GetDatabase(config[1], config[2])  #server , NSF path
    df_sort=df_all.sort_values(by=['工单号','货号'],ascending=True)
    print(df_sort)
    group1 = df_sort.values[0][0] + '#' + df_sort.values[0][1]
    print('first ',group1)

    # strdatetime="20241022134636"
    for dfv in df_sort.values:
        if str(dfv[4]).find(';')>=0:
            seq_list=str(dfv[4]).split(';')
            print("1",dfv[4],dfv[5])
        elif pd.isna(dfv[5]) or pd.isnull(dfv[5]):
            seq_list = [str(dfv[4])]
            print("2",dfv[4],dfv[5])
        else:
            seq_list=[str(seq) for seq in range(int(dfv[4]), int(dfv[5]) + 1)]
            print("3",dfv[4],dfv[5])

        ####處理按組的補數編號, 例如便條簿start
        if config[3]=='Yes' and cnt>0 and layout>0:

            for sl in seq_list:
                seq_list_group = []
                for lay in range(layout):
                    seq_list_group.append(int(sl)+lay*cnt)
                    seq_list=int(sl) + lay * cnt
                seq_list=seq_list_group
                groupseq=sl
                print("Group list seq:",seq_list, groupseq)
                create_doc()
        else:
            groupseq = 0
            if seq_list:
                create_doc()


        ####處理按組的補數編號, 例如便條簿end


        # if pd.isna(dfv[6]) or pd.isnull(dfv[6]):
        #     groupseq=0
        # else:
        #     groupseq = float(dfv[6])



        # if seq_list:
        #     create_doc()
            # doc = db.CreateDocument
            # doc.form = "BsSeq"
            # BsKey = dfv[0] + '#' + dfv[1] + '#' + config[5] + "-" + strdatetime
            # ofname=config[5] + "-" + strdatetime
            # doc.BsKey = BsKey
            # doc.ofname=ofname
            # doc.dp1=config[5]
            # doc.name1=dfv[8]
            # doc.Jobver=dfv[0]
            # doc.item_code = dfv[1]
            # doc.id_code = dfv[3]
            # doc.remark = dfv[12]
            # doc.groupseq=groupseq
            # doc.num1=seq_list
            # doc.qty=len(seq_list)
            # doc.bsreason=dfv[9]
            #
            # # doc.ComputeWithForm(True,True)
            # doc.save(True,False)
            #
            # key.add(BsKey)      #集合去重復





def read_data():
    # os.environ.get("USERNAME")    #pc2002
    s = win32com.client.Dispatch('Notes.NotesSession')
    # print("server:" )
    # hostname = os.environ['COMPUTERNAME']
    hostname = s.CommonUserName     #改了用Lotus 用戶, 不用電腦編號.
    bsdata = connect("./Source/Lite_DB.db")
    cursor = bsdata.cursor()
    cursor.execute("SELECT * FROM `config` where PC_number=?", (hostname,))
    Check_FileName = cursor.fetchone()
    cursor.close()
    bsdata.close()
    return Check_FileName

def read_excel_bak(filelist,config):
    global df_all
    strdatetime=datetime.datetime.strftime(datetime.datetime.now(),"%Y%m%d%H%M%S")
    for num,i in enumerate(filelist):

        df = pd.read_excel(config[4]+'\\'+i)

        if num==0:
            df_all = pd.DataFrame(columns=df.columns)
        df_all=pd.concat([df_all,df])
        # df_all=df_all.append(df,ignore_index=True)
        Write_lotus(df,config,strdatetime)
        # for dfv in df.values:
        #     print(dfv)




class mainwindow(QMainWindow):
    """
    主程序窗口
    """

    def __init__(self, window):
        super().__init__()

        self.ui = ui_main_inputbs.Ui_MainWindow()
        self.ui.setupUi(self)
        self.load_data()
        self.show()
        self.ui_event()
    def close(self):
        config.set("DEFAULT","App_PID","")
        with open('./App.ini', 'w') as configfile:
            config.write(configfile)
    def ui_event(self):
        """
        主窗口事件
        :return:
        """

        self.ui.pushButton_scan.clicked.connect(self.ScanFile)  # Scan
        self.ui.pushButton_process.clicked.connect(self.Process)  # 處理




    def load_data(self):
        config = read_data()
        self.config=config
        print(config)
        self.ui.pushButton_process.setEnabled(False)
        # dir=os.path.dirname(self.config[4])+"\\Finish"
        if config[3]=="Yes":
            self.ui.lineEdit.setVisible(True)
            self.ui.lineEdit_2.setVisible(True)
            self.ui.label_2.setVisible(True)
            self.ui.label_3.setVisible(True)
        else:
            self.ui.lineEdit.setVisible(False)
            self.ui.lineEdit_2.setVisible(False)
            self.ui.label_2.setVisible(False)
            self.ui.label_3.setVisible(False)

        if not os.path.isdir(self.config[4]):
            os.makedirs(self.config[4])
        if not os.path.isdir("Finish"):
            os.makedirs("Finish")

    def Process(self):
        global df_all,joblist
        self.ui.pushButton_process.setEnabled(False)
        if len(self.ui.lineEdit.text())>0:
            cnt = int(self.ui.lineEdit.text())
        else:
            cnt=0
        if len(self.ui.lineEdit_2.text())>0:
            layout = int(self.ui.lineEdit_2.text())
        else:
            layout = 0
        strdatetime = datetime.datetime.strftime(datetime.datetime.now(), "%Y%m%d%H%M%S")
        Write_lotus(self.config, strdatetime,cnt,layout)

        for k, v in df_all.iterrows():
            # print(v['开始号码'], type(v['开始号码']), v['结束号码'], type(v['结束号码']))
            # print('一箱數量:',df_all.iloc[k][6])
            if str(v.iloc[6]).isnumeric():
                cnt_qty = int(v.iloc[6])

            else:
                cnt_qty = 1
            if type(v['开始号码']) == int and not pd.isna(v['结束号码']):
                qty = v['结束号码'] - v['开始号码'] + 1
                # df_all.loc[k, '數量'] = str(int(qty))
                df_all.loc[k, '數量'] = int(qty)
                df_all.loc[k, '~结束号码'] = str(int(v['结束号码']))
                df_all.loc[k, '箱號'] = str(math.ceil(int(v['开始号码'])/cnt_qty)+1)+"-"+str(math.ceil(int(v['结束号码'])/cnt_qty)+1)
            elif type(v['开始号码']) == int and pd.isna(v['结束号码']):
                df_all.loc[k, '數量'] = 1
                df_all.loc[k, '~结束号码'] = ""
                df_all.loc[k, '箱號'] = str(math.ceil(int(v['开始号码']) / cnt_qty) + 1)
            elif type(v['开始号码']) == str:
                qty = str(v['开始号码']).count(';') + 1
                # df_all.loc[k, '數量'] = str(int(qty))
                df_all.loc[k, '數量'] = int(qty)
                df_all.loc[k, '~结束号码'] = ""
                df_all.loc[k, '箱號'] = ""
        dfpdf = df_all[["工单号", "货号", "辅料名称", "识别码", "开始号码", "~结束号码", "箱號", "數量", "跟进人"]]
        dfpdffile=os.getcwd()+"/Finish/"+self.config[5]+'_Bs_'+strdatetime+'.PDF'
        if self.config[3] == "Yes":
            df_to_pdf(dfpdf,dfpdffile)

        file=os.getcwd()+"/Finish/"+self.config[5]+'_Bs_'+strdatetime+'.xlsx'
        df_all["key"]=df_all["工单号"]+'#'+df_all["货号"]+'#'+self.config[5]+'-'+strdatetime
        df_all.to_excel(file, index=False)
        Email_lotus(file,dfpdffile, self.config)
        for f in self.files:
            src=self.config[4]+'\\'+f
            dst=os.getcwd()+"/Finish/"+f
            print("move:",src, dst)
            shutil.move(src, dst)

        self.ui.lineEdit.setText("0")
        self.ui.lineEdit_2.setText("0")
        self.ui.label.setText("狀態: 處理完成!")

    def read_excel(self, filelist, config):
        global df_all

        for num, i in enumerate(filelist):

            df = pd.read_excel(config[4] + '\\' + i)

            if num == 0:
                df_all = pd.DataFrame(columns=df.columns)
            df_all = pd.concat([df_all, df],axis=0)
            # df_all=df_all.append(df,ignore_index=True)

            # Write_lotus(df, config, strdatetime)

    def ScanFile(self):
        global df_all,key,joblist
        df_all=""
        key = set()
        self.files = [f for f in os.listdir(self.config[4]) if not f.startswith('~$') and (f.lower().endswith('xlsx') or f.lower().endswith('xls'))]
        self.ui.textEdit_joblist.setText("")
        self.ui.textEdit_Log.setText("")
        if self.files:
            self.ui.label.setText("狀態: 等待處理!")
            self.ui.pushButton_process.setEnabled(True)
            for f in self.files:
                self.ui.textEdit_joblist.append(f)
            # print(files)
            self.read_excel(self.files, self.config)

            model_view2 = QStandardItemModel()
            model_view2.setHorizontalHeaderLabels(list(df_all.columns.values))
            jobver=set()
            joblist = set()
            for row,dfv in enumerate(df_all.values):
                jobver.add(dfv[0])
                joblist.add(dfv[0]+" "+dfv[1])
                for col, v in enumerate(dfv):
                    if pd.isna(v):
                        tv=""
                    elif type(v)==float and col==6:   #col=6 是整組補數次序欄位
                        print('第幾列',col, '內容',v)
                        tv=str(v)
                    elif type(v)==float and col!=6:
                        tv = str(int(v))
                    else:
                        tv=str(v)
                    model_view2.setItem(row, col, QStandardItem(tv))
                if type(dfv[4])!=str and not pd.isna(dfv[5]):
                    if int(dfv[4])>int(dfv[5]):
                        self.ui.pushButton_process.setEnabled(False)
                        self.ui.textEdit_Log.append(dfv[0]+" "+dfv[1]+' 開始編號:'+str(int(dfv[4]))+' - 結束編號:'+str(int(dfv[5]))+';開始編號不能大于結束編號, 請修改.')
            print('jobver:',jobver)
            self.ui.tableView_data.setModel(model_view2)
            # self.ui.process_tableView2.resizeColumnToContents(0)
            self.ui.tableView_data.resizeColumnsToContents()
            self.ui.tableView_data.resizeRowsToContents()

            #檢查DataDB 是否有此工單 Start
            s = win32com.client.Dispatch('Notes.NotesSession')
            db = s.GetDatabase(self.config[1], r'PublicNSF\DataDBNew.nsf')  # server , NSF path
            view=db.getview("SearchHonourJobVer")
            self.jobver_qty=dict()
            self.jobver_SerialTag = dict()
            self.jobver_MainJobNum = dict()
            for jv in jobver:
                doc=view.GetDocumentByKey(jv,True)
                if doc is not None:
                    TotalMailQty=doc.getitemvalue("TotalMailQty")
                    SerialTag=doc.getitemvalue("SerialTag")
                    MainJobNum = doc.getitemvalue("MainJobNum")
                    print(jv,TotalMailQty,SerialTag)
                    self.jobver_qty[jv]=TotalMailQty[0]
                    self.jobver_SerialTag[jv] = SerialTag[0]
                    self.jobver_MainJobNum[jv]=MainJobNum[0]
                    if MainJobNum[0]!='':
                        # win32api.MessageBox(0, "工單: " + jv + " 是合并單, 不會檢測編號是否超出范圍, 請注意.", "錯誤提示!", win32con.MB_OK)
                        self.ui.textEdit_Log.append(jv + " 是合并單, 不會檢測編號是否超出范圍, 請注意.")
                else:
                    self.ui.pushButton_process.setEnabled(False)
                    # QMessageBox.information(ui_mainwindow, "錯誤提示!", "工單 "+jv+" 不在DataDB, 檢測不到數量. 請檢查!")
                    self.ui.textEdit_Log.append(jv+" 不在DataDB, 檢測不到數量. 請檢查!")


            print(self.jobver_qty)
            # 檢查DataDB 是否有此工單 Start

            #check 識別碼, 數量 start
            df_sort = df_all.sort_values(by=['工单号', '货号'], ascending=True)
            for dfv in df_sort.values:
                if str(dfv[4]).find(';') >= 0:
                    seq_list = str(dfv[4]).strip().split(';')
                    print("1", dfv[4], dfv[5])
                elif pd.isna(dfv[5]) or pd.isnull(dfv[5]):
                    seq_list = [str(dfv[4])]
                    print("2", dfv[4], dfv[5])
                else:
                    seq_list = [str(seq) for seq in range(int(dfv[4]), int(dfv[5]) + 1)]
                    print("3", dfv[4], dfv[5])

                if pd.isna(dfv[6]) or pd.isnull(dfv[6]):
                    groupseq = 0
                else:
                    groupseq = float(dfv[6])
                # print("Group:",groupseq)
                if self.jobver_qty.get(dfv[0]) is not None:
                    max_qty = self.jobver_qty.get(dfv[0])
                    SerialTag = self.jobver_SerialTag.get(dfv[0])
                    MainJobNum = self.jobver_MainJobNum.get(dfv[0])
                else:
                    max_qty = 9999999
                    SerialTag = ""
                if str(SerialTag).strip() != str(dfv[3]).strip():
                    self.ui.pushButton_process.setEnabled(False)
                    # win32api.MessageBox(0, "工單: " + dfv[0] + "識別碼: " + SerialTag + " 與Excel 補數檔: " + str(dfv[3]).strip() + " 不同, 請檢查.", "錯誤提示!", win32con.MB_OK)
                    self.ui.textEdit_Log.append(dfv[0] + " 識別碼: " + SerialTag + " 與Excel 補數檔: " + str(dfv[3]).strip() + " 不同, 請檢查.")
                seq_list=[x for x in seq_list if x!=""]
                for num, sl in enumerate(seq_list):
                    # if sl=="":
                    print("test:",sl)
                    if int(sl) > max_qty:
                        if MainJobNum=="":
                            del seq_list[num]
                            self.ui.pushButton_process.setEnabled(False)
                            # win32api.MessageBox(0, "工單: " + dfv[0] + "補數編號: " + sl + " 大于工單數: " + str(int(max_qty)) + ", 請檢查.", "錯誤提示!", win32con.MB_OK)
                            self.ui.textEdit_Log.append(dfv[0] + " 補數編號: " + sl + " 大于工單數: " + str(int(max_qty)) + ", 請檢查.")

            # check 識別碼, 數量
            bw1=df_all.groupby(['工单号', '货号'])
            for group in bw1.groups:
                print("gp:",group)
                gp = bw1.get_group(group)
                # component_qty = str(gp.iloc[0]["component_qty_y"])
                seq_list2=[]
                for index, row in gp.iterrows():
                    seq1=row["开始号码"]
                    seq2 = row["结束号码"]
                    print("seq1:",seq1, type(seq1),"seq2:",seq2, type(seq2))

                    if pd.isna(seq2):
                        if type(seq1)==str:
                            seq_list = seq1.strip().split(';')
                        else:
                            seq_list = [str(seq1)]
                    else:
                        seq_list = [str(seq) for seq in range(int(seq1), int(seq2) + 1)]
                    for sl_1 in seq_list:
                        seq_list2.append(sl_1)
                count=Counter(seq_list2)
                for k,v in count.items():
                    if v>1:
                        self.ui.pushButton_process.setEnabled(False)
                        self.ui.textEdit_Log.append(group[0]+" "+group[1] + " 補數編號有重復, 重復編號: " + str(k) + ", 重復次數: " +str(v) + ", 請檢查.")
                        print("有重復: ",k,v)


        else:
            self.ui.label.setText("狀態: 沒有補數檔案!")

# def kill_process(name):
#     for proc in psutil.process_iter():
#         if proc.name() == name:
#             proc.kill()


if __name__ == '__main__':
    Honour_Share.kill_process('PyApp_')
    # Program_Name=os.path.splitext(os.path.basename(__file__))[0]
    print('path:',os.getcwd())
    app = QApplication([])
    Honour_Share.update_ver("./ver_inputbs.txt")
    ui_mainwindow = mainwindow(QMainWindow())
    key = set()
    df_all = ""
    config=""
    # read_excel(['1926502欠数表.xlsx'], config)
    # print('key', key)
    # df_all.to_excel('out.xlsx', index=False)
    # Email_lotus(df_all, config)
    #
    app.exec_()