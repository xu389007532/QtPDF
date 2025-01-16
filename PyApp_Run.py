"""
自動運行并更新程序
"""
import os
from PyQt5.QtWidgets import QApplication, QMainWindow,QComboBox,QVBoxLayout,QPushButton,QWidget,QLabel,QMessageBox
from PyQt5.QtGui import QFont
from PyQt5.QtCore import QRect
import win32com.client
import win32api
from time import sleep
import subprocess
from os.path import exists
from os import mkdir
# from os.path import *
import sys

import shutil
import configparser
from Share.Honour_Share import Check_lotus_AppStore,get_lotus_AppStore,get_Lotus_server

def _getCompanyNameAndProductName(file_path):
    """
    Read all properties of the given file return them as a dictionary.
    """
    propNames = ('Comments', 'InternalName', 'ProductName',
                 'CompanyName', 'LegalCopyright', 'ProductVersion',
                 'FileDescription', 'LegalTrademarks', 'PrivateBuild',
                  'FileVersion','OriginalFilename', 'SpecialBuild')

    props = {'FixedFileInfo': None, 'StringFileInfo': None, 'FileVersion': None}


    # backslash as parm returns dictionary of numeric info corresponding to VS_FIXEDFILEINFO struc
    fixedInfo = win32api.GetFileVersionInfo(file_path, '\\')


    props['FixedFileInfo'] = fixedInfo
    props['FileVersion'] = "%d.%d.%d.%d" % (fixedInfo['FileVersionMS'] / 65536,
                                            fixedInfo['FileVersionMS'] % 65536, fixedInfo['FileVersionLS'] / 65536,
                                            fixedInfo['FileVersionLS'] % 65536)

    # \VarFileInfo\Translation returns list of available (language, codepage)
    # pairs that can be used to retreive string info. We are using only the first pair.
    lang, codepage = win32api.GetFileVersionInfo(file_path, '\\VarFileInfo\\Translation')[0]

    # any other must be of the form \StringfileInfo\%04X%04X\parm_name, middle
    # two are language/codepage pair returned from above

    strInfo = {}
    for propName in propNames:
        # strInfoPath = u'\\StringFileInfo\\%04X%04X\\%s' % (lang, codepage, propName)
        strInfoPath = u'\\StringFileInfo\\%04X%04X\\%s' % (lang, codepage, propName)

        # print("strInfoPath:",strInfoPath)
        ## print str_info
        strInfo[propName] = win32api.GetFileVersionInfo(file_path, strInfoPath)

    props['StringFileInfo'] = strInfo
    return props

class MyApp(QMainWindow):
    def __init__(self):
        super().__init__()
        # self.setWindowTitle("App")
        # self.showMinimized()  # 最小化窗口
        self.load()
        if not self.App_Name:
            self.initUI()


    def initUI(self):
        # 设置窗口标题和大小
        self.setWindowTitle('App')
        self.setGeometry(5, 35, 350, 140)
        # 创建一个中心窗口部件和布局
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        self.text1=QLabel("選擇要啟動的應用程序及服務器.")
        self.text1.setStyleSheet("color: rgb(255, 0, 255); font: 75 12pt \"Calibri\"")
        # 创建一个QComboBox
        self.combo_box = QComboBox()
        self.combo_box2 = QComboBox()
        self.combo_box.setStyleSheet("color: rgb(255, 0, 255); font: 75 16pt \"Calibri\"")
        self.combo_box2.setStyleSheet("color: rgb(255, 0, 255); font: 75 16pt \"Calibri\"")
        # 向QComboBox中添加项目
        self.combo_box2.addItem('HMP03/IT/HMP')
        self.combo_box2.addItem('DG-ShengYi01/SHENGYI')
        self.combo_box2.addItem('YaoHui01/IT/YAOHUI')
        self.combo_box2.addItem('Honour01/HONOUR')
        self.combo_box2.setCurrentText(self.Lotus_server)
        for k,v in self.program_name.items():
            # print(c)

            self.combo_box.addItem(v,k)
        self.button = QPushButton('確定', self)
        self.button.setStyleSheet("color: rgb(0, 170, 0); font: 75 16pt \"Calibri\"")
        # 将QComboBox添加到布局中
        layout.addWidget(self.text1)
        layout.addWidget(self.combo_box)
        layout.addWidget(self.combo_box2)
        layout.addWidget(self.button)
        self.button.clicked.connect(self.click_buttom)
        """
        # 创建一个垂直布局
        layout = QVBoxLayout()
        # 创建一个按钮并添加到布局中

        self.comboBox = QComboBox()


        for c in self.program_name:
            self.comboBox.addItem(c)
        # self.button = QPushButton('確定', self)
        #
        # self.button.setStyleSheet("color: red;")
        # self.button.setFixedSize(350,50)
        # font = QFont()
        # font.setFamily("Microsoft Sans Serif")
        # font.setPointSize(12)
        # font.setBold(False)
        # font.setItalic(False)
        # font.setWeight(50)
        # self.button.setFont(font)

        # self.button.clicked.connect(self.run_production)
        layout.addWidget(self.comboBox)
        # layout.addWidget(self.button)

        # 设置窗口的布局.
        self.setLayout(layout)
        """
    def click_buttom(self):
        module_text=self.combo_box.currentData()        #self.combo_box.currentText()
        config = configparser.ConfigParser()
        config.read("./App.ini", "utf-8-sig")  # utf-8-sig  & UTF-8
        config.set('DEFAULT',"App_Name",module_text)
        config.set('DEFAULT', "lotus_server", self.combo_box2.currentText())
        config.set('DEFAULT', "Run_file", module_text+".EXE")
        with open('App.ini', 'w', encoding="utf-8-sig") as configfile:
            config.write(configfile)
        self.showMinimized()
        self.load()

    def load(self):
        if not exists("./Source"):
            mkdir('./Source')
        if not exists("App.ini"):
            s = win32com.client.Dispatch('Notes.NotesSession')
            Lotus_server = get_Lotus_server(s.UserName)
            # if str(s.CommonUserName).startswith("SY."):
            #     Lotus_server="DG-ShengYi01/SHENGYI"
            # else:
            #     Lotus_server = "HMP03/IT/HMP"
            config = configparser.ConfigParser()
            config['DEFAULT'] = {
                'App_Name': '',
                'Lotus_server': Lotus_server,
                'Run_file': '',
                'Update': 'Yes'
            }
            with open('App.ini', 'w') as configfile:
                config.write(configfile)


        config = configparser.ConfigParser()
        update_info=config.read("./App.ini", "utf-8-sig")  # utf-8-sig  & UTF-8
        # if not update_info:
        #     print("App.ini 不存在! 请检查!")
        #     sleep(5)
        #     sys.exit(0)
        self.App_Name = config['DEFAULT']['App_Name']
        self.Lotus_server=config['DEFAULT']['Lotus_server']
        self.Update=config['DEFAULT']['Update']
        Run_file = "./Source/"+config['DEFAULT']['Run_file']
        FileVersion=""
        if self.App_Name:
            if exists(Run_file):
                version_info = _getCompanyNameAndProductName(Run_file)
                FileVersion=version_info.get("FileVersion","")

            self.CheckDBOpen=Check_lotus_AppStore(self.App_Name,FileVersion,self.Update)

            if exists(Run_file) and self.CheckDBOpen:

                print("程序開始運行...")
                startupinfo = subprocess.STARTUPINFO()
                startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                startupinfo.wShowWindow = subprocess.SW_HIDE  # 隐藏窗口，如果你想看到它则设置为 SW_SHOW, SW_HIDE
                process = subprocess.Popen([Run_file], startupinfo=startupinfo)
                # print(process)
            else:
                QMessageBox.information(self, "錯誤提示!", "Lotus 打不開. 請檢查服務器是否正確或網絡是否可以!")

        else:
            self.program_name,self.CheckDBOpen=get_lotus_AppStore()




            # sys.exit(0)
            # os.startfile(os.path.abspath(Run_file))
            # print(os.path.abspath(Run_file))
            # sys.exit(0)



if __name__ == '__main__':
    # print("file name: ",__file__)
    app = QApplication(sys.argv)
    window = MyApp()
    if window.App_Name:
        window.close()
    else:
        window.show()
    # app.quit()
    # 注意：对于PyQt，通常不需要显式调用window.show()，因为showMinimized()已经处理了显示逻辑。
    sys.exit(app.exec_())
