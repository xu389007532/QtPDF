#################实现自动先把ui 文件转为python 文件#################开始
import PyQt5.QtCore
import win32con

import qt_ui_to_py
import ui_main_ms

qt_ui_to_py.runMain()
#################实现自动先把ui 文件转为python 文件#################结束
import fitz
import matplotlib.pyplot as plt
from PIL import Image, ImageDraw, ImageFont
from fitz.utils import getColor,getColorInfoDict
import datetime
import sys
import os
import pymssql as sql
from PyQt5.QtGui import *
from PyQt5 import QAxContainer
from pathlib import Path
import shutil
from PyQt5 import QtCore, QtWidgets, QtWebEngineWidgets
from PyQt5.QtCore import Qt,QVariant
from PyQt5.QtWidgets import QMainWindow,QApplication,QVBoxLayout,QWidget,QMenu,QMessageBox,QColorDialog,QLabel,QInputDialog
from configparser import ConfigParser
from pypdf import PdfReader

class mainwindow(QMainWindow):
    def __init__(self, Window):
        super().__init__()
        # super(mainwindow, self).__init__(parent)
        self.ui = ui_main_ms.Ui_MainWindow()
        self.ui.setupUi(self)
        self.adjust_ui()
        # self.load_data()
        # self.load_data2()

        self.check()
        self.ui_event()

    def adjust_ui(self):
        self.ui.statusbar1 = QLabel("用戶:")
        self.ui.statusbar1.setObjectName("statusbar1")
        self.ui.statusbar2 = QLabel("功能界面:")
        self.ui.statusbar2.setObjectName("statusbar2")
        self.ui.statusbar3 = QLabel("已選擇顏色:          ")
        self.ui.statusbar3.setObjectName("statusbar3")
        self.ui.statusbar4 = QLabel("旋轉角度:            ")
        self.ui.statusbar4.setObjectName("statusbar4")
        self.ui.statusbar5 = QLabel("插入字體大小:")
        self.ui.statusbar5.setObjectName("statusbar5")
        self.ui.statusbar6 = QLabel("頁面像素:")
        self.ui.statusbar6.setObjectName("statusbar6")
        self.ui.statusbar.addWidget(self.ui.statusbar1)
        self.ui.statusbar.addWidget(self.ui.statusbar2)
        self.ui.statusbar.addWidget(self.ui.statusbar3)
        self.ui.statusbar.addWidget(self.ui.statusbar4)
        self.ui.statusbar.addWidget(self.ui.statusbar5)
        self.ui.statusbar.addWidget(self.ui.statusbar6)

        # self.ui.centralwidget = QtWidgets.QWidget(self)
        # self.ui.centralwidget.setObjectName("centralwidget")

        self.ui.horizontalLayout_add1 = QtWidgets.QHBoxLayout()
        self.ui.horizontalLayout_add1.setObjectName("horizontalLayout_add1")

        # self.ui.WebBrowser = QAxContainer.QAxWidget(self.ui.frame1)  #
        self.ui.WebBrowser = QAxContainer.QAxWidget()  #
        self.ui.WebBrowser.setFocusPolicy(Qt.StrongFocus)
        self.ui.WebBrowser.setControl("{8856F961-340A-11D0-A96B-00C04FD705A2}")
        # self.ui.WebBrowser.setProperty("DisplayScrollBars", True)


        self.ui.horizontalLayout.addWidget(self.ui.WebBrowser)
        self.ui.horizontalLayout.setStretch(0, 3)
        self.ui.horizontalLayout.setStretch(1, 7)
        # self.ui.WebBrowser.adjustSize()
        # self.ui.WebBrowser.resize(1280, 900)
        # self.ui.gridLayout_4.addLayout(self.ui.horizontalLayout_add1, 1, 2, 3, 4)

    def check(self):

        folder_path=os.getcwd() + '/temp'
        if not os.path.exists(folder_path):
            os.mkdir(folder_path)
        for filename in os.listdir(folder_path):
            file_path = os.path.join(folder_path, filename)
            if os.path.isfile(file_path):
                os.remove(file_path)

        self.repeat=0
        self.color=(0,0,0)
        config = ConfigParser()
        if not os.path.isfile('./PDF_MergeAndSplit_config.ini'):
            config['UserInfo'] = {
                'User': 'Standard_user',
                'Show': 'Insert',
                'InsertFontsize':10,
                'cropbox' : '0.1, 0.1, 0.1, 0.1',
                'cropbox_insertfilename' : '30,10'
            }

        else:
            config.read('./PDF_MergeAndSplit_config.ini')

        with open('PDF_MergeAndSplit_config.ini', 'w') as configfile:
                config.write(configfile)

        self.fontsize=int(config.get('UserInfo', 'InsertFontsize'))
        self.cropbox = config.get('UserInfo', 'cropbox')
        self.cropbox_insertfilename = config.get('UserInfo', 'cropbox_insertfilename')
        self.User = config.get('UserInfo', 'User')
        if self.User=='Standard_user':
            self.ui.Standard_user.setChecked(True)
            self.ui.Accounting_user.setChecked(False)
        elif self.User =='Accounting_user':
            self.ui.Standard_user.setChecked(False)
            self.ui.Accounting_user.setChecked(True)

        if config.get('UserInfo', 'Show') == 'MergeAndSplit':
            self.ui.frame_insert.hide()
            self.ui.frame_Extract_images.hide()
            self.ui.frame_cropbox.hide()
            self.ui.frame_merge_Split.show()
            self.ui.actiona_2.setChecked(True)
            self.ui.actionb_2.setChecked(False)
            self.ui.action_Extract_images.setChecked(False)
            self.ui.action_cropbox.setChecked(False)
        elif config.get('UserInfo', 'Show') == 'Insert':
            self.ui.actionb_2.setChecked(True)
            self.ui.frame_insert.show()
            self.ui.frame_merge_Split.hide()
            self.ui.frame_Extract_images.hide()
            self.ui.frame_cropbox.hide()
            self.ui.actiona_2.setChecked(False)
            self.ui.action_Extract_images.setChecked(False)
            self.ui.action_cropbox.setChecked(False)
        elif config.get('UserInfo', 'Show') == 'Extract_images':
            self.ui.action_Extract_images.setChecked(True)
            self.ui.frame_Extract_images.show()
            self.ui.frame_merge_Split.hide()
            self.ui.frame_insert.hide()
            self.ui.frame_cropbox.hide()
            self.ui.actiona_2.setChecked(False)
            self.ui.actionb_2.setChecked(False)
            self.ui.action_cropbox.setChecked(False)
        elif config.get('UserInfo', 'Show') == 'cropbox':
            self.ui.frame_Extract_images.hide()
            self.ui.frame_merge_Split.hide()
            self.ui.frame_insert.hide()
            self.ui.frame_cropbox.show()
            self.ui.action_Extract_images.setChecked(False)
            self.ui.actiona_2.setChecked(False)
            self.ui.actionb_2.setChecked(False)
            self.ui.action_cropbox.setChecked(True)
            self.ui.lineEdit_cropbox.setText(self.cropbox)


        self.ui.statusbar1.setText("{:<30}".format("用戶: " + config.get('UserInfo', 'User')))
        self.ui.statusbar2.setText("{:<40}".format("功能界面: " + config.get('UserInfo', 'Show')))
        self.ui.statusbar3.setText("{:<40}".format("已選擇顏色: 黑色" ))
        self.ui.statusbar4.setText("{:<30}".format("旋轉角度: 0度"))
        self.ui.statusbar5.setText("{:<30}".format("插入字體大小: " + config.get('UserInfo', 'insertfontsize')))

        if self.ui.Accounting_user.isChecked():
            self.rotate = 270
            print(self.rotate.__str__()+"Accounting_user")
        elif self.ui.Standard_user.isChecked():
            self.rotate = 0
            print(self.rotate.__str__() + "Standard_user")

        try:
            sql.connect(server='10.2.81.30', user='beginer', password='@fly314', database='HMPSQL01')
            conn=True
        except:
            conn = False

        # if conn==False:
        #     QMessageBox.information(self, "警告", "只能本公司網絡使用!")
        #     sys.exit()


    def ui_event(self):
        """
        主窗口事件
        :return:
        """
        self.ui.pushButton4.clicked.connect(self.Open_File)     #按鈕打開PDF檔
        self.ui.listView.clicked.connect(self.list_item_clicked)

        self.ui.listView.setContextMenuPolicy(3)
        self.ui.listView.customContextMenuRequested[PyQt5.QtCore.QPoint].connect(self.listWidgetContext)
        self.ui.pushButton1.clicked.connect(self.fun_mergePDF)
        self.ui.pushButton2.clicked.connect(self.fun_split)
        self.ui.radioButton.clicked.connect(self.fun_control_split)
        self.ui.radioButton_2.clicked.connect(self.fun_control_split2)
        self.ui.radioButton_3.clicked.connect(self.fun_control_split3)

        self.ui.action0.triggered.connect(self.fun_rotate0)
        self.ui.action90.triggered.connect(self.fun_rotate90)
        self.ui.action180.triggered.connect(self.fun_rotate180)
        self.ui.action270.triggered.connect(self.fun_rotate270)
        self.ui.Standard_user.triggered.connect(self.fun_Standard_user)
        self.ui.Accounting_user.triggered.connect(self.fun_Accounting_user)

        self.ui.action_repeat.triggered.connect(self.fun_repeat)
        self.ui.action_show_imagesXY.triggered.connect(self.fun_show_imagesXY)

        self.ui.actiona_2.triggered.connect(self.fun_MergeAndSplit)
        self.ui.actionb_2.triggered.connect(self.fun_Insert)
        self.ui.action_Extract_images.triggered.connect(self.fun_Extract_images)
        self.ui.action_cropbox.triggered.connect(self.fun_cropbox)
        self.ui.action_PageNumber.triggered.connect(self.fun_Insert_PageNumber)
        self.ui.action_TotalPage.triggered.connect(self.fun_Insert_TotalPage)
        self.ui.action_Date.triggered.connect(self.fun_Insert_Date)
        self.ui.action_Time.triggered.connect(self.fun_Insert_Time)
        self.ui.action_FileName.triggered.connect(self.fun_Insert_FileName)
        self.ui.action_InsertImage.triggered.connect(self.fun_InsertImage)
        self.ui.action_add_background.triggered.connect(self.fun_add_background)


        self.ui.action_BLACK.triggered.connect(lambda: self.fun_select_color(self.ui.action_BLACK))
        self.ui.action_RED.triggered.connect(lambda: self.fun_select_color(self.ui.action_RED))
        self.ui.action_MAGENTA.triggered.connect(lambda: self.fun_select_color(self.ui.action_MAGENTA))
        self.ui.action_ORANGERED.triggered.connect(lambda: self.fun_select_color(self.ui.action_ORANGERED))
        self.ui.action_ORANGE.triggered.connect(lambda: self.fun_select_color(self.ui.action_ORANGE))
        self.ui.action_YELLOW.triggered.connect(lambda: self.fun_select_color(self.ui.action_YELLOW))
        self.ui.action_GREEN.triggered.connect(lambda: self.fun_select_color(self.ui.action_GREEN))
        self.ui.action_BLUE.triggered.connect(lambda: self.fun_select_color(self.ui.action_BLUE))
        self.ui.action_PURPLE.triggered.connect(lambda: self.fun_select_color(self.ui.action_PURPLE))
        self.ui.action_BROWN.triggered.connect(lambda: self.fun_select_color(self.ui.action_BROWN))
        self.ui.action_GRAY.triggered.connect(lambda: self.fun_select_color(self.ui.action_GRAY))
        self.ui.action_WHITE.triggered.connect(lambda: self.fun_select_color(self.ui.action_WHITE))

        self.ui.pushButton_insert.clicked.connect(self.fun_Insert_Process)
        self.ui.pushButton_Extract_images.clicked.connect(self.fun_Extract_images_Process)
        self.ui.pushButton_cropbox.clicked.connect(self.fun_cropbox_Process)

    def fun_add_background(self):
        filename_background, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,  # 父窗口对象
            "选择要添加背景/前景的PDF檔",  # 标题
            # r"./",  # 起始目录
            None,
            "數據類型 (*.PDF)"  # 选择类型过滤项，过滤内容在括号中
        )
        if filename_background:
            print(filename_background)
            list_filename = []
            for i in range(model.rowCount()):
                pindex = model.index(i, 0)
                filepath1 = os.path.dirname(model.data(pindex))
                filename1 = os.path.basename(model.data(pindex))
                filename2 = os.path.splitext(filename1)[0]
                save_filename = filepath1 + '/' + filename2 + "_add_background.pdf"
                list_filename.append(save_filename)
                doc1 = fitz.open(model.data(pindex))
                doc2 = fitz.open(filename_background)
                page = doc1[0]
                rect = page.rect
                # page_count1 = doc1.page_count
                page_count2 = doc2.page_count
                for page in doc1:
                    number = page.number
                    pno = divmod(number, page_count2)[1]
                    # print(pno)
                    page.show_pdf_page(rect, doc2, pno=pno, keep_proportion=True, overlay=True, rotate=self.rotate)
                doc1.save(filepath1 + '/' + filename2 + "_add_background.pdf", garbage=4, deflate=True, clean=True)
            self.add_file(list_filename)
            QMessageBox.information(self, "信息", "處理完成.")



    def fun_cropbox_Process(self):
        cropbox = self.ui.lineEdit_cropbox.text().split(',')
        list_filename=[]
        for i in range(model.rowCount()):
            pindex = model.index(i, 0)
            filepath1 = os.path.dirname(model.data(pindex))
            filename1 = os.path.basename(model.data(pindex))
            filename2 = os.path.splitext(filename1)[0]
            save_filename=filepath1 + '/' + filename2 + "_crop.pdf"
            list_filename.append(save_filename)
            doc = fitz.open(model.data(pindex))
            height=doc[0].rect.height
            width=doc[0].rect.width

            for num, page in enumerate(doc):
                print(num,"cropbox_insertfilename:",self.cropbox_insertfilename)
                if self.cropbox_insertfilename!='':
                    print("Yes",self.rotate)
                    h1=self.cropbox_insertfilename.split(',')
                    # page.insert_text((width/2, height-15), 'Job: ' + filename2, fontsize=10,fontname='helv', rotate=0)
                    rb1 = fitz.Rect(int(h1[0]), height-int(h1[1]), int(h1[2]), height-int(h1[3]))  # 水平
                    rb2 = fitz.Rect(10, 15, 400, 25)  # 水平
                    # 使用 Pillow 创建一个图片，并添加文字
                    image = Image.new('RGB', (400, 15), color=(255, 255, 255))
                    draw = ImageDraw.Draw(image)
                    font = ImageFont.truetype('arial.ttf', 12)
                    draw.text((25, 0), 'Job: '+filename2, font=font, fill=(0, 0, 0))
                    if doc.metadata['producer']=='Microsoft: Print To PDF':
                        image = image.transpose(Image.FLIP_LEFT_RIGHT)  # 镜像文字

                        # 将图片保存到文件
                        image.save('insert_text.png')
                        page.insert_image(rb1, filename="./insert_text.png", rotate=180, keep_proportion=True)
                    else:
                        image.save('insert_text.png')
                        page.insert_image(rb2, filename="./insert_text.png", rotate=0, keep_proportion=True)



                page.set_cropbox(fitz.Rect(float(cropbox[0])*72/2.54, float(cropbox[1])*72/2.54, width-(float(cropbox[2])*72/2.54), height-(float(cropbox[3])*72/2.54)))

            doc.save(save_filename, garbage=0, deflate=True)
        self.add_file(list_filename)
        QMessageBox.information(self, "信息", "處理完成.")

    def fun_Extract_images_Process(self):
        # print(self.ui.listView.currentIndex().data())

        # reader = PdfReader(self.ui.listView.currentIndex().data())
        pagerange = self.ui.textEdit_Extract_PageNumber.toPlainText().split(',')
        for i in range(model.rowCount()):
            pindex = model.index(i, 0)
            filepath1 = os.path.dirname(model.data(pindex))
            filename1 = os.path.basename(model.data(pindex))
            filename2 = os.path.splitext(filename1)[0]
            reader = PdfReader(model.data(pindex))
            for num, page in enumerate(reader.pages):
                print(num)
                if (pagerange[0] != '' and len([k for k in pagerange if k == str(num+1)])>0) or pagerange[0] == '':
                    for count, image_file_object in enumerate(page.images):
                        with open(filepath1+'/'+filename2+'-page' + str(num+1) + '-' + image_file_object.name, "wb") as fp:
                            fp.write(image_file_object.data)
                            print(filename2+'-page' + str(num+1) + '-' + image_file_object.name)
        QMessageBox.information(self, "信息", "處理完成.")

    def fun_show_imagesXY(self):
        print(self.ui.listView.currentIndex().data())
        point_text=self.ui.lineEdit_Point.text().split(',')
        print(point_text.__len__())
        add_text=""
        if point_text.__len__()==4:
            add_text=','+point_text[2]+','+point_text[3]


        # 打开PDF文件
        pdf_file = self.ui.listView.currentIndex().data()
        doc = fitz.open(pdf_file)
        # 获取第一页
        page = doc[0]
        # 获取页面的像素尺寸
        pix = page.get_pixmap()
        # 将像素数据保存为图像文件
        image_file = "ImagesShow.png"
        pix.save(image_file)
        # 关闭PDF文档
        doc.close()

        window = plt.figure("圖像顯示")
        im = Image.open("ImagesShow.png")
        plt.imshow(im, cmap=plt.get_cmap("gray"))

        def on_click(event):
            if event.inaxes and event.button == 1:  # 只处理鼠标左键点击
                x, y = event.xdata, event.ydata
                label = f'X: {x:.2f}, Y: {y:.2f}'
                # plt.title("點擊坐標: "+label)
                self.ui.lineEdit_Point.setText(str(int(x))+','+str(int(y))+add_text)
                print(label)  # 这里你可以将数据保存到其他数据结构中，而不仅仅是打印出来

        plt.gcf().canvas.mpl_connect('button_press_event', on_click)
        plt.show()

    def fun_repeat(self):
        self.repeat,repeat_check=QInputDialog.getInt(None,"Information","請輸入重復PDF頁數:")
        print(self.repeat,repeat_check)

    def fun_select_color(self,color):
        self.color=getColor(color.toolTip())
        print(self.color)

        self.ui.statusbar3.setText("{:<40}".format("已選擇顏色: "+color.text()))

    def fun_Insert_Process(self):
        start=datetime.datetime.now()
        print("fun_Insert_Process start: "+ start.__str__())
        point=self.ui.lineEdit_Point.text().split(',')

        pagerange = self.ui.textEdit_PageNumber.toPlainText().split(',')

        if model.rowCount()>0:
            InertContent = self.ui.textEdit_InertContent.toPlainText()
            fontname = "helv"
            for ic in InertContent:
                if ord(ic)>256:
                    fontname="china-s"
                    break

            # doc = fitz.open()
            if InertContent == "{Image}":
                insert_Rect = fitz.Rect(int(point[0]), int(point[1]), int(point[2]), int(point[3]))
                img = open(self.image, "rb").read()
            else:
                insert_position = fitz.Point(int(point[0]), int(point[1]))

            for i in range(model.rowCount()):
                InertContent = self.ui.textEdit_InertContent.toPlainText()
                pindex=model.index(i,0)
                print(model.data(pindex))
                doc1 = fitz.open(model.data(pindex))
                filepath1 = os.path.dirname(model.data(pindex))
                filename1 = os.path.basename(model.data(pindex))
                filename2 = os.path.splitext(filename1)[0]
                if InertContent.find("{FileName}") >= 0:
                    InertContent = InertContent.replace("{FileName}", filename1)
                if InertContent.find("{Time}") >= 0:
                    InertContent = InertContent.replace("{Time}", datetime.datetime.now().time().strftime("%H:%M:%S"))
                if InertContent.find("{Date}") >= 0:
                    InertContent = InertContent.replace("{Date}", datetime.datetime.now().date().strftime("%Y-%m-%d"))
                if InertContent.find("{TotalPage}") >= 0 and self.repeat==0:
                    InertContent = InertContent.replace("{TotalPage}", str(doc1.page_count))
                # 插入重復頁面.
                # 方法1
                # doc1 = fitz.open(model.data(pindex))
                # doc1_count=doc1.page_count
                # if self.repeat>0:
                #     for time in range(self.repeat):  # 復制PDF頁
                #         for i in range(doc1_count):
                #             insertpg = i
                #             position = i * (self.repeat) + time
                #             doc1.fullcopy_page(insertpg, to=-1)

                # 方法2
                if self.repeat > 0:
                    doc1_count = doc1.page_count
                    InertContent = InertContent.replace("{TotalPage}", str(doc1_count*self.repeat))
                    print(InertContent)
                    width_total, height_total = doc1[0].mediabox.width, doc1[0].mediabox.height
                    r1 = fitz.Rect(0, 0, width_total, height_total)
                    doc = fitz.open()
                    pn=0
                    for i in range(self.repeat):
                        for k in range(doc1_count):
                            doc.new_page(-1, width=width_total, height=height_total)
                            doc[pn].show_pdf_page(r1, doc1, k, keep_proportion=1)
                            InertContent_new = InertContent.replace("{PageNumber}", str(pn + 1))
                            # if InertContent.find("{PageNumber}") >= 0:
                            #     InertContent_new = InertContent.replace("{PageNumber}", str(pn + 1))
                            # else:
                            #     InertContent_new = InertContent

                            if InertContent == "{Image}":
                                doc[pn].insert_image(insert_Rect, stream=img, rotate=self.rotate, keep_proportion=True)
                            else:
                                doc[pn].insert_text(insert_position, InertContent_new, fontname=fontname, fontsize=self.fontsize, rotate=self.rotate, color=self.color)

                            pn=pn+1

                    try:
                        with open(filepath1 + '/' + filename2 + "_insert.pdf", 'w') as f:
                            f.close()
                            doc.save(filepath1 + '/' + filename2 + "_insert.pdf", garbage=0, deflate=True)
                            # QMessageBox.information(self, "信息", "處理完成: 文件名: " + filepath1 + '/' + filename2 + "_insert.pdf")
                            # print(f.mode)
                            # do something with the file
                    except PermissionError:
                        # print(PermissionError.strerror)
                        QMessageBox.information(self, "警告", '文件已经被打开:'+filename2 + "_insert.pdf")
                        return
                        # print("文件已经被打开")

                    end = datetime.datetime.now()
                    print("fun_Insert_Process end: " + end.__str__())
                    print("fun_Insert_Process 處理時間: " , (end-start))

                # 插入重復頁面.



                else:
                    for pl in doc1:
                        if InertContent.find("{PageNumber}") >= 0:
                            InertContent_new = InertContent.replace("{PageNumber}", str(pl.number+1))
                        else:
                            InertContent_new = InertContent
                        # pagel = doc1.load_page(pl.number)
                        if (pagerange[0]!='' and len([i for i in pagerange if i==str(pl.number+1)])>0) or pagerange[0]=='':
                            if InertContent=="{Image}":
                                pl.insert_image(insert_Rect,stream=img,rotate=self.rotate,keep_proportion=True)
                            else:
                                pl.insert_text(insert_position,InertContent_new,fontname=fontname,fontsize=self.fontsize,rotate=self.rotate, color=self.color)
                                # helv
                                # china-t



                    try:
                        with open(filepath1 + '/' + filename2 + "_insert.pdf", 'w') as f:
                            f.close()
                            doc1.save(filepath1 + '/' + filename2 + "_insert.pdf", garbage=0, deflate=True)
                            # QMessageBox.information(self, "信息", "處理完成: 文件名: " + filepath1 + '/' + filename2 + "_insert.pdf")
                            # print(f.mode)
                            # do something with the file
                    except PermissionError:
                        # print(PermissionError.strerror)
                        QMessageBox.information(self, "警告", '文件已经被打开:'+filename2 + "_insert.pdf")
                        return
                        # print("文件已经被打开")
                    # doc1.save(filepath1 + '/' +filename2+ "_insert.pdf", garbage=0, deflate=True)
                QMessageBox.information(self, "信息", "處理完成: 文件名: " + filepath1 + '/' + filename2 + "_insert.pdf")
        else:
            QMessageBox.information(self, "信息", "沒有打開PDF檔!")


    def fun_selectFont(self):
        font, ok = QtWidgets.QFontDialog.getFont()

        if ok:
            print("Selected Font:", font.pointSize(), font.family(),font.bold(),font.italic(),font.styleName())
            self.fontsize=font.pointSize()


    def fun_InsertImage(self):
        self.image, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,  # 父窗口对象
            "选择要插入的圖像檔",  # 标题
            # r"./",  # 起始目录
            None,
            "*.PNG;*.JPG;JPEG;BMP"  # 选择类型过滤项，过滤内容在括号中
        )
        self.ui.textEdit_InertContent.setText("{Image}")

    def fun_Insert_FileName(self):
        if self.ui.textEdit_InertContent.toPlainText()=="{Image}":
            self.ui.textEdit_InertContent.setText("{FileName}")
        else:
            self.ui.textEdit_InertContent.insertPlainText("{FileName}")

    def fun_Insert_Time(self):
        if self.ui.textEdit_InertContent.toPlainText()=="{Image}":
            self.ui.textEdit_InertContent.setText("{Time}")
        else:
            self.ui.textEdit_InertContent.insertPlainText("{Time}")

    def fun_Insert_Date(self):
        if self.ui.textEdit_InertContent.toPlainText()=="{Image}":
            self.ui.textEdit_InertContent.setText("{Date}")
        else:
            self.ui.textEdit_InertContent.insertPlainText("{Date}")

    def fun_Insert_TotalPage(self):
        if self.ui.textEdit_InertContent.toPlainText()=="{Image}":
            self.ui.textEdit_InertContent.setText("{TotalPage}")
        else:
            self.ui.textEdit_InertContent.insertPlainText("{TotalPage}")

    def fun_Insert_PageNumber(self):
        if self.ui.textEdit_InertContent.toPlainText()=="{Image}":
            self.ui.textEdit_InertContent.setText("{PageNumber}")
        else:
            self.ui.textEdit_InertContent.insertPlainText("{PageNumber}")

    def fun_cropbox(self):
        if self.ui.action_cropbox.isChecked():
            self.ui.actiona_2.setChecked(False)
            self.ui.actionb_2.setChecked(False)
            self.ui.action_Extract_images.setChecked(False)
            self.ui.frame_cropbox.show()
            self.ui.frame_Extract_images.hide()
            self.ui.frame_insert.hide()
            self.ui.frame_merge_Split.hide()
            config = ConfigParser()
            config.read('./PDF_MergeAndSplit_config.ini')
            config.set('UserInfo', 'Show', 'cropbox')
            with open('PDF_MergeAndSplit_config.ini', 'w') as configfile:
                config.write(configfile)
            self.ui.statusbar2.setText("{:<40}".format("功能界面: " + config.get('UserInfo', 'Show')))

    def fun_Extract_images(self):
        if self.ui.action_Extract_images.isChecked():
            self.ui.actiona_2.setChecked(False)
            self.ui.actionb_2.setChecked(False)
            self.ui.action_cropbox.setChecked(False)
            self.ui.frame_Extract_images.show()
            self.ui.frame_insert.hide()
            self.ui.frame_cropbox.hide()
            self.ui.frame_merge_Split.hide()
            config = ConfigParser()
            config.read('./PDF_MergeAndSplit_config.ini')
            config.set('UserInfo', 'Show', 'Extract_images')
            with open('PDF_MergeAndSplit_config.ini', 'w') as configfile:
                config.write(configfile)
            self.ui.statusbar2.setText("{:<40}".format("功能界面: " + config.get('UserInfo', 'Show')))

    def fun_Insert(self):
        if self.ui.actionb_2.isChecked():
            self.ui.actiona_2.setChecked(False)
            self.ui.action_Extract_images.setChecked(False)
            self.ui.action_cropbox.setChecked(False)
            self.ui.frame_insert.show()
            self.ui.frame_merge_Split.hide()
            self.ui.frame_cropbox.hide()
            self.ui.frame_Extract_images.hide()
            config = ConfigParser()
            config.read('./PDF_MergeAndSplit_config.ini')
            config.set('UserInfo', 'Show', 'Insert')
            with open('PDF_MergeAndSplit_config.ini', 'w') as configfile:
                config.write(configfile)
            self.ui.statusbar2.setText("{:<40}".format("功能界面: " + config.get('UserInfo', 'Show')))

    def fun_MergeAndSplit(self):
        if self.ui.actiona_2.isChecked():
            self.ui.actionb_2.setChecked(False)
            self.ui.action_Extract_images.setChecked(False)
            self.ui.action_cropbox.setChecked(False)
            self.ui.frame_insert.hide()
            self.ui.frame_Extract_images.hide()
            self.ui.frame_cropbox.hide()
            self.ui.frame_merge_Split.show()
            config = ConfigParser()
            config.read('./PDF_MergeAndSplit_config.ini')
            config.set('UserInfo', 'Show', 'MergeAndSplit')
            with open('PDF_MergeAndSplit_config.ini', 'w') as configfile:
                config.write(configfile)
            self.ui.statusbar2.setText("{:<40}".format("功能界面: " + config.get('UserInfo', 'Show')))

    def fun_Standard_user(self):
        if self.ui.Standard_user.isChecked():
            self.ui.Accounting_user.setChecked(False)

            if self.ui.action0.isChecked():
                self.rotate = 0
            elif self.ui.action90.isChecked():
                self.rotate = 90
            elif self.ui.action180.isChecked():
                self.rotate = 180
            elif self.ui.action270.isChecked():
                self.rotate = 270

            # self.ui.status_label0.setText("已設置: 標準用戶; 旋轉角度" + str(self.rotate))
            # self.ui.status_label0.setStyleSheet(u"color: rgb(255, 0, 255);")
            config = ConfigParser()
            if not os.path.isfile('./PDF_MergeAndSplit_config.ini'):
                config['UserInfo'] = {
                    'User': 'Standard_user'
                }
            else:
                config.read('./PDF_MergeAndSplit_config.ini')
            config.set('UserInfo', 'User', 'Standard_user')
            self.ui.statusbar1.setText("{:<30}".format("用戶: " + config.get('UserInfo', 'User')))
            with open('PDF_MergeAndSplit_config.ini', 'w') as configfile:
                config.write(configfile)

    def fun_Accounting_user(self):
        if self.ui.Accounting_user.isChecked():
            self.ui.Standard_user.setChecked(False)

            if self.ui.action0.isChecked():
                self.rotate=270
                self.rotate_show = 0
            elif self.ui.action90.isChecked():
                self.rotate = 0
                self.rotate_show = 90
            elif self.ui.action180.isChecked():
                self.rotate = 90
                self.rotate_show = 180
            elif self.ui.action270.isChecked():
                self.rotate = 180
                self.rotate_show = 270

            # self.ui.status_label0.setText("已設置: 會計用戶; 旋轉角度"+str(self.rotate_show))
            # self.ui.status_label0.setStyleSheet(u"color: rgb(255, 0, 255);")
            config = ConfigParser()
            if not os.path.isfile('./PDF_MergeAndSplit_config.ini'):
                config['UserInfo'] = {
                    'User': 'Accounting_user'
                }
            else:
                config.read('./PDF_MergeAndSplit_config.ini')
            config.set('UserInfo', 'User', 'Accounting_user')
            self.ui.statusbar1.setText("{:<30}".format("用戶: " + config.get('UserInfo', 'User')))
            with open('PDF_MergeAndSplit_config.ini', 'w') as configfile:
                config.write(configfile)

    def fun_control_split(self):
        self.ui.textEdit.setEnabled(False)
        self.ui.lineEdit.setEnabled(False)
        # self.ui.status_label1.setText("  拆單個: 輸出在處理PDF的目錄, 文件名: PDF檔名+Page+頁碼.pdf")
        # self.ui.status_label1.setStyleSheet(u"color: rgb(0, 0, 255);")

    def fun_control_split2(self):
        self.ui.textEdit.setEnabled(False)
        self.ui.lineEdit.setEnabled(True)
        self.ui.lineEdit.setFocus()
        # self.ui.status_label1.setText("  按頁數拆: 輸出在處理PDF的目錄, 文件名: PDF檔名+Split+順序號.pdf")
        # self.ui.status_label1.setStyleSheet(u"color: rgb(0, 0, 255);")

    def fun_control_split3(self):
        self.ui.lineEdit.setEnabled(False)
        self.ui.textEdit.setEnabled(True)
        self.ui.textEdit.setFocus()
        # self.ui.status_label1.setText("  按頁碼抽: 輸出在處理PDF的目錄, 文件名: PDF檔名+specify_page.pdf")
        # self.ui.status_label1.setStyleSheet(u"color: rgb(0, 0, 255);")

    def fun_rotate0(self):
        if self.ui.action0.isChecked():
            self.ui.action90.setChecked(False)
            self.ui.action180.setChecked(False)
            self.ui.action270.setChecked(False)
            if self.ui.Accounting_user.isChecked():
                self.rotate=270
                self.user="會計用戶"
            elif self.ui.Standard_user.isChecked():
                self.rotate = 0
                self.user = "標準用戶"
            self.ui.statusbar4.setText("{:<30}".format("旋轉角度: 0度"))


            # self.ui.status_label0.setText("已設置: "+self.user+"; 旋轉角度 0 度")
            # self.ui.status_label0.setStyleSheet(u"color: rgb(255, 0, 255);")
    def fun_rotate90(self):
        if self.ui.action90.isChecked():
            self.ui.action0.setChecked(False)
            self.ui.action180.setChecked(False)
            self.ui.action270.setChecked(False)
            if self.ui.Accounting_user.isChecked():
                self.rotate=0
                self.user = "會計用戶"
            elif self.ui.Standard_user.isChecked():
                self.rotate = 90
                self.user = "標準用戶"
            self.ui.statusbar4.setText("{:<30}".format("旋轉角度: 90度"))
            # self.ui.status_label0.setText("已設置: "+self.user+"; 旋轉角度 90 度")
            # self.ui.status_label0.setStyleSheet(u"color: rgb(255, 0, 255);")

    def fun_rotate180(self):
        if self.ui.action180.isChecked():
            self.ui.action0.setChecked(False)
            self.ui.action90.setChecked(False)
            self.ui.action270.setChecked(False)
            if self.ui.Accounting_user.isChecked():
                self.rotate=90
                self.user = "會計用戶"
            elif self.ui.Standard_user.isChecked():
                self.rotate = 180
                self.user = "標準用戶"
            self.ui.statusbar4.setText("{:<30}".format("旋轉角度: 180度"))
            # self.ui.status_label0.setText("已設置: "+self.user+"; 旋轉角度 180 度")
            # self.ui.status_label0.setStyleSheet(u"color: rgb(255, 0, 255);")

    def fun_rotate270(self):
        if self.ui.action270.isChecked():
            self.ui.action0.setChecked(False)
            self.ui.action90.setChecked(False)
            self.ui.action180.setChecked(False)
            if self.ui.Accounting_user.isChecked():
                self.rotate=180
                self.user = "會計用戶"
            elif self.ui.Standard_user.isChecked():
                self.rotate = 270
                self.user = "標準用戶"
            self.ui.statusbar4.setText("{:<30}".format("旋轉角度: 270度"))
            # self.ui.status_label0.setText("已設置: "+self.user+"; 旋轉角度 270 度")
            # self.ui.status_label0.setStyleSheet(u"color: rgb(255, 0, 255);")

    def listWidgetContext(self, point):
        popMenu = QMenu()
        opt1=popMenu.addAction("删除全部")
        opt2=popMenu.addAction("删除選擇行")
        action=popMenu.exec_(QCursor.pos())

        if action == opt1:
            self.ui.listView.model().removeRows(0,self.ui.listView.model().rowCount())
            return
        elif action == opt2:
            self.ui.listView.model().removeRow(self.ui.listView.currentIndex().row())
            return
        else:
            return



    def load_data(self):
        print(
            f"PyQt5 version: {QtCore.PYQT_VERSION_STR}, Qt version: {QtCore.QT_VERSION_STR}"
        )

        # app = QtWidgets.QApplication(sys.argv)
        filename, _ = QtWidgets.QFileDialog.getOpenFileName(None, filter="PDF (*.pdf)")
        if not filename:
            print("please select the .pdf file")
            sys.exit(0)
        view = QtWebEngineWidgets.QWebEngineView(self.ui.frame1)
        settings = view.settings()
        settings.setAttribute(QtWebEngineWidgets.QWebEngineSettings.PluginsEnabled, True)

        url = QtCore.QUrl.fromLocalFile(filename)
        view.load(url)
        view.resize(640, 480)

        # view.show()
        # sys.exit(app.exec_())
    def list_item_clicked(self,index):

        print(self.CurrentIndexData,index.data())
        if self.CurrentIndexData != index.data():
            print(index.row())
            item_text=index.data()

            checkDoc = fitz.open(item_text)
            width, height = checkDoc[0].mediabox.width, checkDoc[0].mediabox.height
            print(width,height)
            self.ui.statusbar6.setText("{:<30}".format("頁面像素: " + str(int(width)) + " X "+str(int(height))))
            checkDoc.close()
            print(os.getcwd() + '/temp/ShowTemp' + str(index.row()) +'-'+times.__str__()+'.PDF')
            shutil.copyfile(item_text,os.getcwd()+'/temp/ShowTemp'+str(index.row())+'-'+times.__str__()+'.PDF')

            f = Path(os.getcwd()+'/temp/ShowTemp'+str(index.row())+'-'+times.__str__()+'.PDF').as_uri()

            # f = Path(item_text).as_uri()

            # if self.User != "Data_Department":
            self.ui.WebBrowser.dynamicCall('Navigate(const QString&)', f)
        self.CurrentIndexData=index.data()

    def add_file(self,filename):
        """

        :param filename: 類型列表
        :return:
        """
        global times
        times = times + 1
        rowCount = model.rowCount()
        # print(rowCount)
        for row_num, fl in enumerate(filename):
            model.setItem(row_num + rowCount, 0, QStandardItem(fl))
        self.ui.listView.setModel(model)
        self.ui.listView.setCurrentIndex(self.ui.listView.model().index(0, 0))

        # self.main_layout = QVBoxLayout(self)
        #
        # self.WebBrowser = QAxContainer.QAxWidget(self.ui.frame1)  #
        # self.WebBrowser.setFocusPolicy(Qt.StrongFocus)
        # self.WebBrowser.setControl("{8856F961-340A-11D0-A96B-00C04FD705A2}")
        # self.WebBrowser.setProperty("DisplayScrollBars",True)
        #
        # self.main_layout.addWidget(self.WebBrowser)

        # self.ui.horizontalLayout.addWidget(self.WebBrowser)
        # self.WebBrowser.resize(800, 680)

        # self.ui.frame1.resize(800, 680)
        # self.ui.frame1.repaint()

        index0 = self.ui.listView.model().data(self.ui.listView.model().index(0, 0))
        self.CurrentIndexData = index0
        checkDoc = fitz.open(index0)
        width, height = checkDoc[0].mediabox.width, checkDoc[0].mediabox.height
        print(width, height)
        self.ui.statusbar6.setText("{:<30}".format("頁面像素: " + str(int(width)) + " X " + str(int(height))))
        checkDoc.close()

        # f = Path(filename[0]).as_uri()

        shutil.copyfile(index0, os.getcwd() + '/temp/ShowTemp0' + '-' + times.__str__() + '.PDF')

        f = Path(os.getcwd() + '/temp/ShowTemp0' + '-' + times.__str__() + '.PDF').as_uri()
        # f = Path(index0).as_uri()
        # load object
        print(f)
        # self.ui.WebBrowser.setControl("Acrobat Reader")
        # self.ui.WebBrowser.dynamicCall('LoadFile(const QString)', f)
        if self.User != "Data_Department":
            self.ui.WebBrowser.dynamicCall('Navigate(const QString&)', f)
        # self.WebBrowser.dynamicCall('Navigate(const QString&)', "https://www.baidu.com/")

    def Open_File(self):
        global times
        # filename, _ = QtWidgets.QFileDialog.getOpenFileName(None, filter="PDF (*.pdf)")
        filename, _ = QtWidgets.QFileDialog.getOpenFileNames(
            self,  # 父窗口对象
            "选择要處理的PDF檔",  # 标题
            # r"./",  # 起始目录
            None,
            "數據類型 (*.PDF)"  # 选择类型过滤项，过滤内容在括号中
        )
        if filename:
            self.add_file(filename)
            # times=times+1
            # rowCount=model.rowCount()
            # for row_num,fl in enumerate(filename):
            #     model.setItem(row_num+rowCount,0,QStandardItem(fl))
            # self.ui.listView.setModel(model)
            # self.ui.listView.setCurrentIndex(self.ui.listView.model().index(0,0))
            # index0=self.ui.listView.model().data(self.ui.listView.model().index(0,0))
            # self.CurrentIndexData=index0
            # checkDoc = fitz.open(index0)
            # width, height = checkDoc[0].mediabox.width, checkDoc[0].mediabox.height
            # print(width,height)
            # self.ui.statusbar6.setText("{:<30}".format("頁面像素: " + str(int(width)) + " X "+str(int(height))))
            # checkDoc.close()
            # shutil.copyfile(index0,os.getcwd()+'/temp/ShowTemp0'+'-'+times.__str__()+'.PDF')
            # f = Path(os.getcwd()+'/temp/ShowTemp0'+'-'+times.__str__()+'.PDF').as_uri()
            # print(f)
            # self.ui.WebBrowser.dynamicCall('Navigate(const QString&)', f)


    def fun_split(self):
        print("fun_split")
        if model.rowCount() == 0:
            QMessageBox.information(self, "信息", "沒有打開PDF檔!")
            return
        if self.ui.radioButton.isChecked():
            for i in range(model.rowCount()):
                pindex=model.index(i,0)
                print(model.data(pindex))
                doc1 = fitz.open(model.data(pindex))
                filepath1=os.path.dirname(model.data(pindex))
                filename1 = os.path.basename(model.data(pindex))
                filename2=os.path.splitext(filename1)[0]
                # width_total, height_total = doc1[0].mediabox_size
                # if self.rotate == 90 or self.rotate == 270:
                #     width_total, height_total = height_total, width_total
                # r1 = fitz.Rect(0, 0, width_total, height_total)
                for pl in range(doc1.page_count):
                    doc = fitz.open()
                    doc.insert_pdf(doc1,from_page=pl, to_page=pl, rotate=self.rotate)
                #     page = doc.new_page(-1, width=width_total, height=height_total)
                #     # pagel = doc.load_page(pl.number)
                #     page.show_pdf_page(r1, doc1, pl.number, keep_proportion=1, rotate=self.rotate)

                    try:
                        with open(filepath1+'/'+filename2+"_Page"+str(pl+1)+".pdf", 'w') as f:
                            f.close()
                            doc.save(filepath1+'/'+filename2+"_Page"+str(pl+1)+".pdf", garbage=0, deflate=True)

                            # print(f.mode)
                            # do something with the file
                    except PermissionError:
                        # print(PermissionError.strerror)
                        QMessageBox.information(self, "警告", '文件已经被打开:'+filename2+"_Page"+str(pl+1)+".pdf")
                        return

            QMessageBox.information(self, "信息", "處理完成: 文件名: "+filepath1+"/XXX_PageXXX.pdf")

        if self.ui.radioButton_2.isChecked():
            if self.ui.lineEdit.text()=='':
                QMessageBox.information(self, "信息", "請輸入拆分頁數.")
                return

            PDF_qty=int(self.ui.lineEdit.text())
            for i in range(model.rowCount()):
                pindex=model.index(i,0)
                print(model.data(pindex))
                doc1 = fitz.open(model.data(pindex))
                filepath1=os.path.dirname(model.data(pindex))
                filename1 = os.path.basename(model.data(pindex))
                filename2=os.path.splitext(filename1)[0]
                # width_total, height_total = doc1[0].mediabox_size
                # if self.rotate == 90 or self.rotate == 270:
                #     width_total, height_total = height_total, width_total
                # r1 = fitz.Rect(0, 0, width_total, height_total)
                doc = fitz.open()
                for pl in range(doc1.page_count):

                    doc.insert_pdf(doc1,from_page=pl, to_page=pl, rotate=self.rotate)
                    # page = doc.new_page(-1, width=width_total, height=height_total)
                    # pagel = doc.load_page(pl.number)
                    # page.show_pdf_page(r1, doc1, pl.number, keep_proportion=1, rotate=self.rotate)
                    if divmod(pl+1,PDF_qty)[1]==0:
                        try:
                            with open(filepath1+'/'+filename2+"_Split"+str(divmod(pl+1,PDF_qty)[0])+".pdf", 'w') as f:
                                f.close()
                                doc.save(filepath1+'/'+filename2+"_Split"+str(divmod(pl+1,PDF_qty)[0])+".pdf", garbage=0, deflate=True)
                                # QMessageBox.information(self, "信息", "處理完成: 文件名: " + filepath1 + '/' + filename2 + "_Split" + str(divmod(pl + 1, PDF_qty)[0]) + ".pdf")
                                # print(f.mode)
                                # do something with the file
                        except PermissionError:
                            # print(PermissionError.strerror)
                            QMessageBox.information(self, "警告", '文件已经被打开:' + filename2+"_Split"+str(divmod(pl+1,PDF_qty)[0])+".pdf")
                            return

                        doc = fitz.open()

                if divmod(doc1.page_count,PDF_qty)[1]!=0:

                    try:
                        with open(filepath1+'/'+filename2+"_Split"+str(divmod(pl+1,PDF_qty)[0]+1)+".pdf", 'w') as f:
                            f.close()
                            doc.save(filepath1+'/'+filename2+"_Split"+str(divmod(pl+1,PDF_qty)[0]+1)+".pdf", garbage=0, deflate=True)
                            # QMessageBox.information(self, "信息", "處理完成: 文件名: " + filepath1+'/'+filename2+"_Split"+str(divmod(pl+1,PDF_qty)[0]+1)+".pdf")
                            # print(f.mode)
                            # do something with the file
                    except PermissionError:
                        # print(PermissionError.strerror)
                        QMessageBox.information(self, "警告", '文件已经被打开:' + filename2+"_Split"+str(divmod(pl+1,PDF_qty)[0]+1)+".pdf")
                        return

            QMessageBox.information(self, "信息", "處理完成: 文件名: " + filepath1 + "/XXX_SplitXXX.pdf")

        if self.ui.radioButton_3.isChecked():
            if self.ui.textEdit.toPlainText() == '':
                QMessageBox.information(self, "信息", "請輸入拆分頁碼, 以逗號分隔!.")
                return
            PDF_num=self.ui.textEdit.toPlainText()
            PDF_num_list=PDF_num.split(',')
            for i in range(model.rowCount()):
                pindex=model.index(i,0)
                print(model.data(pindex))
                doc1 = fitz.open(model.data(pindex))
                filepath1=os.path.dirname(model.data(pindex))
                filename1 = os.path.basename(model.data(pindex))
                filename2=os.path.splitext(filename1)[0]
                # width_total, height_total = doc1[0].mediabox_size
                # if self.rotate == 90 or self.rotate == 270:
                #     width_total, height_total = height_total, width_total
                # r1 = fitz.Rect(0, 0, width_total, height_total)
                doc = fitz.open()

                for row,pl in enumerate(PDF_num_list):
                    print(self.rotate)
                    intpl=int(pl)-1
                    doc.insert_pdf(doc1, from_page=intpl, to_page=intpl, rotate=self.rotate)
                    # page = doc.new_page(-1, width=width_total, height=height_total)
                    # pagel = doc.load_page(introw)
                    # page.show_pdf_page(r1, doc1, intpl, keep_proportion=1,rotate=self.rotate)

                try:
                    with open(filepath1+'/'+filename2+"_specify_page.pdf", 'w') as f:
                        f.close()
                        doc.save(filepath1+'/'+filename2+"_specify_page.pdf", garbage=0, deflate=True)

                        # print(f.mode)
                        # do something with the file
                except PermissionError:
                    # print(PermissionError.strerror)
                    QMessageBox.information(self, "警告", '文件已经被打开:' + filename2+"_specify_page.pdf")
                    return

            QMessageBox.information(self, "信息", "處理完成: 文件名: " + filepath1 + "/XXX_specify_page.pdf")

    def fun_mergePDF(self):
        print("fun_mergePDF")

        if model.rowCount()>0:
            page_count=0
            doc = fitz.open()
            for i in range(model.rowCount()):
                pindex=model.index(i,0)
                print(model.data(pindex))
                doc1 = fitz.open(model.data(pindex))
                doc.insert_pdf(doc1,rotate=self.rotate)
                print(self.rotate)
                # width_total, height_total = doc1[0].mediabox_size
                # width_total, height_total = doc1[0].rect.width,doc1[0].rect.height
                # print(width_total, height_total)
                # if self.rotate==90 or self.rotate==270:
                #     width_total,height_total=height_total,width_total
                #
                # r1 = fitz.Rect(0, 0, width_total, height_total)



                # for pn in range(doc1.page_count):
                #     page = doc.new_page(-1, width=width_total, height=height_total)


                # for pl in doc1:
                #     pagel = doc.load_page(pl.number+page_count)
                #     pagel.show_pdf_page(r1, doc1, pl.number, keep_proportion=1,rotate=self.rotate)

                # page_count=page_count+doc1.page_count
                # print(page_count)

            filepath1 = os.path.dirname(model.data(pindex))
            filename1 = os.path.basename(model.data(pindex))
            filename2 = os.path.splitext(filename1)[0]

            try:
                with open(filepath1 + '/' + "Merge.pdf", 'w') as f:
                    f.close()
                    doc.save(filepath1 + '/' + "Merge.pdf", garbage=0, deflate=True)

                    # print(f.mode)
                    # do something with the file
            except PermissionError:
                # print(PermissionError.strerror)
                QMessageBox.information(self, "警告", '文件已经被打开:' + filepath1 + '/' + "Merge.pdf")
                return

            QMessageBox.information(self, "信息", "處理完成: 文件名: " + filepath1 + '/' + "Merge.pdf")
        else:
            QMessageBox.information(self, "信息", "沒有打開PDF檔!")

    def fun_mergePDF_show_pdf_page(self):
        print("fun_mergePDF")

        if model.rowCount()>0:
            page_count=0
            doc = fitz.open()
            for i in range(model.rowCount()):
                pindex=model.index(i,0)
                print(model.data(pindex))
                doc1 = fitz.open(model.data(pindex))
                width_total, height_total = doc1[0].mediabox_size
                width_total, height_total = doc1[0].rect.width,doc1[0].rect.height
                print(width_total, height_total)
                if self.rotate==90 or self.rotate==270:
                    width_total,height_total=height_total,width_total

                r1 = fitz.Rect(0, 0, width_total, height_total)

                for pn in range(doc1.page_count):
                    page = doc.new_page(-1, width=width_total, height=height_total)


                for pl in doc1:
                    pagel = doc.load_page(pl.number+page_count)
                    pagel.show_pdf_page(r1, doc1, pl.number, keep_proportion=1,rotate=self.rotate)

                page_count=page_count+doc1.page_count
                print(page_count)

            filepath1 = os.path.dirname(model.data(pindex))
            filename1 = os.path.basename(model.data(pindex))
            filename2 = os.path.splitext(filename1)[0]

            try:
                with open(filepath1 + '/' + "Merge.pdf", 'w') as f:
                    f.close()
                    doc.save(filepath1 + '/' + "Merge.pdf", garbage=0, deflate=True)

                    # print(f.mode)
                    # do something with the file
            except PermissionError:
                # print(PermissionError.strerror)
                QMessageBox.information(self, "警告", '文件已经被打开:' + filepath1 + '/' + "Merge.pdf")
                return

            QMessageBox.information(self, "信息", "處理完成: 文件名: " + filepath1 + '/' + "Merge.pdf")
        else:
            QMessageBox.information(self, "信息", "沒有打開PDF檔!")

if __name__ == "__main__":
    times=0
    app = QApplication([])
    model = QStandardItemModel()
    ui_mainwindow = mainwindow(QMainWindow())
    ui_mainwindow.show()
    sys.exit(app.exec_())


