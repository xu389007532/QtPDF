# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'ui_main.ui'
#
# Created by: PyQt5 UI code generator 5.15.10
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(1259, 821)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap(":/ico_file/sourcefile/title.ico"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        MainWindow.setWindowIcon(icon)
        MainWindow.setStyleSheet("QLabel{\n"
"    color: rgb(0, 0, 127);\n"
"    font: 12pt \"Microsoft Sans Serif\";\n"
"}\n"
"QComboBox{\n"
"    font: 12pt \"Microsoft Sans Serif\";\n"
"}\n"
"QLineEdit{\n"
"font: 12pt \"Microsoft Sans Serif\";\n"
"}\n"
"QTableView{\n"
"    border-color: rgb(0, 85, 0);\n"
"    \n"
"    alternate-background-color: rgb(230, 201, 255);\n"
"    background-color: rgb(199, 205, 255);\n"
"/*    selection-background-color: rgb(85, 255, 0);*/\n"
"\n"
"}\n"
"QPushButton{\n"
"    /*background-color: rgb(170, 255, 127);*/\n"
"    \n"
"    font: 16pt \"Microsoft Sans Serif\";\n"
"    \n"
"    background-color: rgb(239, 239, 119);\n"
"}\n"
"QCheckBox{\n"
"font: 12pt \"Microsoft Sans Serif\";\n"
"    color: rgb(255, 85, 0);\n"
"}")
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.process_tableView2 = QtWidgets.QTableView(self.centralwidget)
        self.process_tableView2.setGeometry(QtCore.QRect(10, 490, 1240, 271))
        self.process_tableView2.setAlternatingRowColors(True)
        self.process_tableView2.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.process_tableView2.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.process_tableView2.setObjectName("process_tableView2")
        self.open_tableView1 = QtWidgets.QTableView(self.centralwidget)
        self.open_tableView1.setGeometry(QtCore.QRect(10, 54, 1241, 111))
        self.open_tableView1.setAlternatingRowColors(True)
        self.open_tableView1.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.open_tableView1.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.open_tableView1.setObjectName("open_tableView1")
        self.open_tableView1.horizontalHeader().setCascadingSectionResizes(False)
        self.open_tableView1.horizontalHeader().setDefaultSectionSize(50)
        self.open_tableView1.horizontalHeader().setStretchLastSection(True)
        self.frame_act_size = QtWidgets.QFrame(self.centralwidget)
        self.frame_act_size.setGeometry(QtCore.QRect(430, 298, 431, 121))
        self.frame_act_size.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_act_size.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_act_size.setObjectName("frame_act_size")
        self.lineEdit_r3c3 = QtWidgets.QLineEdit(self.frame_act_size)
        self.lineEdit_r3c3.setEnabled(True)
        self.lineEdit_r3c3.setGeometry(QtCore.QRect(110, 70, 81, 31))
        self.lineEdit_r3c3.setText("")
        self.lineEdit_r3c3.setObjectName("lineEdit_r3c3")
        self.label_15 = QtWidgets.QLabel(self.frame_act_size)
        self.label_15.setGeometry(QtCore.QRect(13, 75, 91, 20))
        self.label_15.setObjectName("label_15")
        self.lineEdit_r2c3 = QtWidgets.QLineEdit(self.frame_act_size)
        self.lineEdit_r2c3.setEnabled(True)
        self.lineEdit_r2c3.setGeometry(QtCore.QRect(110, 15, 81, 31))
        self.lineEdit_r2c3.setObjectName("lineEdit_r2c3")
        self.label_14 = QtWidgets.QLabel(self.frame_act_size)
        self.label_14.setGeometry(QtCore.QRect(13, 20, 91, 20))
        self.label_14.setObjectName("label_14")
        self.label_16 = QtWidgets.QLabel(self.frame_act_size)
        self.label_16.setGeometry(QtCore.QRect(240, 70, 91, 31))
        self.label_16.setObjectName("label_16")
        self.lineEdit_r2c4 = QtWidgets.QLineEdit(self.frame_act_size)
        self.lineEdit_r2c4.setEnabled(True)
        self.lineEdit_r2c4.setGeometry(QtCore.QRect(330, 15, 61, 31))
        self.lineEdit_r2c4.setObjectName("lineEdit_r2c4")
        self.lineEdit_r3c4 = QtWidgets.QLineEdit(self.frame_act_size)
        self.lineEdit_r3c4.setEnabled(True)
        self.lineEdit_r3c4.setGeometry(QtCore.QRect(330, 70, 61, 31))
        self.lineEdit_r3c4.setText("")
        self.lineEdit_r3c4.setObjectName("lineEdit_r3c4")
        self.label_13 = QtWidgets.QLabel(self.frame_act_size)
        self.label_13.setGeometry(QtCore.QRect(240, 15, 91, 31))
        self.label_13.setObjectName("label_13")
        self.frame_bleeding = QtWidgets.QFrame(self.centralwidget)
        self.frame_bleeding.setGeometry(QtCore.QRect(430, 248, 661, 41))
        self.frame_bleeding.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_bleeding.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_bleeding.setObjectName("frame_bleeding")
        self.lineEdit_r1c6 = QtWidgets.QLineEdit(self.frame_bleeding)
        self.lineEdit_r1c6.setGeometry(QtCore.QRect(570, 4, 61, 31))
        self.lineEdit_r1c6.setObjectName("lineEdit_r1c6")
        self.label_18 = QtWidgets.QLabel(self.frame_bleeding)
        self.label_18.setGeometry(QtCore.QRect(500, 4, 81, 31))
        self.label_18.setObjectName("label_18")
        self.lineEdit_r1c5 = QtWidgets.QLineEdit(self.frame_bleeding)
        self.lineEdit_r1c5.setGeometry(QtCore.QRect(420, 4, 61, 31))
        self.lineEdit_r1c5.setObjectName("lineEdit_r1c5")
        self.label_17 = QtWidgets.QLabel(self.frame_bleeding)
        self.label_17.setGeometry(QtCore.QRect(351, 3, 81, 31))
        self.label_17.setObjectName("label_17")
        self.label_7 = QtWidgets.QLabel(self.frame_bleeding)
        self.label_7.setGeometry(QtCore.QRect(0, 0, 81, 31))
        self.label_7.setObjectName("label_7")
        self.lineEdit_r1c3 = QtWidgets.QLineEdit(self.frame_bleeding)
        self.lineEdit_r1c3.setEnabled(False)
        self.lineEdit_r1c3.setGeometry(QtCore.QRect(90, 4, 61, 31))
        self.lineEdit_r1c3.setObjectName("lineEdit_r1c3")
        self.label_8 = QtWidgets.QLabel(self.frame_bleeding)
        self.label_8.setGeometry(QtCore.QRect(160, 1, 81, 31))
        self.label_8.setObjectName("label_8")
        self.lineEdit_r1c4 = QtWidgets.QLineEdit(self.frame_bleeding)
        self.lineEdit_r1c4.setEnabled(False)
        self.lineEdit_r1c4.setGeometry(QtCore.QRect(251, 4, 61, 31))
        self.lineEdit_r1c4.setObjectName("lineEdit_r1c4")
        self.frame = QtWidgets.QFrame(self.centralwidget)
        self.frame.setGeometry(QtCore.QRect(20, 248, 391, 181))
        self.frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame.setObjectName("frame")
        self.lineEdit_r3c2 = QtWidgets.QLineEdit(self.frame)
        self.lineEdit_r3c2.setGeometry(QtCore.QRect(280, 130, 61, 31))
        self.lineEdit_r3c2.setObjectName("lineEdit_r3c2")
        self.label_5 = QtWidgets.QLabel(self.frame)
        self.label_5.setGeometry(QtCore.QRect(1, 10, 64, 20))
        self.label_5.setObjectName("label_5")
        self.lineEdit_r2c2 = QtWidgets.QLineEdit(self.frame)
        self.lineEdit_r2c2.setGeometry(QtCore.QRect(250, 70, 61, 31))
        self.lineEdit_r2c2.setObjectName("lineEdit_r2c2")
        self.label_9 = QtWidgets.QLabel(self.frame)
        self.label_9.setGeometry(QtCore.QRect(169, 70, 81, 31))
        self.label_9.setObjectName("label_9")
        self.label_11 = QtWidgets.QLabel(self.frame)
        self.label_11.setGeometry(QtCore.QRect(0, 130, 81, 31))
        self.label_11.setObjectName("label_11")
        self.label_6 = QtWidgets.QLabel(self.frame)
        self.label_6.setGeometry(QtCore.QRect(168, 12, 74, 20))
        self.label_6.setObjectName("label_6")
        self.label_12 = QtWidgets.QLabel(self.frame)
        self.label_12.setGeometry(QtCore.QRect(180, 130, 81, 31))
        self.label_12.setObjectName("label_12")
        self.lineEdit_r1c1 = QtWidgets.QLineEdit(self.frame)
        self.lineEdit_r1c1.setGeometry(QtCore.QRect(71, 9, 60, 31))
        self.lineEdit_r1c1.setObjectName("lineEdit_r1c1")
        self.lineEdit_r3c1 = QtWidgets.QLineEdit(self.frame)
        self.lineEdit_r3c1.setGeometry(QtCore.QRect(90, 130, 61, 31))
        self.lineEdit_r3c1.setObjectName("lineEdit_r3c1")
        self.lineEdit_r1c2 = QtWidgets.QLineEdit(self.frame)
        self.lineEdit_r1c2.setGeometry(QtCore.QRect(250, 12, 60, 30))
        self.lineEdit_r1c2.setObjectName("lineEdit_r1c2")
        self.label_10 = QtWidgets.QLabel(self.frame)
        self.label_10.setGeometry(QtCore.QRect(0, 70, 81, 31))
        self.label_10.setObjectName("label_10")
        self.lineEdit_r2c1 = QtWidgets.QLineEdit(self.frame)
        self.lineEdit_r2c1.setGeometry(QtCore.QRect(80, 70, 61, 31))
        self.lineEdit_r2c1.setObjectName("lineEdit_r2c1")
        self.frame_2 = QtWidgets.QFrame(self.centralwidget)
        self.frame_2.setGeometry(QtCore.QRect(10, 198, 1251, 41))
        self.frame_2.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_2.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_2.setObjectName("frame_2")
        self.comboBox_rotate = QtWidgets.QComboBox(self.frame_2)
        self.comboBox_rotate.setGeometry(QtCore.QRect(840, 10, 61, 26))
        self.comboBox_rotate.setObjectName("comboBox_rotate")
        self.comboBox_rotate.addItem("")
        self.comboBox_rotate.addItem("")
        self.comboBox_rotate.addItem("")
        self.comboBox_rotate.addItem("")
        self.comboBox1 = QtWidgets.QComboBox(self.frame_2)
        self.comboBox1.setGeometry(QtCore.QRect(81, 10, 135, 26))
        self.comboBox1.setObjectName("comboBox1")
        self.comboBox1.addItem("")
        self.comboBox1.addItem("")
        self.comboBox1.addItem("")
        self.comboBox1.addItem("")
        self.comboBox1.addItem("")
        self.label_2 = QtWidgets.QLabel(self.frame_2)
        self.label_2.setGeometry(QtCore.QRect(11, 10, 64, 20))
        self.label_2.setObjectName("label_2")
        self.label_3 = QtWidgets.QLabel(self.frame_2)
        self.label_3.setGeometry(QtCore.QRect(230, 10, 111, 20))
        self.label_3.setAutoFillBackground(False)
        self.label_3.setObjectName("label_3")
        self.label_4 = QtWidgets.QLabel(self.frame_2)
        self.label_4.setGeometry(QtCore.QRect(460, 10, 141, 20))
        self.label_4.setObjectName("label_4")
        self.checkBox1 = QtWidgets.QCheckBox(self.frame_2)
        self.checkBox1.setGeometry(QtCore.QRect(930, 11, 87, 24))
        font = QtGui.QFont()
        font.setFamily("Microsoft Sans Serif")
        font.setPointSize(12)
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(50)
        self.checkBox1.setFont(font)
        self.checkBox1.setChecked(False)
        self.checkBox1.setObjectName("checkBox1")
        self.comboBox3 = QtWidgets.QComboBox(self.frame_2)
        self.comboBox3.setGeometry(QtCore.QRect(600, 10, 151, 26))
        self.comboBox3.setObjectName("comboBox3")
        self.comboBox3.addItem("")
        self.comboBox3.addItem("")
        self.comboBox3.addItem("")
        self.comboBox3.addItem("")
        self.comboBox3.addItem("")
        self.comboBox3.addItem("")
        self.comboBox3.addItem("")
        self.comboBox3.addItem("")
        self.comboBox3.addItem("")
        self.checkBox_blank2 = QtWidgets.QCheckBox(self.frame_2)
        self.checkBox_blank2.setGeometry(QtCore.QRect(1107, 10, 141, 24))
        font = QtGui.QFont()
        font.setFamily("Microsoft Sans Serif")
        font.setPointSize(12)
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(50)
        self.checkBox_blank2.setFont(font)
        self.checkBox_blank2.setChecked(False)
        self.checkBox_blank2.setObjectName("checkBox_blank2")
        self.label_19 = QtWidgets.QLabel(self.frame_2)
        self.label_19.setGeometry(QtCore.QRect(770, 10, 64, 20))
        self.label_19.setObjectName("label_19")
        self.comboBox2 = QtWidgets.QComboBox(self.frame_2)
        self.comboBox2.setGeometry(QtCore.QRect(337, 10, 103, 26))
        self.comboBox2.setObjectName("comboBox2")
        self.comboBox2.addItem("")
        self.comboBox2.addItem("")
        self.checkBox_blank = QtWidgets.QCheckBox(self.frame_2)
        self.checkBox_blank.setGeometry(QtCore.QRect(1027, 10, 87, 24))
        font = QtGui.QFont()
        font.setFamily("Microsoft Sans Serif")
        font.setPointSize(12)
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(50)
        self.checkBox_blank.setFont(font)
        self.checkBox_blank.setChecked(False)
        self.checkBox_blank.setObjectName("checkBox_blank")
        self.frame_3 = QtWidgets.QFrame(self.centralwidget)
        self.frame_3.setGeometry(QtCore.QRect(20, 428, 1231, 61))
        self.frame_3.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_3.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_3.setObjectName("frame_3")
        self.progressBar = QtWidgets.QProgressBar(self.frame_3)
        self.progressBar.setGeometry(QtCore.QRect(422, 10, 771, 40))
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(12)
        font.setBold(True)
        font.setWeight(75)
        font.setKerning(True)
        self.progressBar.setFont(font)
        self.progressBar.setLayoutDirection(QtCore.Qt.RightToLeft)
        self.progressBar.setLocale(QtCore.QLocale(QtCore.QLocale.English, QtCore.QLocale.UnitedStates))
        self.progressBar.setProperty("value", 0)
        self.progressBar.setObjectName("progressBar")
        self.process_pushButton3 = QtWidgets.QPushButton(self.frame_3)
        self.process_pushButton3.setEnabled(False)
        self.process_pushButton3.setGeometry(QtCore.QRect(229, 5, 151, 51))
        icon1 = QtGui.QIcon()
        icon1.addPixmap(QtGui.QPixmap(":/ico_file/sourcefile/run.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.process_pushButton3.setIcon(icon1)
        self.process_pushButton3.setObjectName("process_pushButton3")
        self.add_pushButton2 = QtWidgets.QPushButton(self.frame_3)
        self.add_pushButton2.setEnabled(False)
        self.add_pushButton2.setGeometry(QtCore.QRect(10, 6, 161, 51))
        icon2 = QtGui.QIcon()
        icon2.addPixmap(QtGui.QPixmap(":/ico_file/sourcefile/add2.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.add_pushButton2.setIcon(icon2)
        self.add_pushButton2.setObjectName("add_pushButton2")
        self.frame_5 = QtWidgets.QFrame(self.centralwidget)
        self.frame_5.setGeometry(QtCore.QRect(0, 1, 1251, 61))
        self.frame_5.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_5.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_5.setObjectName("frame_5")
        self.label_art = QtWidgets.QLabel(self.frame_5)
        self.label_art.setGeometry(QtCore.QRect(760, 16, 511, 21))
        self.label_art.setObjectName("label_art")
        self.open_pushButton1 = QtWidgets.QPushButton(self.frame_5)
        self.open_pushButton1.setGeometry(QtCore.QRect(10, 4, 221, 41))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.open_pushButton1.sizePolicy().hasHeightForWidth())
        self.open_pushButton1.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("Microsoft Sans Serif")
        font.setPointSize(16)
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(50)
        self.open_pushButton1.setFont(font)
        icon3 = QtGui.QIcon()
        icon3.addPixmap(QtGui.QPixmap("C:/Users/ITProg02/.designer/backup/sourcefile/open2.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.open_pushButton1.setIcon(icon3)
        self.open_pushButton1.setObjectName("open_pushButton1")
        self.open_pushButton_folder = QtWidgets.QPushButton(self.frame_5)
        self.open_pushButton_folder.setGeometry(QtCore.QRect(270, 5, 221, 41))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.open_pushButton_folder.sizePolicy().hasHeightForWidth())
        self.open_pushButton_folder.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("Microsoft Sans Serif")
        font.setPointSize(16)
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(50)
        self.open_pushButton_folder.setFont(font)
        self.open_pushButton_folder.setStyleSheet("background-color: rgb(85, 170, 255);")
        self.open_pushButton_folder.setIcon(icon3)
        self.open_pushButton_folder.setObjectName("open_pushButton_folder")
        self.open_pushButton_art = QtWidgets.QPushButton(self.frame_5)
        self.open_pushButton_art.setGeometry(QtCore.QRect(530, 4, 221, 41))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.open_pushButton_art.sizePolicy().hasHeightForWidth())
        self.open_pushButton_art.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("Microsoft Sans Serif")
        font.setPointSize(16)
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(50)
        self.open_pushButton_art.setFont(font)
        self.open_pushButton_art.setStyleSheet("background-color: rgb(115, 170, 130);")
        self.open_pushButton_art.setIcon(icon3)
        self.open_pushButton_art.setObjectName("open_pushButton_art")
        self.checkBox_addpage = QtWidgets.QCheckBox(self.centralwidget)
        self.checkBox_addpage.setGeometry(QtCore.QRect(1120, 238, 87, 24))
        font = QtGui.QFont()
        font.setFamily("Microsoft Sans Serif")
        font.setPointSize(12)
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(50)
        self.checkBox_addpage.setFont(font)
        self.checkBox_addpage.setChecked(False)
        self.checkBox_addpage.setObjectName("checkBox_addpage")
        self.lineEdit_Repeat = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_Repeat.setGeometry(QtCore.QRect(1190, 304, 60, 30))
        self.lineEdit_Repeat.setObjectName("lineEdit_Repeat")
        self.label_20 = QtWidgets.QLabel(self.centralwidget)
        self.label_20.setGeometry(QtCore.QRect(1110, 308, 74, 20))
        self.label_20.setObjectName("label_20")
        self.lineEdit_StartSeq = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_StartSeq.setGeometry(QtCore.QRect(1170, 338, 80, 31))
        self.lineEdit_StartSeq.setObjectName("lineEdit_StartSeq")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(1096, 342, 81, 21))
        self.label.setObjectName("label")
        self.label_21 = QtWidgets.QLabel(self.centralwidget)
        self.label_21.setGeometry(QtCore.QRect(1040, 383, 131, 21))
        self.label_21.setObjectName("label_21")
        self.lineEdit_ColorBlockPosition = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_ColorBlockPosition.setGeometry(QtCore.QRect(1170, 379, 80, 31))
        self.lineEdit_ColorBlockPosition.setObjectName("lineEdit_ColorBlockPosition")
        self.checkBox_addSample = QtWidgets.QCheckBox(self.centralwidget)
        self.checkBox_addSample.setEnabled(False)
        self.checkBox_addSample.setGeometry(QtCore.QRect(1120, 268, 121, 24))
        font = QtGui.QFont()
        font.setFamily("Microsoft Sans Serif")
        font.setPointSize(12)
        font.setBold(False)
        font.setItalic(False)
        font.setWeight(50)
        self.checkBox_addSample.setFont(font)
        self.checkBox_addSample.setChecked(False)
        self.checkBox_addSample.setObjectName("checkBox_addSample")
        self.lineEdit_ERP_Mark = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_ERP_Mark.setGeometry(QtCore.QRect(130, 170, 1121, 31))
        self.lineEdit_ERP_Mark.setObjectName("lineEdit_ERP_Mark")
        self.label_22 = QtWidgets.QLabel(self.centralwidget)
        self.label_22.setGeometry(QtCore.QRect(20, 170, 111, 31))
        self.label_22.setObjectName("label_22")
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 1259, 30))
        font = QtGui.QFont()
        font.setPointSize(14)
        self.menubar.setFont(font)
        self.menubar.setObjectName("menubar")
        self.menu = QtWidgets.QMenu(self.menubar)
        font = QtGui.QFont()
        font.setPointSize(14)
        font.setKerning(True)
        font.setStyleStrategy(QtGui.QFont.PreferDefault)
        self.menu.setFont(font)
        self.menu.setContextMenuPolicy(QtCore.Qt.DefaultContextMenu)
        self.menu.setToolTip("")
        self.menu.setWhatsThis("")
        self.menu.setObjectName("menu")
        self.menu_2 = QtWidgets.QMenu(self.menubar)
        self.menu_2.setObjectName("menu_2")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.action1 = QtWidgets.QAction(MainWindow)
        icon4 = QtGui.QIcon()
        icon4.addPixmap(QtGui.QPixmap(":/ico_file/sourcefile/favicon.ico"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.action1.setIcon(icon4)
        font = QtGui.QFont()
        font.setFamily("Microsoft Sans Serif")
        font.setPointSize(14)
        self.action1.setFont(font)
        self.action1.setObjectName("action1")
        self.action22 = QtWidgets.QAction(MainWindow)
        icon5 = QtGui.QIcon()
        icon5.addPixmap(QtGui.QPixmap(":/ico_file/sourcefile/save.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.action22.setIcon(icon5)
        self.action22.setObjectName("action22")
        self.action2 = QtWidgets.QAction(MainWindow)
        icon6 = QtGui.QIcon()
        icon6.addPixmap(QtGui.QPixmap(":/ico_file/sourcefile/8923926.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.action2.setIcon(icon6)
        self.action2.setObjectName("action2")
        self.action_quit = QtWidgets.QAction(MainWindow)
        self.action_quit.setObjectName("action_quit")
        self.action_dele = QtWidgets.QAction(MainWindow)
        self.action_dele.setObjectName("action_dele")
        self.PDFart = QtWidgets.QAction(MainWindow)
        self.PDFart.setCheckable(True)
        self.PDFart.setObjectName("PDFart")
        self.menu.addAction(self.action1)
        self.menu.addAction(self.action2)
        self.menu.addAction(self.action22)
        self.menu.addAction(self.action_dele)
        self.menu_2.addAction(self.action_quit)
        self.menubar.addAction(self.menu.menuAction())
        self.menubar.addAction(self.menu_2.menuAction())

        self.retranslateUi(MainWindow)
        self.comboBox_rotate.setCurrentIndex(0)
        self.comboBox1.setCurrentIndex(0)
        self.comboBox3.setCurrentIndex(0)
        self.comboBox2.setCurrentIndex(0)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Honour PDF排版及數據處理軟件"))
        self.label_15.setText(_translate("MainWindow", "水平位置"))
        self.label_14.setText(_translate("MainWindow", "實際成品寬"))
        self.label_16.setText(_translate("MainWindow", "垂直位置"))
        self.label_13.setText(_translate("MainWindow", "實際成品高"))
        self.frame_bleeding.setToolTip(_translate("MainWindow", "出血位"))
        self.label_18.setText(_translate("MainWindow", "下邊增減"))
        self.label_17.setText(_translate("MainWindow", "右邊增減"))
        self.label_7.setText(_translate("MainWindow", "左右出血位(寸)"))
        self.label_8.setText(_translate("MainWindow", "上下出血位(寸)"))
        self.label_5.setText(_translate("MainWindow", "橫排個數"))
        self.label_9.setText(_translate("MainWindow", "豎刀位(寸)"))
        self.label_11.setText(_translate("MainWindow", "一箱張數"))
        self.label_6.setText(_translate("MainWindow", "橫刀位(寸)"))
        self.label_12.setText(_translate("MainWindow", "PDF檔頁數"))
        self.label_10.setText(_translate("MainWindow", "豎排個數"))
        self.comboBox_rotate.setItemText(0, _translate("MainWindow", "0"))
        self.comboBox_rotate.setItemText(1, _translate("MainWindow", "90"))
        self.comboBox_rotate.setItemText(2, _translate("MainWindow", "180"))
        self.comboBox_rotate.setItemText(3, _translate("MainWindow", "270"))
        self.comboBox1.setItemText(0, _translate("MainWindow", "1:從小到大"))
        self.comboBox1.setItemText(1, _translate("MainWindow", "2:從大到小"))
        self.comboBox1.setItemText(2, _translate("MainWindow", "3:兜圈順序打印"))
        self.comboBox1.setItemText(3, _translate("MainWindow", "4:分切成品-順序"))
        self.comboBox1.setItemText(4, _translate("MainWindow", "5:分切成品-倒序"))
        self.label_2.setText(_translate("MainWindow", "編號次序"))
        self.label_3.setText(_translate("MainWindow", "PDF輸出方向"))
        self.label_4.setText(_translate("MainWindow", "出血位及顯示位置"))
        self.checkBox1.setText(_translate("MainWindow", "雙面打印"))
        self.comboBox3.setItemText(0, _translate("MainWindow", "0:排版後無出血位"))
        self.comboBox3.setItemText(1, _translate("MainWindow", "1:左上角水平位置"))
        self.comboBox3.setItemText(2, _translate("MainWindow", "2:左上角垂直位置"))
        self.comboBox3.setItemText(3, _translate("MainWindow", "3:右上角水平位置"))
        self.comboBox3.setItemText(4, _translate("MainWindow", "4:右上角垂直位置"))
        self.comboBox3.setItemText(5, _translate("MainWindow", "5:左下角水平位置"))
        self.comboBox3.setItemText(6, _translate("MainWindow", "6:左下角垂直位置"))
        self.comboBox3.setItemText(7, _translate("MainWindow", "7:右下角水平位置"))
        self.comboBox3.setItemText(8, _translate("MainWindow", "8:右下角垂直位置"))
        self.checkBox_blank2.setText(_translate("MainWindow", "單面加稿件底頁"))
        self.label_19.setText(_translate("MainWindow", "PDF角度"))
        self.comboBox2.setItemText(0, _translate("MainWindow", "1:从左到右"))
        self.comboBox2.setItemText(1, _translate("MainWindow", "2:从上到下"))
        self.checkBox_blank.setText(_translate("MainWindow", "加隔紙"))
        self.process_pushButton3.setText(_translate("MainWindow", "處理"))
        self.add_pushButton2.setText(_translate("MainWindow", "添加"))
        self.label_art.setText(_translate("MainWindow", "no_art"))
        self.open_pushButton1.setText(_translate("MainWindow", "打開需排版PDF檔"))
        self.open_pushButton_folder.setText(_translate("MainWindow", "打開需排版文件夾"))
        self.open_pushButton_art.setText(_translate("MainWindow", "打開排版PDF稿件"))
        self.checkBox_addpage.setText(_translate("MainWindow", "加頁碼"))
        self.lineEdit_Repeat.setText(_translate("MainWindow", "1"))
        self.label_20.setText(_translate("MainWindow", "重復數量"))
        self.lineEdit_StartSeq.setText(_translate("MainWindow", "1"))
        self.label.setText(_translate("MainWindow", "開始編號"))
        self.label_21.setText(_translate("MainWindow", "色塊位置增減(寸)"))
        self.lineEdit_ColorBlockPosition.setText(_translate("MainWindow", "0"))
        self.checkBox_addSample.setText(_translate("MainWindow", "增加大貨樣板"))
        self.label_22.setText(_translate("MainWindow", "ERP 備注信息"))
        self.menu.setTitle(_translate("MainWindow", "功能選項"))
        self.menu_2.setTitle(_translate("MainWindow", "系統"))
        self.action1.setText(_translate("MainWindow", "修改PDF成品大小及位置"))
        self.action22.setText(_translate("MainWindow", "修改輸出PDF路徑"))
        self.action2.setText(_translate("MainWindow", "修改右邊和下邊出血位"))
        self.action2.setIconText(_translate("MainWindow", "修改右邊和下邊出血位"))
        self.action_quit.setText(_translate("MainWindow", "退出"))
        self.action_dele.setText(_translate("MainWindow", "刪除行"))
        self.PDFart.setText(_translate("MainWindow", "PDF稿件只有一面"))
import resourcefile_rc
