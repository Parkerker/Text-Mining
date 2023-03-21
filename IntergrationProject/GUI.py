# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'GUI.ui'
#
# Created by: PyQt5 UI code generator 5.13.2
#
# WARNING! All changes made in this file will be lost!

import sys
import time
import threading
import ExcelAccess as EA
import pickle
import nltk
from nltk import word_tokenize
from nltk import sent_tokenize
from nltkprocessobj import *
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QMainWindow, QDialog, QProgressDialog, QMessageBox
from PyQt5.QtCore import Qt, QObject, pyqtSignal, pyqtSlot, QTimer
from qroundprogressbar import QRoundProgressBar
nltk.download('stopwords')
classifierPath = "object/classifier.pickle"
dataSetPath = "object/dataSet.pickle"
FontType_BIG5 = '微軟正黑體'
ProjectName_text = '利用文字探勘與分類器技術建置社群推文檢索工具'         #'Tourism Text Exploration' ,'Community Tweet Retrieval Tool'
SV_Title_text = 'Ettoday旅遊雲'
SV_Content_text = '《Ettoday旅遊雲》是一個台灣多媒介的網路資源'#是一個美國的多語言網路傳媒。也是提供給我們訓練器訓練新聞分類的主要來源，我們從本網站蒐集了26個類別以及20萬多篇的新聞文章，之後經過討論，我們將26個類別又以9個大類別概括之，點擊下面的檢視按鈕即可檢視全部資料。
PDV_tfidfIntrod_text = '一種用於資訊檢索與文字挖掘的常用加權技術。TF – IDF 是一種統計方法，用以評估一字詞對於一個檔案集或一個語料庫中的其中一份檔案的重要程度。字詞的重要性隨著它在檔案中出現的次數成正比增加，但同時會隨著它在語料庫中出現的頻率成反比下降。'
PDV_tfidfContri_text='第一步，我們取出九個類別的單獨類別對其他八個類別做TF – IDF計算，並把結果儲存下來，點擊下方按鈕可檢視結果，最後搭配輸入的文章與預測類別內的所有新聞做TF – IDF計算。'
WA_Title_text = 'Word Relational Vector'

accuracyDict = {"economy": 57,
                "education": 94,
                "environment": 77,
                "health": 74,
                "politics": 82}


class Ui_MainWindow(QMainWindow):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.setFixedSize(1600, 900)
        #///////MainWindow背景顏色///////
        MainWindow.setAutoFillBackground(False)      
        MainWindow.setStyleSheet(
            "background-color: rgba(255, 251, 255, 255)")     
        #///////////////////////////////
        self.CentralWidget = QtWidgets.QWidget(MainWindow)
        self.CentralWidget.setObjectName("CentralWidget")
        self.VerticalLayoutWidget = QtWidgets.QWidget(self.CentralWidget)
        self.VerticalLayoutWidget.setGeometry(QtCore.QRect(0, 0, 1600, 900))
        self.VerticalLayoutWidget.setObjectName("VerticalLayoutWidget")
        self.MainLayout = QtWidgets.QVBoxLayout(self.VerticalLayoutWidget)
        self.MainLayout.setContentsMargins(0, 0, 0, 0)
        self.MainLayout.setSpacing(0)
        self.MainLayout.setObjectName("MainLayout")

        

        # ************************Head Design************************
        self.HeadLayout = QtWidgets.QFrame(self.VerticalLayoutWidget)
        sizePolicy = QtWidgets.QSizePolicy(
            QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(
            self.HeadLayout.sizePolicy().hasHeightForWidth())
        self.HeadLayout.setSizePolicy(sizePolicy)
        self.HeadLayout.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.HeadLayout.setFrameShadow(QtWidgets.QFrame.Raised)
        self.HeadLayout.setObjectName("HeadLayout")
        self.HeadIcon = QtWidgets.QLabel(self.HeadLayout)
        self.HeadIcon.setGeometry(QtCore.QRect(18, 14, 250, 120)) #30 20 200 100    #封面LOGO
        self.HeadIcon.setPixmap(QtGui.QPixmap("imgs/alpha.PNG"))
        self.HeadIcon.setScaledContents(True)
        #標題文字 顏色 大小
        self.ProjectName = self.TextLabel(
            ProjectName_text, 'Bahnschrift SemiCondensed', 32, '#FFBF00', self.HeadLayout)  #black #FFBF00琥珀色
        self.ProjectName.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.ProjectName.setGeometry(QtCore.QRect(300, 0, 1000, 160))        ##(250, 0, 950, 160)
        self.MainLayout.addWidget(self.HeadLayout)
        # *********************End of Head Design*********************

        self.SubLayout = QtWidgets.QHBoxLayout()
        self.SubLayout.setSpacing(0)
        self.SubLayout.setObjectName("SubLayout")
        self.FunctionsWidget = QtWidgets.QWidget(self.VerticalLayoutWidget)
        self.FunctionsWidget.setEnabled(True)
        sizePolicy = QtWidgets.QSizePolicy(
            QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(1)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(
            self.FunctionsWidget.sizePolicy().hasHeightForWidth())
        self.FunctionsWidget.setSizePolicy(sizePolicy)
        self.FunctionsWidget.setStyleSheet(
            "background-color: rgba(125, 125, 125, 255)")
        self.FunctionsWidget.setObjectName("FunctionsWidget")
        self.TA_Button = QtWidgets.QPushButton(self.FunctionsWidget)
        self.TA_Button.setGeometry(QtCore.QRect(0, 0, 320, 152))
        self.TA_Button.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.TA_Button.setMouseTracking(True)
        self.TA_Button.setText("")
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap("imgs/AnalyzerButton.png"),                #初始頁面
                       QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.TA_Button.setIcon(icon)
        self.TA_Button.setIconSize(QtCore.QSize(310, 150))
        self.TA_Button.setObjectName("TA_Button")
        self.SV_Button = QtWidgets.QPushButton(self.FunctionsWidget)
        self.SV_Button.setGeometry(QtCore.QRect(0, 152, 320, 152))
        self.SV_Button.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.SV_Button.setMouseTracking(True)
        self.SV_Button.setText("")
        icon1 = QtGui.QIcon()
        icon1.addPixmap(QtGui.QPixmap("imgs/IDLESourceButton.png"),
                        QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.SV_Button.setIcon(icon1)
        self.SV_Button.setIconSize(QtCore.QSize(310, 150))
        self.SV_Button.setObjectName("SV_Button")
        self.PDV_Button = QtWidgets.QPushButton(self.FunctionsWidget)
        self.PDV_Button.setGeometry(QtCore.QRect(0, 304, 320, 152))
        self.PDV_Button.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.PDV_Button.setMouseTracking(True)
        self.PDV_Button.setText("")
        icon2 = QtGui.QIcon()
        icon2.addPixmap(QtGui.QPixmap("imgs/IDLEPreDataButton.png"),
                        QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.PDV_Button.setIcon(icon2)
        self.PDV_Button.setIconSize(QtCore.QSize(310, 150))
        self.PDV_Button.setObjectName("PDV_Button")
        self.SubLayout.addWidget(self.FunctionsWidget)

        # ************************功能框架比例設置************************
        sizePolicy = QtWidgets.QSizePolicy(
            QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(4)
        sizePolicy.setVerticalStretch(0)

        # ******************************文字分析框架******************************

        self.TA_TabFrame = QtWidgets.QTabWidget(self.VerticalLayoutWidget)
        sizePolicy.setHeightForWidth(
            self.TA_TabFrame.sizePolicy().hasHeightForWidth())
        self.TA_TabFrame.setSizePolicy(sizePolicy)
        self.TA_TabFrame.setAutoFillBackground(True)
        self.TA_TabFrame.setObjectName("TA_Frame")

        # ************************NEWS Tabs************************
        self.TA_News = QtWidgets.QWidget()
        self.TA_News.setAutoFillBackground(True)
        self.TA_News.setObjectName("TA_News")

        # THE following para is for subthread

        self.IsLoadingDone = False
        self.message_obj = Message()
        # end of subthread para

        self.Line = QtWidgets.QFrame(self.TA_News)
        self.Line.setGeometry(QtCore.QRect(50, 110, 230, 20))
        self.Line.setFrameShape(QtWidgets.QFrame.HLine)
        self.Line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.Line.setObjectName("Line")
        self.VerticleLine = QtWidgets.QFrame(self.TA_News)
        self.VerticleLine.setGeometry(QtCore.QRect(600, 20, 20, 680))
        self.VerticleLine.setFrameShape(QtWidgets.QFrame.VLine)
        self.VerticleLine.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.VerticleLine.setObjectName("VerticleLine")
        self.showAccuracy = self.TextLabel(
            '準確度百分比-----', '微軟正黑體', 26, 'violet', self.TA_News)    #With Accuracy ----------   鈷藍色#0047AB
        self.showAccuracy.setGeometry(QtCore.QRect(660, 110, 570, 60))
        self.showAccuracy.setObjectName("showAccuracy")
        self.showCate = self.TextLabel(
            '類別預測:', '微軟正黑體', 26, 'violet', self.TA_News) #Predicted Category:     violet紫羅蘭
        self.showCate.setGeometry(QtCore.QRect(660, 20, 500, 40))
        self.showCate.setObjectName("showCate")
        self.showTFIDF = QtWidgets.QTableWidget(100, 2, self.TA_News)
        self.showTFIDF.setGeometry(QtCore.QRect(660, 220, 546, 470))
        nameHeader = QtWidgets.QTableWidgetItem()
        nameHeader.setText('Name')
        self.showTFIDF.setHorizontalHeaderItem(0, nameHeader)
        valueHeader = QtWidgets.QTableWidgetItem()
        valueHeader.setText('TF-IDF Value')
        self.showTFIDF.setHorizontalHeaderItem(1, valueHeader)
        self.showTFIDF.setColumnWidth(0, 250)
        self.showTFIDF.setColumnWidth(1, 250)
        self.showTFIDF.setObjectName("showTFIDF")
        #********************************************************test textedit
        self.FileTextEdit = QtWidgets.QTextEdit(self.TA_News)
        self.FileTextEdit.setGeometry(QtCore.QRect(50, 25, 250, 30))    #50 50 250 30
        self.FileTextEdit.setFont(QtGui.QFont("Arial", 12))
        self.FileTextEdit.setText("Input File")
        self.FileTextEdit.setEnabled(False)
        self.FileTextEdit.setObjectName("FileTextEdit")
        #********************************************************
        #********************************************************test2
        self.FileTextEdit2 = QtWidgets.QTextEdit(self.TA_News)
        self.FileTextEdit2.setGeometry(QtCore.QRect(50, 73, 250, 30))
        self.FileTextEdit2.setFont(QtGui.QFont("Arial", 12))
        self.FileTextEdit2.setText("Input File")
        self.FileTextEdit2.setEnabled(False)
        self.FileTextEdit2.setObjectName("FileTextEdit2")
        #********************************************************
        self.OR_Line = QtWidgets.QLabel(self.TA_News)
        self.OR_Line.setGeometry(QtCore.QRect(290, 105, 30, 30))
        self.OR_Line.setObjectName("OR_Line")
        self.Line_2 = QtWidgets.QFrame(self.TA_News)
        self.Line_2.setGeometry(QtCore.QRect(330, 110, 230, 20))
        self.Line_2.setFrameShape(QtWidgets.QFrame.HLine)
        self.Line_2.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.Line_2.setAutoFillBackground(False)
        self.Line_2.setObjectName("Line_2")
        #********************************************************************
        self.bt_ChooseFile = QtWidgets.QPushButton(self.TA_News)    # 選擇檔案
        self.bt_ChooseFile.setFont(QtGui.QFont(FontType_BIG5, 12))
        self.bt_ChooseFile.setGeometry(QtCore.QRect(310, 25, 101, 30))  #310 50 101 30
        self.bt_ChooseFile.setCursor(
            QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.bt_ChooseFile.setObjectName("bt_ChooseFile")
        self.bt_ChooseFile.clicked.connect(
            self.TA_bt_ChooseFile_on_pushButton_clicked)
        #/////////////////////////////////////////////////////////////////////
        self.bt_ChooseFile2 = QtWidgets.QPushButton(self.TA_News)    # 選擇檔案2
        self.bt_ChooseFile2.setFont(QtGui.QFont(FontType_BIG5, 12))
        self.bt_ChooseFile2.setGeometry(QtCore.QRect(310, 73, 101, 30))
        self.bt_ChooseFile2.setCursor(
            QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.bt_ChooseFile2.setObjectName("bt_ChooseFile2")
        self.bt_ChooseFile2.clicked.connect(
            self.TA_bt_ChooseFile_on_pushButton_clicked)
        #*///////////////////////////////////////////////////////////////////
        self.TextEditor = QtWidgets.QTextEdit(self.TA_News)
        self.TextEditor.setGeometry(QtCore.QRect(50, 160, 510, 450))    #50 160 510 450
        self.TextEditor.setObjectName("TextEditor")
        #********************************************上傳檔案
        self.bt_Upload_Analyze = self.ViewButton(
            'imgs/TA/upload_analyze.png', 'imgs/TA/upload_analyze.png', self.TA_News)
        self.bt_Upload_Analyze.setEnabled(True)
        self.bt_Upload_Analyze.setGeometry(QtCore.QRect(440, 15, 77, 48))  #420 20 140 60       #(x,y,長,寬)
        self.bt_Upload_Analyze.setIconSize(QtCore.QSize(77, 48))       #7:3   146 60 ,r2:210 90    r3 161 69
        self.bt_Upload_Analyze.setAutoRepeat(False)
        self.bt_Upload_Analyze.setAutoExclusive(False)
        self.bt_Upload_Analyze.setAutoDefault(False)
        self.bt_Upload_Analyze.clicked.connect(
            self.bt_Upload_Analyze_on_pushButton_clicked)
        #*********************************************
        #********************************************上傳檔案2
        self.bt_Upload_Analyze2 = self.ViewButton(
            'imgs/TA/upload_analyze.png', 'imgs/TA/upload_analyze.png', self.TA_News)
        self.bt_Upload_Analyze2.setEnabled(True)
        self.bt_Upload_Analyze2.setGeometry(QtCore.QRect(440, 65, 77, 48))  #420 20 140 60
        self.bt_Upload_Analyze2.setIconSize(QtCore.QSize(77, 48))       #146 60 ,r2:210 90     r3 161 69
        self.bt_Upload_Analyze2.setAutoRepeat(False)
        self.bt_Upload_Analyze2.setAutoExclusive(False)
        self.bt_Upload_Analyze2.setAutoDefault(False)
        self.bt_Upload_Analyze2.clicked.connect(
            self.bt_Upload_Analyze2_on_pushButton_clicked)
        #*********************************************
        self.bt_Send = self.ViewButton(
            'imgs/TA/send.png', 'imgs/TA/send.png', self.TA_News)
        self.bt_Send.setEnabled(True)
        self.bt_Send.setGeometry(QtCore.QRect(230, 630, 210, 90))   ##230 630 140 60
        self.bt_Send.setIconSize(QtCore.QSize(322, 138))     ##146 60
        self.bt_Send.setAutoRepeat(False)
        self.bt_Send.setAutoExclusive(False)
        self.bt_Send.setAutoDefault(False)
        self.bt_Send.clicked.connect(self.bt_Send_on_pushButton_clicked)
        self.CircleProgress = QRoundProgressBar(self.TA_News)
        self.CircleProgress.setGeometry(950, 90, 100, 100)
        self.CircleProgress.setBarStyle(QRoundProgressBar.BarStyle.DONUT)
        palette = QtGui.QPalette()
        brush = QtGui.QBrush(QtGui.QColor(34, 177, 76))
        brush.setStyle(Qt.SolidPattern)
        palette.setBrush(QtGui.QPalette.Active,
                         QtGui.QPalette.Highlight, brush)
        self.CircleProgress.setPalette(palette)
        self.CircleProgress.setValue(0)
        self.TA_TabFrame.addTab(self.TA_News, "")
        # ******************End of NEWS Tabs******************

        # ************************Word Vector Tabs************************
        self.TA_WordVector = QtWidgets.QWidget()
        self.TA_WordVector.setAutoFillBackground(True)
        self.TA_WordVector.setObjectName("TA_Pubmad")

        self.WV_TextEdit = QtWidgets.QTextEdit(self.TA_WordVector)
        self.WV_TextEdit.setGeometry(QtCore.QRect(50, 30, 500, 280))
        self.WV_TextEdit.setFont(QtGui.QFont("Arial", 12))
        self.WV_TextEdit.setObjectName("FileTextEdit")

        self.wordArrNo01 = self.TextLabel(
            '01.', FontType_BIG5, 12, 'black', self.TA_WordVector)
        self.wordArrNo01.setGeometry(600, 30, 30, 30)
        self.wordArrNo02 = self.TextLabel(
            '02.', FontType_BIG5, 12, 'black', self.TA_WordVector)
        self.wordArrNo02.setGeometry(600, 92, 30, 30)
        self.wordArrNo03 = self.TextLabel(
            '03.', FontType_BIG5, 12, 'black', self.TA_WordVector)
        self.wordArrNo03.setGeometry(600, 154, 30, 30)
        self.wordArrNo04 = self.TextLabel(
            '04.', FontType_BIG5, 12, 'black', self.TA_WordVector)
        self.wordArrNo04.setGeometry(600, 216, 30, 30)
        self.wordArrNo05 = self.TextLabel(
            '05.', FontType_BIG5, 12, 'black', self.TA_WordVector)
        self.wordArrNo05.setGeometry(600, 278, 30, 30)
        self.wordArrNo06 = self.TextLabel(
            '06.', FontType_BIG5, 12, 'black', self.TA_WordVector)
        self.wordArrNo06.setGeometry(840, 30, 30, 30)
        self.wordArrNo07 = self.TextLabel(
            '07.', FontType_BIG5, 12, 'black', self.TA_WordVector)
        self.wordArrNo07.setGeometry(840, 92, 30, 30)
        self.wordArrNo08 = self.TextLabel(
            '08.', FontType_BIG5, 12, 'black', self.TA_WordVector)
        self.wordArrNo08.setGeometry(840, 154, 30, 30)
        self.wordArrNo09 = self.TextLabel(
            '09.', FontType_BIG5, 12, 'black', self.TA_WordVector)
        self.wordArrNo09.setGeometry(840, 216, 30, 30)
        self.wordArrNo10 = self.TextLabel(
            '10.', FontType_BIG5, 12, 'black', self.TA_WordVector)
        self.wordArrNo10.setGeometry(840, 278, 30, 30)

        self.wordArrField0 = self.LineEdit('', 630, 30, self.TA_WordVector)
        self.wordArrField1 = self.LineEdit('', 630, 92, self.TA_WordVector)
        self.wordArrField2 = self.LineEdit('', 630, 154, self.TA_WordVector)
        self.wordArrField3 = self.LineEdit('', 630, 216, self.TA_WordVector)
        self.wordArrField4 = self.LineEdit('', 630, 278, self.TA_WordVector)
        self.wordArrField5 = self.LineEdit('', 870, 30, self.TA_WordVector)
        self.wordArrField6 = self.LineEdit('', 870, 92, self.TA_WordVector)
        self.wordArrField7 = self.LineEdit('', 870, 154, self.TA_WordVector)
        self.wordArrField8 = self.LineEdit('', 870, 216, self.TA_WordVector)
        self.wordArrField9 = self.LineEdit('', 870, 278, self.TA_WordVector)

        self.WV_bt_Send = self.ViewButton(
            'imgs/TA/send.png', 'imgs/TA/send_Highlight.png', self.TA_WordVector)
        self.WV_bt_Send.setEnabled(True)
        self.WV_bt_Send.setGeometry(QtCore.QRect(1090, 250, 210, 90))
        self.WV_bt_Send.setSizePolicy(sizePolicy)
        self.WV_bt_Send.setIconSize(QtCore.QSize(322, 138))
        self.WV_bt_Send.setAutoRepeat(False)
        self.WV_bt_Send.setAutoExclusive(False)
        self.WV_bt_Send.setAutoDefault(False)
        self.WV_bt_Send.clicked.connect(self.TA_WV_bt_Send_clicked)

        self.WV_Line = QtWidgets.QFrame(self.TA_WordVector)
        self.WV_Line.setGeometry(QtCore.QRect(50, 329, 1178, 20))
        self.WV_Line.setFrameShape(QtWidgets.QFrame.HLine)
        self.WV_Line.setFrameShadow(QtWidgets.QFrame.Sunken)

        self.WA_TableFrame = QtWidgets.QTableWidget(11, 11, self.TA_WordVector)
        self.WA_TableFrame.setGeometry(QtCore.QRect(49, 368, 1182, 332))
        self.WA_TableFrame.horizontalHeader().setVisible(False)
        self.WA_TableFrame.verticalHeader().setVisible(False)
        for i in range(0, 10):
            headerTextCol = QtWidgets.QTableWidgetItem('Not set')
            headerTextCol.setTextAlignment(Qt.AlignCenter)
            headerTextRow=QtWidgets.QTableWidgetItem('Not set')
            headerTextRow.setTextAlignment(Qt.AlignCenter)
            self.WA_TableFrame.setItem(0, i+1, headerTextCol)
            self.WA_TableFrame.setItem(i+1, 0, headerTextRow)
            self.WA_TableFrame.setColumnWidth(i, 108)
            self.WA_TableFrame.setRowHeight(i, 30)
        self.WA_TableFrame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.WA_TableFrame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.WA_TableFrame.setObjectName("WA_TableFrame")

        self.TA_TabFrame.addTab(self.TA_WordVector, "")

        # ******************End of Word Vector Tabs******************

        # ******************End of 文字分析框架******************

        # ************************來源檢視框架************************
        self.SV_Frame = QtWidgets.QFrame(self.VerticalLayoutWidget)
        sizePolicy.setHeightForWidth(
            self.SV_Frame.sizePolicy().hasHeightForWidth())
        self.SV_Frame.setSizePolicy(sizePolicy)
        self.SV_Frame.setAutoFillBackground(False)              #來源檢視 背景或許
        self.SV_Frame.setStyleSheet(
            "background-color: rgba(0, 0, 0, 255)")
        self.SV_Frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.SV_Frame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.SV_Frame.setObjectName("SV_Frame")
        self.SV_Frame.setVisible(False)
        self.SV_News_Cover = QtWidgets.QLabel(self.SV_Frame)
        self.SV_News_Cover.setGeometry(QtCore.QRect(510, 70, 277, 160))    #(290, -80, 700, 390)
        self.SV_News_Cover.setPixmap(
            QtGui.QPixmap("imgs/SV/huffpost_logo.png"))
        self.SV_News_Cover.setScaledContents(True)
        self.SV_News_Cover.setObjectName("SV_News_Cover")
        self.SV_News_Cover.raise_()

        self.SV_Title = self.TextLabel(
            SV_Title_text, FontType_BIG5, 48, 'white', self.SV_Frame)
        self.SV_Title.setGeometry(140, 240, 1000, 64)
        self.SV_Title.setAlignment(Qt.AlignCenter)
        self.SV_Content = self.TextLabel(
            SV_Content_text, FontType_BIG5, 18, 'white', self.SV_Frame)
        self.SV_Content.setWordWrap(True)  # Auto next line
        self.SV_Content.setGeometry(240, 380, 800, 300)
        self.SV_Content.setAlignment(Qt.AlignTop)

        self.SV_btNews = self.ViewButton(
            "imgs/SV/btView.png", "imgs/SV/btView_Highlight.png", self.SV_Frame)
        self.SV_btNews.setGeometry(QtCore.QRect(440, 580, 400, 121))
        self.SV_btNews.setIconSize(QtCore.QSize(400, 121))
        self.SV_btNews.setObjectName("SV_btNews")
        self.SV_btNews.clicked.connect(self.SV_News_on_pushButton_clicked)

        # ******************End of 來源檢視框架******************

        # ************************資料前處理框架************************
        self.PDV_Frame = QtWidgets.QFrame(self.VerticalLayoutWidget)
        sizePolicy.setHeightForWidth(
            self.PDV_Frame.sizePolicy().hasHeightForWidth())
        self.PDV_Frame.setSizePolicy(sizePolicy)
        self.PDV_Frame.setAutoFillBackground(False)
        self.PDV_Frame.setStyleSheet(
            "background-color: rgba(240, 240, 240, 255)")
        self.PDV_Frame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.PDV_Frame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.PDV_Frame.setObjectName("PDV_Frame")
        self.PDV_Frame.setVisible(False)

        self.PDV_TitlePic=QtWidgets.QLabel(self.PDV_Frame)
        self.PDV_TitlePic.setGeometry(50,5,402,194)
        self.PDV_TitlePic.setPixmap(QtGui.QPixmap("imgs/PDV/Title.png"))
        self.PDV_DefPic=QtWidgets.QLabel(self.PDV_Frame)
        self.PDV_DefPic.setGeometry(105,180,233,70)
        self.PDV_DefPic.setPixmap(QtGui.QPixmap("imgs/PDV/def.png"))
        self.PDV_DefPic.setScaledContents(True)
        self.PDV_ConPic=QtWidgets.QLabel(self.PDV_Frame)
        self.PDV_ConPic.setGeometry(650,180,457,70)
        self.PDV_ConPic.setPixmap(QtGui.QPixmap("imgs/PDV/contribute.png"))
        self.PDV_ConPic.setScaledContents(True)

        self.PDV_tfidfIntro=QtWidgets.QLabel(self.PDV_Frame)
        self.PDV_tfidfIntro.setAlignment(Qt.AlignTop)
        self.PDV_tfidfIntro.setWordWrap(True)
        self.PDV_tfidfIntro.setGeometry(110,280,450,420)
        self.PDV_tfidfContri=QtWidgets.QLabel(self.PDV_Frame)
        self.PDV_tfidfContri.setAlignment(Qt.AlignTop)
        self.PDV_tfidfContri.setWordWrap(True)
        self.PDV_tfidfContri.setGeometry(655,280,500,250)
        

        self.PDV_btNews = self.ViewButton(
            "imgs/PDV/btView.png", "imgs/PDV/btView_Highlight.png", self.PDV_Frame)
        self.PDV_btNews.setGeometry(QtCore.QRect(705, 560, 400, 121))
        self.PDV_btNews.setIconSize(QtCore.QSize(400, 121))
        self.PDV_btNews.clicked.connect(self.on_pushButton_clicked)
        # ******************End of 資料前處理框架******************

        self.SubLayout.addWidget(self.TA_TabFrame)
        self.SubLayout.addWidget(self.SV_Frame)
        self.SubLayout.addWidget(self.PDV_Frame)
        self.MainLayout.addLayout(self.SubLayout)
        self.MainLayout.setStretch(0, 1)
        self.MainLayout.setStretch(1, 5)
        MainWindow.setCentralWidget(self.CentralWidget)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

        self.TA_Button.clicked.connect(self.selectTA)
        self.SV_Button.clicked.connect(self.selectSV)
        self.PDV_Button.clicked.connect(self.selectPDV)

    def loadToolFile(self):
        with open(classifierPath, 'rb') as file:
            self.classifier = pickle.load(file)

        with open(dataSetPath, 'rb') as file:
            self.dataSet = pickle.load(file)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Text Analyzer"))
        self.PDV_tfidfIntro.setText(_translate(
            "MainWindow", "<html><head/><body><p style=\"line-height:2;\"><span style=\"font-size:18pt; font-family:{:s};\">{:s}</span></p></body></html>".format(FontType_BIG5, PDV_tfidfIntrod_text)))
        self.PDV_tfidfContri.setText(_translate(
            "MainWindow", "<html><head/><body><p style=\"line-height:2;\"><span style=\"font-size:18pt; font-family:{:s};\">{:s}</span></p></body></html>".format(FontType_BIG5, PDV_tfidfContri_text)))
        self.OR_Line.setText(_translate(
            "MainWindow", "<html><head/><body><p align=\"center\"><span style=\" font-size:16pt; font-weight:bold; font-family:{:s}; font-weight:600;\">或</span></p></body></html>".format(FontType_BIG5)))
        self.bt_ChooseFile.setText(_translate("MainWindow", "選擇檔案"))
        self.bt_ChooseFile2.setText(_translate("MainWindow", "選擇影片"))   #更改按鈕文字
        self.TA_TabFrame.setTabText(self.TA_TabFrame.indexOf(
            self.TA_News), _translate("MainWindow", "首頁"))                #標籤文字   #News
        self.TA_TabFrame.setTabText(self.TA_TabFrame.indexOf(
            self.TA_WordVector), _translate("MainWindow", "自訂測資"))   #Word Vector

    # **************************************************
    # -------------------首頁控制按鈕--------------------
    # **************************************************
    def selectTA(self):
        icon1 = QtGui.QIcon()
        icon1.addPixmap(QtGui.QPixmap("imgs/AnalyzerButton.png"),
                        QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.TA_Button.setIcon(icon1)
        icon2 = QtGui.QIcon()
        icon2.addPixmap(QtGui.QPixmap("imgs/IDLESourceButton.png"),
                        QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.SV_Button.setIcon(icon2)
        icon3 = QtGui.QIcon()
        icon3.addPixmap(QtGui.QPixmap("imgs/IDLEPreDataButton.png"),
                        QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.PDV_Button.setIcon(icon3)
        self.TA_TabFrame.setVisible(True)
        self.SV_Frame.setVisible(False)
        self.PDV_Frame.setVisible(False)

    def selectSV(self):
        icon1 = QtGui.QIcon()
        icon1.addPixmap(QtGui.QPixmap("imgs/IDLEAnalyzerButton.png"),
                        QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.TA_Button.setIcon(icon1)
        icon2 = QtGui.QIcon()
        icon2.addPixmap(QtGui.QPixmap("imgs/SourceButton.png"),
                        QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.SV_Button.setIcon(icon2)
        icon3 = QtGui.QIcon()
        icon3.addPixmap(QtGui.QPixmap("imgs/IDLEPreDataButton.png"),
                        QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.PDV_Button.setIcon(icon3)
        self.TA_TabFrame.setVisible(False)
        self.SV_Frame.setVisible(True)
        self.PDV_Frame.setVisible(False)

    def selectPDV(self):
        icon1 = QtGui.QIcon()
        icon1.addPixmap(QtGui.QPixmap("imgs/IDLEAnalyzerButton.png"),
                        QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.TA_Button.setIcon(icon1)
        icon2 = QtGui.QIcon()
        icon2.addPixmap(QtGui.QPixmap("imgs/IDLESourceButton.png"),
                        QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.SV_Button.setIcon(icon2)
        icon3 = QtGui.QIcon()
        icon3.addPixmap(QtGui.QPixmap("imgs/PreDataButton.png"),
                        QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.PDV_Button.setIcon(icon3)
        self.TA_TabFrame.setVisible(False)
        self.SV_Frame.setVisible(False)
        self.PDV_Frame.setVisible(True)

    def on_pushButton_clicked(self):
        self.pdv_NewsWindow = PDV_NewsWindow()
        self.pdv_NewsWindow.show()

    def SV_News_on_pushButton_clicked(self):
        self.sv_NewsWindows = SV_NewsWindows()
        self.sv_NewsWindows.show()

    # 選擇檔案按鈕觸發
    def TA_bt_ChooseFile_on_pushButton_clicked(self):
        self.notepad, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Open Text File", r"/home/user/", "(*.txt)")
        Filename = QtCore.QFileInfo(self.notepad).fileName()
        self.FileTextEdit.setText(Filename)

    # 上傳分析按鈕觸發
    def bt_Upload_Analyze_on_pushButton_clicked(self):
        try:
            f = open(self.notepad, "r", encoding="utf-8")
        except FileNotFoundError as identifier:
            if QtCore.QFileInfo(self.notepad).fileName()=='':
                QMessageBox.critical(
                self, 'Error', 'Please choose a text file.', QMessageBox.Ok)
            else:
                QMessageBox.critical (
                    self, 'Error', 'Article must contain more than 50 characters at least.', QMessageBox.Ok)
            return
        except AttributeError as identifier:
            QMessageBox.critical(
                self, 'Error', 'Please choose a text file.', QMessageBox.Ok)
            return

        text = str(f.read())
        f.close()
        if len(text) < 50:
            QMessageBox.warning(
                self, 'Warning', 'Article must contain more than 50 characters at least.', QMessageBox.Ok)
        else:
            self.progress_indicator = QProgressDialog(
                self, QtCore.Qt.WindowStaysOnTopHint)
            self.progress_indicator.setWindowModality(Qt.WindowModal)
            self.progress_indicator.setWindowTitle('Loading...')
            self.progress_indicator.setRange(0, 0)
            self.progress_indicator.setAttribute(Qt.WA_DeleteOnClose)
            self.message_obj.finished.connect(
                self.progress_indicator.close, Qt.QueuedConnection
            )
            self.progress_indicator.show()

            feaText = self.dataSet.setFeature(word_tokenize(text))
            cate = self.classifier.clsfier.classify(feaText)
            self.CircleProgress.setValue(accuracyDict[cate])
            self.showCate.setText('類別預測:' + cate)    #Predicted Category: 

            BarThread = threading.Thread(target=self.Loading_Thread,
                                         args=(self.message_obj, ))
            BarThread.start()
            AnaThread = threading.Thread(target=self.Analyze, args=(text, cate, ))
            AnaThread.start()

    # 上傳分析按鈕觸發 bt_Upload_Analyze2
    def bt_Upload_Analyze2_on_pushButton_clicked(self):
        try:
            f = open(self.notepad, "r", encoding="utf-8")
        except FileNotFoundError as identifier:
            if QtCore.QFileInfo(self.notepad).fileName()=='':
                QMessageBox.critical(
                self, 'Error', 'Please choose a text file.', QMessageBox.Ok)
            else:
                QMessageBox.critical (
                    self, 'Error', 'Article must contain more than 50 characters at least.', QMessageBox.Ok)
            return
        except AttributeError as identifier:
            QMessageBox.critical(
                self, 'Error', 'Please choose a text file.', QMessageBox.Ok)
            return

        text = str(f.read())
        f.close()
        if len(text) < 50:
            QMessageBox.warning(
                self, 'Warning', 'Article must contain more than 50 characters at least.', QMessageBox.Ok)
        else:
            self.progress_indicator = QProgressDialog(
                self, QtCore.Qt.WindowStaysOnTopHint)
            self.progress_indicator.setWindowModality(Qt.WindowModal)
            self.progress_indicator.setWindowTitle('Loading...')
            self.progress_indicator.setRange(0, 0)
            self.progress_indicator.setAttribute(Qt.WA_DeleteOnClose)
            self.message_obj.finished.connect(
                self.progress_indicator.close, Qt.QueuedConnection
            )
            self.progress_indicator.show()

            feaText = self.dataSet.setFeature(word_tokenize(text))
            cate = self.classifier.clsfier.classify(feaText)
            self.CircleProgress.setValue(accuracyDict[cate])
            self.showCate.setText('類別預測:' + cate)    #Predicted Category:

            BarThread = threading.Thread(target=self.Loading_Thread,
                                         args=(self.message_obj, ))
            BarThread.start()
            AnaThread = threading.Thread(target=self.Analyze, args=(text, cate, ))
            AnaThread.start()

    def bt_Send_on_pushButton_clicked(self):
        text = self.TextEditor.toPlainText()
        if len(text) < 50:
            QMessageBox.warning(
                self, 'Warning', 'Article must contain more than 50 characters at least.', QMessageBox.Ok)
        else:
            self.progress_indicator = QProgressDialog(
                self, QtCore.Qt.WindowStaysOnTopHint)
            self.progress_indicator.setWindowModality(Qt.WindowModal)
            self.progress_indicator.setWindowTitle('Loading...')
            self.progress_indicator.setRange(0, 0)
            self.progress_indicator.setAttribute(Qt.WA_DeleteOnClose)
            self.message_obj.finished.connect(
                self.progress_indicator.close, Qt.QueuedConnection
            )
            self.progress_indicator.show()

            feaText = self.dataSet.setFeature(word_tokenize(text))
            cate = self.classifier.clsfier.classify(feaText)
            self.CircleProgress.setValue(accuracyDict[cate])
            self.showCate.setText('類別預測:' + cate)   #Predicted Category:

            BarThread = threading.Thread(target=self.Loading_Thread,
                                         args=(self.message_obj, ))
            BarThread.start()
            AnaThread = threading.Thread(target=self.Analyze, args=(text, cate,))
            AnaThread.start()

    def TA_WV_bt_Send_clicked(self):
        PMIcomputer = NLTKPMIcomputer()
        text = self.WV_TextEdit.toPlainText()
        if len(text) < 50:
            QMessageBox.warning(
                self, 'Warning', 'Article must contain more than 50 characters at least.', QMessageBox.Ok)
            return
        text = sent_tokenize(text.lower())

        arr = []
        if self.wordArrField0.text() != '':
            arr.append(self.wordArrField0.text())
        if self.wordArrField1.text() != '':
            arr.append(self.wordArrField1.text())
        if self.wordArrField2.text() != '':
            arr.append(self.wordArrField2.text())
        if self.wordArrField3.text() != '':
            arr.append(self.wordArrField3.text())
        if self.wordArrField4.text() != '':
            arr.append(self.wordArrField4.text())
        if self.wordArrField5.text() != '':
            arr.append(self.wordArrField5.text())
        if self.wordArrField6.text() != '':
            arr.append(self.wordArrField6.text())
        if self.wordArrField7.text() != '':
            arr.append(self.wordArrField7.text())
        if self.wordArrField8.text() != '':
            arr.append(self.wordArrField8.text())
        if self.wordArrField9.text() != '':
            arr.append(self.wordArrField9.text())
        
        arr = [word.lower() for word in arr]
        if len(arr)==0:
            QMessageBox.warning(
                self, 'Warning', 'You must fill one word in textfield at least.', QMessageBox.Ok)
            return          
     
        for i in range(0, len(arr)):
            headerTextCol = QtWidgets.QTableWidgetItem(arr[i])
            headerTextCol.setTextAlignment(Qt.AlignCenter)
            headerTextRow=QtWidgets.QTableWidgetItem(arr[i])
            headerTextRow.setTextAlignment(Qt.AlignCenter)
            self.WA_TableFrame.setItem(0, i+1, headerTextCol)
            self.WA_TableFrame.setItem(i+1, 0, headerTextRow)
            word1 = arr[i]
            PMIvalue_Standardize1=PMIcomputer.sheetPMI(text, arr[i], arr[i],logMode=False)
            #print(PMIvalue_Standardize)
            for j in range(0, len(arr)):
                word2 = arr[j]
                PMIvalue_Standardize2=PMIcomputer.sheetPMI(text, arr[j], arr[j],logMode=False)
                PMIvalue = PMIcomputer.sheetPMI(text, word1, word2,logMode=False)
                if PMIvalue_Standardize1*PMIvalue_Standardize2 != 0:
                    PMIvalue = (PMIvalue*PMIvalue)/(PMIvalue_Standardize1*PMIvalue_Standardize2)
                value = QtWidgets.QTableWidgetItem(str(round(PMIvalue, 5)))
                value.setTextAlignment(Qt.AlignCenter)
                self.WA_TableFrame.setItem(i+1, j+1, value)

        for i in range(len(arr), 10):
           headerTextCol = QtWidgets.QTableWidgetItem('Not set')
           headerTextCol.setTextAlignment(Qt.AlignCenter)
           headerTextRow = QtWidgets.QTableWidgetItem('Not set')
           headerTextRow.setTextAlignment(Qt.AlignCenter)
           self.WA_TableFrame.setItem(0, i+1, headerTextCol)
           self.WA_TableFrame.setItem(i+1, 0, headerTextRow)
           for j in range(0, 10):
                valueCol = QtWidgets.QTableWidgetItem('')
                valueCol.setTextAlignment(Qt.AlignCenter)
                valueRow = QtWidgets.QTableWidgetItem('')
                valueRow.setTextAlignment(Qt.AlignCenter)
                self.WA_TableFrame.setItem(i+1, j+1, valueCol)
                self.WA_TableFrame.setItem(j+1, i+1, valueRow)
        
        
    def Loading_Thread(self, obj):
        while(self.IsLoadingDone != True):
            time.sleep(0.1)
        self.IsLoadingDone = False
        obj.finished.emit()

    def Analyze(self, text, cate):
        MainWindow.setEnabled(False)
        fileName = "/object/"+str(cate)
        TFIDFComputer = NLTKTFIDFComputer()
        result = TFIDFComputer.TFIDF_Compute(fileName, text, cate)
        count = 0
        for data in result:
            name = QtWidgets.QTableWidgetItem(str(data[0]))
            self.showTFIDF.setItem(count, 0, name)
            tfidf = QtWidgets.QTableWidgetItem(str(data[1]))
            self.showTFIDF.setItem(count, 1, tfidf)
            count += 1
            if count == 100:
                break

        self.IsLoadingDone = True
        MainWindow.setEnabled(True)

    class TextLabel(QtWidgets.QLabel):
        def __init__(self, str, fonyType, fontSize, color, parent=None, flags=Qt.WindowFlags()):
            super().__init__(str, parent=parent, flags=flags)
            self.setFont(QtGui.QFont(fonyType, fontSize))
            self.setStyleSheet('color: {:s}'.format(color))

    class LineEdit(QtWidgets.QLineEdit):
        def __init__(self, str, x, y, parent=None):
            super().__init__(str, parent=parent)
            self.setGeometry(x, y, 180, 30)

    class ViewButton(QtWidgets.QPushButton):
        def __init__(self, pixmap, pixmap_highlight, parent=None,):
            super(QtWidgets.QPushButton, self).__init__(parent)
            self.icon_idleView = QtGui.QIcon()
            self.icon_idleView.addPixmap(QtGui.QPixmap(
                pixmap), QtGui.QIcon.  Normal, QtGui.QIcon.Off)
            self.setIcon(self.icon_idleView)
            self.icon_turnView = QtGui.QIcon()
            self.icon_turnView.addPixmap(QtGui.QPixmap(
                pixmap_highlight), QtGui.QIcon.  Normal, QtGui.QIcon.Off)
            self.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
            self.setStyleSheet("border: none;")

        def enterEvent(self, QEvent):
            self.setIcon(self.icon_turnView)

        def leaveEvent(self, QEvent):
            self.setIcon(self.icon_idleView)

    # ************************End of 主畫面控制按鈕************************


class PDV_NewsWindow(QMainWindow):
    def __init__(self,  *args, **kwargs):
        QMainWindow.__init__(self, *args, **kwargs)

        self.setFixedSize(1024, 629)
        sizePolicy = QtWidgets.QSizePolicy(
            QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        self.setStyleSheet("background-color: rgba(89, 89, 89, 1)")
        self.centralwidget = QtWidgets.QWidget(self)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayout = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout.setSpacing(0)
        self.gridLayout.setObjectName("gridLayout")

        self.HeaButton = self.ViewButton(
            "imgs/Icon/IdelHea.PNG", self.centralwidget)
        self.EcoButton = self.ViewButton(
            "imgs/Icon/IdelEco.PNG", self.centralwidget)
        self.EntButton = self.ViewButton(
            "imgs/Icon/IdelEnt.PNG", self.centralwidget)
        self.PolButton = self.ViewButton(
            "imgs/Icon/IdelPoli.PNG", self.centralwidget)
        self.OtherButton = self.ViewButton(
            "imgs/Icon/IdelOther.PNG", self.centralwidget)
        self.TechButton = self.ViewButton(
            "imgs/Icon/IdelTech.PNG", self.centralwidget)
        self.HomeButton = self.ViewButton(
            "imgs/Icon/IdelHome.PNG", self.centralwidget)
        self.EnvButton = self.ViewButton(
            "imgs/Icon/IdelEnv.PNG", self.centralwidget)
        self.EduButton = self.ViewButton(
            "imgs/Icon/IdelEdu.PNG", self.centralwidget)

        self.gridLayout.addWidget(self.HeaButton, 1, 1, 1, 1)
        self.gridLayout.addWidget(self.EcoButton, 0, 0, 1, 1)
        self.gridLayout.addWidget(self.EntButton, 0, 2, 1, 1)
        self.gridLayout.addWidget(self.PolButton, 2, 0, 1, 1)
        self.gridLayout.addWidget(self.OtherButton, 2, 2, 1, 1)
        self.gridLayout.addWidget(self.TechButton, 2, 1, 1, 1)
        self.gridLayout.addWidget(self.HomeButton, 1, 2, 1, 1)
        self.gridLayout.addWidget(self.EnvButton, 1, 0, 1, 1)
        self.gridLayout.addWidget(self.EduButton, 0, 1, 1, 1)

        # ************************************************************
        # ------------------------按鈕事件連結-------------------------
        # ************************************************************
        self.EcoButton.clicked.connect(self.selectEco,)
        self.EduButton.clicked.connect(self.selectEdu)
        self.EntButton.clicked.connect(self.selectEnt)
        self.EnvButton.clicked.connect(self.selectEnv)
        self.HeaButton.clicked.connect(self.selectHea)
        self.HomeButton.clicked.connect(self.selectHome)
        self.PolButton.clicked.connect(self.selectPol)
        self.TechButton.clicked.connect(self.selectTech)
        self.OtherButton.clicked.connect(self.selectOther)
        # **********************End of 按鈕事件連結**********************

        # ************************************************************
        # --------------------------表格設定---------------------------
        # ************************************************************
        self.TableFrame = self.NewTableFrame(
            "/Source/PD/News/economyRevise.xlsx")
        # ************************End of 表格設定************************

        self.gridLayout.addWidget(self.TableFrame, 0, 3, 3, 1)
        self.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 1024, 21))
        self.menubar.setObjectName("menubar")
        self.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        self.setStatusBar(self.statusbar)

        self.retranslateUi(self)
        QtCore.QMetaObject.connectSlotsByName(self)

    # ************************************************************
    # --------------------------按鈕事件---------------------------
    # ************************************************************

    def selectEco(self):
        self.TableFrame.setVisible(False)
        self.gridLayout.removeWidget(self.TableFrame)
        self.TableFrame = self.NewTableFrame(
            "/Source/PD/News/economyRevise.xlsx")
        self.gridLayout.addWidget(self.TableFrame, 0, 3, 3, 1)

    def selectEdu(self):
        self.TableFrame.setVisible(False)
        self.gridLayout.removeWidget(self.TableFrame)
        self.TableFrame = self.NewTableFrame(
            "/Source/PD/News/educationRevise.xlsx")
        self.gridLayout.addWidget(self.TableFrame, 0, 3, 3, 1)

    def selectEnt(self):
        self.TableFrame.setVisible(False)
        self.gridLayout.removeWidget(self.TableFrame)
        self.TableFrame = self.NewTableFrame(
            "/Source/PD/News/entertainmentRevise.xlsx")
        self.gridLayout.addWidget(self.TableFrame, 0, 3, 3, 1)

    def selectEnv(self):
        self.TableFrame.setVisible(False)
        self.gridLayout.removeWidget(self.TableFrame)
        self.TableFrame = self.NewTableFrame(
            "/Source/PD/News/environmentRevise.xlsx")
        self.gridLayout.addWidget(self.TableFrame, 0, 3, 3, 1)

    def selectHea(self):
        self.TableFrame.setVisible(False)
        self.gridLayout.removeWidget(self.TableFrame)
        self.TableFrame = self.NewTableFrame(
            "/Source/PD/News/healthRevise.xlsx")
        self.gridLayout.addWidget(self.TableFrame, 0, 3, 3, 1)

    def selectHome(self):
        self.TableFrame.setVisible(False)
        self.gridLayout.removeWidget(self.TableFrame)
        self.TableFrame = self.NewTableFrame(
            "/Source/PD/News/homeRevise.xlsx")
        self.TableFrame.setColumnWidth(0, 236)
        self.TableFrame.setColumnWidth(1, 236)
        self.gridLayout.addWidget(self.TableFrame, 0, 3, 3, 1)

    def selectPol(self):
        self.TableFrame.setVisible(False)
        self.gridLayout.removeWidget(self.TableFrame)
        self.TableFrame = self.NewTableFrame(
            "/Source/PD/News/politicsRevise.xlsx")
        self.gridLayout.addWidget(self.TableFrame, 0, 3, 3, 1)

    def selectTech(self):
        self.TableFrame.setVisible(False)
        self.gridLayout.removeWidget(self.TableFrame)
        self.TableFrame = self.NewTableFrame(
            "/Source/PD/News/techRevise.xlsx")
        self.gridLayout.addWidget(self.TableFrame, 0, 3, 3, 1)

    def selectOther(self):
        self.TableFrame.setVisible(False)
        self.gridLayout.removeWidget(self.TableFrame)
        self.TableFrame = self.NewTableFrame(
            "/Source/PD/News/otherRevise.xlsx")
        self.gridLayout.addWidget(self.TableFrame, 0, 3, 3, 1)
    # ************************End of 按鈕事件************************

    def NewTableFrame(self, Subpath):
        InitCells = EA.RetuenSingleXlsxDataFrame(Subpath)
        TableFrame = QtWidgets.QTableWidget(
            InitCells[0]-2, 2, self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(
            QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(5)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(
            TableFrame.sizePolicy().hasHeightForWidth())
        TableFrame.setSizePolicy(sizePolicy)
        TableFrame.setAutoFillBackground(False)
        TableFrame.setStyleSheet(
            "background-color: rgba(38, 38, 38, 200);border-color: yellow;")    
        TableFrame.setObjectName("TableFrame")

        # ************************Header設定************************
        item1 = QtWidgets.QTableWidgetItem('Word')
        item1.setBackground(QtGui.QColor(255, 0, 0))
        TableFrame.setHorizontalHeaderItem(0, item1)
        item2 = QtWidgets.QTableWidgetItem('TF-IDF Value')
        item2.setBackground(QtGui.QColor(255, 0, 0))
        TableFrame.setHorizontalHeaderItem(0, item1)
        TableFrame.setHorizontalHeaderItem(1, item2)
        TableFrame.setColumnWidth(0, 232)
        TableFrame.setColumnWidth(1, 233)

        # *************************cell設定*************************
        for i in range(0, InitCells[0]-2):
            name = QtWidgets.QTableWidgetItem(InitCells[1][i])          #背景顏色
            name.setBackground(QtGui.QColor("black"))
            TableFrame.setItem(i, 0, name)
            tfidf = QtWidgets.QTableWidgetItem(str(InitCells[2][i]))
            tfidf.setBackground(QtGui.QColor("#black"))                #white
            TableFrame.setItem(i, 1, tfidf)

        return TableFrame

    class ViewButton(QtWidgets.QPushButton):
        def __init__(self, pixmap, parent=None,):
            super(QtWidgets.QPushButton, self).__init__(parent)
            self.icon_idleView = QtGui.QIcon()
            self.icon_idleView.addPixmap(QtGui.QPixmap(
                pixmap), QtGui.QIcon.  Normal, QtGui.QIcon.Off)
            self.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
            self.setMouseTracking(True)
            self.setIcon(self.icon_idleView)
            self.setIconSize(QtCore.QSize(150, 190))

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate(
            "MainWindow", "Preprocessing News Data Viewer"))


class SV_NewsWindows(QMainWindow):
    def __init__(self, *args, **kwargs):
        QMainWindow.__init__(self, *args, **kwargs)

        # THE following para is for subthread
        self.IsLoadingDone = False
        self.message_obj = Message()
        self.timer = QTimer(interval=3000, timeout=self.on_timeout)
        # end of subthread para

        self.setFixedSize(1280, 600)
        sizePolicy = QtWidgets.QSizePolicy(
            QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(1)
        sizePolicy.setHeightForWidth(self.sizePolicy().hasHeightForWidth())
        self.setSizePolicy(sizePolicy)
        self.centralwidget = QtWidgets.QWidget(self)
        self.centralwidget.setObjectName("centralwidget")
        self.horizontalLayoutWidget = QtWidgets.QWidget(self.centralwidget)
        self.horizontalLayoutWidget.setGeometry(QtCore.QRect(0, 0, 1281, 581))
        self.horizontalLayoutWidget.setObjectName("horizontalLayoutWidget")
        self.HorizenLayout = QtWidgets.QHBoxLayout(self.horizontalLayoutWidget)
        self.HorizenLayout.setContentsMargins(0, 0, 0, 0)
        self.HorizenLayout.setSpacing(0)
        self.HorizenLayout.setObjectName("HorizenLayout")

        self.treeWidget = QtWidgets.QTreeWidget(self.horizontalLayoutWidget)
        sizePolicy = QtWidgets.QSizePolicy(
            QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(1)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(
            self.treeWidget.sizePolicy().hasHeightForWidth())
        self.treeWidget.setSizePolicy(sizePolicy)
        self.treeWidget.setObjectName("treeWidget")
        self.treeWidget.setFont(QtGui.QFont("Arial", 12, True))
        self.treeWidget.clicked.connect(self.ListView_clicked)
        item_0 = QtWidgets.QTreeWidgetItem(self.treeWidget)
        item_1 = QtWidgets.QTreeWidgetItem(item_0)
        item_1 = QtWidgets.QTreeWidgetItem(item_0)                      #*******************改樹類別內容*******************
        item_0 = QtWidgets.QTreeWidgetItem(self.treeWidget)
        item_1 = QtWidgets.QTreeWidgetItem(item_0)
        item_1 = QtWidgets.QTreeWidgetItem(item_0)
        item_1 = QtWidgets.QTreeWidgetItem(item_0)
        item_0 = QtWidgets.QTreeWidgetItem(self.treeWidget)
        item_1 = QtWidgets.QTreeWidgetItem(item_0)
        item_1 = QtWidgets.QTreeWidgetItem(item_0)
        item_1 = QtWidgets.QTreeWidgetItem(item_0)
        item_1 = QtWidgets.QTreeWidgetItem(item_0)
        item_0 = QtWidgets.QTreeWidgetItem(self.treeWidget)
        item_1 = QtWidgets.QTreeWidgetItem(item_0)
        item_1 = QtWidgets.QTreeWidgetItem(item_0)
        item_0 = QtWidgets.QTreeWidgetItem(self.treeWidget)
        item_1 = QtWidgets.QTreeWidgetItem(item_0)
        item_1 = QtWidgets.QTreeWidgetItem(item_0)
        item_0 = QtWidgets.QTreeWidgetItem(self.treeWidget)
        item_1 = QtWidgets.QTreeWidgetItem(item_0)
        item_1 = QtWidgets.QTreeWidgetItem(item_0)
        item_1 = QtWidgets.QTreeWidgetItem(item_0)
        item_1 = QtWidgets.QTreeWidgetItem(item_0)
        item_0 = QtWidgets.QTreeWidgetItem(self.treeWidget)
        item_1 = QtWidgets.QTreeWidgetItem(item_0)
        item_0 = QtWidgets.QTreeWidgetItem(self.treeWidget)
        item_1 = QtWidgets.QTreeWidgetItem(item_0)
        item_1 = QtWidgets.QTreeWidgetItem(item_0)
        item_0 = QtWidgets.QTreeWidgetItem(self.treeWidget)
        item_1 = QtWidgets.QTreeWidgetItem(item_0)
        item_1 = QtWidgets.QTreeWidgetItem(item_0)
        item_1 = QtWidgets.QTreeWidgetItem(item_0)
        item_1 = QtWidgets.QTreeWidgetItem(item_0)
        item_1 = QtWidgets.QTreeWidgetItem(item_0)
        item_1 = QtWidgets.QTreeWidgetItem(item_0)

        self.HorizenLayout.addWidget(self.treeWidget)

        sizePolicy = QtWidgets.QSizePolicy(
            QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(3)

        # ************************************************************
        # --------------------------表格設定---------------------------
        # ************************************************************
        self.TableFrame = self.NewTableFrame(
            "/Source/SV/News/COLLEGE.xlsx")
        self.TableFrame.setSizePolicy(sizePolicy)
        self.TableFrame.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.TableFrame.setFrameShadow(QtWidgets.QFrame.Raised)
        self.TableFrame.setObjectName("TableFrame")
        self.HorizenLayout.addWidget(self.TableFrame)
        # ************************End of 表格設定************************

        self.textBrowser = QtWidgets.QTextBrowser(self.horizontalLayoutWidget)
        self.textBrowser.setSizePolicy(sizePolicy)
        self.textBrowser.setFont(QtGui.QFont("Arial", 12))
        self.textBrowser.setText(self.InitCells[1][0])
        self.textBrowser.setObjectName("textBrowser")
        self.HorizenLayout.addWidget(self.textBrowser)

        self.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(self)
        self.statusbar.setObjectName("statusbar")
        self.setStatusBar(self.statusbar)

        self.retranslateUi(self)
        QtCore.QMetaObject.connectSlotsByName(self)
        # *************************View Source 分類**********************
    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        self.setWindowTitle(_translate("MainWindow", "News Source Viewer"))
        self.treeWidget.headerItem().setText(
            0, _translate("MainWindow", "All Categories"))
        __sortingEnabled = self.treeWidget.isSortingEnabled()
        self.treeWidget.setSortingEnabled(False)
        self.treeWidget.topLevelItem(0).setText(
            0, _translate("MainWindow", "Economy"))
        self.treeWidget.topLevelItem(0).child(0).setText(
            0, _translate("MainWindow", "Business"))
        self.treeWidget.topLevelItem(0).child(1).setText(
            0, _translate("MainWindow", "Money"))                #這裡
        self.treeWidget.topLevelItem(1).setText(
            0, _translate("MainWindow", "Education"))
        self.treeWidget.topLevelItem(1).child(0).setText(
            0, _translate("MainWindow", "College"))
        self.treeWidget.topLevelItem(1).child(1).setText(
            0, _translate("MainWindow", "Culture & Arts"))
        self.treeWidget.topLevelItem(1).child(2).setText(
            0, _translate("MainWindow", "Education"))
        self.treeWidget.topLevelItem(2).setText(
            0, _translate("MainWindow", "Entertainment"))
        self.treeWidget.topLevelItem(2).child(0).setText(
            0, _translate("MainWindow", "Entertainment"))
        self.treeWidget.topLevelItem(2).child(1).setText(
            0, _translate("MainWindow", "Food & Drink"))
        self.treeWidget.topLevelItem(2).child(2).setText(
            0, _translate("MainWindow", "Style & Beauty"))
        self.treeWidget.topLevelItem(2).child(3).setText(
            0, _translate("MainWindow", "Travel"))
        self.treeWidget.topLevelItem(3).setText(
            0, _translate("MainWindow", "Environment"))
        self.treeWidget.topLevelItem(3).child(0).setText(
            0, _translate("MainWindow", "Environment"))
        self.treeWidget.topLevelItem(3).child(1).setText(
            0, _translate("MainWindow", "Green"))
        self.treeWidget.topLevelItem(4).setText(
            0, _translate("MainWindow", "Health"))
        self.treeWidget.topLevelItem(4).child(0).setText(
            0, _translate("MainWindow", "Healthy living"))
        self.treeWidget.topLevelItem(4).child(1).setText(
            0, _translate("MainWindow", "Wellness"))
        self.treeWidget.topLevelItem(5).setText(
            0, _translate("MainWindow", "Home"))
        self.treeWidget.topLevelItem(5).child(0).setText(
            0, _translate("MainWindow", "Divorce"))
        self.treeWidget.topLevelItem(5).child(1).setText(
            0, _translate("MainWindow", "Home & Living"))
        self.treeWidget.topLevelItem(5).child(2).setText(
            0, _translate("MainWindow", "Parents"))
        self.treeWidget.topLevelItem(5).child(3).setText(
            0, _translate("MainWindow", "Weddings"))
        self.treeWidget.topLevelItem(6).setText(
            0, _translate("MainWindow", "Politics"))
        self.treeWidget.topLevelItem(6).child(0).setText(
            0, _translate("MainWindow", "Politics"))
        self.treeWidget.topLevelItem(7).setText(
            0, _translate("MainWindow", "Technology"))
        self.treeWidget.topLevelItem(7).child(0).setText(
            0, _translate("MainWindow", "Science"))
        self.treeWidget.topLevelItem(7).child(1).setText(
            0, _translate("MainWindow", "Technology"))
        self.treeWidget.topLevelItem(8).setText(
            0, _translate("MainWindow", "Other"))
        self.treeWidget.topLevelItem(8).child(0).setText(
            0, _translate("MainWindow", "Crime"))
        self.treeWidget.topLevelItem(8).child(1).setText(
            0, _translate("MainWindow", "Comedy"))
        self.treeWidget.topLevelItem(8).child(2).setText(
            0, _translate("MainWindow", "Religion"))
        self.treeWidget.topLevelItem(8).child(3).setText(
            0, _translate("MainWindow", "Sports"))
        self.treeWidget.topLevelItem(8).child(4).setText(
            0, _translate("MainWindow", "The Worldpost"))
        self.treeWidget.topLevelItem(8).child(5).setText(
            0, _translate("MainWindow", "World News"))
        self.treeWidget.setSortingEnabled(__sortingEnabled)

    def NewTableFrame(self, Subpath):
        self.InitCells = EA.RetuenSingleXlsxDataFrame(Subpath)
        TableFrame = QtWidgets.QTableWidget(
            self.InitCells[0]-2, 1, self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(
            QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy = QtWidgets.QSizePolicy(
            QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(3)
        TableFrame.setSizePolicy(sizePolicy)
        TableFrame.setAutoFillBackground(False)
        TableFrame.setObjectName("TableFrame")
        TableFrame.clicked.connect(self.TableWidget_clicked)

        # ************************Header設定************************
        item1 = QtWidgets.QTableWidgetItem('Headline')
        item1.setBackground(QtGui.QColor(255, 0, 0))
        TableFrame.setHorizontalHeaderItem(0, item1)
        if self.InitCells[0]-2 >= 10000:
            TableFrame.setColumnWidth(0, 490)
        else:
            TableFrame.setColumnWidth(0, 495)

        # *************************cell設定*************************
        for i in range(0, self.InitCells[0]-2):
            Headline = QtWidgets.QTableWidgetItem(self.InitCells[2][i])
            Headline.setBackground(QtGui.QColor("white"))
            TableFrame.setItem(i, 0, Headline)

        return TableFrame

    def Loading_Thread(self, obj):
        while(self.IsLoadingDone != True):
            time.sleep(0.1)
        self.IsLoadingDone = False
        obj.finished.emit()

    def ListView_clicked(self):
        self.progress_indicator = QProgressDialog(self)
        self.progress_indicator.setWindowModality(Qt.WindowModal)
        self.progress_indicator.setWindowTitle('Loading...')
        self.progress_indicator.setRange(0, 0)
        self.progress_indicator.setAttribute(Qt.WA_DeleteOnClose)
        self.message_obj.finished.connect(
            self.progress_indicator.close, Qt.QueuedConnection
        )
        self.progress_indicator.show()

        BarThread = threading.Thread(target=self.Loading_Thread,
                                     args=(self.message_obj, ))
        BarThread.start()
        self.timer.start()

    @pyqtSlot()
    def on_timeout(self):
        item = self.treeWidget.currentItem()
        if str(item.text(0)) == "Technology":
            string = "/Source/SV/News/TECH"+".xlsx"
        else:
            string = "/Source/SV/News/"+str(item.text(0)).upper()+".xlsx"

        self.HorizenLayout.removeWidget(self.TableFrame)
        self.HorizenLayout.removeWidget(self.textBrowser)

        try:
            self.TableFrame = self.NewTableFrame(string)
        except FileNotFoundError as identifier:
            msg = QMessageBox.warning(
                self, 'Warning', 'This category doesn\'t exist. Please select the child node of category to view since parent node may be an abstract category.', QMessageBox.Ok)
            self.IsLoadingDone = True
            self.timer.stop()
            return

        self.HorizenLayout.addWidget(self.TableFrame)
        sizePolicy = QtWidgets.QSizePolicy(
            QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(3)
        self.textBrowser = QtWidgets.QTextBrowser(self.horizontalLayoutWidget)
        self.textBrowser.setSizePolicy(sizePolicy)
        self.textBrowser.setObjectName("textBrowser")
        self.textBrowser.setFont(QtGui.QFont("Arial", 12))
        self.HorizenLayout.addWidget(self.textBrowser)

        self.IsLoadingDone = True
        self.timer.stop()

    def TableWidget_clicked(self):
        currentRow = self.TableFrame.currentRow()
        self.textBrowser.setText(self.InitCells[1][currentRow])


class Message(QObject):
    finished = pyqtSignal()


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    app.setStyle("fusion")
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    ui.loadToolFile()
    MainWindow.show()
    sys.exit(app.exec_())
