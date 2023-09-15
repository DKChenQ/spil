#!/usr/bin/python
# -*- coding: utf-8 -*

import sys, pyodbc
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtGui import QPainter,QPen,QColor, QBrush, QFont
from PyQt5.QtCore import Qt
from PyQt5.QtChart import QChart, QChartView, QLineSeries, QScatterSeries, QCategoryAxis, QValueAxis

class Ui_MainWindow(object):
    # def __init__(self):
        # super().__init__()

    def database_connect(self):
        # 連接Access資料庫
        self.conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=\\NAS27\LB-Repair-Room\02 系統資料\04_維修管理系統\QCT_DataBase.accdb;Uid=Admin;Pwd=;')
        #self.conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=D:\SPIL_DATA\QCT_DataBase.accdb;Uid=Admin;Pwd=;')
        self.cursor = self.conn.cursor()

    def closeEvent(self, event):
     # 視窗關閉時關閉資料庫連接
        self.conn.close()
        event.accept()

    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.setFixedSize(1122, 718) #鎖住視窗放大功能
       # MainWindow.resize(1122, 718)
        MainWindow.setIconSize(QtCore.QSize(32, 32))
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")

        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setGeometry(QtCore.QRect(830, 60, 131, 41))
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(12)
        self.pushButton.setFont(font)
        icon = QtGui.QIcon()
        icon.addPixmap(QtGui.QPixmap("1450991424_Magnifier.ico"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.pushButton.setIcon(icon)
        self.pushButton.setIconSize(QtCore.QSize(32, 32))
        self.pushButton.setObjectName("pushButton")

        self.groupBox = QtWidgets.QGroupBox(self.centralwidget)
        self.groupBox.setGeometry(QtCore.QRect(0, 120, 1111, 251))
        self.groupBox.setObjectName("groupBox")

        self.graphicsView = QChartView(self.groupBox)
        self.graphicsView.setGeometry(QtCore.QRect(20, 20, 1081, 221))
        self.graphicsView.setObjectName("graphicsView")
       
        self.comboBox = QtWidgets.QComboBox(self.centralwidget)
        self.comboBox.setGeometry(QtCore.QRect(200, 70, 200, 22))
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(12)
        self.comboBox.setFont(font)
        self.comboBox.setObjectName("comboBox")

        self.comboBox_1 = QtWidgets.QComboBox(self.centralwidget)
        self.comboBox_1.setGeometry(QtCore.QRect(40, 70, 151, 22))
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(12)
        self.comboBox_1.setFont(font)
        self.comboBox_1.setObjectName("comboBox")

        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(40, 40, 91, 31))
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(12)
        self.label.setFont(font)
        self.label.setObjectName("label")

        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(200, 40, 101, 31))
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(12)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")

        self.tableWidget = QtWidgets.QTableWidget(self.centralwidget)
        self.tableWidget.setGeometry(QtCore.QRect(20, 380, 1091, 301))
        self.tableWidget.setObjectName("tableWidget")
        self.tableWidget.itemSelectionChanged.connect(self.update_chart)

        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 1122, 21))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        self.database_connect()
        self.inp_data()
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

        self.pushButton.clicked.connect(self.search_database) #搜尋按鈕
        
    def retranslateUi(self, MainWindow): #部件名稱
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowIcon(QtGui.QIcon("design.ico"))
        MainWindow.setWindowTitle(_translate("MainWindow", "HEALTH RATE WEEKLY TREND - LOADBOARD"))
        self.pushButton.setText(_translate("MainWindow", "搜尋"))
        self.groupBox.setTitle(_translate("MainWindow", ""))
        self.label.setText(_translate("MainWindow", "Report Year"))
        self.label_2.setText(_translate("MainWindow", "Device Name"))

    def inp_data(self):
        # 連接資料表的Device_Name欄位資料
        self.cursor.execute("SELECT DISTINCT [Device_Name] FROM [Qualcomm_Data]")
        results = self.cursor.fetchall()
        self.comboBox.addItem('(ALL)')
        for row in results:
            self.comboBox.addItem(row[0])

        # 連接資料表的Y_Date欄位資料
        self.cursor.execute("SELECT DISTINCT Format([CH_DATE],'yyyy') as Y_Date FROM [Qualcomm_Data]")
        results = self.cursor.fetchall()
        self.comboBox_1.addItem('(ALL)')
        for row in results:
            self.comboBox_1.addItem(row[0])
        
        self.conn.close()

    def populate_table(self, data): #寫入Table
        # 設定表格欄位名稱--手動添加
        columns = [column[0] for column in self.cursor.description]
        ##columns = ['Device_Name','WeekNumber']
        # 設定表格欄位樣式
        header = self.tableWidget.horizontalHeader()
        header.setStyleSheet("QHeaderView::section {background-color: rgb(30, 138, 138); color: white; font-weight: bold;}")
        header.setHighlightSections(False)
        # 設定表格大小為資料表大小
        self.tableWidget.setColumnCount(len(columns))
        self.tableWidget.setHorizontalHeaderLabels(columns)
        self.tableWidget.setRowCount(len(data))
        #self.tableWidget.setColumnHidden(0, True) # 隱藏第0欄
        # 將資料寫入表格
        for row in range(len(data)):
            for col in range(len(data[0])):
                item = QtWidgets.QTableWidgetItem(str(data[row][col]))
                self.tableWidget.setItem(row, col, item)
                item.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter) #文字置中
        
        # 設置表格第一欄的單元格為當前單元格
        self.tableWidget.setCurrentCell(0, 0)
        # 更新曲线图
        self.update_chart()
        
    def search_database(self):
        self.database_connect()
        device_name = self.comboBox.currentText()
        Year_date = self.comboBox_1.currentText()

        # 建立基本的查詢指令
        query = "SELECT * FROM Weekly_P"
        params = ()
        # 判斷是否需要加入WHERE條件
        if device_name !='(ALL)' :
            query += " WHERE Device_Name=?"
            params = (device_name,)       
        else:
            self.cursor.execute(query, params) 
  
        # 判斷是否需要加入WHERE條件
        # if device_name != '(ALL)' and Year_date != '(ALL)':
            # query += " WHERE Device_Name=? AND Y_Date=?"
            # params = (device_name, Year_date)
        # elif device_name != '(ALL)':
            # query += " WHERE Device_Name=?"
            # params = (device_name,)
        # elif Year_date != '(ALL)':
            # query += " WHERE Y_Date=?"
            # params = (Year_date,)

        # 執行查詢
        self.cursor.execute(query, params)    
        data = self.cursor.fetchall()
        self.populate_table(data)
        self.update_chart() # 更新曲线图
        self.conn.close() # 關閉資料庫

    def update_chart(self):
        # 取得當前選擇的行列
        items = self.tableWidget.selectedItems()
        #firstColumnItems = [item for item in items if item.column() == 0]
        if not items:
            return
        row = items[0].row()
        col = items[0].column()
        # 取得所有欄位名稱
        labels = []
        for i in range(self.tableWidget.columnCount()):
            item = self.tableWidget.horizontalHeaderItem(i)
            labels.append(item.text())
            #print(labels[0])
        # 取得選擇的行的所有數值
        data = []
        for i in range(1,self.tableWidget.columnCount()):
            item = self.tableWidget.item(row, i)
            data.append(item.text())
            data_float = [float(d.strip('%')) for d in data]
        cname = []
        for i in range(self.tableWidget.columnCount()):
            item = self.tableWidget.item(row, i)
            cname.append(item.text())
            Device_name = cname[0]
        # 創建數據序列
        series = QLineSeries()
        
        scatter = QScatterSeries()
        # 添加數據點
        for i in range(len(data_float)):
            series.append(i+1, float(data_float[i]))
            scatter.append(i+1, float(data_float[i]))

        # 創建圖表並添加數據序列
        self.chart = QChart()
        # ...設置 chartView 大小和其他屬性...
        self.chart.legend().setVisible(True)
        self.chart.legend().setAlignment(Qt.AlignRight)
        
        # 顯示/背景顏色
        self.chart.setBackgroundVisible(True)
        #self.chart.setBackgroundBrush(QBrush(Qt.transparent)) # 背景透明
        #self.chart.setBackgroundBrush(QBrush(QColor(30,150,155)))
        # 將 LineSeries 添加到圖表中
        self.chart.removeAllSeries()
        self.chart.addSeries(series)
        self.chart.addSeries(scatter)
        
        # 創建 X 軸和 Y 軸
        xaxis = QCategoryAxis()
        yaxis = QValueAxis()

        # 將欄位名稱添加到 X 軸
        for i in range(len(labels)):
            if (i==0):{}
            else: {
            xaxis.append(labels[i],i)}
                
        # 將 X 軸添加到圖表中
        self.chart.addAxis(xaxis, Qt.AlignBottom)
        series.attachAxis(xaxis)
        scatter.attachAxis(xaxis)
        # 將 Y 軸添加到圖表中
        self.chart.addAxis(yaxis, Qt.AlignLeft)
        series.attachAxis(yaxis)
        scatter.attachAxis(yaxis)

        # 設置圖表標題/背景主題
        #self.chart.setTitle("Weekly Quantity")
        #self.chart.setTheme(QChart.ChartThemeBrownSand)
        #self.chart.setTheme(QChart.ChartThemeHighContrast)
        #self.chart.setTheme(QChart.ChartThemeBlueCerulean)
        #self.chart.setTheme(QChart.ChartThemeLight)
        #self.chart.setTheme(QChart.ChartThemeDark)
        # 設置折線圖的線條顏色和名稱
        pen = QPen()
        pen.setWidth(2)
        pen.setColor(QColor(0, 150, 200))
        series.setPen(pen)
        series.setName(Device_name)
    
        # 設置數據點標記的大小和樣式
        scatter.setVisible(True)
        scatter.setMarkerSize(6)
        scatter.setMarkerShape(QScatterSeries.MarkerShapeRectangle) #圖標形狀
        scatter.setColor(QColor(0, 150, 200))
        scatter.setPointLabelsVisible(True)
        scatter.setPointLabelsColor(QColor(32, 78, 152))
        scatter.setPointLabelsFormat("@yPoint%") #標記數值
        scatter.setPointLabelsClipping(False)
        font = QFont("Arial", 8)
        scatter.setPointLabelsFont(font)
            
        # 隱藏散點圖的圖例
        marker = self.chart.legend().markers(scatter)[0]
        marker.setVisible(False)

        # 90% 標準線
        series2 = QLineSeries()
        series2.setVisible(True)
        series2.setName("Standard line")
        series2.setPen(QPen(QColor(0, 180, 80),1.5, Qt.DashLine))
        series2.append(0, 90)
        series2.append(len(labels), 90)
        self.chart.addSeries(series2)
        series2.attachAxis(yaxis)
       # series2.setPointLabelsFormat("90%") #標記數值
       # series2.setPointLabelsClipping(False)
        marker = self.chart.legend().markers(series2)[0]
        marker.setVisible(False)
        
        # 設置 X 軸的範圍和標籤
        xaxis.setLabelsVisible(True) 
        xaxis.setRange(0, len(labels))
        xaxis.setLabelsPosition(QCategoryAxis.AxisLabelsPositionOnValue)# 對齊格線
        xaxis.setTitleVisible(False)
        xaxis.setTitleText("Weekly")
        
        # 設置 Y 軸範圍
        yaxis.setLabelFormat("%.2f%")
        yaxis.setTickCount(3)
        yaxis.setMinorTickCount(4)
        yaxis.setRange(0, 100)
        yaxis.setTitleText("Health %")

        # 設置圖表視圖的屬性和添加圖表到視圖
        self.graphicsView.setRenderHint(QPainter.Antialiasing) # 抗巨齒
        self.graphicsView.setChart(self.chart)
        self.graphicsView.raise_()

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
