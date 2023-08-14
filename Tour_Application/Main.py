import sys
import numpy as np
import pandas as pd
from PySide6.QtWidgets import (QMainWindow,QApplication,QWidget,
                                QHBoxLayout,QVBoxLayout,QFileDialog,QLineEdit,
                                QPushButton,QLabel,QTableView,QHeaderView,QScrollBar,
                                QScrollArea)
from PySide6.QtCore import (Qt,QAbstractTableModel,QPointF,QEvent)
from PySide6.QtGui import QColor
import pyautogui
import matplotlib.pyplot as plt
import seaborn as sb
from PySide6.QtCharts import (QChartView,QChart,QPieSeries,QBarSeries,QBarSet,QValueAxis,QLineSeries,QAbstractSeries)
from wsproto import ConnectionType





WIN_WIDTH, WIN_HEIGHT= pyautogui.size()
#Main windown
class mainwindow(QMainWindow):
    def __init__(self, widget=None):
        QMainWindow.__init__(self)
        self.setWindowTitle("Tour Analyze Application")

        self.setCentralWidget(widget)

class TableModel(QAbstractTableModel):
    def __init__(self, data):
        super(TableModel, self).__init__()
        self._data = data

    def data(self, index, role=Qt.DisplayRole):
        if index.isValid():
            if role == Qt.DisplayRole:
                return str(self._data.iloc[index.row()][index.column()])
        return None

    def rowCount(self, index):
        # The length of the outer list.
        return len(self._data.values)

    def columnCount(self, index):
        # The following takes the first sub-list, and returns
        # the length (only works if all rows are an equal length)
        return self._data.columns.size

    def headerData(self, col, orientation, role):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return self._data.columns[col]
        return None

#Main widget
class widget(QWidget):
    def __init__(self):
        QWidget.__init__(self)

        #File input
        self.file_path = QLineEdit()
        #Button get path file
        self.input_button = QPushButton("Browse") 
        self.input_button.clicked.connect(self.get_file)

        #Button analyze data
        self.analyze_button = QPushButton("Start Analyze")
        self.analyze_button.clicked.connect(self.show_data)

        #file input layout
        self.input_layout = QHBoxLayout()
        self.input_layout.addWidget(self.file_path)
        self.input_layout.addWidget(self.input_button)


        #Table view
        self.table = QTableView()

        #left layout
        self.left_layout = QVBoxLayout()
        self.left_layout.addWidget(QLabel("Enter file path"))
        self.left_layout.addLayout(self.input_layout)
        self.left_layout.addWidget(self.analyze_button)
        self.left_layout.addSpacing(20)
        self.left_layout.addWidget(self.table)


        # Scroll bar
        self.scroll_area =QScrollArea()
        self.scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.scroll_area.setWidgetResizable(True)
        #right layout
        self.right_layout = QVBoxLayout()
        self.right_widget = QWidget()
        #Char view
        self.pie_chart = QChartView()
        self.pie_chart.setFixedHeight(WIN_HEIGHT/3)
        self.line_char = QChartView()
        self.line_char.setFixedHeight(WIN_HEIGHT/3)
        self.bar_char = QChartView()
        self.bar_char.setFixedHeight(WIN_HEIGHT/3)
        self.bar_char_2 = QChartView()
        self.bar_char_2.setFixedHeight(WIN_HEIGHT/3)
        


        self.pie_chart_volunteer_work = QChart()
        self.bar_chart_tour = QChart()
        self.pie_chart_calling = QChart()
        self.line_chart_customer = QChart()
        
        self.right_layout.addWidget(self.pie_chart)
        self.right_layout.addWidget(self.line_char)
        self.right_layout.addWidget(self.bar_char)
        self.right_layout.addWidget(self.bar_char_2)
        self.right_widget.setLayout(self.right_layout)
        self.scroll_area.setWidget(self.right_widget)




        #Set main layout
        self.layout = QHBoxLayout()
        self.layout.addLayout(self.left_layout)
        self.layout.addWidget(self.scroll_area)
        self.setLayout(self.layout)

    def get_file(self):
        fileName, _ = QFileDialog.getOpenFileName(self, 'Single File', '', '')
        self.file_path.setText(fileName)

    def show_data(self):
        if self.file_path.text()[-4:] =="xlsx":
            self.data = pd.read_excel(self.file_path.text())
        else:
            self.data = pd.read_csv(self.file_path.text())
        self.process_data()
        self.model = TableModel(self.data)
        self.table.setModel(self.model)
        self.file_path.setText("")
        
    def process_data(self):
        Column_name = ['Name','Tour','Duration','Rating','Volunteer work','Calling']
        self.data = self.data.iloc[:,[1,2,3,4,5,6]]
        self.data.columns = Column_name
        self.data['Duration'] = self.data['Duration'].replace("/2022","").replace("không",np.nan)
        self.data['Duration'] = self.data['Duration'].str.replace("/2022","")
        self.data['Duration'] = self.data['Duration'] + "/2022"
        try:
            self.data['Duration'] = self.data["Duration"].replace('Hứa Văn Đức/2022',np.nan)
            self.data['Duration'] = self.data["Duration"].replace('Đinh Hải Dương/2022',np.nan)
            self.data['Duration'] = self.data["Duration"].replace('không /2022',np.nan)
            self.data['Duration'] = self.data["Duration"].replace('Không/2022',np.nan)
            self.data['Duration'] = self.data["Duration"].replace('Không /2022',np.nan)
            self.data['Duration'] = self.data["Duration"].replace('không có /2022',np.nan)
        except:
            pass
        self.data['Duration'] = pd.to_datetime(self.data['Duration'])
        self.data['Tour'] = self.data['Tour'].replace("10/4/2022","")

        self.data.dropna(axis=0,inplace=True)
        # visualize the charts
        #Calling 
        self.call_pie_series = QPieSeries()
        for key,value in self.data['Calling'].value_counts().squeeze().items():
            self.call_pie_series.append(key,value)
        self.pie_chart_calling.legend().setAlignment(Qt.AlignRight)
        self.pie_chart_calling.setTitle("Tỷ lệ mong muốn nhận cuộc gọi")
        self.pie_chart_calling.addSeries(self.call_pie_series)
        self.pie_chart_calling.setAnimationOptions(QChart.SeriesAnimations)
        self.pie_chart_calling.setTheme(QChart.ChartThemeQt)
        self.bar_char_2.setChart(self.pie_chart_calling)

        self.customer_duration = pd.DataFrame(self.data.iloc[:,[0,2]].drop_duplicates().sort_values(by="Duration"))
        self.distribute_customer = self.customer_duration.groupby(by = self.customer_duration["Duration"].dt.month).count().iloc[:,0]
        
        #Distribute the customer
        Axis_X = QValueAxis()
        Axis_Y = QValueAxis()
        self.customer_series = QLineSeries()
        for month,customer_count in self.distribute_customer.squeeze().items():
            self.customer_series.append(month,customer_count)

        Axis_X.setRange(0,12)
        Axis_X.setTickCount(1)
        Axis_X.setTitleText("Tháng")
        Axis_Y.setRange(0,max(self.distribute_customer.squeeze().values)) 
        Axis_Y.setTitleText("Số lượng khách đi tour")
        self.line_chart_customer.setTitle("Biểu đồ phân bố khách hàng theo tháng")
        self.line_chart_customer.addSeries(self.customer_series)
        self.line_chart_customer.setAxisX(Axis_X)
        self.line_chart_customer.setAxisY(Axis_Y)
        self.line_chart_customer.setTheme(QChart.ChartThemeQt)
        self.line_chart_customer.setAnimationOptions(QChart.SeriesAnimations)
        self.line_char.setChart(self.line_chart_customer)
        
        #Volunteer work
        self.volunteer_pie_series =QPieSeries()
        for key,value in self.data['Volunteer work'].value_counts().squeeze().items():
            self.volunteer_pie_series.append(key,value)
        
        self.pie_chart_volunteer_work.legend().setAlignment(Qt.AlignRight)
        self.pie_chart_volunteer_work.setTitle("Tỷ lệ phân phối các hoạt động thiện nguyện")
        self.pie_chart_volunteer_work.addSeries(self.volunteer_pie_series)
        self.pie_chart_volunteer_work.setAnimationOptions(QChart.SeriesAnimations)
        self.pie_chart_volunteer_work.setTheme(QChart.ChartThemeQt)
        self.pie_chart.setChart(self.pie_chart_volunteer_work)

        #Tour count
        self.tour_count = pd.DataFrame(self.data['Tour'].value_counts())
        self.tour_count = self.tour_count.drop_duplicates(keep="first")

        self.bar_array =[]
        for value in self.tour_count.squeeze().keys():
            Tour_set = QBarSet(value)
            self.bar_array.append(Tour_set)

        for set_name,value in zip(self.bar_array,self.tour_count.squeeze().values):
            set_name << value
        
        self.bar_series = QBarSeries()
        for _set in self.bar_array:
            self.bar_series.append(_set)
        
        axisY = QValueAxis()
        axisY.setRange(0,max(self.tour_count.squeeze().values))
        self.bar_chart_tour.addAxis(axisY, Qt.AlignLeft)
        self.bar_series.attachAxis(axisY)

        self.bar_chart_tour.addSeries(self.bar_series)
        self.bar_chart_tour.setTheme(QChart.ChartThemeQt)
        self.bar_char.setChart(self.bar_chart_tour)
        self.bar_chart_tour.legend().setVisible(True)
        self.bar_chart_tour.legend().setAlignment(Qt.AlignRight)

        self.bar_chart_tour.setMinimumWidth(100)
        self.bar_chart_tour.setTitle("Phân bố tour du lịch")
        self.bar_chart_tour.setAnimationOptions(QChart.SeriesAnimations)

        self.data.insert(0, 'ID', np.array([x for x in range (1,len(self.data)+1)]))
    


        

if __name__ =="__main__":
    app = QApplication()

    widget = widget()
    window = mainwindow(widget)

    window.resize(int(WIN_WIDTH),int(WIN_HEIGHT))
    window.show()


    sys.exit(app.exec())