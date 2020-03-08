"""

Real Estate Transaction Analysis Tools
SNU 220-344

"""

# Package Import
import pip
import sys
import time

def install(package):
    pip.main(['install', package])

try:
    from PyQt5.QtWidgets import *
    from PyQt5 import uic, QtGui
except ImportError:
    print("Installing PyQt")
    install('PyQt5')
    from PyQt5.QtWidgets import *
    from PyQt5 import uic, QtGui

try:
    import win32com.client
except ImportError:
    print("Installing PyWin32")
    install('pywin32')
    import win32com.client

try:
    import pandas as pd
except ImportError:
    print("Installing Pandas")
    install('pandas')
    import pandas as pd

try:
    import numpy as np
except ImportError:
    print("Installing NumPy")
    install('numpy')
    import numpy as np

try:
    from statsmodels.formula.api import ols
    from statsmodels.formula.api import quantreg as qreg
    import statsmodels.api as sm
except ImportError:
    print("Installing StatsModels")
    install('statsmodels')
    from statsmodels.formula.api import ols
    from statsmodels.formula.api import quantreg as qreg
    import statsmodels.api as sm

try:
    import patsy
except ImportError:
    print("Installing Patsy")
    install('patsy')
    import patsy

form_class = uic.loadUiType("gui/pyqt5.ui")[0]
excel = win32com.client.Dispatch("Excel.Application")
app_ver = 191227

class Window(QMainWindow, form_class):
    def __init__(self):
        # form_class inheriance
        super().__init__()
        self.setupUi(self)

        # Qt GUI Design
        stylesheet = """ QTabWidget>QWidget>QWidget{background: #f0f0f0;} """
        self.setStyleSheet(stylesheet)
        self.setWindowIcon(QtGui.QIcon("gui/icon.png"))

        ''' Main Screen '''
        # Tab Connection
        self.Main_Connect1.clicked.connect(self.connect_tab1)
        self.Main_Connect2.clicked.connect(self.connect_tab2)
        self.Main_Connect3.clicked.connect(self.connect_tab3)
        self.Main_Connect4.clicked.connect(self.connect_tab4)

        # Main Logo
        self.Main_logo.setPixmap(QtGui.QPixmap("gui/logo_small.png"))

        ''' Rawdata Updater '''
        # Variable Initialization
        self.RU_path1 = None
        self.RU_path2 = None
        self.RU_data1 = None
        self.RU_data2 = None
        self.RU_append_data = None

        # Event Connection
        self.RU_button_open1.clicked.connect(self.ru_open1)
        self.RU_button_open2.clicked.connect(self.ru_open2)
        self.RU_button_import1.clicked.connect(self.ru_import1)
        self.RU_button_import2.clicked.connect(self.ru_import2)
        self.RU_button_append.clicked.connect(self.ru_append)
        self.RU_button_save.clicked.connect(self.ru_save)

        ''' Data Processor '''
        # Variable Initialization
        self.DP_path_rawdata = None
        self.DP_path_key = None
        self.DP_path_dongcode = None
        self.DP_path_cpi = None
        self.DP_data = None
        self.DP_key = None
        self.DP_dongcode = None
        self.DP_cpi = None
        self.DP_data_edited = None
        self.DP_wb = None
        self.DP_ws = None

        # Event Connection
        self.DP_button_open_rawdata.clicked.connect(self.dp_open_rawdata)
        self.DP_button_open_key.clicked.connect(self.dp_open_key)
        self.DP_button_open_dongcode.clicked.connect(self.dp_open_dongcode)
        self.DP_button_open_cpi.clicked.connect(self.dp_open_cpi)
        self.DP_button_import_rawdata.clicked.connect(self.dp_import_rawdata)
        self.DP_button_import_key.clicked.connect(self.dp_import_key)
        self.DP_button_import_dongcode.clicked.connect(self.dp_import_dongcode)
        self.DP_button_import_cpi.clicked.connect(self.dp_import_cpi)
        self.DP_button_analyze.clicked.connect(self.dp_analysis)
        self.DP_button_viewdata.clicked.connect(self.dp_view_data)
        self.DP_button_save.clicked.connect(self.dp_save)

        ''' Pivot Table Maker '''
        # Variable Initialization
        self.PTM_path = None
        self.PTM_data = None
        self.PTM_column_list = None
        self.PTM_pivot = None
        self.PTM_raw_wb = None
        self.PTM_raw_ws = None
        self.PTM_pivot_wb = None
        self.PTM_pivot_ws = None

        # Event Connection
        self.PTM_button_open.clicked.connect(self.ptm_open_file)
        self.PTM_button_import.clicked.connect(self.ptm_import_file)
        self.PTM_button_analyze.clicked.connect(self.ptm_analysis)
        self.PTM_button_save.clicked.connect(self.ptm_save)
        self.PTM_button_viewraw.clicked.connect(self.view_rawdata)
        self.PTM_button_viewpivot.clicked.connect(self.view_pivot)

        ''' Regression '''
        # Variable Initialization
        self.reg_path = None
        self.reg_data = None
        self.reg_column_list = None
        self.reg_frame = None

        # Event Connection
        self.reg_button_open.clicked.connect(self.reg_open_file)
        self.reg_button_import.clicked.connect(self.reg_import_file)
        self.reg_button_analyze.clicked.connect(self.reg_analyze)
        self.reg_button_save.clicked.connect(self.reg_save)

        ''' Menu Bar '''
        self.action_exit.triggered.connect(self.close_app)
        self.action_about.triggered.connect(self.about_app)

    # Function by Tabs from here

    ''' Main Screen '''
    def connect_tab1(self):
        QTabWidget.setCurrentIndex(self.tabWidget, 1)

    def connect_tab2(self):
        QTabWidget.setCurrentIndex(self.tabWidget, 2)

    def connect_tab3(self):
        QTabWidget.setCurrentIndex(self.tabWidget, 3)

    def connect_tab4(self):
        QTabWidget.setCurrentIndex(self.tabWidget, 4)

    ''' Rawdata Updater '''
    def ru_open1(self):
        ru_open_path1 = QFileDialog.getOpenFileName(self,
                                                    caption='Open Existing Rawdata',
                                                    filter='CSV Files (*.csv)')
        if ru_open_path1[0] == "":
            pass
        else:
            self.RU_path1 = ru_open_path1[0]
            self.RU_text_status.setText("Existing rawdata file selected.")
            self.RU_text_path1.setText("<span style=\"color:#000000;\" >" + self.RU_path1 + "</span>")

    def ru_open2(self):
        ru_open_path2 = QFileDialog.getOpenFileName(self,
                                                    caption='Open New Rawdata',
                                                    filter='Excel Files (*.xlsx)')
        if ru_open_path2[0] == "":
            pass
        else:
            self.RU_path2 = ru_open_path2[0]
            self.RU_text_status.setText("New rawdata file selected.")
            self.RU_text_path2.setText("<span style=\"color:#000000;\" >" + self.RU_path2 + "</span>")

    def ru_import1(self):
        if self.RU_path1 is None:
            QMessageBox.about(self, "Error", "Existing rawdata isn't selected.")
        else:
            try:
                self.RU_data1 = pd.read_csv(self.RU_path1, encoding="CP949", engine="python")
                try:
                    range_list = list(self.RU_data1.groupby(self.RU_data1["계약년월"]).describe().index)
                    self.RU_text_status.setText(
                        "Existing rawdata has been imported.\nRange: " + str(min(range_list)) + " to " + str(
                            max(range_list)))
                except:
                    self.RU_text_status.setText(
                        "Existing rawdata has been imported.\nAn error occurred while checking time range.")
                self.RU_text_path1.setText("<span style=\"color:#707070;\" >" + self.RU_path1 + "</span>")
            except UnicodeDecodeError:
                self.RU_data1 = pd.read_csv(self.RU_path1, encoding="UTF-8", engine="python")
                self.RU_text_status.setText("Existing rawdata has been imported.\nThis file is encoded as unicode.")
                self.RU_text_path1.setText("<span style=\"color:#707070;\" >" + self.RU_path1 + "</span>")

    def ru_import2(self):
        if self.RU_path2 is None:
            QMessageBox.about(self, "Error", "New rawdata isn't selected.")
        else:
            try:
                self.RU_data2 = pd.read_excel(self.RU_path2, encoding="CP949", engine=None)
                try:
                    range_list = list(self.RU_data2.groupby(self.RU_data2["계약년월"]).describe().index)
                    self.RU_text_status.setText(
                        "New rawdata has been imported.\nRange: " + str(min(range_list)) + " to " + str(
                            max(range_list)))
                except:
                    self.RU_text_status.setText(
                        "New rawdata has been imported.\nAn error occurred while checking time range.")
                self.RU_text_path2.setText("<span style=\"color:#707070;\" >" + self.RU_path2 + "</span>")
            except UnicodeDecodeError:
                self.RU_data2 = pd.read_excel(self.RU_path2, encoding="UTF-8", engine="python")
                self.RU_text_status.setText("New rawdata has been imported.\nThis file is encoded as unicode.")
                self.RU_text_path2.setText("<span style=\"color:#707070;\" >" + self.RU_path2 + "</span>")

    def ru_append(self):
        if self.RU_data1 is None:
            QMessageBox.about(self, "Error", "Existing rawdata isn't imported.")
        elif self.RU_data2 is None:
            QMessageBox.about(self, "Error", "New rawdata isn't imported.")
        elif not list(self.RU_data1.columns) == list(self.RU_data2.columns):
            QMessageBox.about(self, "Error", "Columns between data table aren't equal.\n" +
                              "Please check the data one more time.")
        else:
            try:
                self.RU_append_data = self.RU_data1.append(self.RU_data2)
                self.RU_text_status.setText("Successfully appended.\n" +
                                            "Number of total data is " + str(len(self.RU_append_data)))
            except:
                QMessageBox.about(self, "Error", "An unexpected error occurred.")

    def ru_save(self):
        if self.RU_append_data is None:
            QMessageBox.about(self, "Error", "There isn't appended data.")
        else:
            ru_save_path = QFileDialog.getSaveFileName(self,
                                                       caption='Save File',
                                                       filter='CSV Files (*.csv)')
            if ru_save_path[0] == "":
                pass
            else:
                self.RU_append_data.to_csv(ru_save_path[0], index=False, encoding="CP949")
                self.RU_text_status.setText("Rawdata has been saved to the following location:\n"
                                            + str(ru_save_path[0]))

    ''' Data Processor '''
    def dp_open_rawdata(self):
        dp_open_path = QFileDialog.getOpenFileName(self, caption='Open File', filter='CSV Files (*.csv)')
        if dp_open_path[0] == "":
            pass
        else:
            self.DP_path_rawdata = dp_open_path[0]
            self.DP_text_status.setText("Rawdata CSV file selected.")
            self.DP_text_path_rawdata.setText("<span style=\"color:#000000;\" >" + self.DP_path_rawdata + "</span>")

    def dp_import_rawdata(self):
        if self.DP_path_rawdata is None:
            QMessageBox.about(self, "Error", "No rawdata CSV file selected.")
        else:
            self.DP_data = pd.read_csv(self.DP_path_rawdata, encoding="CP949", engine="python")
            self.DP_text_status.setText("Rawdata import has been completed.")
            self.DP_text_path_rawdata.setText("<span style=\"color:#707070;\" >" + self.DP_path_rawdata + "</span>")

    def dp_open_key(self):
        dp_open_path = QFileDialog.getOpenFileName(self, caption='Open File', filter='CSV Files (*.csv)')
        if dp_open_path[0] == "":
            pass
        else:
            self.DP_path_key = dp_open_path[0]
            self.DP_text_status.setText("Key CSV file selected.")
            self.DP_text_path_key.setText("<span style=\"color:#000000;\" >" + self.DP_path_key + "</span>")

    def dp_import_key(self):
        if self.DP_path_key is None:
            QMessageBox.about(self, "Error", "Key CSV file isn't selected.")
        else:
            self.DP_key = pd.read_csv(self.DP_path_key, encoding="CP949", engine="python")
            self.DP_text_status.setText("Key import has been completed.")
            self.DP_text_path_key.setText("<span style=\"color:#707070;\" >" + self.DP_path_key + "</span>")

    def dp_open_dongcode(self):
        dp_open_path = QFileDialog.getOpenFileName(self, caption='Open File', filter='CSV Files (*.csv)')
        if dp_open_path[0] == "":
            pass
        else:
            self.DP_path_dongcode = dp_open_path[0]
            self.DP_text_status.setText("Dongcode CSV file selected.")
            self.DP_text_path_dongcode.setText("<span style=\"color:#000000;\" >" + self.DP_path_dongcode + "</span>")

    def dp_import_dongcode(self):
        if self.DP_path_dongcode is None:
            QMessageBox.about(self, "Error", "Dongcode CSV file isn't selected.")
        else:
            self.DP_dongcode = pd.read_csv(self.DP_path_dongcode, encoding="CP949", engine="python")
            self.DP_text_status.setText("Dongcode import has been completed.")
            self.DP_text_path_dongcode.setText("<span style=\"color:#707070;\" >" + self.DP_path_dongcode + "</span>")

    def dp_open_cpi(self):
        dp_open_path = QFileDialog.getOpenFileName(self, caption='Open File', filter='CSV Files (*.csv)')
        if dp_open_path[0] == "":
            pass
        else:
            self.DP_path_cpi = dp_open_path[0]
            self.DP_text_status.setText("CPI CSV file selected.")
            self.DP_text_path_cpi.setText("<span style=\"color:#000000;\" >" + self.DP_path_cpi + "</span>")

    def dp_import_cpi(self):
        if self.DP_path_cpi is None:
            QMessageBox.about(self, "Error", "CPI CSV file isn't selected.")
        else:
            self.DP_cpi = pd.read_csv(self.DP_path_cpi, encoding="CP949", engine="python")
            self.DP_text_status.setText("CPI import has been completed.")
            self.DP_text_path_cpi.setText("<span style=\"color:#707070;\" >" + self.DP_path_cpi + "</span>")

    def dp_view_data(self):
        if self.DP_data is None:
            QMessageBox.about(self, "Error", "No rawdata imported.")
        else:
            self.DP_wb = excel.Workbooks.Add()
            self.DP_ws = self.DP_wb.Worksheets("Sheet1")
            num1 = 1
            for i in self.DP_data.head().columns:
                self.DP_ws.Cells(1, num1).Value = i
                num1 = num1 + 1
            num2 = 1
            for i in list(self.DP_data.head().columns):
                for j in range(len(list(self.DP_data.head().index))):
                    if type(self.DP_data.head()[i][j]).__module__ == 'numpy':
                        self.DP_ws.Cells(j + 2, num2).Value = int(self.DP_data.head()[i][j])
                    else:
                        self.DP_ws.Cells(j + 2, num2).Value = self.DP_data.head()[i][j]
                num2 = num2 + 1
            del num1, num2
            excel.Visible = True

    def dp_analysis(self):
        try:

            ''' # Temporary
            def sort_quarter(df, column_name):
                january = df[df[column_name].astype(str).str[4:] == "01"]
                february = df[df[column_name].astype(str).str[4:] == "02"]
                march = df[df[column_name].astype(str).str[4:] == "03"]
                april = df[df[column_name].astype(str).str[4:] == "04"]
                may = df[df[column_name].astype(str).str[4:] == "05"]
                june = df[df[column_name].astype(str).str[4:] == "06"]
                july = df[df[column_name].astype(str).str[4:] == "07"]
                august = df[df[column_name].astype(str).str[4:] == "08"]
                september = df[df[column_name].astype(str).str[4:] == "09"]
                october = df[df[column_name].astype(str).str[4:] == "10"]
                november = df[df[column_name].astype(str).str[4:] == "11"]
                december = df[df[column_name].astype(str).str[4:] == "12"]

                q1 = january.append(february).append(march)
                q2 = april.append(may).append(june)
                q3 = july.append(august).append(september)
                q4 = october.append(november).append(december)

                q1["Quarter"] = q1[column_name].astype(str).str[0:4] + "Q1"
                q2["Quarter"] = q2[column_name].astype(str).str[0:4] + "Q2"
                q3["Quarter"] = q3[column_name].astype(str).str[0:4] + "Q3"
                q4["Quarter"] = q4[column_name].astype(str).str[0:4] + "Q4"

                sorted_df = q1.append(q2).append(q3).append(q4).sort_index()
                assert isinstance(sorted_df, object)
                return sorted_df
            '''

            if str(self.DP_combobox_transaction_type.currentText()) == "APT_Trade":
                data = self.DP_data.copy()
                key = self.DP_key.copy()
                del key["Sigungu"], key["Dong"]
                data.columns = ["Location", "Address", "Address1", "Address2", "APT_name", "Area", "ContractMonth",
                                "ContractDay", "Price", "Floor", "ConstructionYear", "RoadName"]
                data["Sido"] = data["Location"].str.split(" ", expand=True)[0]
                data["Sigungu"] = data["Location"].str.split(" ", expand=True)[1]
                data["Dong"] = data["Location"].str.split(" ", expand=True)[2]
                data = data[data["Sido"]=="서울특별시"]
                data["ContractMonth_date"] = pd.to_datetime(data["ContractMonth"], format="%Y%m")
                data["Area_pyung"] = data["Area"] / 3.3
                data["Area_dummy"] = np.where(data["Area"]>60, (np.where(data["Area"]>85, (np.where(data["Area"]>135, "대형", "중대형")), "중소형")), "소형")
                data["Floor_dummy"] = np.where(data["Floor"]<3, "저층부", "3층 이상")
                data["Price"] = data["Price"].astype(str).str.replace(",", "").astype(int)
                data["PricePerPyung"] = data["Price"] / data["Area_pyung"]
                data = pd.merge(data, self.DP_dongcode, left_on="Location", right_on="DongName")
                data["Key"] = data["Dong"] + data["Address"] + data["APT_name"]
                del data["DongName"], data["Maintenance"]
                data = pd.merge(data, key, left_on="Key", right_on="Key", how="inner")
                self.DP_data_edited = pd.merge(data, self.DP_cpi, left_on="ContractMonth", right_on="ContractMonth")
                self.DP_data_edited["RealPricePerPyung"] = self.DP_data_edited["PricePerPyung"] / (self.DP_data_edited["CPI"] / 100)
                self.DP_data_edited = self.DP_data_edited[self.DP_data_edited["Area_pyung"] > 10]
                self.DP_data_edited = self.DP_data_edited[[
                    "Location", "Address", "Address1", "Address2", "Sido", "Sigungu", "Dong", "DongCode", "APT_name",
                     "Area", "Area_pyung", "Area_dummy","ContractMonth", "ContractMonth_date", "CPI", "ContractDay",
                     "Price", "Floor", "Floor_dummy","ConstructionYear", "RoadName", "PricePerPyung", "RealPricePerPyung",
                     "APT_name_cleared_key", "APT_unit"]]

                self.DP_text_status.setText("Data processing is complete. Please save the DataFrame.\nCondition: "
                                            "Seoul, Over 10 pyung, APT name cleared.\nTime range is {} to {}.".format(
                    min(self.DP_data_edited["ContractMonth"]), max(self.DP_data_edited["ContractMonth"])))

            elif str(self.DP_combobox_transaction_type.currentText()) == "APT_Rent":
                data = self.DP_data.copy()
                key = self.DP_key.copy()
                del key["Sigungu"], key["Dong"]
                data.columns = ["Location", "Address", "Address1", "Address2", "APT_name", "RentType", "Area",
                                        "ContractMonth", "ContractDay", "Deposit", "MonthlyRent", "Floor", "ConstructionYear", "RoadName"]
                data["Sido"] = data["Location"].str.split(" ", expand=True)[0]
                data["Sigungu"] = data["Location"].str.split(" ", expand=True)[1]
                data["Dong"] = data["Location"].str.split(" ", expand=True)[2]
                data = data[(data["Sido"]=="서울특별시")&(data["RentType"]=="전세")]
                data["ContractMonth_date"] = pd.to_datetime(data["ContractMonth"], format="%Y%m")
                data["Area_pyung"] = data["Area"] / 3.3
                data["Area_dummy"] = np.where(data["Area"]>60, (np.where(data["Area"]>85, (np.where(data["Area"]>135, "대형", "중대형")), "중소형")), "소형")
                data["Floor_dummy"] = np.where(data["Floor"]<3, "저층부", "3층 이상")
                data["Deposit"] = data["Deposit"].astype(str).str.replace(",", "").astype(int)
                data["MonthlyRent"] = data["MonthlyRent"].astype(str).str.replace(",", "").astype(int)
                data["DepositPerPyung"] = data["Deposit"] / data["Area_pyung"]
                data = pd.merge(data, self.DP_dongcode, left_on="Location", right_on="DongName")
                data["Key"] = data["Dong"] + data["Address"] + data["APT_name"]
                del data["DongName"], data["Maintenance"]
                data = pd.merge(data, key, left_on="Key", right_on="Key", how="inner")
                self.DP_data_edited = pd.merge(data, self.DP_cpi, left_on="ContractMonth", right_on="ContractMonth")
                self.DP_data_edited["RealDepositPerPyung"] = self.DP_data_edited["DepositPerPyung"] / (self.DP_data_edited["CPI"] / 100)
                self.DP_data_edited = self.DP_data_edited[self.DP_data_edited["Area_pyung"] > 10]
                self.DP_data_edited = self.DP_data_edited[["Location", "Address", "Address1", "Address2", "Sido", "Sigungu",
                                                           "Dong","DongCode", "APT_name", "RentType","Area", "Area_pyung", "Area_dummy",
                                                           "ContractMonth", "ContractMonth_date", "CPI","ContractDay",
                                                           "Deposit", "Floor", "Floor_dummy","ConstructionYear", "RoadName",
                                                           "DepositPerPyung", "RealDepositPerPyung", "APT_name_cleared_key", "APT_unit"]]

                self.DP_text_status.setText("Data processing is complete. Please save the DataFrame.\nCondition: "
                                            "Seoul, Jeonse, Over 10 pyung, APT name cleared.\nTime range is {} to {}.".format(min(self.DP_data_edited["ContractMonth"]), max(self.DP_data_edited["ContractMonth"])))

            else:
                QMessageBox.about(self, "Error", "At present, it is available only APT trade or APT rent.")

        except Exception as e:
            QMessageBox.about(self, "Error", "An error occurred. Detail: " + str(e))

    def dp_save(self):
        if self.DP_data_edited is None:
            QMessageBox.about(self, "Error", "There isn't edited data.")
        else:
            DP_save_path = QFileDialog.getSaveFileName(self,
                                                       caption='Save File',
                                                       filter='CSV Files (*.csv)')
            if DP_save_path[0] == "":
                pass
            else:
                self.DP_data_edited.to_csv(DP_save_path[0], index=False, encoding="CP949")
                self.DP_text_status.setText("DataFrame has been saved to the following location:\n"
                                            + str(DP_save_path[0]))
                self.DP_data_edited = None

    ''' Pivot Table Maker '''
    def ptm_open_file(self):
        ptm_open_path = QFileDialog.getOpenFileName(self,
                                                    caption='Open File',
                                                    filter='CSV Files (*.csv)')
        if ptm_open_path[0] == "":
            pass
        else:
            self.PTM_path = ptm_open_path[0]
            self.PTM_text_status.setText("CSV file selected")
            self.PTM_text_path.setText("<span style=\"color:#000000;\" >" + self.PTM_path + "</span>")

    def ptm_import_file(self):
        if self.PTM_path is None:
            QMessageBox.about(self, "Error", "No CSV file selected.")
        else:
            self.PTM_data = pd.read_csv(self.PTM_path, encoding="CP949", engine="python")
            self.PTM_text_status.setText("Data import has been completed.")
            self.PTM_text_path.setText("<span style=\"color:#707070;\" >" + self.PTM_path + "</span>")
            self.PTM_column_list = list(self.PTM_data.columns)
            self.PTM_combo_column.clear()
            self.PTM_combo_column.addItems(self.PTM_column_list)
            self.PTM_combo_row.clear()
            self.PTM_combo_row.addItems(self.PTM_column_list)
            self.PTM_combo_subject.clear()
            self.PTM_combo_subject.addItems(self.PTM_column_list)

    def ptm_analysis(self):
        ptm_table_reset = None
        try:
            stats = str(self.PTM_combo_stats.currentText())
            row = str(self.PTM_combo_row.currentText())
            column = str(self.PTM_combo_column.currentText())
            subject = str(self.PTM_combo_subject.currentText())

            # Calc Stats and Making Pivot Table
            if stats == "Mean":
                self.PTM_pivot = pd.pivot_table(self.PTM_data, index=row, columns=column, aggfunc="mean")[subject]
            elif stats == "Median":
                self.PTM_pivot = pd.pivot_table(self.PTM_data, index=row, columns=column, aggfunc="median")[subject]
            elif stats == "Min":
                self.PTM_pivot = pd.pivot_table(self.PTM_data, index=row, columns=column, aggfunc="min")[subject]
            elif stats == "Max":
                self.PTM_pivot = pd.pivot_table(self.PTM_data, index=row, columns=column, aggfunc="max")[subject]
            elif stats == "Standard Deviation":
                self.PTM_pivot = pd.pivot_table(self.PTM_data, index=row, columns=column, aggfunc="std")[subject]
            elif stats == "Number of Data":
                self.PTM_pivot = pd.pivot_table(self.PTM_data, index=row, columns=column, aggfunc="count")[subject]
            elif stats == "Sum":
                self.PTM_pivot = pd.pivot_table(self.PTM_data, index=row, columns=column, aggfunc="sum")[subject]
            self.PTM_text_status.setText("Pivot table has been made.")
        except:
            QMessageBox.about(self, "Error", "An error occurred.")

    def view_rawdata(self):
        if self.PTM_data is None:
            QMessageBox.about(self, "Error", "No rawdata imported.")
        else:
            self.PTM_raw_wb = excel.Workbooks.Add()
            self.PTM_raw_ws = self.PTM_raw_wb.Worksheets("Sheet1")
            num1 = 1
            for i in self.PTM_data.head().columns:
                self.PTM_raw_ws.Cells(1, num1).Value = i
                num1 = num1 + 1
            num2 = 1
            for i in list(self.PTM_data.head().columns):
                for j in range(len(list(self.PTM_data.head().index))):
                    if type(self.PTM_data.head()[i][j]).__module__ == 'numpy':
                        self.PTM_raw_ws.Cells(j + 2, num2).Value = int(self.PTM_data.head()[i][j])
                    else:
                        self.PTM_raw_ws.Cells(j + 2, num2).Value = self.PTM_data.head()[i][j]
                num2 = num2 + 1
            del num1, num2
            excel.Visible = True

    def view_pivot(self):
        if self.PTM_pivot is None:
            QMessageBox.about(self, "Error", "There isn't pivot table.")
        else:
            self.PTM_pivot_wb = excel.Workbooks.Add()
            self.PTM_pivot_ws = self.PTM_pivot_wb.Worksheets("Sheet1")
            num1 = 2
            num2 = 2
            num3 = 2
            for i in self.PTM_pivot.head().columns.tolist():
                self.PTM_pivot_ws.Cells(1, num1).Value = i[1]
                num1 = num1 + 1
            for i in list(self.PTM_pivot.head().index.tolist()):
                self.PTM_pivot_ws.Cells(num3, 1).Value = i
                num3 = num3 + 1
            for i in list(self.PTM_pivot.head().columns):
                for j in range(len(list(self.PTM_pivot.head().index))):
                    if type(self.PTM_pivot.head()[i][j]).__module__ == 'numpy':
                        if self.PTM_pivot.head()[i][j] != self.PTM_pivot.head()[i][j]:
                            self.PTM_pivot_ws.Cells(j + 2, num2).Value = "NaN"
                        else:
                            self.PTM_pivot_ws.Cells(j + 2, num2).Value = float(self.PTM_pivot.head()[i][j])
                    else:
                        self.PTM_pivot_ws.Cells(j + 2, num2).Value = self.PTM_pivot.head()[i][j]
                num2 = num2 + 1
            del num1, num2, num3
            excel.Visible = True

    def ptm_save(self):
        if self.PTM_pivot is None:
            QMessageBox.about(self, "Error", "There isn't pivot table.")
        else:
            ptm_save_path = QFileDialog.getSaveFileName(self,
                                                        caption='Save File',
                                                        filter='CSV Files (*.csv)')
            if ptm_save_path[0] == "":
                pass
            else:
                self.PTM_pivot.to_csv(ptm_save_path[0], index=True, encoding="CP949")
                self.PTM_text_status.setText("Pivot table has been saved to the following location:\n"
                                             + str(ptm_save_path[0]))

    ''' Regression '''
    def reg_open_file(self):
        reg_open_path = QFileDialog.getOpenFileName(self,
                                                    caption='Open File',
                                                    filter='CSV Files (*.csv)')
        if reg_open_path[0] == "":
            pass
        else:
            self.reg_path = reg_open_path[0]
            self.reg_text_status.setText("CSV file selected")
            self.reg_text_path.setText("<span style=\"color:#000000;\" >" + self.reg_path + "</span>")

    def reg_import_file(self):
        if self.reg_path is None:
            QMessageBox.about(self, "Error", "No CSV file selected.")
        else:
            self.reg_data = pd.read_csv(self.reg_path, encoding="CP949", engine="python")
            self.reg_text_status.setText("Data import has been completed.")
            self.reg_text_path.setText("<span style=\"color:#707070;\" >" + self.reg_path + "</span>")
            self.reg_column_list = list(self.reg_data.columns)
            self.reg_combo_deplog.clear()
            self.reg_combo_deplog.addItems(self.reg_column_list)
            self.reg_list_internal.clear()
            self.reg_list_internal.addItems(self.reg_column_list)
            self.reg_combo_columns.clear()
            self.reg_combo_columns.addItems(self.reg_column_list)
            self.reg_combo_time.clear()
            self.reg_combo_time.addItems(self.reg_column_list)
            self.reg_combo_apt.clear()
            self.reg_combo_apt.addItems(self.reg_column_list)

    def reg_analyze(self):
        try:
            def areg(formula, data=None, absorb=None):
                y, X = patsy.dmatrices(formula, data, return_type='dataframe')
                ybar = y.mean()
                y = y - y.groupby(data[absorb]).transform('mean') + ybar
                Xbar = X.mean()
                X = X - X.groupby(data[absorb]).transform('mean') + Xbar
                reg = sm.OLS(y, X)
                reg.df_resid -= (data[absorb].nunique() - 1)
                return reg

            if self.reg_data is None:
                QMessageBox.about(self, "Error", "No data imported.")

            elif str(self.reg_combo_regtype.currentText()) == "OLS Regression":

                start = time.time()
                rows_name = str(self.reg_combo_time.currentText())
                columns_name = str(self.reg_combo_columns.currentText())
                dep = str(self.reg_combo_deplog.currentText())
                indep_internal = ") +C(".join(str(x.text()) for x in self.reg_list_internal.selectedItems())
                time_dummy = str(self.reg_combo_time.currentText())
                apt_dummy = str(self.reg_combo_apt.currentText())
                times = self.reg_data[rows_name].drop_duplicates()
                time_length = len(times)
                locations = self.reg_data[columns_name].drop_duplicates().sort_values()

                # Seoul Total with areg
                formula = "np.log(" + dep + ") ~ " + "C(" + time_dummy + ") " + " + C(" + indep_internal + ")"
                model = areg(formula, data=self.reg_data, absorb=apt_dummy).fit()
                self.reg_frame = pd.DataFrame(model.params, columns=["서울특별시_coef"]).ix[0:time_length + 4]
                self.reg_frame["서울특별시_coef"][0] = 0
                self.reg_frame["서울특별시"] = np.exp(self.reg_frame["서울특별시_coef"]) * 100
                del self.reg_frame["서울특별시_coef"]

                for location_name in locations:
                    data = self.reg_data[self.reg_data[columns_name] == location_name]
                    formula = "np.log(" + dep + ") ~ " + "C(" + time_dummy + ") " + " + C(" + indep_internal + ")" + "+C(" + apt_dummy + ")"
                    model = ols(formula, data=data).fit()
                    frame_temp = pd.DataFrame(model.params, columns=[str(location_name) + "_coef"]).ix[0:time_length+4]
                    frame_temp[str(location_name) + "_coef"][0] = 0
                    frame_temp[location_name] = np.exp(frame_temp[str(location_name) + "_coef"]) * 100
                    del frame_temp[str(location_name) + "_coef"]
                    self.reg_frame = pd.merge(self.reg_frame, frame_temp, left_index=True, right_index=True)

                self.reg_text_status.setText(
                    "OLS Regression Complete.\nRunning time: {} seconds.\nTime range is {} to {}.".format(
                        int(time.time() - start), min(times), max(times)))

            elif str(self.reg_combo_regtype.currentText()) == "Median Regression":
                start = time.time()
                rows_name = str(self.reg_combo_time.currentText())
                columns_name = str(self.reg_combo_columns.currentText())
                dep = str(self.reg_combo_deplog.currentText())
                indep_internal = " + ".join(str(x.text()) for x in self.reg_list_internal.selectedItems())
                time_dummy = str(self.reg_combo_time.currentText())
                apt_dummy = str(self.reg_combo_apt.currentText())
                times = self.reg_data[rows_name].drop_duplicates()
                time_length = len(times)
                locations = self.reg_data[columns_name].drop_duplicates()

                # Seoul Total with areg
                formula = "np.log(" + dep + ") ~ " + "C(" + time_dummy + ") " + " + C(" + indep_internal + ")"
                model = areg(formula, data=self.reg_data, absorb=apt_dummy).fit()
                self.reg_frame = pd.DataFrame(model.params, columns=["서울특별시_coef"]).ix[0:time_length + 4]
                self.reg_frame["서울특별시_coef"][0] = 0
                self.reg_frame["서울특별시"] = np.exp(self.reg_frame["서울특별시_coef"]) * 100
                del self.reg_frame["서울특별시_coef"]

                for location_name in locations:
                    data = self.reg_data[self.reg_data[columns_name] == location_name]
                    formula = "np.log(" + dep + ") ~ " + "C(" + time_dummy + ") " + " + C(" + indep_internal + ")" + "+C(" + apt_dummy + ")"
                    model = qreg(formula, data=data).fit(q=0.5)
                    frame_temp = pd.DataFrame(model.params, columns=[str(location_name) + "_coef"]).ix[0:time_length+4]
                    frame_temp[str(location_name) + "_coef"][0] = 0
                    frame_temp[location_name] = np.exp(frame_temp[str(location_name) + "_coef"]) * 100
                    del frame_temp[str(location_name) + "_coef"]
                    self.reg_frame = pd.merge(self.reg_frame, frame_temp, left_index=True, right_index=True)

                self.reg_text_status.setText(
                    "Median Regression Complete.\nRunning time: {} seconds.\nTime range is {} to {}.".format(
                        int(time.time() - start), min(times), max(times)))
        except:
            QMessageBox.about(self, "Error", "An error occurred while regression analyzing.")

    def reg_save(self):
        if self.reg_frame is None:
            QMessageBox.about(self, "Error", "There isn't regression results.")
        else:
            reg_save_path = QFileDialog.getSaveFileName(self,
                                                        caption='Save File',
                                                        filter='CSV Files (*.csv)')
            if reg_save_path[0] == "":
                pass
            else:
                self.reg_frame.to_csv(reg_save_path[0], index=True, encoding="CP949")
                self.reg_text_status.setText("Indices have been saved to the following location:\n"
                                             + str(reg_save_path[0]))
                self.reg_frame = None

    ''' Menu Bar '''
    @staticmethod
    def close_app():
        app.exit()
        excel.Quit()

    def about_app(self):
        QMessageBox.about(self, "About", "Contact SNU 220-344\nShared City Lab\nVer." + str(app_ver))


if __name__ == "__main__":
    app = QApplication(sys.argv)
    RETAT = Window()
    RETAT.show()
    app.exec_()
