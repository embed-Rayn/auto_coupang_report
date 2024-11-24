import openpyxl
from datetime import datetime, timedelta
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import *
from PyQt6 import QtGui
from gui_larz import Ui_Dialog
import shutil
import os
import sys
import re
import pandas as pd


def is_xlsx(file_path):
    # 파일이고 확장자가 .xlsx인지 확인
    return os.path.isfile(file_path) and file_path.lower().endswith('.xlsx')


class WindowClass(QMainWindow, Ui_Dialog):
    def __init__(self):
        super().__init__()
        self.ui = Ui_Dialog()
        self.setupUi(self)
        self.setAcceptDrops(True)
        self.setWindowTitle("Auto Coupang report - 플렙")

        self.setFixedWidth(915)
        self.setFixedHeight(554)

        font1 = QtGui.QFont("새굴림", 11)
        font2 = QtGui.QFont("새굴림", 15)
        h1_list = [self.label_1, self.label_2]
        h2_list = [self.label_3, self.label_4, self.label_5]
        [obj.setFont(font1) for obj in h1_list]
        [obj.setFont(font2) for obj in h2_list]
        self.label_6.setStyleSheet("color: red;")
        
        self.input_file_path_1 = ""
        self.input_file_path_2 = ""
        self.input_file_path_3 = ""
        self.input_file_path_4 = ""
        self.output_file_path = ""

        self.btn_search_excel_1.clicked.connect(self.browse_1)
        self.btn_search_excel_2.clicked.connect(self.browse_2)
        self.btn_search_excel_3.clicked.connect(self.browse_3)
        self.btn_search_excel_4.clicked.connect(self.browse_4)
        self.btn_execute.clicked.connect(self.push_execute)
        self.gui_home_dir = os.path.abspath(".")
        self.is_successed = True

    def browse_1(self):
        try:
            file_filter = "Excel Files (*.xlsx)"
            path, _ = QFileDialog.getOpenFileName(None, 'Select an Excel File', '', file_filter)
            self.input_file_path_1 = path
        except Exception as e:
            print(e)
            self.input_file_path_1 = ""
        self.textEdit_log.setText(f"{self.textEdit_log.toPlainText()}쿠팡 일일 데이터 엑셀 위치: {self.input_file_path_1}\n")
        self.lineEdit_excel_1.setText(f"{self.input_file_path_1}")


    def browse_2(self):
        try:
            file_filter = "Excel Files (*.xlsx)"
            path, _ = QFileDialog.getOpenFileName(None, 'Select an Excel File', '', file_filter)
            self.input_file_path_2 = path
        except:
            self.input_file_path_2 = ""
        self.textEdit_log.setText(f"{self.textEdit_log.toPlainText()}쿠팡 누적 데이터 엑셀 위치: {self.input_file_path_2}\n")
        self.lineEdit_excel_2.setText(f"{self.input_file_path_2}")


    def browse_3(self):
        try:
            file_filter = "Excel Files (*.xlsx)"
            path, _ = QFileDialog.getOpenFileName(None, 'Select an Excel File', '', file_filter)
            self.input_file_path_3 = path
        except:
            self.input_file_path_3 = ""
        self.textEdit_log.setText(f"{self.textEdit_log.toPlainText()}적용할 리포트 엑셀 위치: {self.input_file_path_3}\n")
        self.lineEdit_excel_3.setText(f"{self.input_file_path_3}")


    def browse_4(self):
        try:
            file_filter = "Excel Files (*.xlsx)"
            path, _ = QFileDialog.getOpenFileName(None, 'Select an Excel File', '', file_filter)
            self.input_file_path_4 = path
        except:
            self.input_file_path_4 = ""
        self.textEdit_log.setText(f"{self.textEdit_log.toPlainText()}적용할 리포트 엑셀 위치: {self.input_file_path_4}\n")
        self.lineEdit_excel_4.setText(f"{self.input_file_path_4}")


    def finish_task(self, rst):
        if rst and self.is_successed:
            self.textEdit_log.setText(f"{self.textEdit_log.toPlainText()}작업이 완료 됐습니다.\n생성한 리포트 파일 위치: {self.output_file_path}\n")
        else:
            self.textEdit_log.setText(f"{self.textEdit_log.toPlainText()}{self.input_file_path_1}sheet가 여러개 입니다. 실행 종료\n")
        self.textEdit_log.setText(f"{self.textEdit_log.toPlainText()}================================================================\n")
        self.is_successed = True


    def process_file_path(self, report_day):
        pattern = r"\/\d{6}_"  # 6자리 숫자 (YYMMDD) 패턴
        match = re.search(pattern, self.input_file_path_3)

        if match:
            old_date = match.group()
            new_date = report_day.strftime("%y%m%d")
            self.output_file_path = self.input_file_path_3.replace(old_date, f"/{new_date}_")


    def auto_report(self):
        df_daily_o = pd.read_excel(self.input_file_path_1)
        df_daily_o['날짜'] = pd.to_datetime(df_daily_o['날짜'], format='%Y%m%d')
        df_daily_o = df_daily_o[["날짜", "캠페인명", "노출수", "클릭수", "광고비", "총 판매수량(14일)", "총 전환매출액(14일)"]]
        df_daily_o.columns = ["날짜", "캠페인명", "노출수", "클릭수", "광고비", "전환수", "전환매출"]

        df_daily_x = pd.read_excel(self.input_file_path_2)
        df_daily_x['날짜'] = pd.to_datetime(df_daily_x['날짜'], format='%Y%m%d')
        df_daily_x = df_daily_x[["날짜", "캠페인명", "노출수", "클릭수", "광고비", "총 판매수량(14일)", "총 전환매출액(14일)"]]
        df_daily_x.columns = ["날짜", "캠페인명", "노출수", "클릭수", "광고비", "전환수", "전환매출"]

        df_daily_total = pd.read_excel(self.input_file_path_3)
        df_daily_total['날짜'] = pd.to_datetime(df_daily_total['날짜'], format='%Y%m%d')
        df_daily_total = df_daily_total[["날짜", "캠페인명", "노출수", "클릭수", "광고비", "총 판매수량(14일)", "총 전환매출액(14일)"]]

        result_campaign = df_daily_total.groupby("campaign_name")[["exposure_num", "click_num", "ad_cost", "convert_num", "convert_cost"]].sum()
        result_by_keyword = df_daily_total.groupby("keyword")[["exposure_num", "click_num", "ad_cost", "convert_num", "convert_cost"]].sum()
        result_by_keyword = result_by_keyword.sort_values(by=["convert_num", "ad_cost"], ascending=[False, False]).reset_index()

        report_day = max(df_daily_o['날짜'].max, df_daily_x['날짜'].max)
        file_day = report_day + timedelta(days=1)
        self.process_file_path(file_day)
        try:
            shutil.copy(self.input_file_path_3, self.output_file_path)
        except PermissionError:
            self.textEdit_log.setText(f"{self.textEdit_log.toPlainText()}{self.output_file_path}파일이 이미 존재하거나 열려있어 실패.\n")
            self.is_successed = False
        #################################################### 리포트 파일
        #################################################### ["요약"] 시트
        report_wb = openpyxl.load_workbook(self.output_file_path)
        report_sheet = report_wb["요약"]
        report_sheet["C110"] = report_day
        report_sheet["C110"].number_format = "YYYY-MM-DD"
        #################################################### ["쿠팡_일일"] 시트
        report_sheet = report_wb["쿠팡_일일"]
        key_col_dict_o = {"노출수": "D", "클릭수": "E", "광고비": "H", "전환수": "I", "전환매출": "L"}
        for idx, row in df_daily_o.iterrows():
            date_obj = row['날짜']###############
            for key, col in key_col_dict_o.items():
                report_sheet[f"{col}{start_row+idx}"] = row[key]

        key_col_dict_x = {"노출수": "AF", "클릭수": "AG", "광고비": "AJ", "전환수": "AK", "전환매출": "AN"}
        for idx, row in df_daily_x.iterrows():
            for key, col in key_col_dict_x.items():
                report_sheet[f"{col}{start_row+idx}"] = row[key]


        for date_obj in dataset_daily.keys():
            for row_idx, cell in enumerate(report_sheet["B"]):
                try:
                    cv_date = cell.value.date()
                except AttributeError:
                    continue
                if cv_date.strftime("%y-%m-%d") == date_obj.strftime("%y-%m-%d"):
                    data_idx = row_idx + 1
                    report_sheet[f"D{data_idx}"] = sum(campaign["exposure_num"] for campaign in dataset_daily[date_obj])
                    report_sheet[f"E{data_idx}"] = sum(campaign["click_num"] for campaign in dataset_daily[date_obj])
                    report_sheet[f"H{data_idx}"] = sum(campaign["ad_cost"] for campaign in dataset_daily[date_obj])
                    report_sheet[f"I{data_idx}"] = sum(campaign["convert_num"] for campaign in dataset_daily[date_obj])
                    report_sheet[f"L{data_idx}"] = sum(campaign["convert_cost"] for campaign in dataset_daily[date_obj])
                    continue
        #################################################### ["쿠팡_누적"] 시트
        report_sheet = report_wb["쿠팡_누적"]
        key_col_dict_c = {"campaign_name": "C", "exposure_num": "D", "click_num": "E", "ad_cost": "H", "convert_num": "I", "convert_cost": "L"}
        rows = [9, 10]

        for row in rows:
            # Get the campaign name from the Excel sheet
            campaign = report_sheet[f"{key_col_dict_c['campaign_name']}{row}"].value
            
            if campaign in result_campaign.index:
                for key, col in key_col_dict_c.items():
                    if key != "campaign_name":  # Skip writing the campaign name back
                        # Populate the cell with the corresponding value from the DataFrame
                        report_sheet[f"{col}{row}"] = result_campaign.loc[campaign, key]

        key_col_dict_k = {"keyword": "C", "exposure_num": "D", "click_num": "E", "ad_cost": "H", "convert_num": "I", "convert_cost": "L"}
        start_row = 33
        for idx, row in result_by_keyword.iterrows():
            for key, col in key_col_dict_k.items():
                report_sheet[f"{col}{start_row+idx}"] = row[key]
        try:
            report_wb.save(self.output_file_path)
        except PermissionError:
            self.textEdit_log.setText(f"{self.textEdit_log.toPlainText()}{self.output_file_path}파일이 이미 존재하거나 열려있어 실패.\n")
            self.is_successed = False
        report_wb.close()
        return True

    def push_execute(self):
        rst = self.auto_report()
        self.finish_task(rst)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = WindowClass()
    myWindow.show()
    app.exec()