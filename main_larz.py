import openpyxl
from datetime import datetime, timedelta
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
        self.setWindowTitle("Auto Coupang report - 라르츠")

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
        self.textEdit_log.setText(f"{self.textEdit_log.toPlainText()}일일 데이터(캐디O) 위치: {self.input_file_path_1}\n")
        self.lineEdit_excel_1.setText(f"{self.input_file_path_1}")


    def browse_2(self):
        try:
            file_filter = "Excel Files (*.xlsx)"
            path, _ = QFileDialog.getOpenFileName(None, 'Select an Excel File', '', file_filter)
            self.input_file_path_2 = path
        except Exception as e:
            print(e)
            self.input_file_path_2 = ""
        self.textEdit_log.setText(f"{self.textEdit_log.toPlainText()}일일 데이터(캐디X) 엑셀 위치: {self.input_file_path_2}\n")
        self.lineEdit_excel_2.setText(f"{self.input_file_path_2}")


    def browse_3(self):
        try:
            file_filter = "Excel Files (*.xlsx)"
            path, _ = QFileDialog.getOpenFileName(None, 'Select an Excel File', '', file_filter)
            self.input_file_path_3 = path
        except Exception as e:
            print(e)
            self.input_file_path_3 = ""
        self.textEdit_log.setText(f"{self.textEdit_log.toPlainText()}월간 데이터 엑셀 위치: {self.input_file_path_3}\n")
        self.lineEdit_excel_3.setText(f"{self.input_file_path_3}")


    def browse_4(self):
        try:
            file_filter = "Excel Files (*.xlsx)"
            path, _ = QFileDialog.getOpenFileName(None, 'Select an Excel File', '', file_filter)
            self.input_file_path_4 = path
        except Exception as e:
            print(e)
            self.input_file_path_4 = ""
        self.textEdit_log.setText(f"{self.textEdit_log.toPlainText()}적용할 리포트 엑셀 위치: {self.input_file_path_4}\n")
        self.lineEdit_excel_4.setText(f"{self.input_file_path_4}")


    def finish_task(self, rst):
        if rst and self.is_successed:
            self.textEdit_log.setText(f"{self.textEdit_log.toPlainText()}작업이 완료 됐습니다.\n생성한 리포트 파일 위치: {self.output_file_path}\n")
        else:
            self.textEdit_log.setText(f"{self.textEdit_log.toPlainText()}{self.input_file_path_1}sheet가 여러개 입니다. 실행 종료\n")
        self.textEdit_log.setText(f"{self.textEdit_log.toPlainText()}===============================================================================================\n")
        self.is_successed = True


    def process_file_path(self, report_day):
        pattern = r"\/\d{6}_"  # 6자리 숫자 (YYMMDD) 패턴
        match = re.search(pattern, self.input_file_path_4)

        if match:
            old_date = match.group()
            new_date = report_day.strftime("%y%m%d")
            self.output_file_path = self.input_file_path_4.replace(old_date, f"/{new_date}_")


    def auto_report(self):
        df_daily_o = pd.read_excel(self.input_file_path_1)
        df_daily_x = pd.read_excel(self.input_file_path_2)
        df_monthly_total = pd.read_excel(self.input_file_path_3)
        df_daily_o['날짜'] = pd.to_datetime(df_daily_o['날짜'], format='%Y%m%d')
        df_daily_x['날짜'] = pd.to_datetime(df_daily_x['날짜'], format='%Y%m%d')
        df_daily_total = pd.concat([df_daily_o, df_daily_x], ignore_index=True)
        report_day = df_daily_total['날짜'].max()
        file_day = report_day + timedelta(days=1)
        self.process_file_path(file_day)
        try:
            shutil.copy(self.input_file_path_4, self.output_file_path)
        except PermissionError:
            self.textEdit_log.setText(f"{self.textEdit_log.toPlainText()}{self.output_file_path}파일이 이미 존재하거나 열려있어 실패.\n")
            self.is_successed = False
        report_wb = openpyxl.load_workbook(self.output_file_path)
        #################################################### 리포트 파일
        #################################################### ["쿠팡_일일"] 시트
        report_sheet = report_wb["쿠팡_일일"] 
        if not df_daily_o.empty:
            df_daily_o = df_daily_o[["날짜", "캠페인명", "노출수", "클릭수", "광고비", "총 판매수량(14일)", "총 전환매출액(14일)"]]
            df_daily_o.columns = ["날짜", "캠페인명", "노출수", "클릭수", "광고비", "전환수", "전환매출"]
            df_daily_o = df_daily_o.groupby('날짜', as_index=False).sum()
            key_col_dict_o = {"노출수": "R", "클릭수": "S", "광고비": "V", "전환수": "W", "전환매출": "Z"}
            for data_row_idx, row in df_daily_o.iterrows():
                date_obj = row['날짜']
                for row_idx, cell in enumerate(report_sheet["P"]):
                    try:
                        cv_date = cell.value.date()
                        if cv_date.strftime("%y-%m-%d") == date_obj.strftime("%y-%m-%d"):
                            break
                    except AttributeError:
                        continue
                data_idx = row_idx + 1
                for key, col in key_col_dict_o.items():
                    report_sheet[f"{col}{data_idx}"] = row[key]
        else:
            self.textEdit_log.setText(f"{self.textEdit_log.toPlainText()}{self.input_file_path_1}일일 데이터(캐디O) 파일이 비어있습니다.\n")

        if not df_daily_x.empty:
            df_daily_x = df_daily_x[["날짜", "캠페인명", "노출수", "클릭수", "광고비", "총 판매수량(14일)", "총 전환매출액(14일)"]]
            df_daily_x.columns = ["날짜", "캠페인명", "노출수", "클릭수", "광고비", "전환수", "전환매출"]
            df_daily_x = df_daily_x.groupby('날짜', as_index=False).sum()
            key_col_dict_x = {"노출수": "AF", "클릭수": "AG", "광고비": "AJ", "전환수": "AK", "전환매출": "AN"}
            for data_row_idx, row in df_daily_x.iterrows():
                date_obj = row['날짜']
                for row_idx, cell in enumerate(report_sheet["AD"]):
                    try:
                        cv_date = cell.value.date()
                        if cv_date.strftime("%y-%m-%d") == date_obj.strftime("%y-%m-%d"):
                            break
                    except AttributeError:
                        continue
                data_idx = row_idx + 1
                for key, col in key_col_dict_x.items():
                    report_sheet[f"{col}{data_idx}"] = row[key]
        else:
            self.textEdit_log.setText(f"{self.textEdit_log.toPlainText()}{self.input_file_path_2}일일 데이터(캐디X) 파일이 비어있습니다.\n")

        df_daily_total = pd.concat([df_daily_o, df_daily_x], ignore_index=True)
        if not df_daily_total.empty:
            df_daily_total = df_daily_total.groupby('날짜', as_index=False).sum()
            key_col_dict_tot = {"노출수": "D", "클릭수": "E", "광고비": "H", "전환수": "I", "전환매출": "L"}
            for data_row_idx, row in df_daily_total.iterrows():
                date_obj = row['날짜']
                for row_idx, cell in enumerate(report_sheet["B"]):
                    try:
                        cv_date = cell.value.date()
                        
                        if cv_date.strftime("%y-%m-%d") == date_obj.strftime("%y-%m-%d"):
                            break
                    except AttributeError:
                        continue
                data_idx = row_idx + 1
                for key, col in key_col_dict_tot.items():
                    report_sheet[f"{col}{data_idx}"] = row[key]
        #################################################### ["요약"] 시트
        report_sheet = report_wb["요약"]
        report_sheet["C110"] = report_day
        report_sheet["C110"].number_format = "YYYY-MM-DD"
        #################################################### ["쿠팡_누적"] 시트
        report_sheet = report_wb["쿠팡_누적"]
        if not df_monthly_total.empty:
            df_monthly_total = df_monthly_total[["광고집행 상품명", "키워드", "노출수", "클릭수", "광고비", "총 판매수량(14일)", "총 전환매출액(14일)"]]
            df_monthly_total.columns = ["상품명", "키워드", "노출수", "클릭수", "광고비", "전환수", "전환매출"]
            df_monthly_total["키워드"] = df_monthly_total["키워드"].fillna("비검색 영역")
            
            summary_product = df_monthly_total.groupby("상품명")[["노출수", "클릭수", "광고비", "전환수", "전환매출"]].sum()
            summary_product = summary_product.sort_values(by=["전환수", "광고비"], ascending=[False, False]).reset_index()
            contains_x = df_monthly_total[df_monthly_total['상품명'].str.contains("X", na=False)]
            contains_o = df_monthly_total[~df_monthly_total['상품명'].str.contains("X", na=False)]

            contains_o = contains_o.groupby("키워드")[["노출수", "클릭수", "광고비", "전환수", "전환매출"]].sum()
            contains_o = contains_o.sort_values(by=["전환수", "광고비"], ascending=[False, False]).reset_index()
            contains_x = contains_x.groupby("키워드")[["노출수", "클릭수", "광고비", "전환수", "전환매출"]].sum()
            contains_x = contains_x.sort_values(by=["전환수", "광고비"], ascending=[False, False]).reset_index()
            summary_start_row = 9
            key_col_dict_s = {"상품명": "C", "노출수": "D", "클릭수": "E", "광고비": "H", "전환수": "I", "전환매출": "L"}
            for data_row_idx, data_row in summary_product.iterrows():
                for key, col in key_col_dict_s.items():
                    report_sheet[f"{col}{summary_start_row+data_row_idx}"] = data_row[key]

            start_row = 36
            key_col_dict_o = {"키워드": "C", "노출수": "D", "클릭수": "E", "광고비": "H", "전환수": "I", "전환매출": "L"}
            for data_row_idx, data_row in contains_o.iterrows():
                for key, col in key_col_dict_o.items():
                    report_sheet[f"{col}{start_row+data_row_idx}"] = data_row[key]

            key_col_dict_x = {"키워드": "Q", "노출수": "R", "클릭수": "S", "광고비": "V", "전환수": "W", "전환매출": "Z"}
            for data_row_idx, data_row in contains_x.iterrows():
                for key, col in key_col_dict_x.items():
                    report_sheet[f"{col}{start_row+data_row_idx}"] = data_row[key]
        else:
            self.textEdit_log.setText(f"{self.textEdit_log.toPlainText()}{self.input_file_path_3} 파일이 비어있습니다.\n")

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