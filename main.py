import openpyxl
from datetime import datetime, timedelta

def date_minus_day_to_str(date_obj, days = 1, str_format = "%Y년 %m월 %d일"):
    before_oneday_obj = date_obj - timedelta(days=days)
    formatted_date = before_oneday_obj.strftime(str_format)
    return formatted_date

####################################################
data_xls_name = "data/20241105_data.xlsx"
workbook = openpyxl.load_workbook(data_xls_name)
sheet = workbook.active  # 기본 활성화된 시트를 선택합니다. 특정 시트가 있으면 sheet = workbook["Sheet1"] 과 같이 지정합니다.

row_values = []
for row_idx, row in enumerate(sheet):
    if row_idx == 0:
        continue
    row_data = [col.value for col in row]
    row_values.append(row_data)

date_list = []
campaign_list = []
exposure_num_list = []
click_num_list = []
advertisement_list = []
for row in row_values:
    date_list.append(str(int(row[0])))
    campaign_list.append(row[6])
    exposure_num_list.append(int(row[7]))
    click_num_list.append(int(row[8]))
    advertisement_list.append(int(row[9]))
workbook.close()

assert len(set(date_list)) == 1, "날짜는 하나만..."
date_obj = datetime.strptime(date_list[0], "%Y%m%d").date()
#################################################### 리포트 파일
####################################################["요약"]
report_template_xls_name = "data/241106_dearcamp.xlsx"
report_wb = openpyxl.load_workbook(report_template_xls_name)
report_sheet = report_wb["요약"]
date_str = date_minus_day_to_str(date_obj, 1, "%Y년 %m월 %d일")
report_sheet["B31"] = date_str
date_str = date_minus_day_to_str(date_obj, 1, "%Y-%m-%d")
report_sheet["C104"] = date_str
####################################################
report_sheet = report_wb["쿠팡_일일"]
date_str = date_minus_day_to_str(date_obj, 1, "%Y년 %m월 %d일")
report_sheet["B31"] = date_str
date_str = date_minus_day_to_str(date_obj, 1, "%Y-%m-%d")
report_sheet["C104"] = date_str
####################################################
report_sheet = report_wb["쿠팡_누적"]
# 변경사항 저장
report_wb.save(report_template_xls_name)
report_wb.close()
