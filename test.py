from datetime import datetime, timedelta
import openpyxl


# 문자열을 date 객체로 변환
date_string = "20241105"
date_obj = datetime.strptime(date_string, "%Y%m%d").date()
before_oneday_obj = date_obj - timedelta(days=1)
# 원하는 형식으로 변환
formatted_date = before_oneday_obj.strftime("%Y년 %m월 %d일")
print(formatted_date)



data_xls_name = "data/20241105_data.xlsx"
# data_xls_name = "data/8910.xlsx"
workbook = openpyxl.load_workbook(data_xls_name)
sheet = workbook.active  # 기본 활성화된 시트를 선택합니다. 특정 시트가 있으면 sheet = workbook["Sheet1"] 과 같이 지정합니다.

row_values = []
for row_idx, row in enumerate(sheet):
    if row_idx == 0:
        continue
    row_data = [col.value for col in row]
    row_values.append(row_data)

dataset = {}
for row in row_values:
    date_obj = datetime.strptime(str(int(row[0])), "%Y%m%d").date()
    temp = {}
    temp["campaign"] = row[6]
    temp["exposure_num"] = row[7]
    temp["click_num"] = row[8]
    temp["ad_cost"] = row[9]
    temp["convert_num"] = row[15]
    temp["convert_cost"] = row[18]
    try:
        dataset[date_obj].append(temp)
    except KeyError:
        dataset[date_obj] = [temp]
workbook.close()

print(dataset)
report_day = max(dataset.keys())
# date_obj = datetime.strptime(report_day, "%Y%m%d").date()
date_obj = report_day
print(date_obj, type(date_obj))

print(sum(campaign["exposure_num"] for campaign in dataset[date_obj]))
print(campaign["exposure_num"] for campaign in dataset[date_obj])


# report_template_xls_name = "data/241106_dearcamp.xlsx"
# report_wb = openpyxl.load_workbook(report_template_xls_name)

# #########################
# report_sheet = report_wb["쿠팡_일일"]
# before_oneday_obj = date_obj - timedelta(days=1)
# def asdf(before_oneday_obj):
#     for row_idx, cell in enumerate(report_sheet["B"]):
#         try:
#             cv = cell.value.date()
#         except AttributeError:
#             continue
#         if cv == before_oneday_obj:
                    
#             break
