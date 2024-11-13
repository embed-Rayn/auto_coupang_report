import openpyxl
from datetime import datetime, timedelta

def date_minus_day_to_str(date_obj, days = 1, str_format = "%Y년 %m월 %d일"):
    before_oneday_obj = date_obj - timedelta(days=days)
    formatted_date = before_oneday_obj.strftime(str_format)
    return formatted_date

def das():
    return chr
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

report_day = max(dataset.keys())
#################################################### 리포트 파일
#################################################### ["요약"] 시트
report_template_xls_name = "data/241106_dearcamp.xlsx"
report_wb = openpyxl.load_workbook(report_template_xls_name)
report_sheet = report_wb["요약"]
date_str = date_minus_day_to_str(report_day, 1, "%Y년 %m월 %d일")
report_sheet["B31"] = date_str
date_str = date_minus_day_to_str(report_day, 1, "%Y-%m-%d")
report_sheet["C104"] = date_str
#################################################### ["쿠팡_일일"] 시트
report_sheet = report_wb["쿠팡_일일"]
for date_obj in dataset.keys():
    before_oneday_obj = date_obj - timedelta(days=1)
    for row_idx, cell in enumerate(report_sheet["B"]):
        try:
            cv_date = cell.value.date()
        except AttributeError:
            continue
        if cv_date == before_oneday_obj:
            data_idx = row_idx + 1
            report_sheet[f"D{data_idx}"] = sum(campaign["exposure_num"] for campaign in dataset[date_obj])
            report_sheet[f"E{data_idx}"] = sum(campaign["click_num"] for campaign in dataset[date_obj])
            report_sheet[f"H{data_idx}"] = sum(campaign["ad_cost"] for campaign in dataset[date_obj])
            report_sheet[f"I{data_idx}"] = sum(campaign["convert_num"] for campaign in dataset[date_obj])
            report_sheet[f"L{data_idx}"] = sum(campaign["convert_cost"] for campaign in dataset[date_obj])
            # print(report_sheet[f"D{data_idx}"].value, report_sheet[f"E{data_idx}"].value, 
            #       report_sheet[f"H{data_idx}"].value, report_sheet[f"I{data_idx}"].value, 
            #       report_sheet[f"L{data_idx}"].value)
#################################################### ["쿠팡_누적"] 시트
report_sheet = report_wb["쿠팡_누적"]
campaign_totals = {}
for campaigns in dataset.values():
    for campaign in campaigns:
        campaign_name = campaign["campaign"]
        exposure_num = campaign["exposure_num"]
        click_num = campaign["click_num"]
        ad_cost = campaign["ad_cost"]
        convert_num = campaign["convert_num"]
        convert_cost = campaign["convert_cost"]
        metric = [exposure_num, click_num, ad_cost, convert_num, convert_cost]
        if campaign_name in campaign_totals:
            campaign_totals[campaign_name] = [sum(x) for x in zip(campaign_totals[campaign_name], metric)]
        else:
            campaign_totals[campaign_name] = metric

c1 = report_sheet["C9"].value
metrics = campaign_totals[c1]
cols = ["C", "D", "E", "H", "I"]
rows = [9, 10]
for row in rows:
    for idx, col in enumerate(cols):
        cell = report_sheet[f"{col}{row}"]
        cell = cell.value + metrics[idx]


# 변경사항 저장
report_wb.save(report_template_xls_name)
report_wb.close()
