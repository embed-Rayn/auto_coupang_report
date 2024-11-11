from datetime import datetime, timedelta

# 문자열을 date 객체로 변환
date_string = "20241105"
date_obj = datetime.strptime(date_string, "%Y%m%d").date()
before_oneday_obj = date_obj - timedelta(days=1)
# 원하는 형식으로 변환
formatted_date = before_oneday_obj.strftime("%Y년 %m월 %d일")
print(formatted_date)

# import openpyxl
# data_xls_name = "data/20241105_data.xlsx"
# workbook = openpyxl.load_workbook(data_xls_name)
# sheet = workbook.active  # 기본 활성화된 시트를 선택합니다. 특정 시트가 있으면 sheet = workbook["Sheet1"] 과 같이 지정합니다.

# # 1행의 모든 값 읽기
# row_values = []
# for row_idx, row in enumerate(sheet):
#     if row_idx == 0:
#         continue
#     row_data = [col.value for col in row]
#     row_values.append(row_data)

# date_list = []
# campaign_list = []
# exposure_num_list = []
# click_num_list = []
# advertisement_list = []
# for row in row_values:
#     date_list.append(str(int(row[0])))
#     campaign_list.append(row[6])
#     exposure_num_list.append(int(row[7]))
#     click_num_list.append(int(row[8]))
#     advertisement_list.append(int(row[9]))
# workbook.close()
# print(date_list)
# print(campaign_list)
# print(exposure_num_list)
# print(click_num_list)
# print(advertisement_list)
# workbook.close()
print(len(set([1,1,1,1])))