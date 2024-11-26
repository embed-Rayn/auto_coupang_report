from datetime import datetime, timedelta
import openpyxl
import pandas as pd


excel_path = r"F:\2.code\git_repo\auto_coupang_report\라르츠 제작 샘플 파일 전달 - 복사본\A01217844_pa_daily_campaign_20241121_20241121_포켓캐디.xlsx"
excel_path = r"larz2\A01217844_pa_daily_campaign_20241122_20241124 (6)_포켓캐디.xlsx"
df2 = pd.read_excel(excel_path)
df2['날짜'] = pd.to_datetime(df2['날짜'], format='%Y%m%d')
df2 = df2[["날짜", "노출수", "클릭수", "광고비", "총 판매수량(14일)", "총 전환매출액(14일)"]] # "캠페인명"


excel_path = r"F:\2.code\git_repo\auto_coupang_report\라르츠 제작 샘플 파일 전달 - 복사본\A01217844_pa_daily_campaign_20241121_20241121_포켓캐디X.xlsx"
excel_path = r"larz2\A01217844_pa_daily_campaign_20241122_20241124 (7)_포켓캐디X.xlsx"
df1 = pd.read_excel(excel_path)
df1['날짜'] = pd.to_datetime(df1['날짜'], format='%Y%m%d')
#df1 = df1[["날짜", "노출수", "클릭수", "광고비", "총 판매수량(14일)", "총 전환매출액(14일)"]] # "캠페인명"
combined_df = pd.concat([df1, df2], ignore_index=True)
df2 = df2.groupby('날짜', as_index=False).sum()
print(df2)
df1 = df1.groupby('날짜', as_index=False).sum()
print(df1)


print("--------------")
print(combined_df)
xls_path = "larz2/241122_라르츠 거리측정기_오픈마켓_보고서(이베이 11번가 쿠팡).xlsx"
report_wb = openpyxl.load_workbook(xls_path)
report_sheet = report_wb["쿠팡_일일"]
df_daily_x = df1.groupby('날짜', as_index=False).sum()
df_daily_x = df_daily_x[["날짜", "캠페인명", "노출수", "클릭수", "광고비", "총 판매수량(14일)", "총 전환매출액(14일)"]]
df_daily_x.columns = ["날짜", "캠페인명", "노출수", "클릭수", "광고비", "전환수", "전환매출"]
key_col_dict_x = {"노출수": "AF", "클릭수": "AG", "광고비": "AJ", "전환수": "AK", "전환매출": "AN"}
print("--------------")
print(report_sheet.columns)
print(report_sheet.column_groups)
print(report_sheet.col_breaks)
ad_column = [report_sheet[f"P{row}"].value for row in range(20, 60)]
print(ad_column)
ad_column = [report_sheet[f"AD{row}"].value for row in range(20, 60)]
print(ad_column)
# for data_row_idx, row in df_daily_x.iterrows():
#     date_obj = row['날짜']
#     for row_idx, cell in enumerate(report_sheet["AD"]):
#         try:
#             cv_date = cell.value.date()
            
#             if cv_date.strftime("%y-%m-%d") == date_obj.strftime("%y-%m-%d"):
#                 break
#         except AttributeError:
#             continue
#     data_idx = row_idx + 1
#     for key, col in key_col_dict_x.items():
#         report_sheet[f"{col}{data_idx}"] = row[key]
#         print(f"{col}{data_idx}", row[key])