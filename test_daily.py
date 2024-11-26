from datetime import datetime, timedelta
import openpyxl
import pandas as pd


excel_path = r"F:\2.code\git_repo\auto_coupang_report\라르츠 제작 샘플 파일 전달 - 복사본\A01217844_pa_daily_campaign_20241121_20241121_포켓캐디.xlsx"
excel_path = r"C:\Users\RyanPark\codeview\auto_coupang_report\larz\A01217844_pa_daily_campaign_20241121_20241121_포켓캐디.xlsx"
df2 = pd.read_excel(excel_path)
df2['날짜'] = pd.to_datetime(df2['날짜'], format='%Y%m%d')
df2 = df2[["날짜", "노출수", "클릭수", "광고비", "총 판매수량(14일)", "총 전환매출액(14일)"]] # "캠페인명"


excel_path = r"F:\2.code\git_repo\auto_coupang_report\라르츠 제작 샘플 파일 전달 - 복사본\A01217844_pa_daily_campaign_20241121_20241121_포켓캐디X.xlsx"
excel_path = r"C:\Users\RyanPark\codeview\auto_coupang_report\larz\A01217844_pa_daily_campaign_20241121_20241121_포켓캐디X.xlsx"
df1 = pd.read_excel(excel_path)
df1['날짜'] = pd.to_datetime(df1['날짜'], format='%Y%m%d')
df1 = df1[["날짜", "노출수", "클릭수", "광고비", "총 판매수량(14일)", "총 전환매출액(14일)"]] # "캠페인명"
combined_df = pd.concat([df1, df2], ignore_index=True)
df2 = df2.groupby('날짜', as_index=False).sum()
print(df2)
df1 = df1.groupby('날짜', as_index=False).sum()
print(df1)


print("--------------")
print(combined_df)