from datetime import datetime, timedelta
import openpyxl
import pandas as pd


df2 = pd.read_excel(r"F:\2.code\git_repo\auto_coupang_report\라르츠 제작 샘플 파일 전달 - 복사본\A01217844_pa_daily_campaign_20241121_20241121_포켓캐디.xlsx")
df2['날짜'] = pd.to_datetime(df2['날짜'], format='%Y%m%d')
df2 = df2[["날짜", "캠페인명", "노출수", "클릭수", "광고비", "총 판매수량(14일)", "총 전환매출액(14일)"]]
print(df2)
df1 = pd.read_excel(r"F:\2.code\git_repo\auto_coupang_report\라르츠 제작 샘플 파일 전달 - 복사본\A01217844_pa_daily_campaign_20241121_20241121_포켓캐디X.xlsx")
df1['날짜'] = pd.to_datetime(df1['날짜'], format='%Y%m%d')
df1 = df1[["날짜", "캠페인명", "노출수", "클릭수", "광고비", "총 판매수량(14일)", "총 전환매출액(14일)"]]
print(df1)
