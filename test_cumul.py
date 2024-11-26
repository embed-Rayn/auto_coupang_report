from datetime import datetime, timedelta
import openpyxl
import pandas as pd


class Test():
    def func_xl(self):
        data_xls_name = "larz/A01217844_pa_total_campaign_20241101_20241121.xlsx"
        df_monthly_total = pd.read_excel(data_xls_name)
        df_monthly_total = df_monthly_total[["광고집행 상품명", "키워드", "노출수", "클릭수", "광고비", "총 판매수량(14일)", "총 전환매출액(14일)"]]
        df_monthly_total.columns = ["상품명", "키워드", "노출수", "클릭수", "광고비", "전환수", "전환매출"]

        print(df_monthly_total.head())
        summary_product = df_monthly_total.groupby("상품명")[["노출수", "클릭수", "광고비", "전환수", "전환매출"]].sum()
        print(summary_product)
        contains_x = df_monthly_total[df_monthly_total['상품명'].str.contains("X", na=False)]
        contains_o = df_monthly_total[~df_monthly_total['상품명'].str.contains("X", na=False)]
        print(contains_o.head())
        print(contains_x.head())
        print("="*50)
        contains_o = contains_o.groupby("키워드")[["노출수", "클릭수", "광고비", "전환수", "전환매출"]].sum()
        contains_o = contains_o.sort_values(by=["전환수", "광고비"], ascending=[False, False]).reset_index()
        contains_x = contains_x.groupby("키워드")[["노출수", "클릭수", "광고비", "전환수", "전환매출"]].sum()
        contains_x = contains_x.sort_values(by=["전환수", "광고비"], ascending=[False, False]).reset_index()
        print(contains_o.head())
        print(len(contains_o))
        print(contains_x.head())
        print(len(contains_x))
        test_wb = "larz/241121_라르츠 거리측정기_오픈마켓_보고서(이베이 11번가 쿠팡).xlsx"
        # wb = openpyxl.load_workbook(test_wb)  # Replace with your Excel file name
        # report_sheet = wb.active  # Replace with the correct sheet name if needed

        # key_col_dict_c = {"campaign_name": "C", "exposure_num": "D", "click_num": "E", "ad_cost": "H", "convert_num": "I", "convert_cost": "L"}
        # rows = [9, 10]

        # for row in rows:
        #     # Get the campaign name from the Excel sheet
        #     campaign = report_sheet[f"{key_col_dict_c['campaign_name']}{row}"].value
            
        #     if campaign in result_campaign.index:
        #         for key, col in key_col_dict_c.items():
        #             if key != "campaign_name":  # Skip writing the campaign name back
        #                 # Populate the cell with the corresponding value from the DataFrame
        #                 report_sheet[f"{col}{row}"] = result_campaign.loc[campaign, key]


        # key_col_dict_k = {"keyword": "C", "exposure_num": "D", "click_num": "E", "ad_cost": "H", "convert_num": "I", "convert_cost": "L"}
        # start_row = 33
        # for idx, row in result_by_keyword.iterrows():
        #     for key, col in key_col_dict_k.items():
        #         report_sheet[f"{col}{start_row+idx}"] = row[key]

        # # Save the updated Excel file
        # wb.save(test_wb)  # Replace with your desired file name

        print("complete")
t = Test()
t.func_xl()