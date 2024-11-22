from datetime import datetime, timedelta
import openpyxl
import pandas as pd


class Test():
    def func_xl(self):
        data_xls_name = "data_2.1/A01217844_pa_total_campaign_20241101_20241120.xlsx"
        workbook_cum = openpyxl.load_workbook(data_xls_name)
        sheet = workbook_cum.active 
        row_values_cumul = []
        for row_idx, row in enumerate(sheet):
            if row_idx == 0:
                continue
            row_data = [col.value for col in row]
            row_values_cumul.append(row_data)
        workbook_cum.close()


        dataset_cumul = []
        for row in row_values_cumul:
            temp = {}
            temp["campaign_name"] = row[4]
            if row[11]:
                keyword = row[11]
            else:
                keyword = "비검색 영역"
            temp["keyword"] = keyword
            temp["exposure_num"] = row[12]
            temp["click_num"] = row[13]
            temp["ad_cost"] = row[14]
            temp["convert_num"] = row[28]
            temp["convert_cost"] = row[31]
            dataset_cumul.append(temp)
        df = pd.DataFrame(dataset_cumul)

        result_campaign = df.groupby("campaign_name")[["exposure_num", "click_num", "ad_cost", "convert_num", "convert_cost"]].sum()

        # Group by "keyword" and calculate the sum for the selected columns
        result_by_keyword = df.groupby("keyword")[["exposure_num", "click_num", "ad_cost", "convert_num", "convert_cost"]].sum()
        result_by_keyword = result_by_keyword.sort_values(by=["ad_cost", "convert_num"], ascending=[False, False]).reset_index()

        from openpyxl import load_workbook

        # Assuming `result_campaign` is the DataFrame with campaign-wise sums
        result_campaign = df.groupby("campaign_name")[["exposure_num", "click_num", "ad_cost", "convert_num", "convert_cost"]].sum()

        # Load the workbook and get the sheet
        test_wb = "data_2.1/test.xlsx"
        wb = load_workbook(test_wb)  # Replace with your Excel file name
        report_sheet = wb.active  # Replace with the correct sheet name if needed

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

        # Save the updated Excel file
        wb.save(test_wb)  # Replace with your desired file name


        # cumul_summary = {}
        # for key, records in dataset_cumul.items():
        #     total = {
        #         "exposure_num": 0,
        #         "click_num": 0,
        #         "ad_cost": 0,
        #         "convert_num": 0,
        #         "convert_cost": 0
        #     }
        #     for record in records:
        #         for sub_key in total:
        #             total[sub_key] += int(record[sub_key])
        #     cumul_summary[key] = total


        #####################
        # output_file_path = "./data_2.1/241119_디어캠프_오픈마켓_보고서(이베이 11번가 쿠팡).xlsx"
        # report_wb = openpyxl.load_workbook(output_file_path)
        # report_sheet = report_wb["쿠팡_누적"]
        # key_col_dict = {"exposure_num": "D", "click_num": "E", "ad_cost": "H", "convert_num": "I", "convert_cost": "L"}
        # rows = [9, 10]
        # for row in rows:
        #     campaign = report_sheet[f"C{row}"].value
        #     for key, col in key_col_dict.items():
        #         report_sheet[f"{col}{row}"] = dataset_cumul[campaign][key]

        print("complete")
t = Test()
t.func_xl()