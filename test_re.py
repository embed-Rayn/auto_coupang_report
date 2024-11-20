as_dict = {
	"(자동) 해먹": [
		{"exposure_num": 215572.0, "click_num": 221.0, "ad_cost": 47999.0, "convert_num": 2.0, "convert_cost": 198000.0}, 
		{"exposure_num": 236109.0, "click_num": 238.0, "ad_cost": 54251.0, "convert_num": 4.0, "convert_cost": 393030.0}, 
		{"exposure_num": 175017.0, "click_num": 212.0, "ad_cost": 52535.0, "convert_num": 2.0, "convert_cost": 195030.0}, 
		{"exposure_num": 218913.0, "click_num": 172.0, "ad_cost": 33837.0, "convert_num": 0.0, "convert_cost": 0.0}, 
		{"exposure_num": 230412.0, "click_num": 214.0, "ad_cost": 38665.0, "convert_num": 0.0, "convert_cost": 0.0}
	], 
	"(수동) 해먹": [
		{"exposure_num": 36916.0, "click_num": 87.0, "ad_cost": 36208.0, "convert_num": 1.0, "convert_cost": 99000.0}, 
		{"exposure_num": 17057.0, "click_num": 64.0, "ad_cost": 25254.0, "convert_num": 0.0, "convert_cost": 0.0}, 
		{"exposure_num": 88878.0, "click_num": 154.0, "ad_cost": 50970.0, "convert_num": 0.0, "convert_cost": 0.0}, 
		{"exposure_num": 11629.0, "click_num": 42.0, "ad_cost": 14803.0, "convert_num": 0.0, "convert_cost": 0.0}, 
		{"exposure_num": 6823.0, "click_num": 55.0, "ad_cost": 35246.0, "convert_num": 0.0, "convert_cost": 0.0}
	]
}
result = {}

# 최상위 key별 합계 계산
for key, records in as_dict.items():
    # 초기화
    total = {
        "exposure_num": 0,
        "click_num": 0,
        "ad_cost": 0,
        "convert_num": 0,
        "convert_cost": 0
    }
    # 각 항목의 값을 합산
    for record in records:
        for sub_key in total:
            total[sub_key] += int(record[sub_key])
    # 결과 저장
    result[key] = total

# 결과 출력
for key, totals in result.items():
    print(f"{key}: {totals}")