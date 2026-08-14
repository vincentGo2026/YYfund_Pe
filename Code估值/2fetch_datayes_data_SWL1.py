# -*- coding: utf-8 -*-
import os
from datetime import date

import pandas as pd
import requests

def fetch_datayes_data():
    begin_date = "20210501"
    end_date = date.today().strftime("%Y%m%d")

    # Read the industry codes from A and B columns
    industry_file = r'D:\CC\DB\data\申银万国一级行业指数代码.xlsx'
    try:
        ind_df = pd.read_excel(industry_file, usecols="A:B")
        ind_df.columns = ["代码", "行业名称"]
    except Exception as e:
        print(f"[ERROR] 读取行业指数代码失败: {e}")
        return

    # Clean the ticker codes (remove .SI, .SH etc if present)
    ind_df["代码"] = ind_df["代码"].astype(str).str.replace(r'\..*', '', regex=True)

    # Token (从环境变量读取, 避免密钥泄露到仓库)
    token = os.environ.get("DATAYES_TOKEN")
    if not token:
        print("[ERROR] 请先设置环境变量 DATAYES_TOKEN")
        return
    
    # 创建头信息,传入token
    headers = {
        "Authorization": "Bearer " + token,
        "Accept-Encoding": "gzip, deflate"
    }

    all_data = []

    for index, row in ind_df.iterrows():
        ticker = row["代码"]
        name = row["行业名称"]
        if not ticker or ticker == "nan" or str(ticker).strip() == "" or "数据来源" in str(name):
            continue

        # Dynamic URL construction
        api_url = f'/api/market/getMktIdxdSw.json?field=&beginDate={begin_date}&endDate={end_date}&ticker={ticker}&tradeDate=&exchangeCD='
    
        url = 'https://api.datayes.com/data/v1' + api_url
        print(f"正在获取 [{name} - {ticker}] 数据 ...")
        
        try:
            # 访问api, 获取数据
            res = requests.get(url, headers=headers, timeout=30)
            code = res.status_code
            
            if code == 200:
                result_json = res.json()
                if result_json.get('retCode') == 1:
                    data_list = result_json['data']
                    if data_list:
                        df_part = pd.DataFrame(data_list)
                        df_part['Fetched_Ind_Name'] = name
                        all_data.append(df_part)
                        print(f"  [OK] 成功获取 {len(df_part)} 行记录")
                    else:
                        print(f"  [WARN] 无数据返回")
                else:
                    print("  [ERROR] 请求成功但API返回错误:", result_json.get('retMsg'))
            else:
                print("  [ERROR] HTTP请求失败，状态码:", code)
        except Exception as e:
            print(f"  [ERROR] 发生异常: {e}")

    if all_data:
        final_df = pd.concat(all_data, ignore_index=True)
        # 顺便导出到 CSV 方便查看
        output_path = rf"D:\CC\DB\data\datayes_all_SW_Industries_{begin_date}_{end_date}.csv"
        final_df.to_csv(output_path, index=False, encoding='utf-8-sig')
        print(f"\n[SAVED] 所有行业数据已成功保存至: {output_path} (共 {len(final_df)} 行)")
    else:
        print("\n[ERROR] 未能成功获取任何行业数据。")

if __name__ == "__main__":
    fetch_datayes_data()
