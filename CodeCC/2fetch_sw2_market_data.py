# -*- coding: utf-8 -*-
"""
参照 2fetch_datayes_data_SWL2，提取申万二级行业指数行情（getMktIdxd.json）
输出：D:\CC\DB\data\sw2_market_data_{begin_date}_{end_date}.csv
"""

import pandas as pd
import requests
import os
import sys
from datetime import date

# ============ 配置 ============
# Token (从环境变量读取, 避免密钥泄露到仓库)
TOKEN = os.environ.get("DATAYES_TOKEN")
if not TOKEN:
    print("[ERROR] 请先设置环境变量 DATAYES_TOKEN")
    sys.exit(1)
BEGIN_DATE = "20210501"
END_DATE = date.today().strftime("%Y%m%d")

# 行业代码文件（从已有行情数据提取的 ticker→名称 映射）
TICKER_MAPPING = r"D:\CC\DB\data\sw2_ticker_mapping.csv"
# 输出目录
OUT_DIR = r"D:\CC\DB\data"

HEADERS = {
    "Authorization": "Bearer " + TOKEN,
    "Accept-Encoding": "gzip, deflate"
}

# ============ 读取行业代码 ============
try:
    ind_df = pd.read_csv(TICKER_MAPPING, dtype={"ticker": str}, encoding="gbk")
    ind_df.columns = ["代码", "行业名称"]
    print(f"读取行业代码: {len(ind_df)} 个二级行业")
except Exception as e:
    print(f"读取行业代码失败: {e}")
    sys.exit(1)

# ============ 循环请求 ============
all_data = []

for idx, row in ind_df.iterrows():
    ticker = row["代码"]
    name = row["行业名称"]
    if not ticker or str(ticker).strip() == "" or ticker == "nan":
        continue

    # getMktIdxd.json: ticker + exchangeCD=ZICN
    api_url = (
        f"/api/market/getMktIdxd.json?"
        f"field=&beginDate={BEGIN_DATE}&endDate={END_DATE}"
        f"&indexID=&ticker={ticker}&exchangeCD=ZICN&tradeDate="
    )
    url = "https://api.datayes.com/data/v1" + api_url
    print(f"正在获取 [{name} - {ticker}] ...", end=" ")

    try:
        res = requests.get(url, headers=HEADERS)
        code = res.status_code
        if code == 200:
            result_json = res.json()
            if result_json.get("retCode") == 1:
                data_list = result_json["data"]
                if data_list:
                    df_part = pd.DataFrame(data_list)
                    df_part["申万二级行业"] = name
                    all_data.append(df_part)
                    print(f"✅ {len(df_part)} 行")
                else:
                    print("⚠️ 无数据")
            else:
                print(f"❌ API 错误: {result_json.get('retMsg')}")
        else:
            print(f"❌ HTTP {code}")
    except Exception as e:
        print(f"❌ 异常: {e}")

# ============ 输出 ============
if all_data:
    final_df = pd.concat(all_data, ignore_index=True)
    os.makedirs(OUT_DIR, exist_ok=True)
    output_path = os.path.join(
        OUT_DIR,
        f"sw2_market_data_{BEGIN_DATE}_{END_DATE}.csv"
    )
    final_df.to_csv(output_path, index=False, encoding="utf-8-sig")
    print(f"\n完成! {len(ind_df['代码'].unique())} 个行业, {len(final_df)} 行")
    print(f"输出: {output_path}")
else:
    print("\n未获取到任何数据。")
