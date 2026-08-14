"""
为 datayes_all_SW_Industries_Level2 CSV 补充"申万一级行业"列。

映射来源：
  1. D:\CC\DB\data\申万三级和中信一级行业板块分类参考.xlsx → "二级" sheet (B列:一级, C列:二级)
  2. 若 Excel 中无映射，回退到已有的 _mapped 版本

输入: D:\CC\DB\data\datayes_all_SW_Industries_Level2.csv
输出: D:\CC\DB\data\datayes_all_SW_Industries_Level2_mapped.csv
"""
import pandas as pd
import os

# ============================================================
# 1. 加载映射表
# ============================================================

# 主线映射：Excel "二级" sheet 的 B列(申万一级行业) ← C列(申万二级行业)
xls_path = r"D:\CC\DB\data\申万三级和中信一级行业板块分类参考.xlsx"
xls_df = pd.read_excel(xls_path, sheet_name="二级")

excel_map = {}  # {申万二级行业: 申万一级行业}
for _, row in xls_df.iterrows():
    sw1, sw2 = row["申万一级行业"], row["申万二级行业"]
    if pd.isna(sw1) or pd.isna(sw2):
        continue
    excel_map[str(sw2).rstrip('Ⅱ')] = sw1

print(f"Excel 映射表加载: {len(excel_map)} 条")

# 兜底映射：已有的 _mapped 版本（覆盖 Excel 中未收录的细分行业）
old_mapped_path = r"D:\CC\DB\data\datayes_all_SW_Industries_Level2_mapped.csv"
fallback_map = {}
if os.path.exists(old_mapped_path):
    old = pd.read_csv(old_mapped_path)
    old_uniq = old[["Fetched_Ind_Name", "申万一级行业"]].dropna().drop_duplicates()
    for _, row in old_uniq.iterrows():
        clean = row["Fetched_Ind_Name"].replace("(申万)", "")
        if clean not in excel_map:
            fallback_map[clean] = row["申万一级行业"]
    print(f"兜底映射补充: {len(fallback_map)} 条")

# ============================================================
# 2. 读取源 CSV
# ============================================================
csv_path = r"D:\CC\DB\data\datayes_all_SW_Industries_Level2.csv"
print(f"\n读取 CSV: {csv_path}")
df = pd.read_csv(csv_path)
print(f"行数: {len(df)}, 列数: {len(df.columns)}")

# ============================================================
# 3. 名称清理 → 映射
# ============================================================
# Fetched_Ind_Name 格式: "半导体(申万)" → 去掉 "(申万)" → "半导体"
df["_clean"] = df["Fetched_Ind_Name"].str.replace("(申万)", "", regex=False)

# 优先 Excel，再兜底
df["申万一级行业"] = df["_clean"].map(excel_map)
mask_na = df["申万一级行业"].isna()
if mask_na.any() and fallback_map:
    df.loc[mask_na, "申万一级行业"] = df.loc[mask_na, "_clean"].map(fallback_map)

still_na = df["申万一级行业"].isna().sum()
if still_na > 0:
    print(f"\n[WARN] 仍有 {still_na} 行未映射:")
    for name in sorted(df[df["申万一级行业"].isna()]["_clean"].unique()):
        print(f"  - {name}")
else:
    print("映射完成: 100% 覆盖，0 缺失")

# ============================================================
# 4. 保存
# ============================================================
df = df.drop(columns=["_clean"])

output_path = r"D:\CC\DB\data\datayes_all_SW_Industries_Level2_mapped.csv"
df.to_csv(output_path, index=False, encoding="utf-8-sig")

print(f"\n已保存: {output_path}")
print(f"列数: {len(df.columns)} (新增: 申万一级行业)")
print(f"申万一级行业数量: {df['申万一级行业'].nunique()}")
print("\n=== 完成 ===")
