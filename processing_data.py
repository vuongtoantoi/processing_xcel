"""
Nhân sự 1: Xử lý và chuẩn hóa file Viettel - KPI - Telecom.xlsx
"""
 

import pandas as pd
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
 
INPUT_FILE = BASE_DIR / "Viettel - KPI - Telecom.xlsx"
OUTPUT_FILE_XLSX = BASE_DIR / "Viettel_-_KPI_-_Telecom_normalized.xlsx"
OUTPUT_FILE_CSV = BASE_DIR / "Viettel_-_KPI_-_Telecom_normalized.csv"

# %% 3. Đọc dữ liệu
df = pd.read_excel(INPUT_FILE, sheet_name="KPI")
print("Số dòng, số cột:", df.shape)
print(df.head())
 
# %% 4. Đổi tên cột
RENAME_MAP = {
    "KPI_001": "BEARER_MME_UTIL",
    "KPI_002": "BEARER_4G",
    "KPI_003": "CREATE_DEDICATED_BEARER_SR",
    "KPI_004": "CSFB_SR",
    "KPI_005": "DEDICATED_BEARER_MME",
    "KPI_006": "EASR",
    "KPI_007": "INTRA_TAU_SR",
    "KPI_008": "PGW_BEARER_SR",
    "KPI_009": "PGW_BEARER_UTIL",
    "KPI_010": "S1_INTRA_HANDOVER_SR",
    "KPI_011": "S1_PAGGING_SR",
    "KPI_012": "SAU_4G",
    "KPI_013": "SAU_UTIL_4G",
    "KPI_014": "SERVICE_REQUEST_SR",
    "KPI_015": "SGW_BEARER_SR",
    "KPI_016": "SG_LU_SR",
    "KPI_017": "SMS_PAGGING_SR",
    "KPI_018": "THROUGHPUT_4G",
    "KPI_019": "THROUGHPUT_UTIL",
    "KPI_020": "NO_PGW_IMS_BEARER",
    "KPI_021": "NO_PGW_IMS_SUBS",
    "KPI_022": "NO_PGW_SUBS",
}
df = df.rename(columns=RENAME_MAP)
print("\nCột sau khi đổi tên:")
print(df.columns.tolist())
 
# %% 5. Chuẩn hóa cột SYS_DATETIME
# 5.1 Ép kiểu datetime chuẩn (nếu có giá trị lỗi/text sẽ thành NaT thay vì crash)
df["SYS_DATETIME"] = pd.to_datetime(df["SYS_DATETIME"], errors="coerce")
 
# Kiểm tra có dòng nào parse lỗi không
n_bad = df["SYS_DATETIME"].isna().sum()
print(f"\nSố dòng SYS_DATETIME không hợp lệ (NaT): {n_bad}")
if n_bad > 0:
    print(df[df["SYS_DATETIME"].isna()])
 
# 5.2 Tách thêm cột Ngày / Giờ để dễ lọc và pivot
df["SYS_DATE"] = df["SYS_DATETIME"].dt.date
df["SYS_TIME"] = df["SYS_DATETIME"].dt.time
 
# 5.3 Kiểm tra trùng lặp theo (NE_NAME, SYS_DATETIME)
dup_mask = df.duplicated(subset=["NE_NAME", "SYS_DATETIME"], keep=False)
df_dup = df[dup_mask].sort_values(["NE_NAME", "SYS_DATETIME"])
print(f"\nSố dòng bị trùng (NE_NAME, SYS_DATETIME): {len(df_dup)}")
 
# 5.4 Sắp xếp lại theo thiết bị rồi theo thời gian: có nghĩa là để xem device đó xuất hiện theo thời gian như thế nào
df = df.sort_values(["NE_NAME", "SYS_DATETIME"]).reset_index(drop=True)
 
# 5.5 Loại bỏ bản ghi trùng, giữ bản ghi đầu tiên 
# Bỏ comment dòng dưới nếu muốn loại trùng:
# df = df.drop_duplicates(subset=["NE_NAME", "SYS_DATETIME"], keep="first")
 
# %% 6. Xuất file kết quả
with pd.ExcelWriter(OUTPUT_FILE_XLSX, engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="KPI_normalized", index=False)
    if len(df_dup) > 0:
        df_dup.to_excel(writer, sheet_name="Duplicates_flagged", index=False)
 
df.to_csv(OUTPUT_FILE_CSV, index=False, encoding="utf-8-sig")
print(f"\nĐã lưu file chuẩn hóa (Excel): {OUTPUT_FILE_XLSX.resolve()}")
print(f"Đã lưu file chuẩn hóa (CSV):   {OUTPUT_FILE_CSV.resolve()}")
 