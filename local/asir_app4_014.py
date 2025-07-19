import streamlit as st
from PIL import Image, ExifTags
import os
import io
import zipfile
import pandas as pd
from pathlib import Path
import time
import xlsxwriter
import shutil
import math

st.set_page_config(page_title="📘 相片索引表 v1.7", layout="wide")
st.title("📘 相片索引表 v1.7")

uploaded_zip = st.file_uploader("📦 上傳圖片資料夾（zip格式）", type="zip")
uploaded_excel = st.file_uploader("📄 上傳相片索引表（.xlsx格式）", type=["xlsx"])
update_btn = st.button("🛠️ 依新檔名更新")

def get_exif_datetime_and_status(img_path):
    try:
        img = Image.open(img_path)
        exif = img._getexif()
        if exif:
            for tag, value in exif.items():
                tag_name = ExifTags.TAGS.get(tag)
                if tag_name == "DateTimeOriginal":
                    return value.replace(":", "-", 2), "✅ 有拍攝時間"
            return "", "⚠️ 無 DateTimeOriginal"
        return "", "⚠️ 無 EXIF 資訊"
    except Exception as e:
        return "", f"❌ 讀取失敗：{e}"

def extract_zip(file):
    extract_path = Path("temp_images/extract_temp")
    if extract_path.exists():
        shutil.rmtree(extract_path)
    extract_path.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(file, "r") as zip_ref:
        zip_ref.extractall(extract_path)
    image_paths = list(extract_path.rglob("*"))
    return [p for p in image_paths if p.suffix.lower() in [".jpg", ".jpeg", ".png", ".webp"]], extract_path

def get_next_res_dir():
    base_dir = Path("temp_images")
    base_dir.mkdir(exist_ok=True)
    existing = [int(p.name[3:]) for p in base_dir.glob("res*") if p.name[3:].isdigit()]
    next_id = max(existing + [0]) + 1
    res_dir = base_dir / f"res{next_id}"
    img_dir = res_dir / "更新後圖片"
    img_dir.mkdir(parents=True, exist_ok=True)
    return res_dir, img_dir, next_id

if update_btn:
    if uploaded_zip and uploaded_excel:
        try:
            df = pd.read_excel(uploaded_excel)
            if "原檔名" not in df.columns or "新/舊檔名" not in df.columns:
                st.error("❌ 相片索引表缺少必要欄位（原檔名 或 新/舊檔名）")
            else:
                image_paths, _ = extract_zip(uploaded_zip)
                res_dir, res_img_dir, res_id = get_next_res_dir()
                valid_rows = []

                for i, row in df.iterrows():
                    old_name = row["原檔名"]
                    new_name = row["新/舊檔名"]
                    if pd.isna(new_name) or str(new_name).strip() == "":
                        continue
                    match = [p for p in image_paths if p.name == old_name]
                    if match:
                        img_path = match[0]
                        new_img_path = res_img_dir / new_name
                        shutil.copy(img_path, new_img_path)

                        # 建立縮圖 BytesIO
                        thumb_io = io.BytesIO()
                        img = Image.open(img_path)
                        img.thumbnail((120, 120))
                        img.save(thumb_io, format='PNG')

                        # 建立行資料
                        row["新/舊檔名"] = f"{old_name} → {new_name}"
                        row["原檔名"] = new_name
                        row["原圖路徑"] = f"./更新後圖片/{new_name}"
                        row["縮圖"] = thumb_io
                        valid_rows.append(row)

                # 寫入 Excel（含縮圖）
                xlsx_path = res_dir / "相片索引表.xlsx"
                with xlsxwriter.Workbook(xlsx_path) as workbook:
                    worksheet = workbook.add_worksheet("索引表")
                    header = ["縮圖", "原檔名", "新/舊檔名", "相片說明", "原圖路徑",
                              "修改時間", "拍攝時間", "EXIF狀態", "檔案大小 (KB)", "gx", "gy", "gz"]
                    for col_num, h in enumerate(header):
                        worksheet.write(0, col_num, h)
                        worksheet.set_column(col_num, col_num, 28)
                    for row_num, row in enumerate(valid_rows, start=1):
                        worksheet.set_row(row_num, 100)
                        worksheet.insert_image(row_num, 0, row["原檔名"], {'image_data': row["縮圖"]})
                        for col_num, key in enumerate(header[1:], start=1):
                            val = row.get(key, "")
                            if pd.isna(val) or (isinstance(val, float) and (math.isinf(val) or math.isnan(val))):
                                val = ""
                            worksheet.write(row_num, col_num, val)

                st.success(f"✅ 相片與索引表已儲存於：{res_dir.resolve()}")
                st.info("請直接到本機資料夾查看檔案，無需點擊額外下載按鈕。")

        except Exception as e:
            st.error(f"❌ 錯誤：{e}")
    else:
        st.warning("⚠️ 請同時上傳圖片壓縮包與相片索引表")
