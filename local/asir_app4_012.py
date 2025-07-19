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

st.set_page_config(page_title="📘 相片索引表 v1.5", layout="wide")
st.title("📘 相片索引表 v1.5")

uploaded_zip = st.file_uploader("📦 上傳圖片資料夾（zip格式）", type="zip")
uploaded_excel = st.file_uploader("📄 上傳相片索引表（.xlsx格式）", type=["xlsx"])

generate_btn = st.button("🧾 生成相片索引表")
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
    extract_path = Path("temp_images")
    if extract_path.exists():
        shutil.rmtree(extract_path)
    extract_path.mkdir(exist_ok=True)
    with zipfile.ZipFile(file, "r") as zip_ref:
        zip_ref.extractall(extract_path)
    image_paths = list(extract_path.rglob("*"))
    return [p for p in image_paths if p.suffix.lower() in [".jpg", ".jpeg", ".png", ".webp"]], extract_path

def save_bytesio_to_file(bytes_io, filepath):
    with open(filepath, "wb") as f:
        f.write(bytes_io.getvalue())

if update_btn:
    if uploaded_zip and uploaded_excel:
        try:
            df = pd.read_excel(uploaded_excel)
            if "原檔名" not in df.columns or "新/舊檔名" not in df.columns:
                st.error("❌ 相片索引表缺少必要欄位（原檔名 或 新/舊檔名）")
            else:
                image_paths, extract_path = extract_zip(uploaded_zip)
                valid_rows = []
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w") as zip_out:
                    for i, row in df.iterrows():
                        old_name = row["原檔名"]
                        new_name = row["新/舊檔名"]
                        if pd.isna(new_name) or str(new_name).strip() == "":
                            continue
                        match = [p for p in image_paths if p.name == old_name]
                        if match:
                            img_path = match[0]
                            new_path = img_path.parent / new_name
                            shutil.copy(img_path, new_path)

                            # 準備縮圖 BytesIO
                            thumb_io = io.BytesIO()
                            img = Image.open(img_path)
                            img.thumbnail((120, 120))
                            img.save(thumb_io, format='PNG')

                            zip_out.write(new_path, arcname=new_name)
                            new_path.unlink()

                            row["新/舊檔名"] = f"{old_name} → {new_name}"
                            row["原檔名"] = new_name
                            row["縮圖"] = thumb_io
                            valid_rows.append(row)

                st.success("✅ 已完成檔名更新與索引表調整")

                zip_path = Path("temp_images/更新後圖片.zip")
                save_bytesio_to_file(zip_buffer, zip_path)

                xlsx_path = Path("temp_images/更新後_相片索引表.xlsx")
                with xlsxwriter.Workbook(xlsx_path) as workbook:
                    worksheet = workbook.add_worksheet("更新索引表")
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

                combo_path = Path("temp_images/更新成果.zip")
                with zipfile.ZipFile(combo_path, "w") as z:
                    z.write(zip_path, arcname="更新後圖片.zip")
                    z.write(xlsx_path, arcname="更新後_相片索引表.xlsx")

                with open(combo_path, "rb") as f:
                    st.download_button("📦 一鍵下載更新成果", data=f.read(),
                                       file_name="更新成果.zip", mime="application/zip")
        except Exception as e:
            st.error(f"❌ 錯誤：{e}")
    else:
        st.warning("⚠️ 請同時上傳圖片壓縮包與相片索引表")
