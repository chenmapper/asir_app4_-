import streamlit as st
from PIL import Image, ExifTags
import os
import io
import zipfile
import pandas as pd
from pathlib import Path
import time
import xlsxwriter
import math
import shutil

st.set_page_config(page_title="asir_app4_通用相片索引", layout="wide")
st.title("📘 asir_app4_通用相片索引 v0.25（修正 full_file_path 未定義）")

uploaded_zip = st.file_uploader("📦 上傳圖片資料夾（zip格式）", type="zip")
st.markdown("📁 預設圖片儲存根路徑：")
default_path = "C:/Users/User/Downloads/table_app/case4/temp_images"
custom_root = st.text_input("🔧 自訂根目錄（供原圖路徑欄位寫入 file:///）", value=default_path)
generate_btn = st.button("🧾 產生索引＋整包下載")

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

def get_next_res_dir():
    base_dir = Path("temp_images")
    base_dir.mkdir(exist_ok=True)
    existing = [int(p.name[3:]) for p in base_dir.glob("res*") if p.name[3:].isdigit()]
    next_id = max(existing + [0]) + 1
    res_dir = base_dir / f"res{next_id}"
    res_dir.mkdir(parents=True, exist_ok=True)
    return res_dir, next_id

if generate_btn and uploaded_zip:
    extract_path = Path("temp_images/_extract_temp")
    if extract_path.exists():
        shutil.rmtree(extract_path)
    extract_path.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(uploaded_zip, "r") as zip_ref:
        zip_ref.extractall(extract_path)

    image_paths = list(extract_path.rglob("*"))
    image_paths = [p for p in image_paths if p.suffix.lower() in [".jpg", ".jpeg", ".png", ".webp"]]

    st.success(f"✅ 解壓完成，共 {len(image_paths)} 張圖像")

    res_dir, res_id = get_next_res_dir()
    data = []

    for img_path in image_paths:
        try:
            new_path = res_dir / img_path.name
            shutil.copy(img_path, new_path)
            img = Image.open(img_path)
            stat = img_path.stat()
            exif_time, exif_status = get_exif_datetime_and_status(img_path)

            full_file_path = str(Path(custom_root) / f"res{res_id}" / img_path.name)
            file_url = "file:///" + full_file_path.replace("\\", "/")
 

            data.append({
                "原檔名": img_path.name,
                "新/舊檔名": img_path.name,
                "相片說明": "",
                "原圖路徑": file_url,
                "修改時間": time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(stat.st_mtime)),
                "拍攝時間": exif_time,
                "EXIF狀態": exif_status,
                "檔案大小 (KB)": round(stat.st_size / 1024, 2),
                "gx": "",
                "gy": "",
                "gz": "",
                "圖檔": new_path
            })
        except Exception as e:
            st.warning(f"{img_path.name} 載入失敗：{e}")

    xlsx_path = res_dir / "相片索引表.xlsx"
    with xlsxwriter.Workbook(xlsx_path) as workbook:
        worksheet = workbook.add_worksheet("相片索引表")
        header = ["縮圖", "原檔名", "新/舊檔名", "相片說明", "原圖路徑",
                  "修改時間", "拍攝時間", "EXIF狀態", "檔案大小 (KB)", "gx", "gy", "gz"]

        for col_num, h in enumerate(header):
            worksheet.write(0, col_num, h)
            worksheet.set_column(col_num, col_num, 28)

        for row_num, row in enumerate(data, start=1):
            worksheet.set_row(row_num, 100)
            img = Image.open(row["圖檔"])
            img.thumbnail((120, 120))
            img_bytes = io.BytesIO()
            img.save(img_bytes, format='PNG')
            worksheet.insert_image(row_num, 0, row["原檔名"], {'image_data': img_bytes})
            for col_num, key in enumerate(header[1:], start=1):
                val = row[key]
                if key == "原圖路徑":
                    worksheet.write_url(row_num, col_num, val, string=val)
                else:
                    worksheet.write(row_num, col_num, val)

    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
        for file in res_dir.glob("*"):
            zipf.write(file, arcname=file.name)

    st.success(f"📦 res{res_id} 資料夾建立完成，包含圖檔與索引表")
    st.download_button("⬇️ 下載整包 res 資料夾（zip）", data=zip_buffer.getvalue(),
                       file_name=f"res{res_id}.zip", mime="application/zip")
