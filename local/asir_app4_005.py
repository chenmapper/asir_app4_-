
import streamlit as st
from PIL import Image, ExifTags
import os
import io
import zipfile
import pandas as pd
from pathlib import Path
import time
import xlsxwriter

st.set_page_config(page_title="📘 通用索引相片簿 v1.0", layout="wide")
st.title("📘 通用索引相片簿 v1.0")
st.markdown("上傳你的相片資料夾（ZIP），建立圖片索引、說明與預覽")

uploaded_zip = st.file_uploader("📦 上傳圖片資料夾（zip格式）", type="zip")

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

if uploaded_zip:
    extract_path = Path("temp_images")
    extract_path.mkdir(exist_ok=True)

    for file in extract_path.glob("*"):
        if file.is_file():
            file.unlink()
        else:
            for subfile in file.rglob("*"):
                if subfile.is_file():
                    subfile.unlink()
            file.rmdir()

    with zipfile.ZipFile(uploaded_zip, "r") as zip_ref:
        zip_ref.extractall(extract_path)

    image_paths = list(extract_path.rglob("*"))
    image_paths = [p for p in image_paths if p.suffix.lower() in [".jpg", ".jpeg", ".png", ".webp"]]

    st.success(f"✅ 解壓縮完成，共有 {len(image_paths)} 張圖像")

    data = []
    cols = st.columns(4)

    for i, img_path in enumerate(image_paths):
        try:
            img = Image.open(img_path)
            img.thumbnail((200, 200))
            with cols[i % 4]:
                st.image(img, caption=img_path.name)
                desc = st.text_input(f"說明 - {img_path.name}", key=img_path.name)
            stat = img_path.stat()
            exif_time, exif_status = get_exif_datetime_and_status(img_path)
            data.append({
                "圖檔": img_path,
                "檔名": img_path.name,
                "說明": desc,
                "原圖路徑": str(img_path.resolve()),
                "修改時間": time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(stat.st_mtime)),
                "拍攝時間": exif_time,
                "EXIF狀態": exif_status,
                "檔案大小 (KB)": round(stat.st_size / 1024, 2)
            })
        except Exception as e:
            st.warning(f"{img_path.name} 載入失敗：{e}")

    if data:
        df = pd.DataFrame(data)
        st.markdown("### 📋 圖片索引表")
        st.dataframe(df.drop(columns=["圖檔"]))

        with io.BytesIO() as output:
            workbook = xlsxwriter.Workbook(output, {'in_memory': True})
            worksheet = workbook.add_worksheet("相片索引表")
            header = ["縮圖", "檔名", "說明", "原圖路徑", "修改時間", "拍攝時間", "EXIF狀態", "檔案大小 (KB)"]
            for col_num, h in enumerate(header):
                worksheet.write(0, col_num, h)
                worksheet.set_column(col_num, col_num, 25)

            for row_num, row in enumerate(data, start=1):
                img = Image.open(row["圖檔"])
                img.thumbnail((120, 120))
                img_bytes = io.BytesIO()
                img.save(img_bytes, format='PNG')
                worksheet.set_row(row_num, 100)
                worksheet.insert_image(row_num, 0, row["檔名"], {'image_data': img_bytes})
                worksheet.write(row_num, 1, row["檔名"])
                worksheet.write(row_num, 2, row["說明"])
                worksheet.write(row_num, 3, row["原圖路徑"])
                worksheet.write(row_num, 4, row["修改時間"])
                worksheet.write(row_num, 5, row["拍攝時間"])
                worksheet.write(row_num, 6, row["EXIF狀態"])
                worksheet.write(row_num, 7, row["檔案大小 (KB)"])

            workbook.close()
            st.download_button("📥 下載含縮圖 Excel 索引", data=output.getvalue(),
                               file_name="相片索引_with_img_v5.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
