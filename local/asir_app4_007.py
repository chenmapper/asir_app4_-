import streamlit as st
from PIL import Image, ExifTags
import os
import io
import zipfile
import pandas as pd
from pathlib import Path
import time
import xlsxwriter

st.set_page_config(page_title="📘 相片索引表 v1.1", layout="wide")
st.title("📘 相片索引表 v1.1")
st.markdown("上傳你的相片資料夾（ZIP），建立圖片索引、說明表格，便於日後批次改名與定位")

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
                "原檔名": img_path.name,
                "新/舊檔名": img_path.name,
                "相片說明": desc,
                "原圖路徑": str(img_path.resolve()),
                "修改時間": time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(stat.st_mtime)),
                "拍攝時間": exif_time,
                "EXIF狀態": exif_status,
                "檔案大小 (KB)": round(stat.st_size / 1024, 2),
                "gx": "",
                "gy": "",
                "gz": "",
                "圖檔": img_path
            })
        except Exception as e:
            st.warning(f"{img_path.name} 載入失敗：{e}")

    if data:
        st.markdown("### 📋 相片索引預覽")
        df_preview = pd.DataFrame(data).drop(columns=["圖檔"])
        st.dataframe(df_preview)

        with io.BytesIO() as output:
            workbook = xlsxwriter.Workbook(output, {'in_memory': True})
            worksheet = workbook.add_worksheet("相片索引表")

            header = ["縮圖", "原檔名", "新/舊檔名", "相片說明", "原圖路徑",
                      "修改時間", "拍攝時間", "EXIF狀態", "檔案大小 (KB)", "gx", "gy", "gz"]

            for col_num, h in enumerate(header):
                worksheet.write(0, col_num, h)
                worksheet.set_column(col_num, col_num, 28)

            for row_num, row in enumerate(data, start=1):
                img = Image.open(row["圖檔"])
                img.thumbnail((120, 120))
                img_bytes = io.BytesIO()
                img.save(img_bytes, format='PNG')
                worksheet.set_row(row_num, 100)
                worksheet.insert_image(row_num, 0, row["原檔名"], {'image_data': img_bytes})

                values = [row.get(k, "") for k in header[1:]]
                for col_num, val in enumerate(values, start=1):
                    worksheet.write(row_num, col_num, val)

            workbook.close()
            st.download_button("📥 下載相片索引表", data=output.getvalue(),
                               file_name="相片索引表.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
