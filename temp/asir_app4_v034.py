import streamlit as st
import zipfile, shutil, io, os
from pathlib import Path
import pandas as pd
import xlsxwriter
from PIL import Image
import math

st.set_page_config(page_title="📘 asir_app4_通用相片索引 v0.34", layout="wide")
st.title("📘 asir_app4_通用相片索引 v0.34（雲端重構版）")

st.markdown("### 📥 上傳圖片資料夾（zip格式）")
uploaded_zip = st.file_uploader("ZIP 檔案", type="zip")

st.markdown("### 📄 上傳相片索引表（.xlsx格式）")
uploaded_xlsx = st.file_uploader("相片索引表", type="xlsx")

def process_new_name(old_name, new_name_raw):
    if not new_name_raw or str(new_name_raw).strip() == "":
        return f"({old_name})"
    name = str(new_name_raw).strip()
    if "." not in name:
        ext = Path(old_name).suffix
        name += ext
    if not name.lower().endswith((".jpg", ".jpeg", ".png", ".webp")):
        return f"({name})"
    return name

if uploaded_zip and uploaded_xlsx:
    run_btn = st.button("✅ 執行更名與輸出")

    if run_btn:
        with zipfile.ZipFile(uploaded_zip, "r") as zip_ref:
            zip_ref.extractall("temp_zip_extract")

        df = pd.read_excel(uploaded_xlsx)
        df = df.fillna("")
        old_names = df["原檔名"].tolist()
        new_names = [process_new_name(old, new) for old, new in zip(df["原檔名"], df["新/舊檔名"])]

        base = Path("temp_images/res")
        base.parent.mkdir(parents=True, exist_ok=True)
        i = 1
        while (base / f"res{i}").exists():
            i += 1
        res_path = base / f"res{i}"
        res_path.mkdir(parents=True, exist_ok=True)

        log = []

        for idx, (old, new) in enumerate(zip(old_names, new_names)):
            src_path = Path("temp_zip_extract") / old
            dst_path = res_path / new
            if src_path.exists():
                try:
                    shutil.copy(src_path, dst_path)
                    df.at[idx, "新/舊檔名"] = f"{new} ← {old}"
                    df.at[idx, "原檔名"] = new
                    df.at[idx, "原圖路徑"] = f"file:///{dst_path.resolve().as_posix()}"
                    log.append(f"✅ {old} → {new}")
                except Exception as e:
                    log.append(f"❌ {old} 無法更名：{e}")
            else:
                log.append(f"⚠️ {old} 未更名（可能不存在或相同）")

        df["縮圖"] = ""

        output_excel = io.BytesIO()
        workbook = xlsxwriter.Workbook(output_excel, {"in_memory": True})
        worksheet = workbook.add_worksheet("索引表")

        for col_num, col in enumerate(df.columns):
            worksheet.write(0, col_num, col)

        for row_num, row in enumerate(df.itertuples(index=False), start=1):
            for col_num, val in enumerate(row):
                if df.columns[col_num] == "縮圖":
                    img_name = row[df.columns.get_loc("原檔名")]
                    img_path = res_path / img_name
                    if img_path.exists():
                        try:
                            with Image.open(img_path) as im:
                                im.thumbnail((80, 80))
                                thumb_path = res_path / f"thumb_{img_name}"
                                im.save(thumb_path)
                                worksheet.insert_image(row_num, col_num, str(thumb_path), {"x_scale": 0.5, "y_scale": 0.5})
                        except:
                            pass
                else:
                    if isinstance(val, float) and (math.isnan(val) or math.isinf(val)):
                        worksheet.write(row_num, col_num, "")
                    else:
                        worksheet.write(row_num, col_num, val)

        workbook.close()
        output_excel.seek(0)
        df_output_path = res_path / "相片索引表.xlsx"
        with open(df_output_path, "wb") as f:
            f.write(output_excel.getbuffer())

        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
            for file in res_path.glob("*"):
                zipf.write(file, arcname=file.name)
        zip_buffer.seek(0)

        st.success(f"📦 res{i} 資料夾建立完成，包含圖檔與索引表")
        st.download_button("⬇️ 下載整包圖庫與索引表", data=zip_buffer, file_name=f"res{i}.zip", mime="application/zip")
        st.download_button("⬇️ 下載相片索引表.xlsx", data=output_excel, file_name="相片索引表.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        st.markdown("### 📋 更名紀錄")
        for line in log:
            st.write(line)
