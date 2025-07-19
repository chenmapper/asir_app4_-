import streamlit as st
import pandas as pd
from pathlib import Path
import zipfile
import shutil
import math
import io
from PIL import Image
import xlsxwriter
import time

st.set_page_config(page_title="asir_app4_通用相片索引 v0.27", layout="wide")
st.title("📘 asir_app4_通用相片索引 v0.27（任務二：依索引表更新圖片檔名與路徑）")

uploaded_zip = st.file_uploader("📦 上傳圖片資料夾（zip格式）", type="zip")
uploaded_xlsx = st.file_uploader("📄 上傳相片索引表（.xlsx 格式）", type="xlsx")
st.markdown("📁 預設圖片儲存根路徑：")
default_path = "C:/Users/User/Downloads/table_app/case4/temp_images"
custom_root = st.text_input("🔧 自訂根目錄（供原圖路徑欄位寫入 file:///）", value=default_path)


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



def clear_extract_path(path):
    if path.exists() and path.is_dir():
        shutil.rmtree(path)
    path.mkdir(parents=True, exist_ok=True)

if "output_zip" not in st.session_state:
    st.session_state["output_zip"] = None
if "output_excel" not in st.session_state:
    st.session_state["output_excel"] = None


run_btn = st.button("✅ 依索引表更新圖片與產出結果")

if run_btn and uploaded_zip and uploaded_xlsx:
    zip_name = Path(uploaded_zip.name).stem
    extract_path = Path(f"temp_images/{zip_name}")
    clear_extract_path(extract_path)
    extract_path.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(uploaded_zip, "r") as zip_ref:
        zip_ref.extractall(extract_path)

    df = pd.read_excel(uploaded_xlsx)
    required_cols = ["原檔名", "新/舊檔名"]
    if not all(col in df.columns for col in required_cols):
        st.error("❌ 索引表必須包含欄位：原檔名、新/舊檔名")
    else:
        rename_log = []
        for i, row in df.iterrows():
            old_name = str(row["原檔名"]).strip()
            new_name = process_new_name(old_name, row["新/舊檔名"])
            old_path = extract_path / old_name
            new_path = extract_path / new_name
            if old_path.exists() and new_name and new_name != old_name:
                try:
                    old_path.rename(new_path)
                    rename_log.append(f"✅ {old_name} → {new_name}")
                    df.at[i, "原檔名"] = new_name  # ✅ 實際更新原檔名欄位
                    df.at[i, "新/舊檔名"] = f"{new_name} ← {old_name}"  # ✅ 更新新舊檔名對照
                except Exception as e:
                    rename_log.append(f"❌ {old_name} 無法重新命名：{e}")
            else:
                rename_log.append(f"⚠️ {old_name} 未更名（可能不存在或相同）")

        st.code("\n".join(rename_log), language="text")

        # 更新原圖路徑欄位
        def get_url(name):
            full_path = str(Path(custom_root) / zip_name / name)
            return "file:///" + full_path.replace("\\", "/")

        df["原圖路徑"] = df["原檔名"].apply(get_url)

        # 匯出新版 Excel（含縮圖）
        df = df.fillna("")  # 避免 NaN 寫入失敗
        output_excel = io.BytesIO()
        workbook = xlsxwriter.Workbook(output_excel, {'in_memory': True})
        worksheet = workbook.add_worksheet("相片索引表")

        headers = ["縮圖"] + df.columns.tolist()
        for col_num, h in enumerate(headers):
            worksheet.write(0, col_num, h)
            worksheet.set_column(col_num, col_num, 28)

        for row_num, row in enumerate(df.itertuples(index=False), start=1):
            img_path = extract_path / row._asdict()["原檔名"]
            if img_path.exists():
                worksheet.set_row(row_num, 100)
                try:
                    img = Image.open(img_path)
                    img.thumbnail((120, 120))
                    img_bytes = io.BytesIO()
                    img.save(img_bytes, format='PNG')
                    worksheet.insert_image(row_num, 0, img_path.name, {'image_data': img_bytes})
                except:
                    pass
            for col_num, val in enumerate(row, start=1):
                if isinstance(val, float) and (math.isnan(val) or math.isinf(val)):
                    worksheet.write(row_num, col_num, "")
                else:
                    worksheet.write(row_num, col_num, val)

        workbook.close()

        # 存入 session_state
        st.session_state["output_zip"] = zip_buffer.getvalue()
        st.session_state["output_excel"] = output_excel.getvalue()

        # 壓縮圖片
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
            for file in extract_path.glob("*"):
                if file.suffix.lower() in [".jpg", ".jpeg", ".png", ".webp"]:
                    zipf.write(file, arcname=file.name)

        if st.session_state["output_excel"]:
            st.download_button("⬇️ 下載更新後相片索引表 (.xlsx)",
                           data=output_excel.getvalue(),
                           file_name=f"{zip_name}_相片索引表_更新後.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        if st.session_state["output_zip"]:
            st.download_button("⬇️ 下載更新後圖庫 (.zip)",
                           data=zip_buffer.getvalue(),
                           file_name=f"{zip_name}_更新後圖庫.zip",
                           mime="application/zip")
