
import streamlit as st
from PIL import Image
import os
import io
import zipfile
import pandas as pd
from pathlib import Path

st.set_page_config(page_title="📘 通用索引相片簿 v1.0", layout="wide")
st.title("📘 通用索引相片簿 v1.0")
st.markdown("上傳你的相片資料夾（ZIP），建立圖片索引、說明與預覽")

# 上傳 zip 檔案
uploaded_zip = st.file_uploader("📦 上傳圖片資料夾（zip格式）", type="zip")

if uploaded_zip:
    extract_path = Path("temp_images")
    extract_path.mkdir(exist_ok=True)

    # 清空目錄
    for file in extract_path.glob("*"):
        if file.is_file():
            file.unlink()
        else:
            for subfile in file.rglob("*"):
                if subfile.is_file():
                    subfile.unlink()
            file.rmdir()

    # 解壓縮
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
            data.append({
                "檔名": img_path.name,
                "說明": desc,
                "原圖路徑": str(img_path.resolve())
            })
        except Exception as e:
            st.warning(f"{img_path.name} 載入失敗：{e}")

    if data:
        df = pd.DataFrame(data)
        st.markdown("### 📋 圖片索引表")
        st.dataframe(df)

        # 匯出索引表
        with io.BytesIO() as output:
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name="索引表")
            st.download_button("📥 下載索引表 Excel", data=output.getvalue(),
                               file_name="相片索引表.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
