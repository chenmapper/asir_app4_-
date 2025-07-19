import streamlit as st
import zipfile
import tempfile
import io
import shutil
import time
from pathlib import Path
import pandas as pd
import xlsxwriter
from PIL import Image, ExifTags

# --- Configuration ---
st.set_page_config(page_title="📘 asir_app4_通用相片索引 v0.39一體化", layout="wide")
st.title("📘 asir_app4_通用相片索引 v0.39（初始索引 + 批次更名）")

# Sidebar: 選擇模式
mode = st.sidebar.selectbox("🔄 選擇功能模式", ["生成索引表與圖庫", "從 XLSX 批次更名"])

# 共用：根目錄設定，用於 file:/// 路徑
st.sidebar.header("⚙️ 根目錄設定")
root_dir = st.sidebar.text_input(
    "根目錄路徑",
    value=r"C:\Users\User\Downloads\table_app\case4\temp_images"
)

# Initialize session state
for key in ['zip_data','excel_data','logs','orig_zip_name','orig_xlsx_name']:
    if key not in st.session_state:
        st.session_state[key] = None if key in ['zip_data','excel_data','orig_zip_name','orig_xlsx_name'] else []

if mode == "生成索引表與圖庫":
    # 上傳 ZIP
    st.markdown("### 📥 上傳圖片資料夾（zip格式）")
    uploaded_zip = st.file_uploader("ZIP 檔案", type="zip", key="gen_zip")
    if uploaded_zip:
        st.session_state['orig_zip_name'] = uploaded_zip.name
        # 上傳完成按鈕
        if st.button("🧾 生成索引表與圖庫", key="gen_btn"):
            with st.spinner("處理中..."):
                zip_name = Path(uploaded_zip.name).stem
                # 解壓到暫存
                extract_dir = tempfile.mkdtemp()
                Path(extract_dir, zip_name).mkdir(parents=True, exist_ok=True)
                with zipfile.ZipFile(uploaded_zip, 'r') as z:
                    z.extractall(Path(extract_dir)/zip_name)
                image_paths = list(Path(extract_dir, zip_name).rglob("*.jpg")) + \
                              list(Path(extract_dir, zip_name).rglob("*.png")) + \
                              list(Path(extract_dir, zip_name).rglob("*.jpeg")) + \
                              list(Path(extract_dir, zip_name).rglob("*.webp"))
                # 構建資料
                data=[]
                def get_exif_datetime(img_path):
                    try:
                        img = Image.open(img_path)
                        exif = img._getexif()
                        if exif:
                            for tag, val in exif.items():
                                if ExifTags.TAGS.get(tag) == "DateTimeOriginal":
                                    return val.replace(":", "-", 2), "✅ 有拍攝時間"
                            return "", "⚠️ 無 DateTimeOriginal"
                        return "", "⚠️ 無 EXIF 資訊"
                    except:
                        return "", "❌ 讀取失敗"
                for p in image_paths:
                    mtime = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(p.stat().st_mtime))
                    exif_time, exif_status = get_exif_datetime(p)
                    full = Path(root_dir) / zip_name / p.name
                    file_url = "file:///" + str(full).replace("\\", "/")
                    data.append({
                        "原檔名": p.name,
                        "新/舊檔名": p.name,
                        "相片說明": "",
                        "原圖路徑": file_url,
                        "修改時間": mtime,
                        "拍攝時間": exif_time,
                        "EXIF狀態": exif_status,
                        "檔案大小 (KB)": round(p.stat().st_size/1024, 2),
                        "gx": "", "gy": "", "gz": "",
                        "圖檔": p
                    })
                # 輸出 Excel
                xlsx_buf = io.BytesIO()
                wb = xlsxwriter.Workbook(xlsx_buf, {'in_memory': True})
                ws = wb.add_worksheet("相片索引表")
                header = ["縮圖", "原檔名", "新/舊檔名", "相片說明", "原圖路徑", "修改時間", "拍攝時間", "EXIF狀態", "檔案大小 (KB)", "gx", "gy", "gz"]
                for i, h in enumerate(header): ws.write(0, i, h)
                for r, row in enumerate(data, 1):
                    img = Image.open(row["圖檔"])
                    img.thumbnail((120, 120))
                    buf2 = io.BytesIO()
                    img.save(buf2, format='PNG')
                    ws.set_row(r, 100)
                    ws.insert_image(r, 0, row["原檔名"], {'image_data': buf2})
                    for c, key in enumerate(header[1:], 1):
                        val = row[key]
                        if key == "原圖路徑":
                            ws.write_url(r, c, val, string=val)
                        else:
                            ws.write(r, c, val)
                wb.close()
                xlsx_buf.seek(0)
                # 壓縮圖檔
                zip_buf = io.BytesIO()
                with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as zf:
                    for p in image_paths:
                        zf.write(p, arcname=p.name)
                zip_buf.seek(0)
                st.session_state['zip_data'] = zip_buf.getvalue()
                st.session_state['excel_data'] = xlsx_buf.getvalue()
                st.session_state['logs'] = [f"✅ 共處理 {len(data)} 張圖片"]
                st.success("完成！")

elif mode == "從 XLSX 批次更名":
    # 上傳 ZIP 與 XLSX
    st.markdown("### 📥 上傳原始圖庫 ZIP")
    uploaded_zip = st.file_uploader("ZIP 檔案", type="zip", key="upd_zip")
    st.markdown("### 📄 上傳索引表 XLSX (含 原檔名 與 新/舊檔名)")
    uploaded_xlsx = st.file_uploader("相片索引表", type="xlsx", key="upd_xlsx")
    if uploaded_zip: st.session_state['orig_zip_name'] = uploaded_zip.name
    if uploaded_xlsx: st.session_state['orig_xlsx_name'] = uploaded_xlsx.name
    if uploaded_zip and uploaded_xlsx and st.button("✅ 執行批次更名", key="upd_btn"):
        with st.spinner("批次處理中..."):
            def process_batch2(zip_file, xlsx_file, root_dir):
                with tempfile.TemporaryDirectory() as ext, tempfile.TemporaryDirectory() as out:
                    with zipfile.ZipFile(zip_file) as z:
                        names = [i.filename for i in z.infolist() if not i.is_dir()]
                        roots = {Path(n).parts[0] for n in names if len(Path(n).parts) > 1}
                        upload_dir = roots.pop() if len(roots) == 1 else ''
                        z.extractall(ext)
                    base = Path(ext) / upload_dir if upload_dir else Path(ext)
                    df = pd.read_excel(xlsx_file).fillna("")
                    if '新/舊檔名' in df.columns:
                        df.rename(columns={'新/舊檔名':'新檔名'}, inplace=True)
                    df['原檔名'] = df['原檔名'].astype(str).str.strip()
                    df['新檔名'] = df['新檔名'].astype(str).str.strip()
                    olds = df['原檔名'].tolist()
                    def extf(n, o): return o if not n else (n if Path(n).suffix else f"{n}{Path(o).suffix}")
                    news = [extf(n, o) for n, o in zip(df['新檔名'], olds)]
                    df['原檔名'] = news
                    df['新/舊檔名'] = [f"({n}/{o})" for n, o in zip(news, olds)]
                    df['原圖路徑'] = [f"file:///{Path(root_dir)/(upload_dir or '')/n}" for n in news]
                    logs = []
                    for o, n in zip(olds, news):
                        src = base / o
                        dst = Path(out) / n
                        if src.exists():
                            shutil.copy(src, dst)
                            logs.append(f"✅ {o} → {n}")
                        else:
                            logs.append(f"⚠️ 檔案不存在: {o}")
                    eb = io.BytesIO()
                    df.to_excel(eb, sheet_name='更新索引表', index=False)
                    eb.seek(0)
                    zb = io.BytesIO()
                    with zipfile.ZipFile(zb, 'w', zipfile.ZIP_DEFLATED) as zf:
                        for f in Path(out).iterdir():
                            zf.write(f, arcname=f.name)
                    zb.seek(0)
                    return zb.getvalue(), eb.getvalue(), logs
            zip_data, excel_data, logs = process_batch2(uploaded_zip, uploaded_xlsx, root_dir)
            st.session_state['zip_data'] = zip_data
            st.session_state['excel_data'] = excel_data
            st.session_state['logs'] = logs
            st.session_state['orig_zip_name'] = uploaded_zip.name
            st.session_state['orig_xlsx_name'] = uploaded_xlsx.name
            st.success('批次更名完成！')

# 顯示日誌與下載按鈕
if st.session_state['logs']:
    st.markdown('### 📜 執行日誌')
    for ln in st.session_state['logs']:
        st.write(ln)
if st.session_state['zip_data']:
    st.download_button('⬇️ 下載圖庫 ZIP', data=st.session_state['zip_data'], file_name=st.session_state['orig_zip_name'], mime='application/zip')
if st.session_state['excel_data']:
    st.download_button('⬇️ 下載索引表 XLSX', data=st.session_state['excel_data'], file_name=st.session_state['orig_xlsx_name'], mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
