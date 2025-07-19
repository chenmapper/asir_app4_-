import streamlit as st
import zipfile, tempfile, io, shutil
from pathlib import Path
import pandas as pd

# --- Configuration ---
st.set_page_config(page_title="📘 asir_app4_通用相片索引 v0.38 子任務2：原欄位更新", layout="wide")
st.title("📘 asir_app4_通用相片索引 v0.38（子任務2：從 XLSX 原欄位更新）")

# Initialize session state
for key in ['zip_data', 'excel_data', 'logs', 'orig_zip_name', 'orig_xlsx_name']:
    if key not in st.session_state:
        st.session_state[key] = None if key in ['zip_data','excel_data','orig_zip_name','orig_xlsx_name'] else []

# File uploads
st.markdown("### 📥 上傳原始圖片資料夾（zip 格式）")
uploaded_zip = st.file_uploader("ZIP 檔案", type="zip")
if uploaded_zip:
    st.session_state['orig_zip_name'] = uploaded_zip.name

st.markdown("### 📄 上傳填寫後的相片索引表（XLSX 格式，含「原檔名」與「新/舊檔名」兩欄）")
uploaded_xlsx = st.file_uploader("相片索引表", type="xlsx")
if uploaded_xlsx:
    st.session_state['orig_xlsx_name'] = uploaded_xlsx.name

# Processing function
def process_batch(zip_file, xlsx_file):
    # 解壓 ZIP
    with tempfile.TemporaryDirectory() as extract_dir, tempfile.TemporaryDirectory() as output_dir:
        with zipfile.ZipFile(zip_file) as z:
            names = [info.filename for info in z.infolist() if not info.is_dir()]
            roots = {Path(n).parts[0] for n in names if len(Path(n).parts) > 1}
            upload_dir = roots.pop() if len(roots)==1 else ''
            z.extractall(extract_dir)
        base_path = Path(extract_dir)/upload_dir if upload_dir else Path(extract_dir)

        # 讀取索引表
        df = pd.read_excel(xlsx_file).fillna("")
        # 支援「新/舊檔名」欄位
        if '新/舊檔名' in df.columns:
            df.rename(columns={'新/舊檔名':'新檔名'}, inplace=True)
        # 取出原/新檔名列表
        df['原檔名'] = df['原檔名'].astype(str).str.strip()
        df['新檔名'] = df['新檔名'].astype(str).str.strip()
        old_names = df['原檔名'].tolist()
        # 補齊副檔名
        def ensure_ext(new, old): return old if not new else (new if Path(new).suffix else f"{new}{Path(old).suffix}")
        new_names = [ensure_ext(n,o) for n,o in zip(df['新檔名'], old_names)]

        # 更新欄位內容
        df['原檔名'] = new_names
        df['新檔名'] = [f"({new}/{old})" for new,old in zip(new_names, old_names)]
        df['原圖路徑'] = [f"{upload_dir}/{new}" if upload_dir else new for new in new_names]

        # 複製並更名圖片
        logs=[]
        for old,new in zip(old_names, new_names):
            src = base_path/old; dst = Path(output_dir)/new
            if src.exists(): shutil.copy(src,dst); logs.append(f"✅ {old} → {new}")
            else: logs.append(f"⚠️ 檔案不存在: {old}")

        # 輸出更新後索引表（保留原欄位順序及名稱）
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='更新索引表', index=False)
        excel_buffer.seek(0)

        # 打包新圖庫 ZIP
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer,'w',zipfile.ZIP_DEFLATED) as out_z:
            for file in Path(output_dir).iterdir(): out_z.write(file, arcname=file.name)
        zip_buffer.seek(0)

        return zip_buffer.getvalue(), excel_buffer.getvalue(), logs

# 執行按鈕
if uploaded_zip and uploaded_xlsx:
    if st.button('✅ 執行更新'):
        zip_data, excel_data, logs = process_batch(uploaded_zip, uploaded_xlsx)
        st.session_state['zip_data']=zip_data
        st.session_state['excel_data']=excel_data
        st.session_state['logs']=logs
        st.success('📸 圖片更名並更新索引表完成！')

# 顯示日誌與下載按鈕（持久化）
if st.session_state['logs']:
    st.markdown('### 📜 執行日誌')
    for ln in st.session_state['logs']: st.write(ln)
if st.session_state['zip_data']:
    st.download_button('⬇️ 下載新圖庫 ZIP', data=st.session_state['zip_data'],
                       file_name=st.session_state['orig_zip_name'], mime='application/zip')
if st.session_state['excel_data']:
    st.download_button('⬇️ 下載更新索引表 XLSX', data=st.session_state['excel_data'],
                       file_name=st.session_state['orig_xlsx_name'], mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
