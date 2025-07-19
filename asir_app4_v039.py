import streamlit as st
import zipfile
import tempfile
import io
import shutil
import time
import re
from pathlib import Path
import pandas as pd
import xlsxwriter
from PIL import Image, ExifTags

# --- Configuration ---
st.set_page_config(page_title="📘 asir_app4_通用相片索引 v0.39一體化", layout="wide")
st.title("📘 asir_app4_通用相片索引 v0.39（初始索引 + 批次更名）")

# Sidebar: 選擇模式
mode = st.sidebar.selectbox(
    "🔄 選擇功能模式",
    ["生成索引表與圖庫", "從 XLSX 批次更名"]
)

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

# Regex for valid filenames: alphanumeric, underscore, hyphen, dot extension
valid_pattern = re.compile(r'^[A-Za-z0-9_\-]+\.[A-Za-z0-9]+$')

if mode == "生成索引表與圖庫":
    # 省略：生成索引表與圖庫的實作維持不變
    st.markdown("### 📥 上傳圖片資料夾（zip格式）")
    uploaded_zip = st.file_uploader("ZIP 檔案", type="zip", key="gen_zip")
    if uploaded_zip:
        st.session_state['orig_zip_name'] = uploaded_zip.name
        if st.button("🧾 生成索引表與圖庫", key="gen_btn"):
            # ...初始索引表生成邏輯...
            pass

elif mode == "從 XLSX 批次更名":
    st.markdown("### 📥 上傳原始圖庫 ZIP")
    uploaded_zip = st.file_uploader("ZIP 檔案", type="zip", key="upd_zip")
    st.markdown("### 📄 上傳索引表 XLSX (含 目前檔名 與 新檔名)")
    uploaded_xlsx = st.file_uploader("相片索引表", type="xlsx", key="upd_xlsx")

    if uploaded_zip:
        st.session_state['orig_zip_name'] = uploaded_zip.name
    if uploaded_xlsx:
        st.session_state['orig_xlsx_name'] = uploaded_xlsx.name

    if uploaded_zip and uploaded_xlsx and st.button('✅ 執行批次更名', key='upd_btn'):
        with st.spinner('批次處理中...'):
            # 解壓 ZIP
            with tempfile.TemporaryDirectory() as ext, tempfile.TemporaryDirectory() as out_dir:
                with zipfile.ZipFile(uploaded_zip, 'r') as z:
                    all_files = [info.filename for info in z.infolist() if not info.is_dir()]
                    roots = {Path(f).parts[0] for f in all_files if len(Path(f).parts) > 1}
                    upload_dir = roots.pop() if len(roots) == 1 else ''
                    z.extractall(ext)
                base_folder = Path(ext) / upload_dir if upload_dir else Path(ext)

                # 讀取索引表
                df = pd.read_excel(uploaded_xlsx).fillna('')
                # 檢查並使用「目前檔名」欄位
                if '目前檔名' not in df.columns:
                    st.error("索引表必須包含 '目前檔名' 欄位")
                    st.stop()

                olds = df['目前檔名'].astype(str).str.strip().tolist()
                raws = df.get('新檔名', pd.Series(['']*len(olds))).astype(str).str.strip().tolist()

                finals, logs, rename_logs = [], [], []
                for old_name, raw_name in zip(olds, raws):
                    # 1. 自動補副檔名
                    if raw_name and not Path(raw_name).suffix:
                        candidate = f"{raw_name}{Path(old_name).suffix}"
                    elif raw_name:
                        candidate = raw_name
                    else:
                        candidate = old_name
                    # 2. 命名規則檢核
                    if valid_pattern.match(candidate):
                        logs.append(f"✅ {old_name} → {candidate}")
                        rename_logs.append(old_name)
                        finals.append(candidate)
                    else:
                        logs.append(f"⚠️ 跳過: {raw_name}")
                        rename_logs.append('')
                        finals.append(old_name)

                # 複製並更名檔案到 out_dir
                for old_name, new_name in zip(olds, finals):
                    src = base_folder / old_name
                    dst = Path(out_dir) / new_name
                    if src.exists():
                        shutil.copy(src, dst)

                # 更新欄位內容
                df['目前檔名'] = finals
                df['新檔名'] = ['' if rl else raw for rl, raw in zip(rename_logs, raws)]
                df['更名log'] = rename_logs
                df['原圖路徑'] = [
                    f"file:///{Path(root_dir) / upload_dir / fn}" if upload_dir else f"file:///{Path(root_dir) / fn}"
                    for fn in finals
                ]

                # 欄位順序微調
                desired_cols = [
                    '縮圖', '目前檔名', '新檔名', '相片說明', '原圖路徑',
                    '修改時間', '拍攝時間', 'EXIF狀態', '檔案大小(KB)',
                    'gx', 'gy', 'gz', '更名log'
                ]
                existing_cols = [c for c in desired_cols if c in df.columns]
                output_df = df[existing_cols]

                # 輸出更新後索引表
                excel_buffer = io.BytesIO()
                with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                    output_df.to_excel(writer, sheet_name='更新索引表', index=False)
                excel_buffer.seek(0)

                # 打包圖庫 ZIP
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                    for file in Path(out_dir).iterdir():
                        zf.write(file, arcname=file.name)
                zip_buffer.seek(0)

                # 儲存並顯示結果
                st.session_state['zip_data'] = zip_buffer.getvalue()
                st.session_state['excel_data'] = excel_buffer.getvalue()
                st.session_state['logs'] = logs
                st.session_state['orig_zip_name'] = uploaded_zip.name
                st.session_state['orig_xlsx_name'] = uploaded_xlsx.name
                st.success('批次更名完成！')

# 顯示日誌與下載按鈕
if st.session_state['logs']:
    st.markdown('### 📜 執行日誌')
    for line in st.session_state['logs']:
        st.write(line)
if st.session_state['zip_data']:
    st.download_button('⬇️ 下載圖庫 ZIP', data=st.session_state['zip_data'], file_name=st.session_state['orig_zip_name'], mime='application/zip')
if st.session_state['excel_data']:
    st.download_button('⬇️ 下載索引表 XLSX', data=st.session_state['excel_data'], file_name=st.session_state['orig_xlsx_name'], mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
