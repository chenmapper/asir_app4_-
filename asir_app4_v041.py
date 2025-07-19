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
st.set_page_config(page_title="📘 asir_app4_通用相片索引 v0.41一體化", layout="wide")
st.title("📘 asir_app4_通用相片索引 v0.41（初始索引 + 批次更名 + 縮圖嵌入）")

# Sidebar: 選擇模式及根目錄設定
mode = st.sidebar.selectbox("🔄 選擇功能模式", ["生成索引表與圖庫", "從 XLSX 批次更名", "批次截圖展示"])  # 新增第三模式
st.sidebar.header("⚙️ 根目錄設定")
root_dir = st.sidebar.text_input("根目錄路徑", value=r"C:/Users/User/Downloads/table_app/case4/temp_images")

# Initialize session state
for key in ['zip_data','excel_data','logs','orig_zip_name','orig_xlsx_name']:
    if key not in st.session_state:
        st.session_state[key] = None if key in ['zip_data','excel_data','orig_zip_name','orig_xlsx_name'] else []

# Filename validity regex
valid_pattern = re.compile(r'^[A-Za-z0-9_\-]+\.[A-Za-z0-9]+$')

# EXIF reader
def get_exif_datetime_and_status(path):
    try:
        img = Image.open(path)
        exif = img._getexif()
        if exif:
            for tag, val in exif.items():
                if ExifTags.TAGS.get(tag) == 'DateTimeOriginal':
                    return val.replace(':','-',2), '✅ 有拍攝時間'
            return '', '⚠️ 無 DateTimeOriginal'
        return '', '⚠️ 無 EXIF 資訊'
    except:
        return '', '❌ 讀取失敗'

# --- Mode: 生成索引表與圖庫 ---
if mode == "生成索引表與圖庫":
    st.markdown("### 📥 上傳圖片資料夾（zip格式）")
    uploaded_zip = st.file_uploader("ZIP 檔案", type="zip", key="gen_zip")
    if uploaded_zip:
        st.session_state['orig_zip_name'] = uploaded_zip.name
        if st.button("🧾 生成索引表與圖庫", key="gen_btn"):
            with st.spinner("生成索引表與圖庫中..."):
                # 使用扁平化處理避免路徑巢狀問題
                excel_bytes, zip_bytes = None, None
                with tempfile.TemporaryDirectory() as tmpdir:
                    flat_dir = Path(tmpdir) / "flat"
                    flat_dir.mkdir()
                    # 打開 ZIP 並提取所有圖片到 flat_dir
                    with zipfile.ZipFile(uploaded_zip, 'r') as z:
                        for info in z.infolist():
                            if not info.is_dir() and info.filename.lower().endswith(('.jpg','.jpeg','.png','.webp')):
                                name = Path(info.filename).name
                                dest = flat_dir / name
                                with z.open(info) as source, open(dest, 'wb') as target:
                                    shutil.copyfileobj(source, target)
                    imgs = list(flat_dir.iterdir())
                    # 準備資料
                    data = []
                    for p in imgs:
                        try:
                            mtime = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(p.stat().st_mtime))
                        except:
                            mtime = ''
                        exif_time, exif_status = get_exif_datetime_and_status(p)
                        file_url = f"file:///{(Path(root_dir)/p.name).as_posix()}"
                        size_kb = round(p.stat().st_size/1024, 2)
                        data.append({
                            '縮圖': p,
                            '目前檔名': p.name,
                            '新檔名': p.name,
                            '相片說明': '',
                            '原圖路徑': file_url,
                            '修改時間': mtime,
                            '拍攝時間': exif_time,
                            'EXIF狀態': exif_status,
                            '檔案大小(KB)': size_kb,
                            'gx': '', 'gy': '', 'gz': '',
                            '更名log': ''
                        })
                    # 生成 Excel
                    excel_buf = io.BytesIO()
                    wb = xlsxwriter.Workbook(excel_buf, {'in_memory': True})
                    ws = wb.add_worksheet('相片索引表')
                    headers = ['縮圖','目前檔名','新檔名','相片說明','原圖路徑','修改時間','拍攝時間','EXIF狀態','檔案大小(KB)','gx','gy','gz','更名log']
                    for i, h in enumerate(headers): ws.write(0, i, h)
                    for r, row in enumerate(data, start=1):
                        img = Image.open(row['縮圖']); img.thumbnail((120,120))
                        buf_img = io.BytesIO(); img.save(buf_img, 'PNG')
                        ws.set_row(r, 100)
                        ws.insert_image(r, 0, row['目前檔名'], {'image_data': buf_img})
                        for c, key in enumerate(headers[1:], start=1):
                            val = row[key]
                            if key == '原圖路徑': ws.write_url(r, c, val, string=val)
                            else: ws.write(r, c, val)
                    wb.close(); excel_buf.seek(0)
                    excel_bytes = excel_buf.getvalue()
                    # 打包圖片庫
                    zip_buf = io.BytesIO()
                    with zipfile.ZipFile(zip_buf, 'w', zipfile.ZIP_DEFLATED) as out_z:
                        for p in imgs:
                            out_z.write(p, arcname=p.name)
                    zip_buf.seek(0)
                    zip_bytes = zip_buf.getvalue()
                # 寫入 session
                st.session_state['excel_data'] = excel_bytes
                st.session_state['zip_data'] = zip_bytes
                st.session_state['logs'] = [f"✅ 共處理 {len(data)} 張圖片"]
                st.success("索引表與圖庫生成完成！")

elif mode == "從 XLSX 批次更名":

    st.markdown("### 📥 上傳原始圖庫 ZIP")
    uploaded_zip = st.file_uploader("ZIP 檔案", type="zip", key="upd_zip")
    st.markdown("### 📄 上傳索引表 XLSX（含 目前檔名 與 新檔名）")
    uploaded_xlsx = st.file_uploader("相片索引表", type="xlsx", key="upd_xlsx")
    if uploaded_zip: st.session_state['orig_zip_name'] = uploaded_zip.name
    if uploaded_xlsx: st.session_state['orig_xlsx_name'] = uploaded_xlsx.name
    if uploaded_zip and uploaded_xlsx and st.button('✅ 執行批次更名', key='upd_btn'):
        with st.spinner('批次處理中...'):
            with tempfile.TemporaryDirectory() as ext, tempfile.TemporaryDirectory() as outd:
                with zipfile.ZipFile(uploaded_zip, 'r') as z:
                    flist = [i.filename for i in z.infolist() if not i.is_dir()]
                    roots = {Path(f).parts[0] for f in flist if len(Path(f).parts) > 1}
                    updir = roots.pop() if len(roots) == 1 else ''
                    z.extractall(ext)
                base = Path(ext)/updir if updir else Path(ext)
                df = pd.read_excel(uploaded_xlsx).fillna('')
                if '目前檔名' not in df.columns:
                    st.error("索引表必須包含「目前檔名」欄位"); st.stop()
                olds = df['目前檔名'].astype(str).str.strip().tolist()
                raws = df.get('新檔名', pd.Series(['']*len(olds))).astype(str).str.strip().tolist()
                finals, logs, rlogs = [], [], []
                for old, raw in zip(olds, raws):
                    if raw and not Path(raw).suffix: cand = f"{raw}{Path(old).suffix}"
                    elif raw: cand = raw
                    else: cand = old
                    if valid_pattern.match(cand): logs.append(f"✅ {old} → {cand}"); rlogs.append(old); finals.append(cand)
                    else: logs.append(f"⚠️ 跳過: {raw}"); rlogs.append(''); finals.append(old)
                for old, new in zip(olds, finals):
                    src = base/old; dst = Path(outd)/new
                    if src.exists(): shutil.copy(src, dst)
                excel_buf = io.BytesIO()
                wb2 = xlsxwriter.Workbook(excel_buf, {'in_memory': True})
                ws2 = wb2.add_worksheet('更新索引表')
                hdrs = ['縮圖','目前檔名','新檔名','相片說明','原圖路徑','修改時間','拍攝時間','EXIF狀態','檔案大小(KB)','gx','gy','gz','更名log']
                for i, h in enumerate(hdrs): ws2.write(0, i, h)
                for r, (old, new, raw, lg) in enumerate(zip(olds, finals, raws, rlogs), start=1):
                    dstp = Path(outd)/new
                    try:
                        img = Image.open(dstp); img.thumbnail((80,80))
                        buf2 = io.BytesIO(); img.save(buf2, 'PNG')
                        ws2.set_row(r, 60)
                        ws2.insert_image(r, 0, new, {'image_data': buf2, 'x_scale': 1, 'y_scale': 1})
                    except:
                        ws2.write(r, 0, '')
                    url2 = f"file:///{(Path(root_dir)/updir/new).as_posix()}" if updir else f"file:///{(Path(root_dir)/new).as_posix()}"
                    m2 = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime((Path(outd)/new).stat().st_mtime))
                    e2, s2 = get_exif_datetime_and_status(Path(outd)/new)
                    sz2 = round((Path(outd)/new).stat().st_size/1024, 2)
                    gx2 = df.at[r-1, 'gx'] if 'gx' in df.columns else ''
                    gy2 = df.at[r-1, 'gy'] if 'gy' in df.columns else ''
                    gz2 = df.at[r-1, 'gz'] if 'gz' in df.columns else ''
                    vals = [new, raw, '', url2, m2, e2, s2, sz2, gx2, gy2, gz2, lg]
                    for c, v in enumerate(vals, start=1): ws2.write(r, c, v)
                wb2.close(); excel_buf.seek(0)
                zip_buf2 = io.BytesIO()
                with zipfile.ZipFile(zip_buf2, 'w', zipfile.ZIP_DEFLATED) as zf2:
                    for f in Path(outd).iterdir(): zf2.write(f, arcname=f.name)
                zip_buf2.seek(0)
                st.session_state['excel_data'] = excel_buf.getvalue()
                st.session_state['zip_data'] = zip_buf2.getvalue()
                st.session_state['logs'] = logs
                st.session_state['orig_zip_name'] = uploaded_zip.name
                st.session_state['orig_xlsx_name'] = uploaded_xlsx.name
                st.success('批次更名並嵌入縮圖完成！')

# --- Mode: 批次截圖展示 ---
elif mode == "批次截圖展示":
    st.markdown("### 📥 上傳圖片資料夾以進行批次正方形裁切展示（zip格式）")
    zip_file = st.file_uploader("ZIP 檔案", type="zip", key="sq_zip")
    display_size = st.number_input("展示大圖邊長 (px)", min_value=100, value=800)
    if zip_file and st.button("🖼️ 生成截圖展示", key="sq_btn"):
        with st.spinner("生成批次截圖展示中..."):
            # 解壓並裁切所有圖片於暫存內
            with tempfile.TemporaryDirectory() as tmp_scr:
                flat = Path(tmp_scr)/"flat"
                flat.mkdir()
                with zipfile.ZipFile(zip_file, 'r') as z:
                    for info in z.infolist():
                        if not info.is_dir() and info.filename.lower().endswith(('.jpg','.jpeg','.png','.webp')):
                            fn = Path(info.filename).name
                            dest = flat/fn
                            with z.open(info) as src, open(dest,'wb') as dst:
                                shutil.copyfileobj(src,dst)
                imgs = list(flat.iterdir())
                n = len(imgs)
                if n == 0:
                    st.error("沒有檢測到任何圖片！")
                    st.stop()
                # 裁切正方形
                squares = []
                for p in imgs:
                    img = Image.open(p)
                    w,h = img.size
                    if w >= h:
                        left = (w - h) // 2
                        box = (left, 0, left + h, h)
                    else:
                        top = (h - w) // 2
                        box = (0, top, w, top + w)
                    squares.append(img.crop(box))
                # 計算格數與尺寸
                import math
                q = int(math.sqrt(n))
                if q * q == n:
                    cols = rows = q
                else:
                    cols = q + 1
                    rows = math.ceil(n / cols)
                k = int(display_size / cols)
                # 建立畫布並貼上
                canvas = Image.new('RGB', (display_size, display_size), (255,255,255))
                for idx, sq in enumerate(squares):
                    sqr = sq.resize((k, k))
                    x = (idx % cols) * k
                    y = (idx // cols) * k
                    canvas.paste(sqr, (x, y))
                # 完成後暫存自動清除，canvas 存於記憶體
            # 顯示與下載
            st.image(canvas, caption=f"共 {n} 張裁切後正方圖", use_column_width=False)
            buf = io.BytesIO()
            canvas.save(buf, 'PNG')
            buf.seek(0)
            st.download_button('⬇️ 下載展示圖片', buf, file_name='batch_squares.png', mime='image/png')

# 顯示日誌與下載按鈕
if st.session_state['logs']:
    st.markdown('### 📜 執行日誌')
    for ln in st.session_state['logs']:
        st.write(ln)
if st.session_state['zip_data']:
    st.download_button('⬇️ 下載圖庫 ZIP', data=st.session_state['zip_data'], file_name=st.session_state['orig_zip_name'], mime='application/zip')
if st.session_state['excel_data']:
    st.download_button('⬇️ 下載索引表 XLSX', data=st.session_state['excel_data'], file_name=st.session_state['orig_xlsx_name'], mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
