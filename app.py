"""
TRỘN ĐỀ WORD - THPT MINH ĐỨC (FIXED)
Sửa lỗi: IndentationError & Hiển thị HTML
Logic: Chia 3 phần (1-18, 19-22, 23-hết)
"""

import streamlit as st
import re
import random
import zipfile
import io
from xml.dom import minidom

# ==================== 1. CẤU HÌNH TRANG ====================
st.set_page_config(
    page_title="Trộn Đề Word - THPT Minh Đức",
    page_icon="🔥",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# ==================== 2. CSS & HTML (KHÔNG THỤT DÒNG) ====================
# Lưu ý: Các dòng trong biến chuỗi phải sát lề trái để tránh lỗi hiển thị

CUSTOM_CSS = """
<style>
    header {visibility: hidden;}
    .stApp > header {display: none;}
    .block-container {
        padding-top: 0rem !important;
        padding-bottom: 5rem !important;
        max-width: 100% !important;
        padding-left: 0 !important;
        padding-right: 0 !important;
    }
    [data-testid="stAppViewContainer"] {
        background-color: #f4f6f9;
    }
    .header-wrapper {
        background: linear-gradient(135deg, #ff9966 0%, #ff5e62 100%);
        padding-top: 3rem;
        padding-bottom: 6rem;
        text-align: center;
        color: white;
        border-bottom-left-radius: 50px;
        border-bottom-right-radius: 50px;
        box-shadow: 0 10px 20px rgba(255, 94, 98, 0.3);
        margin-bottom: -80px;
    }
    .school-tag {
        background: rgba(255, 255, 255, 0.25);
        border: 1px solid rgba(255,255,255,0.4);
        padding: 8px 25px;
        border-radius: 30px;
        font-weight: 900;
        font-size: 1.2rem;
        text-transform: uppercase;
        letter-spacing: 1px;
        display: inline-block;
        margin-bottom: 10px;
    }
    .main-h1 {
        font-size: 2.2rem;
        font-weight: 800;
        margin: 0;
        text-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    .sub-text {
        font-size: 1rem;
        opacity: 0.95;
        margin-top: 5px;
    }
    .btn-group-fake {
        margin-top: 20px;
        display: flex;
        justify-content: center;
        gap: 15px;
    }
    .btn-outline {
        border: 1px solid white;
        padding: 6px 20px;
        border-radius: 20px;
        font-weight: 600;
        font-size: 0.9rem;
    }
    .btn-filled {
        background: white;
        color: #ff5e62;
        padding: 6px 20px;
        border-radius: 20px;
        font-weight: 700;
        font-size: 0.9rem;
        box-shadow: 0 4px 10px rgba(0,0,0,0.1);
    }
    .info-gv {
        margin-top: 25px;
        background: rgba(255, 255, 255, 0.2);
        padding: 8px 25px;
        border-radius: 15px;
        display: inline-block;
        font-weight: 600;
        border: 1px dashed rgba(255,255,255,0.6);
    }
    .main-card {
        background: white;
        border-radius: 30px;
        width: 90%;
        max-width: 600px;
        margin: 0 auto;
        padding: 30px;
        box-shadow: 0 20px 50px rgba(0,0,0,0.1);
        position: relative;
        z-index: 99;
    }
    .stButton > button {
        background: linear-gradient(90deg, #ff9966, #ff5e62);
        color: white;
        border: none;
        height: 50px;
        border-radius: 12px;
        font-weight: bold;
        width: 100%;
        margin-top: 15px;
        font-size: 1.1rem;
    }
    .stButton > button:hover {
        transform: scale(1.02);
        color: white;
    }
    .footer {
        text-align: center;
        color: #aaa;
        margin-top: 40px;
        font-size: 0.85rem;
    }
</style>
"""

HEADER_HTML = """
<div class="header-wrapper">
<div class="school-tag">TRƯỜNG THPT MINH ĐỨC</div>
<div class="main-h1">TRỘN ĐỀ TRẮC NGHIỆM</div>
<div class="sub-text">Cấu trúc: P1(1-18), P2(19-22), P3(23-28)</div>
<div class="btn-group-fake">
<span class="btn-outline">Đăng nhập</span>
<span class="btn-filled">Đăng ký</span>
</div>
<div class="info-gv">GV: Nguyễn Văn Hà • Zalo: 0913968302</div>
</div>
"""

# ==================== 3. BACKEND (LOGIC XỬ LÝ) ====================
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

def get_text(block):
    texts = []
    for t in block.getElementsByTagNameNS(W_NS, "t"):
        if t.firstChild: texts.append(t.firstChild.nodeValue)
    return "".join(texts).strip()

def style_run_blue_bold(run):
    doc = run.ownerDocument
    rPr = run.getElementsByTagNameNS(W_NS, "rPr")
    if not rPr:
        rPr = doc.createElementNS(W_NS, "w:rPr")
        run.insertBefore(rPr, run.firstChild)
    else: rPr = rPr[0]
    
    color = rPr.getElementsByTagNameNS(W_NS, "color")
    if not color:
        color_el = doc.createElementNS(W_NS, "w:color")
        rPr.appendChild(color_el)
        color_el.setAttributeNS(W_NS, "w:val", "0000FF")
    
    bold = rPr.getElementsByTagNameNS(W_NS, "b")
    if not bold:
        b_el = doc.createElementNS(W_NS, "w:b")
        rPr.appendChild(b_el)

def parse_all_questions(blocks):
    intro = []
    questions = []
    i = 0
    while i < len(blocks):
        if re.match(r'^Câu\s*\d+', get_text(blocks[i])): break
        intro.append(blocks[i])
        i += 1
    
    while i < len(blocks):
        if re.match(r'^Câu\s*\d+', get_text(blocks[i])):
            group = [blocks[i]]
            i += 1
            while i < len(blocks):
                txt = get_text(blocks[i])
                if re.match(r'^Câu\s*\d+', txt) or "PHẦN" in txt.upper(): break
                group.append(blocks[i])
                i += 1
            questions.append(group)
        else:
            i += 1
    return intro, questions

def shuffle_options(q_block, mode="mcq"):
    pat = r'^\s*[A-D][\.\)]' if mode == "mcq" else r'^\s*[a-d][\.\)]'
    lbls = ["A.", "B.", "C.", "D."] if mode == "mcq" else ["a)", "b)", "c)", "d)"]
    
    indices = [x for x, b in enumerate(q_block) if re.match(pat, get_text(b))]
    if len(indices) >= 2:
        opts = [q_block[x] for x in indices]
        random.shuffle(opts)
        for x_idx, opt in zip(indices, opts):
            q_block[x_idx] = opt
            t_nodes = opt.getElementsByTagNameNS(W_NS, "t")
            for t in t_nodes:
                if t.firstChild:
                    cur_lbl = lbls[indices.index(x_idx)] if indices.index(x_idx) < 4 else ""
                    t.firstChild.nodeValue = re.sub(pat, cur_lbl, t.firstChild.nodeValue, 1)
                    style_run_blue_bold(t.parentNode)
                    break
    return q_block

def shuffle_docx(file_bytes, num_copies):
    out_zip = io.BytesIO()
    # SỬA LỖI INDENTATION TẠI ĐÂY: KHỐI WITH PHẢI THẲNG HÀNG
    with zipfile.ZipFile(out_zip, 'w', zipfile.ZIP_DEFLATED) as z_out:
        in_io = io.BytesIO(file_bytes)
        with zipfile.ZipFile(in_io, 'r') as z_in:
            xml = z_in.read("word/document.xml")
            
            for copy_i in range(num_copies):
                dom = minidom.parseString(xml)
                body = dom.getElementsByTagNameNS(W_NS, "body")[0]
                blocks = [n for n in body.childNodes if n.localName in ['p', 'tbl']]
                
                # Tách câu hỏi
                intro, all_qs = parse_all_questions(blocks)
                
                # CHIA 3 PHẦN
                # P1: 0-17 (18 câu), P2: 18-21 (4 câu), P3: 22-hết
                part1 = all_qs[0:18]
                part2 = all_qs[18:22]
                part3 = all_qs[22:]
                
                # Xử lý P1 (TN)
                p1_proc = [shuffle_options(q, "mcq") for q in part1]
                random.shuffle(p1_proc)
                
                # Xử lý P2 (Đ/S)
                p2_proc = [shuffle_options(q, "tf") for q in part2]
                random.shuffle(p2_proc)
                
                # Xử lý P3 (TLN) - Chỉ đảo câu
                p3_proc = [q for q in part3]
                random.shuffle(p3_proc)
                
                # Gộp lại
                final_qs = p1_proc + p2_proc + p3_proc
                final_blocks = intro[:]
                
                for idx, q_grp in enumerate(final_qs):
                    t_list = q_grp[0].getElementsByTagNameNS(W_NS, "t")
                    for t in t_list:
                        if t.firstChild and re.match(r'^Câu\s*\d+', t.firstChild.nodeValue):
                            t.firstChild.nodeValue = re.sub(r'^Câu\s*\d+', f"Câu {idx+1}", t.firstChild.nodeValue)
                            style_run_blue_bold(t.parentNode)
                            break
                    final_blocks.extend(q_grp)
                
                # Rebuild XML
                for n in list(body.childNodes):
                    if n.localName in ['p', 'tbl']: body.removeChild(n)
                for b in final_blocks: body.appendChild(b)
                
                # Save into Zip
                new_xml = dom.toxml().encode('utf-8')
                docx_io = io.BytesIO()
                with zipfile.ZipFile(docx_io, 'w', zipfile.ZIP_DEFLATED) as z_d:
                    for item in z_in.infolist():
                        if item.filename == "word/document.xml":
                            z_d.writestr(item.filename, new_xml)
                        else:
                            z_d.writestr(item.filename, z_in.read(item.filename))
                
                z_out.writestr(f"De_MinhDuc_Ma_{copy_i+1}.docx", docx_io.getvalue())
                
    return out_zip.getvalue()

# ==================== 4. RENDERING (HIỂN THỊ) ====================

# Kích hoạt CSS
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)
# Hiển thị Header
st.markdown(HEADER_HTML, unsafe_allow_html=True)

# Hiển thị Card
st.markdown('<div class="main-card">', unsafe_allow_html=True)

cols = st.columns(2)
with cols[0]:
    st.markdown('<div style="text-align:center; color:#ff5e62; font-weight:bold; border-bottom:3px solid #ff5e62; padding-bottom:5px;">⚡ Trộn đề</div>', unsafe_allow_html=True)
with cols[1]:
    st.markdown('<div style="text-align:center; color:#ccc; font-weight:500;">📷 QR Code</div>', unsafe_allow_html=True)

st.write("")
uploaded_file = st.file_uploader("Kéo thả file .docx vào đây", type=['docx'])

if uploaded_file:
    st.success(f"✅ Đã chọn: {uploaded_file.name}")

st.write("")
num = st.number_input("Số lượng đề", 1, 50, 4)

if st.button("🚀 TRỘN ĐỀ NGAY"):
    if not uploaded_file:
        st.warning("Vui lòng chọn file!")
    else:
        try:
            with st.spinner("Đang xử lý..."):
                res = shuffle_docx(uploaded_file.read(), num)
                st.success("Thành công!")
                st.download_button("📥 Tải xuống (ZIP)", res, "KetQua_MinhDuc.zip", "application/zip")
                st.balloons()
        except Exception as e:
            st.error(f"Lỗi: {e}")

st.markdown('</div>', unsafe_allow_html=True)
st.markdown('<div class="footer">© 2024 THPT Minh Đức</div>', unsafe_allow_html=True)
