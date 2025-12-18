"""
TRỘN ĐỀ WORD - THPT MINH ĐỨC
Phiên bản: Chia 3 phần cố định (1-18, 19-22, 23-28)
"""

import streamlit as st
import re
import random
import zipfile
import io
from xml.dom import minidom

# 1. CẤU HÌNH TRANG
st.set_page_config(
    page_title="Trộn Đề Word - THPT Minh Đức",
    page_icon="🔥",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# 2. CSS GIAO DIỆN (GIỮ NGUYÊN BẢN ĐẸP)
st.markdown("""
<style>
    /* Ẩn header mặc định */
    header {visibility: hidden;}
    .stApp > header {display: none;}
    
    /* Xóa lề */
    .block-container {
        padding-top: 0rem !important;
        padding-bottom: 5rem !important;
        max-width: 100% !important;
        padding-left: 0 !important;
        padding-right: 0 !important;
    }
    
    /* Nền trang */
    [data-testid="stAppViewContainer"] {
        background-color: #f4f6f9;
    }

    /* HEADER */
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

    /* Tên trường */
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
        font-size: 2.5rem;
        font-weight: 800;
        margin: 0;
        text-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    
    .sub-text {
        font-size: 1rem;
        opacity: 0.95;
        margin-top: 5px;
    }

    /* Nút giả lập */
    .btn-group-fake {
        margin-top: 20px;
        display: flex;
        justify-content: center;
        gap: 15px;
    }
    .btn-outline {
        border: 1px solid white;
        padding: 8px 20px;
        border-radius: 20px;
        font-weight: 600;
        cursor: pointer;
    }
    .btn-filled {
        background: white;
        color: #ff5e62;
        padding: 8px 20px;
        border-radius: 20px;
        font-weight: 700;
        box-shadow: 0 4px 10px rgba(0,0,0,0.1);
        cursor: pointer;
    }

    /* Thông tin GV */
    .info-gv {
        margin-top: 25px;
        background: rgba(255, 255, 255, 0.2);
        padding: 10px 25px;
        border-radius: 15px;
        display: inline-block;
        font-weight: 600;
        border: 1px dashed rgba(255,255,255,0.6);
    }

    /* CARD CHÍNH */
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
""", unsafe_allow_html=True)

# 3. HTML HEADER
HEADER_HTML = """
<div class="header-wrapper">
<div class="school-tag">TRƯỜNG THPT MINH ĐỨC</div>
<div class="main-h1">TRỘN ĐỀ TRẮC NGHIỆM</div>
<div class="sub-text">Chia 3 phần: TN (1-18), Đ/S (19-22), TLN (23-28)</div>
<div class="btn-group-fake">
<span class="btn-outline">Đăng nhập</span>
<span class="btn-filled">Đăng ký</span>
</div>
<div class="info-gv">GV: Nguyễn Văn Hà • Zalo: 0913968302</div>
</div>
"""
st.markdown(HEADER_HTML, unsafe_allow_html=True)

# 4. GIAO DIỆN CARD
st.markdown('<div class="main-card">', unsafe_allow_html=True)

# Tab giả lập
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
col_num, col_empty = st.columns([1, 1])
with col_num:
    num = st.number_input("Số lượng đề", 1, 50, 4)
with col_empty:
    st.info("Cấu trúc cố định:\nP1: Câu 1-18\nP2: Câu 19-22\nP3: Câu 23-28")

# ==================== LOGIC XỬ LÝ 3 PHẦN ====================
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
    """Tách toàn bộ file thành danh sách các câu hỏi"""
    intro = []
    questions = []
    i = 0
    # Lấy phần đầu (Intro)
    while i < len(blocks):
        if re.match(r'^Câu\s*\d+', get_text(blocks[i])): break
        intro.append(blocks[i])
        i += 1
    
    # Lấy các câu hỏi
    while i < len(blocks):
        if re.match(r'^Câu\s*\d+', get_text(blocks[i])):
            group = [blocks[i]]
            i += 1
            while i < len(blocks):
                txt = get_text(blocks[i])
                # Dừng nếu gặp Câu mới hoặc chữ PHẦN
                if re.match(r'^Câu\s*\d+', txt) or "PHẦN" in txt.upper(): break
                group.append(blocks[i])
                i += 1
            questions.append(group)
        else:
            i += 1
    return intro, questions

def shuffle_options_mcq(q_block):
    """Trộn A,B,C,D (Cho phần 1)"""
    indices = [x for x, b in enumerate(q_block) if re.match(r'^\s*[A-D][\.\)]', get_text(b))]
    if len(indices) >= 2:
        opts = [q_block[x] for x in indices]
        random.shuffle(opts)
        lbls = ["A.", "B.", "C.", "D."]
        for x_idx, opt in zip(indices, opts):
            q_block[x_idx] = opt
            t_nodes = opt.getElementsByTagNameNS(W_NS, "t")
            for t in t_nodes:
                if t.firstChild:
                    cur_lbl = lbls[indices.index(x_idx)] if indices.index(x_idx) < 4 else ""
                    # Thay thế ký tự đầu
                    t.firstChild.nodeValue = re.sub(r'^\s*[A-D][\.\)]', cur_lbl, t.firstChild.nodeValue, 1)
                    style_run_blue_bold(t.parentNode)
                    break
    return q_block

def shuffle_options_tf(q_block):
    """Trộn a,b,c,d (Cho phần 2 - Đúng Sai)"""
    # Tìm các dòng a) b) c) d)
    indices = [x for x, b in enumerate(q_block) if re.match(r'^\s*[a-d][\.\)]', get_text(b))]
    if len(indices) >= 2:
        opts = [q_block[x] for x in indices]
        random.shuffle(opts)
        lbls = ["a)", "b)", "c)", "d)"]
        for x_idx, opt in zip(indices, opts):
            q_block[x_idx] = opt
            t_nodes = opt.getElementsByTagNameNS(W_NS, "t")
            for t in t_nodes:
                if t.firstChild:
                    cur_lbl = lbls[indices.index(x_idx)] if indices.index(x_idx) < 4 else ""
                    t.firstChild.nodeValue = re.sub(r'^\s*[a-d][\.\)]', cur_lbl, t.firstChild.nodeValue, 1)
                    style_run_blue_bold(t.parentNode)
                    break
    return q_block

def shuffle_docx(file_bytes, num_copies):
    out_zip = io.BytesIO()
    with zipfile.ZipFile(out_zip, 'w', zipfile.ZIP_DEFLATED) as z_out:
