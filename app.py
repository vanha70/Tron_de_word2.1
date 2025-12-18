"""
TRỘN ĐỀ WORD - THPT MINH ĐỨC
Phiên bản sửa lỗi hiển thị HTML vĩnh viễn
"""

import streamlit as st
import re
import random
import zipfile
import io
from xml.dom import minidom

# ==================== 1. CẤU HÌNH TRANG (BẮT BUỘC ĐẦU TIÊN) ====================
st.set_page_config(
    page_title="Trộn Đề Word - THPT Minh Đức",
    page_icon="🔥",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# ==================== 2. BIẾN HTML (ĐỂ RIÊNG TRÁNH LỖI) ====================

# CSS ĐÃ TỐI ƯU ĐỂ KHÔNG CÒN KHOẢNG TRẮNG
CUSTOM_CSS = """
<style>
    /* Ẩn header mặc định của Streamlit */
    header {visibility: hidden;}
    .stApp > header {display: none;}
    
    /* Xóa lề trắng mặc định để Header full màn hình */
    .block-container {
        padding-top: 0rem !important;
        padding-bottom: 2rem !important;
        max-width: 100% !important;
        padding-left: 0 !important;
        padding-right: 0 !important;
    }
    
    /* Thiết lập nền */
    [data-testid="stAppViewContainer"] {
        background-color: #f4f6f9;
    }

    /* HEADER FULL WIDTH MÀU CAM */
    .header-wrapper {
        width: 100vw;
        background: linear-gradient(135deg, #ff9966 0%, #ff5e62 100%); /* Màu cam đỏ sunset */
        padding-top: 3rem;
        padding-bottom: 6rem; /* Để dư chỗ cho Card đè lên */
        text-align: center;
        color: white;
        border-bottom-left-radius: 50px;
        border-bottom-right-radius: 50px;
        box-shadow: 0 10px 20px rgba(255, 94, 98, 0.3);
        margin-bottom: -70px; /* Kéo nội dung lên trên */
    }

    /* Tên trường */
    .school-tag {
        background: rgba(255, 255, 255, 0.25);
        border: 1px solid rgba(255,255,255,0.4);
        padding: 8px 25px;
        border-radius: 30px;
        font-weight: 900; /* ĐẬM */
        font-size: 1.2rem;
        text-transform: uppercase; /* IN HOA */
        letter-spacing: 1px;
        display: inline-block;
        margin-bottom: 10px;
        backdrop-filter: blur(5px);
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

    /* Nút Đăng nhập/Đăng ký (Giả lập) */
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
        font-size: 0.9rem;
    }
    .btn-filled {
        background: white;
        color: #ff5e62;
        padding: 8px 20px;
        border-radius: 20px;
        font-weight: 700;
        font-size: 0.9rem;
        box-shadow: 0 4px 10px rgba(0,0,0,0.1);
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

    /* CARD GIAO DIỆN CHÍNH */
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

    /* Custom Streamlit Elements */
    .stButton > button {
        background: linear-gradient(90deg, #ff9966, #ff5e62);
        color: white;
        border: none;
        height: 50px;
        border-radius: 12px;
        font-weight: bold;
        font-size: 1.1rem;
        width: 100%;
        margin-top: 15px;
        transition: transform 0.2s;
    }
    .stButton > button:hover {
        transform: scale(1.02);
        color: white;
    }
    
    /* Footer */
    .footer {
        text-align: center;
        color: #aaa;
        margin-top: 40px;
        font-size: 0.85rem;
    }
</style>
"""

# HTML HEADER (KHÔNG ĐƯỢC THỤT ĐẦU DÒNG NỘI DUNG)
HEADER_HTML = """
<div class="header-wrapper">
    <div class="school-tag">TRƯỜNG THPT MINH ĐỨC</div>
    <div class="main-h1">TRỘN ĐỀ TRẮC NGHIỆM</div>
    <div class="sub-text">Chuẩn bị tài liệu định dạng đúng để trộn đề nhanh chóng</div>
    
    <div class="btn-group-fake">
        <span class="btn-outline">Đăng nhập</span>
        <span class="btn-filled">Đăng ký</span>
    </div>
    
    <div class="info-gv">
        GV: Nguyễn Văn Hà • Zalo: 0913968302
    </div>
</div>
"""

# ==================== 3. LOGIC XỬ LÝ (BACKEND) ====================
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

def get_text(block):
    texts = []
    t_nodes = block.getElementsByTagNameNS(W_NS, "t")
    for t in t_nodes:
        if t.firstChild and t.firstChild.nodeValue:
            texts.append(t.firstChild.nodeValue)
    return "".join(texts).strip()

def shuffle_array(arr):
    out = arr.copy()
    for i in range(len(out) - 1, 0, -1):
        j = random.randint(0, i)
        out[i], out[j] = out[j], out[i]
    return out

def style_run_blue_bold(run):
    doc = run.ownerDocument
    rPr_list = run.getElementsByTagNameNS(W_NS, "rPr")
    if rPr_list: rPr = rPr_list[0]
    else:
        rPr = doc.createElementNS(W_NS, "w:rPr")
        run.insertBefore(rPr, run.firstChild)
    color_list = rPr.getElementsByTagNameNS(W_NS, "color")
    if not color_list:
        color_el = doc.createElementNS(W_NS, "w:color")
        rPr.appendChild(color_el)
        color_el.setAttributeNS(W_NS, "w:val", "0000FF")
    b_list = rPr.getElementsByTagNameNS(W_NS, "b")
    if not b_list:
        b_el = doc.createElementNS(W_NS, "w:b")
        rPr.appendChild(b_el)

def update_label(paragraph, new_text):
    t_nodes = paragraph.getElementsByTagNameNS(W_NS, "t")
    if not t_nodes: return
    for t in t_nodes:
        if t.firstChild:
            t.firstChild.nodeValue = new_text
            if t.parentNode.localName == "r":
                style_run_blue_bold(t.parentNode)
            break

def parse_questions_in_range(blocks):
    intro, questions = [], []
    i = 0
    while i < len(blocks):
        txt = get_text(blocks[i])
        if re.match(r'^Câu\s*\d+\b', txt): break
        intro.append(blocks[i])
        i += 1
    while i < len(blocks):
        txt = get_text(blocks[i])
        if re.match(r'^Câu\s*\d+\b', txt):
            group = [blocks[i]]
            i += 1
            while i < len(blocks):
                t2 = get_text(blocks[i])
                if re.match(r'^Câu\s*\d+\b', t2) or "PHẦN" in t2.upper(): break
                group.append(blocks[i])
                i += 1
            questions.append(group)
        else:
            i += 1
    return intro, questions

def simple_shuffle_options(question_blocks, mode="mcq"):
    indices = []
    pattern = r'^\s*[A-D][\.\)]' if mode == "mcq" else r'^\s*[a-d][\.\)]'
    for k, block in enumerate(question_blocks):
        if re.match(pattern, get_text(block), re.IGNORECASE):
            indices.append(k)
    if len(indices) >= 2:
        opts = [question_blocks[k] for k in indices]
        random.shuffle(opts)
        for idx_in_q, opt_block in zip(indices, opts):
            question_blocks[idx_in_q] = opt_block
            labels = ["A.", "B.", "C.", "D."] if mode == "mcq" else ["a)", "b)", "c)", "d)"]
            current_label = labels[indices.index(idx_in_q)] if indices.index(idx_in_q) < 4 else ""
            t_nodes = opt_block.getElementsByTagNameNS(W_NS, "t")
            for t in t_nodes:
                if t.firstChild and t.firstChild.nodeValue:
                    old_txt = t.firstChild.nodeValue
                    new_txt = re.sub(pattern, current_label, old_txt, 1)
                    t.firstChild.nodeValue = new_txt
                    if t.parentNode.localName == "r": style_run_blue_bold(t.parentNode)
                    break
    return question_blocks

def process_docx(file_bytes, num_copies, mode):
    output_zip = io.BytesIO()
    with zipfile.ZipFile(output_zip, 'w', zipfile.ZIP_DEFLATED) as z:
        input_io = io.BytesIO(file_bytes)
        with zipfile.ZipFile(input_io, 'r') as zin:
            xml_content = zin.read("word/document.xml")
            for copy_i in range(num_copies):
                dom = minidom.parseString(xml_content)
                body = dom.getElementsByTagNameNS(W_NS, "body")[0]
                blocks = [n for n in body.childNodes if n.nodeType == n.ELEMENT_NODE and n.localName in ['p', 'tbl']]
                intro, questions = parse_questions_in_range(blocks)
                random.shuffle(questions)
                final_blocks = intro[:]
                for q_idx, q_blocks in enumerate(questions):
                    first_p = q_blocks[0]
                    t_nodes = first_p.getElementsByTagNameNS(W_NS, "t")
                    for t in t_nodes:
                        if t.firstChild:
                            val = t.firstChild.nodeValue
                            if re.match(r'^Câu\s*\d+', val):
                                t.firstChild.nodeValue = re.sub(r'^Câu\s*\d+', f"Câu {q_idx+1}", val)
                                style_run_blue_bold(t.parentNode)
                                break
                    q_blocks = simple_shuffle_options(q_blocks, mode)
                    final_blocks.extend(q_blocks)
                for n in list(body.childNodes):
                     if n.nodeType == n.ELEMENT_NODE and n.localName in ['p', 'tbl']:
                         body.removeChild(n)
                for b in final_blocks:
                    body.appendChild(b)
                new_xml = dom.toxml().encode('utf-8')
                filename = f"De_Tron_Ma_{copy_i+1}.docx"
                single_docx_io = io.BytesIO()
                with zipfile.ZipFile(single_docx_io, 'w', zipfile.ZIP_DEFLATED) as z_single:
                    for item in zin.infolist():
                        if item.filename == "word/document.xml":
                            z_single.writestr(item.filename, new_xml)
                        else:
                            z_single.writestr(item.filename, zin.read(item.filename))
                z.writestr(filename, single_docx_io.getvalue())
    return output_zip.getvalue()

# ==================== 4. GIAO DIỆN CHÍNH (MAIN) ====================
def main():
    # 1. Kích hoạt CSS
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)
    
    # 2. Hiển thị Header (QUAN TRỌNG: unsafe_allow_html=True)
    st.markdown(HEADER_HTML, unsafe_allow_html=True)
    
    # 3. Card Giao diện
    st.markdown('<div class="main-card">', unsafe_allow_html=True)
    
    # --- Tab giả lập ---
    cols = st.columns(2)
    with cols[0]:
        st.markdown('<div style="text-align:center; color:#ff5e62; font-weight:bold; border-bottom:2px solid #ff5e62; padding-bottom:5px;">⚡ Trộn đề</div>', unsafe_allow_html=True)
    with cols[1]:
        st.markdown('<div style="text-align:center; color:#ccc; font-weight:500;">📷 QR Code</div>', unsafe_allow_html=True)
    
    st.write("") # Khoảng cách

    # --- Upload File ---
    uploaded_file = st.file_uploader("Kéo thả file .docx vào đây", type=['docx'])
    
    if uploaded_file:
        st.success(f"✅ Đã chọn: {uploaded_file.name}")

    st.write("")
    
    # --- Tùy chọn ---
    col1, col2 = st.columns(2)
    with col1:
        mode = st.selectbox("Chế độ", ["Trắc nghiệm (MCQ)", "Đúng/Sai (TF)"])
        mode_val = "mcq" if "MCQ" in mode else "tf"
    with col2:
        num = st.number_input("Số lượng đề", 1, 50, 4)
        
    # --- Nút xử lý ---
    if st.button("🚀 BẮT ĐẦU TRỘN"):
        if not uploaded_file:
            st.warning("⚠️ Vui lòng chọn file Word trước!")
        else:
            try:
                with st.spinner("⏳ Đang xử lý..."):
                    file_bytes = uploaded_file.read()
                    result_zip = process_docx(file_bytes, num, mode_val)
                    st.success("✅ Thành công!")
                    st.download_button(
                        label="📥 Tải xuống (ZIP)",
                        data=result_zip,
                        file_name="KetQua_MinhDuc.zip",
                        mime="application/zip"
                    )
                    st.balloons()
            except Exception as e:
                st.error(f"❌ Lỗi: {str(e)}")

    st.markdown('</div>', unsafe_allow_html=True) # End card
    
    # Footer
    st.markdown('<div class="footer">© 2024 THPT Minh Đức</div>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
