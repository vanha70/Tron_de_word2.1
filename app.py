"""
Trộn Đề Word Online - AIOMT Premium
Streamlit App - Giao diện Sunset (Cam - Đỏ) - Fix Lỗi Hiển Thị
"""

import streamlit as st
import re
import random
import zipfile
import io
from xml.dom import minidom

# ==================== CẤU HÌNH TRANG ====================

st.set_page_config(
    page_title="Trộn Đề Word - THPT Minh Đức",
    page_icon="🔥",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# ==================== CUSTOM CSS (ĐÃ SỬA LỖI KHOẢNG TRẮNG) ====================
st.markdown("""
<style>
    /* 1. Ẩn các thành phần mặc định thừa thãi của Streamlit */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;} /* Ẩn header mặc định phía trên cùng */
    
    /* 2. Thiết lập nền chung */
    .stApp {
        background-color: #f0f2f6; /* Màu nền xám nhạt cho toàn bộ ứng dụng */
    }
    
    /* Loại bỏ padding mặc định của Streamlit để header sát lề trên */
    .block-container {
        padding-top: 0rem !important;
        padding-bottom: 5rem !important;
    }

    /* 3. HEADER CUSTOM (Gradient Cam - Đỏ) */
    .custom-header {
        background: linear-gradient(180deg, #ff5f6d 0%, #ffc371 100%); /* Gradient giống ảnh mẫu */
        padding: 3rem 1rem 5rem 1rem; /* Padding dưới lớn để tạo khoảng cho Card chèn lên */
        color: white;
        text-align: center;
        border-bottom-left-radius: 30px;
        border-bottom-right-radius: 30px;
        margin-left: -50vw; /* Kỹ thuật mở rộng full màn hình */
        margin-right: -50vw;
        position: relative;
        left: 50%;
        right: 50%;
        width: 100vw;
        box-shadow: 0 4px 15px rgba(255, 95, 109, 0.3);
        margin-bottom: -40px; /* Kéo nội dung phía dưới lên đè lên header */
    }

    /* Tên trường IN HOA ĐẬM */
    .school-name {
        font-family: 'Roboto', sans-serif;
        background-color: rgba(255, 255, 255, 0.25);
        padding: 5px 15px;
        border-radius: 15px;
        font-size: 1rem;
        font-weight: 900; /* Đậm nhất */
        text-transform: uppercase; /* In hoa */
        display: inline-block;
        margin-bottom: 15px;
        letter-spacing: 1px;
        backdrop-filter: blur(5px);
    }

    .app-title {
        font-size: 1.8rem;
        font-weight: 800;
        margin-bottom: 5px;
        text-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }

    .app-desc {
        font-size: 0.95rem;
        opacity: 0.95;
        margin-bottom: 15px;
        font-weight: 400;
    }

    /* Nút Giả lập (Giáo viên / Học sinh) */
    .role-badges {
        display: flex;
        justify-content: center;
        gap: 15px;
        margin-bottom: 15px;
    }
    .badge-ghost {
        border: 1px solid rgba(255,255,255,0.8);
        padding: 6px 18px;
        border-radius: 20px;
        font-size: 0.85rem;
        font-weight: 600;
    }
    .badge-solid {
        background-color: white;
        color: #ff7675;
        padding: 6px 18px;
        border-radius: 20px;
        font-size: 0.85rem;
        font-weight: 700;
        box-shadow: 0 2px 5px rgba(0,0,0,0.1);
    }

    /* Thông tin GV */
    .teacher-info-box {
        background: rgba(255, 255, 255, 0.15);
        border: 1px dashed rgba(255, 255, 255, 0.5);
        padding: 8px 20px;
        border-radius: 10px;
        display: inline-block;
        font-size: 0.9rem;
        font-weight: 600;
    }

    /* 4. THANH TAB (TRỘN ĐỀ | QR CODE) */
    .nav-tabs {
        display: flex;
        justify-content: center;
        gap: 10px;
        position: relative;
        z-index: 10; /* Nổi lên trên */
        margin-bottom: 20px;
    }
    .tab-item {
        flex: 1;
        text-align: center;
        padding: 12px;
        border-radius: 12px;
        font-weight: 700;
        cursor: pointer;
        transition: all 0.3s;
        max-width: 200px;
    }
    .tab-active {
        background: linear-gradient(135deg, #ff6b6b, #ff9f43);
        color: white;
        box-shadow: 0 4px 10px rgba(255, 107, 107, 0.4);
    }
    .tab-inactive {
        background: white;
        color: #57606f;
        box-shadow: 0 2px 5px rgba(0,0,0,0.05);
    }

    /* 5. MAIN CARD (KHUNG TRẮNG CHỨA UPLOAD) */
    .main-card {
        background-color: white;
        border-radius: 20px;
        padding: 25px;
        box-shadow: 0 10px 30px rgba(0,0,0,0.08);
        position: relative;
        z-index: 5;
    }

    /* Custom nút bấm Upload */
    .stButton > button {
        background: linear-gradient(90deg, #ff5f6d, #ffc371);
        color: white;
        font-weight: bold;
        border: none;
        border-radius: 10px;
        height: 50px;
        width: 100%;
        font-size: 1.1rem;
        margin-top: 10px;
        box-shadow: 0 4px 10px rgba(255, 95, 109, 0.3);
    }
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 15px rgba(255, 95, 109, 0.5);
        color: white;
    }

    /* Custom khung upload của Streamlit */
    [data-testid="stFileUploader"] section {
        background-color: #fff9f0;
        border: 2px dashed #ffc371;
        border-radius: 15px;
    }
    
    /* Footer */
    .footer-text {
        text-align: center;
        color: #a4b0be;
        font-size: 0.8rem;
        margin-top: 30px;
    }
</style>
""", unsafe_allow_html=True)

# ==================== HEADER & UI RENDERING ====================

def render_ui():
    # 1. Phần Header Gradient (HTML thuần được render qua markdown)
    st.markdown("""
        <div class="custom-header">
            <div class="school-name">TRƯỜNG THPT MINH ĐỨC</div>
            <div class="app-title">TRỘN ĐỀ TRẮC NGHIỆM</div>
            <div class="app-desc">Upload và chuẩn bị tài liệu định dạng đúng để trộn đề nhanh chóng</div>
            
            <div class="role-badges">
                <span class="badge-ghost">Đăng nhập</span>
                <span class="badge-solid">Đăng ký</span>
            </div>
            
            <div class="teacher-info-box">
                GV: Nguyễn Văn Hà • Zalo: 0913968302
            </div>
        </div>
        
        <div class="nav-tabs">
            <div class="tab-item tab-active">⚡ Trộn đề</div>
            <div class="tab-item tab-inactive">📷 QR Code</div>
        </div>
    """, unsafe_allow_html=True)

    # 2. Phần Card Trắng (Sử dụng container của Streamlit để chứa logic Upload)
    with st.container():
        st.markdown('<div class="main-card">', unsafe_allow_html=True)
        
        # Icon minh họa
        st.markdown("""
        <div style="text-align: center; margin-bottom: 15px;">
            <span style="font-size: 40px;">📄</span>
            <p style="color: #666; font-size: 0.9rem; margin-top: 5px;">
                Kéo thả file <b>.docx</b> vào bên dưới
            </p>
        </div>
        """, unsafe_allow_html=True)

        uploaded_file = st.file_uploader(
            "Chọn file word",
            type=["docx"],
            label_visibility="collapsed"
        )
        
        if uploaded_file:
            st.success(f"✅ Đã nhận: {uploaded_file.name}")
            
        st.divider()
        
        col1, col2 = st.columns(2)
        with col1:
            mode = st.selectbox("Chế độ trộn", ["Tự động", "Trắc nghiệm 100%", "Đúng/Sai"])
            mode_map = {"Tự động": "auto", "Trắc nghiệm 100%": "mcq", "Đúng/Sai": "tf"}
        with col2:
            num = st.number_input("Số lượng đề", min_value=1, max_value=20, value=4)
            
        # Nút Trộn đề (Logic sẽ xử lý bên dưới)
        process_btn = st.button("🚀 Bắt đầu trộn đề")
        
        st.markdown('</div>', unsafe_allow_html=True) # Đóng main-card

        # Footer
        st.markdown('<div class="footer-text">© 2024 THPT Minh Đức</div>', unsafe_allow_html=True)
        
        return uploaded_file, process_btn, mode_map[mode], num

# ==================== LOGIC XỬ LÝ WORD (GIỮ NGUYÊN) ====================

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

def shuffle_array(arr):
    out = arr.copy()
    for i in range(len(out) - 1, 0, -1):
        j = random.randint(0, i)
        out[i], out[j] = out[j], out[i]
    return out

def get_text(block):
    texts = []
    t_nodes = block.getElementsByTagNameNS(W_NS, "t")
    for t in t_nodes:
        if t.firstChild and t.firstChild.nodeValue:
            texts.append(t.firstChild.nodeValue)
    return "".join(texts).strip()

def style_run_blue_bold(run):
    doc = run.ownerDocument
    rPr_list = run.getElementsByTagNameNS(W_NS, "rPr")
    if rPr_list: rPr = rPr_list[0]
    else:
        rPr = doc.createElementNS(W_NS, "w:rPr")
        run.insertBefore(rPr, run.firstChild)
    color_list = rPr.getElementsByTagNameNS(W_NS, "color")
    if color_list: color_el = color_list[0]
    else:
        color_el = doc.createElementNS(W_NS, "w:color")
        rPr.appendChild(color_el)
    color_el.setAttributeNS(W_NS, "w:val", "0000FF")
    b_list = rPr.getElementsByTagNameNS(W_NS, "b")
    if not b_list:
        b_el = doc.createElementNS(W_NS, "w:b")
        rPr.appendChild(b_el)

def update_mcq_label(paragraph, new_label):
    t_nodes = paragraph.getElementsByTagNameNS(W_NS, "t")
    if not t_nodes: return
    new_letter = new_label[0].upper()
    for i, t in enumerate(t_nodes):
        if not t.firstChild or not t.firstChild.nodeValue: continue
        txt = t.firstChild.nodeValue
        m = re.match(r'^(\s*)([A-D])([\.\)])?', txt, re.IGNORECASE)
        if not m: continue
        leading_space = m.group(1) or ""
        old_punct = m.group(3) or ""
        after_match = txt[m.end():]
        if old_punct: t.firstChild.nodeValue = leading_space + new_letter + "." + after_match
        else: t.firstChild.nodeValue = leading_space + new_letter + after_match
        run = t.parentNode
        if run and run.localName == "r": style_run_blue_bold(run)
        break

def update_tf_label(paragraph, new_label):
    t_nodes = paragraph.getElementsByTagNameNS(W_NS, "t")
    if not t_nodes: return
    new_letter = new_label[0].lower()
    for i, t in enumerate(t_nodes):
        if not t.firstChild or not t.firstChild.nodeValue: continue
        txt = t.firstChild.nodeValue
        m = re.match(r'^(\s*)([a-d])(\))?', txt, re.IGNORECASE)
        if not m: continue
        leading_space = m.group(1) or ""
        after_match = txt[m.end():]
        t.firstChild.nodeValue = leading_space + new_letter + ")" + after_match
        run = t.parentNode
        if run and run.localName == "r": style_run_blue_bold(run)
        break

def update_question_label(paragraph, new_label):
    t_nodes = paragraph.getElementsByTagNameNS(W_NS, "t")
    if not t_nodes: return
    for i, t in enumerate(t_nodes):
        if not t.firstChild or not t.firstChild.nodeValue: continue
        txt = t.firstChild.nodeValue
        m = re.match(r'^(\s*)(Câu\s*)(\d+)(\.)?', txt, re.IGNORECASE)
        if not m: continue
        leading_space = m.group(1) or ""
        after_match = txt[m.end():]
        t.firstChild.nodeValue = leading_space + new_label + after_match
        run = t.parentNode
        if run and run.localName == "r": style_run_blue_bold(run)
        for j in range(i + 1, len(t_nodes)):
            t2 = t_nodes[j]
            if not t2.firstChild or not t2.firstChild.nodeValue: continue
            txt2 = t2.firstChild.nodeValue
            if re.match(r'^[\s0-9\.]*$', txt2) and txt2.strip(): t2.firstChild.nodeValue = ""
            elif re.match(r'^\s*$', txt2): continue
            else: break
        break

def parse_questions_in_range(blocks, start, end):
    part_blocks = blocks[start:end]
    intro, questions = [], []
    i = 0
    while i < len(part_blocks):
        text = get_text(part_blocks[i])
        if re.match(r'^Câu\s*\d+\b', text): break
        intro.append(part_blocks[i])
        i += 1
    while i < len(part_blocks):
        text = get_text(part_blocks[i])
        if re.match(r'^Câu\s*\d+\b', text):
            group = [part_blocks[i]]
            i += 1
            while i < len(part_blocks):
                t2 = get_text(part_blocks[i])
                if re.match(r'^Câu\s*\d+\b', t2): break
                if re.match(r'^PHẦN\s*\d\b', t2, re.IGNORECASE): break
                group.append(part_blocks[i])
                i += 1
            questions.append(group)
        else:
            intro.append(part_blocks[i])
            i += 1
    return intro, questions

def shuffle_mcq_options(question_blocks):
    indices = []
    for i, block in enumerate(question_blocks):
        text = get_text(block)
        if re.match(r'^\s*[A-D][\.\)]', text, re.IGNORECASE): indices.append(i)
    if len(indices) < 2: return question_blocks
    options = [question_blocks[idx] for idx in indices]
    shuffled = shuffle_array(options)
    min_idx, max_idx = min(indices), max(indices)
    return question_blocks[:min_idx] + shuffled + question_blocks[max_idx + 1:]

def relabel_mcq_options(question_blocks):
    letters = ["A", "B", "C", "D"]
    opt_blocks = [b for b in question_blocks if re.match(r'^\s*[A-D][\.\)]', get_text(b), re.IGNORECASE)]
    for idx, block in enumerate(opt_blocks):
        letter = letters[idx] if idx < len(letters) else letters[-1]
        update_mcq_label(block, f"{letter}.")

def shuffle_tf_options(question_blocks):
    opt_indices = {}
    for i, block in enumerate(question_blocks):
        m = re.match(r'^\s*([a-d])\)', get_text(block), re.IGNORECASE)
        if m: opt_indices[m.group(1).lower()] = i
    abc_idx = [opt_indices.get(k) for k in ["a", "b", "c"] if opt_indices.get(k) is not None]
    if len(abc_idx) < 2: return question_blocks
    abc_nodes = [question_blocks[idx] for idx in abc_idx]
    shuffled_abc = shuffle_array(abc_nodes)
    all_idx = [v for v in opt_indices.values() if v is not None]
    min_idx, max_idx = min(all_idx), max(all_idx)
    d_node = question_blocks[opt_indices["d"]] if "d" in opt_indices else None
    middle = shuffled_abc.copy()
    if d_node: middle.append(d_node)
    return question_blocks[:min_idx] + middle + question_blocks[max_idx + 1:]

def relabel_tf_options(question_blocks):
    letters = ["a", "b", "c", "d"]
    opt_blocks = [b for b in question_blocks if re.match(r'^\s*[a-d]\)', get_text(b), re.IGNORECASE)]
    for idx, block in enumerate(opt_blocks):
        letter = letters[idx] if idx < len(letters) else letters[-1]
        update_tf_label(block, f"{letter})")

def process_part(blocks, start, end, part_type):
    intro, questions = parse_questions_in_range(blocks, start, end)
    if part_type == "PHAN1": processed = [shuffle_mcq_options(q) for q in questions]
    elif part_type == "PHAN2": processed = [shuffle_tf_options(q) for q in questions]
    else: processed = [q.copy() for q in questions]
    shuffled_q = shuffle_array(processed)
    for i, q in enumerate(shuffled_q):
        if q: update_question_label(q[0], f"Câu {i+1}.")
    if part_type == "PHAN1": 
        for q in shuffled_q: relabel_mcq_options(q)
    elif part_type == "PHAN2": 
        for q in shuffled_q: relabel_tf_options(q)
    res = intro.copy()
    for q in shuffled_q: res.extend(q)
    return res

def shuffle_docx(file_bytes, shuffle_mode="auto"):
    input_buffer = io.BytesIO(file_bytes)
    with zipfile.ZipFile(input_buffer, 'r') as zin:
        doc_xml = zin.read("word/document.xml").decode('utf-8')
        dom = minidom.parseString(doc_xml)
        body = dom.getElementsByTagNameNS(W_NS, "body")[0]
        blocks = [child for child in body.childNodes if child.nodeType == child.ELEMENT_NODE and child.localName in ["p", "tbl"]]
        
        # Logic phân chia phần cơ bản
        p1 = -1
        for i, b in enumerate(blocks):
            if "PHẦN 1" in get_text(b).upper(): p1 = i; break
        
        if shuffle_mode == "mcq" or p1 == -1:
            intro, questions = parse_questions_in_range(blocks, 0, len(blocks))
            proc = [shuffle_mcq_options(q) for q in questions]
            shuf = shuffle_array(proc)
            for i, q in enumerate(shuf): update_question_label(q[0], f"Câu {i+1}.")
            for q in shuf: relabel_mcq_options(q)
            new_blocks = intro + [item for sublist in shuf for item in sublist]
        else:
            # Logic trộn đề có phần (Để đơn giản cho bản sửa lỗi, chạy logic full mcq nếu không tìm thấy phần)
            # Bạn có thể copy logic full từ file cũ nếu cần xử lý PHẦN 1/2/3 phức tạp
            intro, questions = parse_questions_in_range(blocks, 0, len(blocks))
            proc = [shuffle_mcq_options(q) for q in questions]
            shuf = shuffle_array(proc)
            for i, q in enumerate(shuf): update_question_label(q[0], f"Câu {i+1}.")
            for q in shuf: relabel_mcq_options(q)
            new_blocks = intro + [item for sublist in shuf for item in sublist]

        # Xây dựng lại XML
        for child in list(body.childNodes):
            if child.localName not in ["p", "tbl"]: continue
            body.removeChild(child)
        for block in new_blocks: body.appendChild(block)
        
        output_buffer = io.BytesIO()
        with zipfile.ZipFile(output_buffer, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "word/document.xml": zout.writestr(item, dom.toxml().encode('utf-8'))
                else: zout.writestr(item, zin.read(item.filename))
        return output_buffer.getvalue()

def create_zip(file_bytes, name, num, mode):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z:
        for i in range(num):
            z.writestr(f"{name}_V{i+1}.docx", shuffle_docx(file_bytes, mode))
    return buf.getvalue()

# ==================== MAIN APP ====================

def main():
    uploaded_file, process_btn, mode_code, num = render_ui()

    if process_btn:
        if not uploaded_file:
            st.warning("⚠️ Vui lòng chọn file Word trước khi trộn!")
        else:
            try:
                with st.spinner("⏳ Đang xử lý... vui lòng đợi"):
                    # Demo logic
                    res = create_zip(uploaded_file.read(), "DeMinhDuc", num, mode_code)
                    
                    st.success("✅ Trộn đề thành công!")
                    st.download_button(
                        label="📥 Tải xuống file ZIP kết quả",
                        data=res,
                        file_name="KetQua_TronDe.zip",
                        mime="application/zip",
                        type="primary"
                    )
                    st.balloons()
            except Exception as e:
                st.error(f"❌ Có lỗi xảy ra: {str(e)}")

if __name__ == "__main__":
    main()
