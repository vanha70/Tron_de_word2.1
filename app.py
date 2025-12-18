"""
TNMic PRO - PHẦN MỀM TRỘN ĐỀ (RE-BUILT CORE)
------------------------------------------------
Phiên bản này được viết lại để khắc phục triệt để:
1. Lỗi trùng lặp đáp án.
2. Lỗi không đồng bộ màu sắc/in đậm.
3. Cơ chế: Xóa sạch gốc rễ nhãn cũ -> Chèn nhãn mới chuẩn xác.
"""

import streamlit as st
import re
import random
import zipfile
import io
import csv
from xml.dom import minidom

# ==================== 1. GIAO DIỆN & CẤU HÌNH ====================
st.set_page_config(
    page_title="TNMic - Trộn Đề Trắc Nghiệm",
    page_icon="🧬",
    layout="centered",
    initial_sidebar_state="collapsed"
)

CUSTOM_CSS = """
<style>
    /* Reset & Base */
    header {visibility: hidden;}
    .stApp > header {display: none;}
    .block-container {padding-top: 1rem; padding-bottom: 5rem;}
    
    [data-testid="stAppViewContainer"] {
        background-color: #f0fdfa; /* Teal nhạt */
        background-image: radial-gradient(#99f6e4 1px, transparent 1px);
        background-size: 24px 24px;
        font-family: 'Segoe UI', Arial, sans-serif;
    }

    /* HEADER */
    .header-box {
        background: white;
        padding: 2.5rem;
        border-radius: 20px;
        text-align: center;
        box-shadow: 0 10px 30px -10px rgba(13, 148, 136, 0.2);
        border-top: 6px solid #0d9488; /* Teal đậm */
        margin-bottom: 30px;
    }

    .school-name {
        color: #115e59;
        font-family: 'Times New Roman', serif;
        font-weight: 900;
        font-size: 2rem;
        text-transform: uppercase;
        margin-bottom: 10px;
        letter-spacing: 1px;
    }

    .app-title-badge {
        display: inline-block;
        background: linear-gradient(135deg, #0d9488 0%, #115e59 100%);
        color: white;
        padding: 10px 40px;
        border-radius: 50px;
        font-weight: 800;
        font-size: 1.5rem;
        text-transform: uppercase;
        box-shadow: 0 4px 15px rgba(13, 148, 136, 0.3);
        margin: 10px 0;
    }

    .teacher-info {
        margin-top: 15px;
        font-weight: 600;
        color: #0f766e;
        background: #ccfbf1;
        padding: 8px 25px;
        border-radius: 12px;
        display: inline-block;
        border: 1px solid #99f6e4;
    }

    /* MAIN AREA */
    .main-card {
        background: rgba(255, 255, 255, 0.95);
        backdrop-filter: blur(10px);
        padding: 30px;
        border-radius: 20px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);
    }

    /* BUTTON */
    .stButton > button {
        background: linear-gradient(90deg, #0d9488, #0f766e);
        color: white;
        border: none;
        padding: 14px;
        border-radius: 10px;
        font-weight: 700;
        font-size: 1.1rem;
        text-transform: uppercase;
        width: 100%;
        box-shadow: 0 4px 12px rgba(13, 148, 136, 0.3);
        transition: all 0.2s;
    }
    .stButton > button:hover {
        transform: translateY(-2px);
        background: linear-gradient(90deg, #0f766e, #115e59);
    }
    
    .footer { text-align: center; margin-top: 40px; color: #64748b; font-size: 0.85rem; }
</style>
"""

HEADER_HTML = """
<div class="header-box">
    <div class="school-name">TRƯỜNG THPT MINH ĐỨC</div>
    <div class="app-title-badge">PHẦN MỀM TRỘN ĐỀ</div><br>
    <div class="teacher-info">GV: Nguyễn Văn Hà • Zalo: 0913968302</div>
</div>
"""

# ==================== 2. KỸ THUẬT XỬ LÝ WORD (CORE ENGINE) ====================
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

def get_full_text(paragraph_node):
    """Lấy toàn bộ nội dung text của một đoạn văn"""
    texts = []
    for t in paragraph_node.getElementsByTagNameNS(W_NS, "t"):
        if t.firstChild: texts.append(t.firstChild.nodeValue)
    return "".join(texts)

def create_styled_run(doc, text, color_hex="0070C0", is_bold=True):
    """
    Tạo một thẻ Run (<w:r>) mới với đầy đủ định dạng:
    - Font: Times New Roman
    - Color: Xanh Dương (#0070C0)
    - Style: In Đậm (Bold)
    """
    r = doc.createElementNS(W_NS, "w:r")
    rPr = doc.createElementNS(W_NS, "w:rPr")
    
    # 1. Font
    rFonts = doc.createElementNS(W_NS, "w:rFonts")
    rFonts.setAttributeNS(W_NS, "w:ascii", "Times New Roman")
    rFonts.setAttributeNS(W_NS, "w:hAnsi", "Times New Roman")
    rPr.appendChild(rFonts)
    
    # 2. In đậm (Quan trọng)
    if is_bold:
        b = doc.createElementNS(W_NS, "w:b")
        rPr.appendChild(b)
        
    # 3. Màu sắc
    if color_hex:
        color = doc.createElementNS(W_NS, "w:color")
        color.setAttributeNS(W_NS, "w:val", color_hex)
        rPr.appendChild(color)
        
    r.appendChild(rPr)
    
    # 4. Nội dung text (Có bảo toàn khoảng trắng)
    t = doc.createElementNS(W_NS, "w:t")
    t.setAttribute("xml:space", "preserve")
    t.appendChild(doc.createTextNode(text))
    r.appendChild(t)
    
    return r

def strip_label_prefix(paragraph, regex_pattern):
    """
    Hàm 'Ăn mòn' nhãn cũ:
    Tìm nhãn cũ (VD: 'A.') và xóa sạch nó khỏi các thẻ XML đầu tiên.
    """
    full_txt = get_full_text(paragraph)
    match = re.match(regex_pattern, full_txt)
    
    if match:
        chars_to_delete = len(match.group(0)) # Số ký tự cần xóa
        
        # Duyệt qua các thẻ text để xóa dần
        t_nodes = paragraph.getElementsByTagNameNS(W_NS, "t")
        for t in t_nodes:
            if not t.firstChild: continue
            
            val = t.firstChild.nodeValue
            len_val = len(val)
            
            if chars_to_delete > 0:
                if len_val <= chars_to_delete:
                    # Xóa toàn bộ nội dung node này
                    t.firstChild.nodeValue = ""
                    chars_to_delete -= len_val
                else:
                    # Cắt phần đầu
                    t.firstChild.nodeValue = val[chars_to_delete:]
                    chars_to_delete = 0
            
            if chars_to_delete == 0:
                break

def prepend_new_label(paragraph, doc, label_text):
    """Chèn nhãn mới đã được định dạng vào đầu đoạn văn"""
    # Tạo Run mới (Xanh + Đậm)
    new_run = create_styled_run(doc, label_text + " ", "0070C0", True)
    
    # Chèn vào vị trí đầu tiên
    if paragraph.hasChildNodes():
        paragraph.insertBefore(new_run, paragraph.firstChild)
    else:
        paragraph.appendChild(new_run)

def clean_answer_hints(paragraph):
    """Xóa gạch chân và màu đỏ (để ẩn đáp án đúng)"""
    runs = paragraph.getElementsByTagNameNS(W_NS, "r")
    for r in runs:
        rPr = r.getElementsByTagNameNS(W_NS, "rPr")
        if rPr:
            # Xóa các thẻ định dạng lộ đáp án
            for tag in ["u", "color", "b", "i"]: # Xóa cả bold/italic cũ để đồng bộ
                for node in rPr[0].getElementsByTagNameNS(W_NS, tag):
                    rPr[0].removeChild(node)

def is_correct_answer(paragraph):
    """Check xem câu này có phải đáp án đúng không (Gạch chân/Đỏ)"""
    runs = paragraph.getElementsByTagNameNS(W_NS, "r")
    for r in runs:
        rPr = r.getElementsByTagNameNS(W_NS, "rPr")
        if rPr:
            if rPr[0].getElementsByTagNameNS(W_NS, "u"): return True
            color = rPr[0].getElementsByTagNameNS(W_NS, "color")
            if color and color[0].getAttributeNS(W_NS, "val") in ["FF0000", "RED"]: return True
    return False

# ==================== 3. LOGIC TRỘN ĐỀ (PROCESSORS) ====================

def process_mcq(questions, doc):
    """Xử lý Phần 1 (A,B,C,D)"""
    processed_qs = []
    keys = []
    
    for q_block in questions:
        # Regex tìm A., B. ...
        pat = r'^\s*[A-D][\.\)]'
        # Tìm các index của dòng chứa đáp án
        opt_indices = [i for i, n in enumerate(q_block) if re.match(pat, get_full_text(n))]
        
        correct_char = "X"
        
        if len(opt_indices) >= 2:
            # Tách các options ra
            opts = [q_block[i] for i in opt_indices]
            
            # 1. Xác định đáp án đúng & Làm sạch định dạng
            target_opt = None
            for opt in opts:
                if is_correct_answer(opt): target_opt = opt
                clean_answer_hints(opt) # Xóa gạch chân/đỏ
            
            # 2. Trộn
            random.shuffle(opts)
            labels = ["A.", "B.", "C.", "D."]
            
            # 3. Gán lại vào vị trí & Thay nhãn
            for i, idx in enumerate(opt_indices):
                q_block[idx] = opts[i] # Đặt lại vị trí
                
                # Check Key
                if opts[i] == target_opt:
                    correct_char = labels[i][0] # Lấy chữ cái A,B,C...
                
                # THAY NHÃN: Xóa cũ -> Chèn mới
                strip_label_prefix(opts[i], pat)
                prepend_new_label(opts[i], doc, labels[i])
        
        keys.append(correct_char)
        processed_qs.append(q_block)
        
    return processed_qs, keys

def process_tf(questions, doc):
    """Xử lý Phần 2 (Đúng/Sai)"""
    processed_qs = []
    keys = []
    
    for q_block in questions:
        pat = r'^\s*[a-d][\.\)]'
        opt_indices = [i for i, n in enumerate(q_block) if re.match(pat, get_full_text(n))]
        res_str = []
        
        if len(opt_indices) >= 2:
            opts = [q_block[i] for i in opt_indices]
            
            # Map trạng thái
            status_map = {}
            for opt in opts:
                is_true = is_correct_answer(opt)
                status_map[opt] = "Đ" if is_true else "S"
                clean_answer_hints(opt)
            
            random.shuffle(opts)
            labels = ["a)", "b)", "c)", "d)"]
            
            for i, idx in enumerate(opt_indices):
                q_block[idx] = opts[i]
                res_str.append(f"{labels[i][:-1]}{status_map[opts[i]]}")
                
                # THAY NHÃN
                strip_label_prefix(opts[i], pat)
                prepend_new_label(opts[i], doc, labels[i])
                
        keys.append(" - ".join(res_str))
        processed_qs.append(q_block)
        
    return processed_qs, keys

def process_short(questions):
    """Xử lý Phần 3 (Key)"""
    processed_qs = []
    keys = []
    
    for q_block in questions:
        key_val = ""
        full_txt = "".join([get_full_text(n) for n in q_block])
        
        # Tìm Key <key=...>
        m = re.search(r'<\s*key\s*=\s*(.*?)\s*>', full_txt, re.IGNORECASE)
        if m:
            key_val = m.group(1).strip()
            # Xóa thẻ key trong XML
            for p in q_block:
                t_nodes = p.getElementsByTagNameNS(W_NS, "t")
                for t in t_nodes:
                    if t.firstChild and '<' in t.firstChild.nodeValue:
                        val = t.firstChild.nodeValue
                        val = re.sub(r'<\s*key\s*=\s*.*?>', '', val, flags=re.IGNORECASE)
                        t.firstChild.nodeValue = val
                        
        keys.append(key_val)
        processed_qs.append(q_block)
        
    return processed_qs, keys

# ==================== 4. LOGIC TỔNG HỢP ====================

def parse_docx(dom):
    body = dom.getElementsByTagNameNS(W_NS, "body")[0]
    blocks = [n for n in body.childNodes if n.localName in ['p', 'tbl']]
    intro, questions = [], []
    
    i = 0
    # Intro
    while i < len(blocks):
        if re.match(r'^Câu\s*\d+', get_full_text(blocks[i])): break
        intro.append(blocks[i])
        i += 1
    # Questions
    while i < len(blocks):
        if re.match(r'^Câu\s*\d+', get_full_text(blocks[i])):
            grp = [blocks[i]]
            i += 1
            while i < len(blocks):
                txt = get_full_text(blocks[i])
                if re.match(r'^Câu\s*\d+', txt) or "PHẦN" in txt.upper(): break
                grp.append(blocks[i])
                i += 1
            questions.append(grp)
        else: i += 1
    return intro, questions, body

def create_header_p(doc, text, align="left", bold=False):
    """Tạo đoạn văn Header"""
    p = doc.createElementNS(W_NS, "w:p")
    pPr = doc.createElementNS(W_NS, "w:pPr")
    jc = doc.createElementNS(W_NS, "w:jc")
    jc.setAttributeNS(W_NS, "w:val", align)
    pPr.appendChild(jc)
    p.appendChild(pPr)
    
    # Run
    r = create_styled_run(doc, text, "000000", bold)
    p.appendChild(r)
    return p

def generate_mix(file_bytes, num_copies):
    outer_zip = io.BytesIO()
    csv_data = []
    
    with zipfile.ZipFile(outer_zip, 'w', zipfile.ZIP_DEFLATED) as z_out:
        in_io = io.BytesIO(file_bytes)
        with zipfile.ZipFile(in_io, 'r') as z_in:
            xml_content = z_in.read("word/document.xml")
            
            for i in range(num_copies):
                exam_code = str(random.randint(1001, 9999))
                dom = minidom.parseString(xml_content)
                intro, all_qs, body = parse_docx(dom)
                
                # Chia 3 Phần
                p1 = all_qs[0:18]
                p2 = all_qs[18:22]
                p3 = all_qs[22:]
                
                row_key = [exam_code]
                
                # Xử lý từng phần
                p1_fin, k1 = process_mcq(p1, dom)
                c1 = list(zip(p1_fin, k1))
                random.shuffle(c1)
                p1_fin, k1 = zip(*c1) if c1 else ([],[])
                row_key.extend(k1)
                
                p2_fin, k2 = process_tf(p2, dom)
                c2 = list(zip(p2_fin, k2))
                random.shuffle(c2)
                p2_fin, k2 = zip(*c2) if c2 else ([],[])
                row_key.extend(k2)
                
                p3_fin, k3 = process_short(p3)
                c3 = list(zip(p3_fin, k3))
                random.shuffle(c3)
                p3_fin, k3 = zip(*c3) if c3 else ([],[])
                row_key.extend(k3)
                
                csv_data.append(row_key)
                
                # --- REBUILD WORD BODY ---
                for n in list(body.childNodes):
                    if n.localName in ['p', 'tbl']: body.removeChild(n)
                
                # Header Trường
                body.appendChild(create_header_p(dom, "TRƯỜNG THPT MINH ĐỨC", "center", True))
                body.appendChild(create_header_p(dom, "ĐỀ KIỂM TRA ĐỊNH KỲ", "center", True))
                body.appendChild(create_header_p(dom, f"MÃ ĐỀ: {exam_code}", "right", True))
                body.appendChild(create_header_p(dom, "Họ tên:.......................................................... Lớp:..........", "left"))
                body.appendChild(create_header_p(dom, "", "left"))
                
                # P1
                body.appendChild(create_header_p(dom, "PHẦN I. Trắc nghiệm (18 câu)", "left", True))
                for idx, q in enumerate(p1_fin):
                    # Thay đổi số câu (Câu 1, Câu 2...) và tô xanh
                    strip_label_prefix(q[0], r'^Câu\s*\d+[\.\:]')
                    prepend_new_label(q[0], dom, f"Câu {idx+1}.")
                    for n in q: body.appendChild(n)
                    
                # P2
                body.appendChild(create_header_p(dom, "PHẦN II. Đúng Sai (4 câu)", "left", True))
                for idx, q in enumerate(p2_fin):
                    strip_label_prefix(q[0], r'^Câu\s*\d+[\.\:]')
                    prepend_new_label(q[0], dom, f"Câu {idx+1}.")
                    for n in q: body.appendChild(n)
                    
                # P3
                body.appendChild(create_header_p(dom, "PHẦN III. Trả lời ngắn (6 câu)", "left", True))
                for idx, q in enumerate(p3_fin):
                    strip_label_prefix(q[0], r'^Câu\s*\d+[\.\:]')
                    prepend_new_label(q[0], dom, f"Câu {idx+1}.")
                    for n in q: body.appendChild(n)
                
                # Save Docx
                new_xml = dom.toxml().encode('utf-8')
                docx_io = io.BytesIO()
                with zipfile.ZipFile(docx_io, 'w', zipfile.ZIP_DEFLATED) as z:
                    for it in z_in.infolist():
                        if it.filename == "word/document.xml":
                            z.writestr(it.filename, new_xml)
                        else:
                            z.writestr(it.filename, z_in.read(it.filename))
                z_out.writestr(f"De_Thi/De_{exam_code}.docx", docx_io.getvalue())
        
        # Save CSV
        csv_io = io.StringIO()
        w = csv.writer(csv_io)
        w.writerow(["Mã đề"] + [str(i) for i in range(1,19)] + [f"II_{i}" for i in range(1,5)] + [f"III_{i}" for i in range(1,7)])
        w.writerows(csv_data)
        z_out.writestr("Dap_An.csv", csv_io.getvalue().encode('utf-8-sig'))

    return outer_zip.getvalue()

# ==================== 5. UI LOGIC ====================
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)
st.markdown(HEADER_HTML, unsafe_allow_html=True)

st.markdown('<div class="main-card">', unsafe_allow_html=True)
uploaded_file = st.file_uploader("Kéo thả file .docx vào đây", type=['docx'])

if uploaded_file:
    st.success(f"✅ Đã nhận: {uploaded_file.name}")
    st.info("Quy tắc: P1, P2 gạch chân đáp án đúng. P3 dùng thẻ <key=...>")

st.write("")
num = st.number_input("Số lượng đề cần tạo", 1, 50, 4)

if st.button("🚀 BẮT ĐẦU TRỘN"):
    if not uploaded_file:
        st.warning("Vui lòng chọn file!")
    else:
        try:
            with st.spinner("Đang xử lý (Cơ chế làm sạch sâu)..."):
                final_zip = generate_mix(uploaded_file.read(), num)
                st.success("✅ Thành công! Tải xuống bên dưới.")
                st.download_button(
                    label="📥 Tải về (ZIP)",
                    data=final_zip,
                    file_name="KetQua_TNMic_Final.zip",
                    mime="application/zip"
                )
        except Exception as e:
            st.error(f"Lỗi: {e}")

st.markdown('</div>', unsafe_allow_html=True)
st.markdown('<div class="footer">© 2024 TNMic • Ultimate Edition</div>', unsafe_allow_html=True)
