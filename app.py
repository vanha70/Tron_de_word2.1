"""
TNMic PRO - PHẦN MỀM TRỘN ĐỀ (FINAL V4)
Tính năng:
1. Fix lỗi mất phương án (Do Tab/Newline).
2. Fix lỗi mất câu hỏi (Do gộp dòng).
3. Giao diện: Xanh Ngọc (Teal).
4. Logic: Tách hạt nội dung -> Xây dựng lại (Re-build).
"""

import streamlit as st
import re
import random
import zipfile
import io
import csv
from xml.dom import minidom

# ==================== 1. CẤU HÌNH & GIAO DIỆN ====================
st.set_page_config(
    page_title="TNMic - Trộn Đề Trắc Nghiệm",
    page_icon="💎",
    layout="centered",
    initial_sidebar_state="collapsed"
)

CUSTOM_CSS = """
<style>
    header {visibility: hidden;}
    .stApp > header {display: none;}
    .block-container {padding-top: 1rem; padding-bottom: 5rem;}
    
    [data-testid="stAppViewContainer"] {
        background-color: #f0fdfa; 
        background-image: radial-gradient(#99f6e4 1px, transparent 1px);
        background-size: 24px 24px;
        font-family: 'Segoe UI', sans-serif;
    }

    .header-box {
        background: white;
        padding: 2.5rem;
        border-radius: 20px;
        text-align: center;
        box-shadow: 0 10px 30px -10px rgba(13, 148, 136, 0.2);
        border-top: 6px solid #0d9488;
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

    .app-badge {
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

    .main-card {
        background: rgba(255, 255, 255, 0.95);
        backdrop-filter: blur(10px);
        padding: 30px;
        border-radius: 20px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);
    }

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
    <div class="app-badge">PHẦN MỀM TRỘN ĐỀ</div><br>
    <div class="teacher-info">GV: Nguyễn Văn Hà • Zalo: 0913968302</div>
</div>
"""

# ==================== 2. CORE ENGINE (XML PROCESSING) ====================
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

def get_full_text_safe(node):
    """
    Lấy text an toàn:
    - Thay thế <w:tab/> bằng khoảng trắng.
    - Thay thế <w:br/> bằng khoảng trắng/xuống dòng.
    - Lấy nội dung <w:t>.
    Điều này giúp tách các đáp án dính liền nhau.
    """
    text_parts = []
    # Duyệt đệ quy hoặc duyệt childNodes phẳng (Paragraph thường phẳng)
    for child in node.childNodes:
        if child.localName == "r": # Run
            for r_child in child.childNodes:
                if r_child.localName == "t":
                    if r_child.firstChild: text_parts.append(r_child.firstChild.nodeValue)
                elif r_child.localName == "tab":
                    text_parts.append(" ") # Tab -> Space
                elif r_child.localName == "br":
                    text_parts.append(" ") # Break -> Space
    return "".join(text_parts)

def clear_node_content(node):
    """Xóa sạch nội dung bên trong một paragraph"""
    while node.hasChildNodes():
        node.removeChild(node.firstChild)

def create_styled_run(doc, text, color_hex="0070C0", is_bold=True):
    """Tạo Run mới: Font Times, Màu Xanh, In Đậm"""
    r = doc.createElementNS(W_NS, "w:r")
    rPr = doc.createElementNS(W_NS, "w:rPr")
    
    # Font
    rFonts = doc.createElementNS(W_NS, "w:rFonts")
    rFonts.setAttributeNS(W_NS, "w:ascii", "Times New Roman")
    rFonts.setAttributeNS(W_NS, "w:hAnsi", "Times New Roman")
    rPr.appendChild(rFonts)
    
    # Bold
    if is_bold:
        rPr.appendChild(doc.createElementNS(W_NS, "w:b"))
        
    # Color
    if color_hex:
        color = doc.createElementNS(W_NS, "w:color")
        color.setAttributeNS(W_NS, "w:val", color_hex)
        rPr.appendChild(color)
        
    r.appendChild(rPr)
    
    # Text
    t = doc.createElementNS(W_NS, "w:t")
    t.setAttribute("xml:space", "preserve")
    t.appendChild(doc.createTextNode(text))
    r.appendChild(t)
    return r

def check_correct_in_text(paragraph_node):
    """Check đúng sai dựa trên định dạng của paragraph gốc (Gạch chân/Đỏ)"""
    runs = paragraph_node.getElementsByTagNameNS(W_NS, "r")
    for r in runs:
        rPr = r.getElementsByTagNameNS(W_NS, "rPr")
        if rPr:
            if rPr[0].getElementsByTagNameNS(W_NS, "u"): return True
            c = rPr[0].getElementsByTagNameNS(W_NS, "color")
            if c and c[0].getAttributeNS(W_NS, "val") in ["FF0000", "RED"]: return True
    return False

# ==================== 3. XỬ LÝ NỘI DUNG (PARSING & CLEANING) ====================

def extract_and_parse_options(q_block, mode="mcq"):
    """
    Hút toàn bộ nội dung, tách câu hỏi và đáp án.
    Trả về: (Danh sách paragraph chứa đáp án, Danh sách option sạch, Text câu hỏi thừa nếu có)
    """
    pat = r'([A-D][\.\)])' if mode == "mcq" else r'([a-d][\.\)])'
    
    opt_paragraphs = []
    full_text_buffer = ""
    
    # 1. Gom text từ các dòng có khả năng chứa đáp án
    for p in q_block:
        txt = get_full_text_safe(p)
        # Nếu dòng chứa pattern đáp án (A., B....)
        if re.search(pat, txt):
            opt_paragraphs.append(p)
            full_text_buffer += " " + txt 
    
    if not opt_paragraphs: return [], [], ""

    # 2. Tách bằng Regex
    # Split giữ lại delimiter (A., B...)
    # Thêm lookahead để xử lý dính chữ (vd: A.ĐúngB.Sai) -> Regex phải khéo
    # Dùng split đơn giản, sau đó clean
    parts = re.split(pat, full_text_buffer)
    
    # parts[0] là text TRƯỚC đáp án A (thường là phần đuôi của câu hỏi bị rớt xuống)
    pre_text = parts[0].strip()
    
    clean_options = []
    
    # Duyệt từ 1, bước nhảy 2 (Label, Content)
    # parts: [pre, 'A.', 'Nội dung', 'B.', 'Nội dung'...]
    for i in range(1, len(parts), 2):
        if i+1 >= len(parts): break
        label = parts[i].strip()
        content = parts[i+1].strip()
        
        # Check đúng sai: Quét các paragraph gốc xem có cái nào gạch chân chứa content này ko
        # (Logic tương đối, nhưng an toàn hơn việc parse run từng tí)
        is_correct = False
        for p in opt_paragraphs:
            # Nếu nội dung này nằm trong 1 paragraph có gạch chân -> Đúng
            if content in get_full_text_safe(p) and check_correct_in_text(p):
                is_correct = True
                break
        
        clean_options.append({"text": content, "correct": is_correct})
        
    return opt_paragraphs, clean_options, pre_text

# ==================== 4. LOGIC TÁI TẠO (REBUILDER) ====================

def rebuild_paragraph_options(doc, paragraphs, options, labels, pre_text=""):
    """
    Xóa sạch paragraph cũ, viết lại paragraph mới.
    """
    if not paragraphs: return

    # Neo vào paragraph đầu tiên để chèn
    target_p = paragraphs[0]
    parent = target_p.parentNode
    
    # Xóa tất cả paragraph cũ
    for p in paragraphs:
        # Nếu là paragraph đầu tiên thì giữ lại (clear content) để làm mốc chèn, sau đó xóa sau
        # Cách tốt nhất: Chèn cái mới trước cái đầu tiên, rồi xóa hết cái cũ.
        pass
        
    # Tạo các paragraph mới
    new_paragraphs = []
    
    # 1. Nếu có phần dư của câu hỏi (Pre-text), tạo 1 dòng riêng cho nó
    if pre_text:
        p_pre = doc.createElementNS(W_NS, "w:p")
        # Copy style từ target_p
        if target_p.getElementsByTagNameNS(W_NS, "pPr"):
            p_pre.appendChild(target_p.getElementsByTagNameNS(W_NS, "pPr")[0].cloneNode(True))
        p_pre.appendChild(create_styled_run(doc, pre_text, "000000", False))
        new_paragraphs.append(p_pre)
        
    # 2. Tạo các dòng đáp án (Mỗi đáp án 1 dòng cho đẹp và an toàn)
    for i, opt in enumerate(options):
        new_p = doc.createElementNS(W_NS, "w:p")
        # Copy style
        if target_p.getElementsByTagNameNS(W_NS, "pPr"):
            new_p.appendChild(target_p.getElementsByTagNameNS(W_NS, "pPr")[0].cloneNode(True))
            
        # Nhãn (Xanh + Đậm)
        new_p.appendChild(create_styled_run(doc, labels[i] + " ", "0070C0", True))
        
        # Nội dung (Đen + Thường)
        new_p.appendChild(create_styled_run(doc, opt['text'], "000000", False))
        
        new_paragraphs.append(new_p)
        
    # 3. Thực hiện chèn và xóa
    # Chèn tất cả cái mới trước cái cũ đầu tiên
    for np in new_paragraphs:
        parent.insertBefore(np, target_p)
        
    # Xóa tất cả cái cũ
    for p in paragraphs:
        if p.parentNode == parent: # Check tồn tại
            parent.removeChild(p)

def process_mcq_rebuild(questions, doc):
    processed_qs = []
    keys = []
    labels = ["A.", "B.", "C.", "D."]
    
    for q_block in questions:
        # 1. Parse
        paragraphs, options_data, pre_text = extract_and_parse_options(q_block, "mcq")
        
        # Nếu không đủ đáp án, giữ nguyên
        if len(options_data) < 2:
            processed_qs.append(q_block)
            keys.append("X")
            continue
            
        # 2. Trộn
        random.shuffle(options_data)
        
        # 3. Lấy Key
        correct_char = "X"
        for i, opt in enumerate(options_data):
            if opt['correct']: correct_char = labels[i][0]
        keys.append(correct_char)
        
        # 4. Rebuild
        rebuild_paragraph_options(doc, paragraphs, options_data, labels, pre_text)
        processed_qs.append(q_block)
        
    return processed_qs, keys

def process_tf_rebuild(questions, doc):
    processed_qs = []
    keys = []
    labels = ["a)", "b)", "c)", "d)"]
    
    for q_block in questions:
        paragraphs, options_data, pre_text = extract_and_parse_options(q_block, "tf")
        
        if len(options_data) < 2:
            processed_qs.append(q_block)
            keys.append("")
            continue
            
        random.shuffle(options_data)
        
        res_str = []
        for i, opt in enumerate(options_data):
            status = "Đ" if opt['correct'] else "S"
            res_str.append(f"{labels[i][:-1]}{status}")
        keys.append(" - ".join(res_str))
        
        rebuild_paragraph_options(doc, paragraphs, options_data, labels, pre_text)
        processed_qs.append(q_block)
        
    return processed_qs, keys

def process_short_clean(questions):
    """Làm sạch key trong P3"""
    keys = []
    for q_block in questions:
        full_text = "".join([get_full_text_safe(p) for p in q_block])
        # Regex tìm key linh hoạt
        m = re.search(r'<\s*key\s*=\s*(.*?)\s*>', full_text, re.IGNORECASE)
        k = m.group(1).strip() if m else ""
        keys.append(k)
        
        # Xóa thẻ key trong text
        for p in q_block:
            txt = get_full_text_safe(p)
            if '<' in txt and 'key' in txt:
                clean = re.sub(r'<\s*key\s*=\s*.*?>', '', txt, flags=re.IGNORECASE)
                clear_node_content(p)
                p.appendChild(create_styled_run(p.ownerDocument, clean, "000000", False))
    
    return questions, keys

# ==================== 5. MAIN LOGIC ====================

def create_header_p(doc, text, align="left", bold=False):
    p = doc.createElementNS(W_NS, "w:p")
    pPr = doc.createElementNS(W_NS, "w:pPr")
    jc = doc.createElementNS(W_NS, "w:jc")
    jc.setAttributeNS(W_NS, "w:val", align)
    pPr.appendChild(jc)
    p.appendChild(pPr)
    p.appendChild(create_styled_run(doc, text, "000000", bold))
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
                body = dom.getElementsByTagNameNS(W_NS, "body")[0]
                
                # Parse Blocks
                blocks = [n for n in body.childNodes if n.localName in ['p', 'tbl']]
                intro, questions = [], []
                idx = 0
                
                # Skip intro
                while idx < len(blocks):
                    if re.match(r'^Câu\s*\d+', get_full_text_safe(blocks[idx]).strip()): break
                    intro.append(blocks[idx])
                    idx += 1
                
                # Collect Questions
                while idx < len(blocks):
                    txt = get_full_text_safe(blocks[idx]).strip()
                    if re.match(r'^Câu\s*\d+', txt):
                        grp = [blocks[idx]]
                        idx += 1
                        while idx < len(blocks):
                            sub = get_full_text_safe(blocks[idx]).strip()
                            if re.match(r'^Câu\s*\d+', sub) or "PHẦN" in sub.upper(): break
                            grp.append(blocks[idx])
                            idx += 1
                        questions.append(grp)
                    else: idx += 1
                
                # Slice
                p1 = questions[0:18]
                p2 = questions[18:22]
                p3 = questions[22:]
                
                row_key = [exam_code]
                
                # Process
                p1_fin, k1 = process_mcq_rebuild(p1, dom)
                c1 = list(zip(p1_fin, k1))
                random.shuffle(c1)
                p1_fin, k1 = zip(*c1) if c1 else ([],[])
                row_key.extend(k1)
                
                p2_fin, k2 = process_tf_rebuild(p2, dom)
                c2 = list(zip(p2_fin, k2))
                random.shuffle(c2)
                p2_fin, k2 = zip(*c2) if c2 else ([],[])
                row_key.extend(k2)
                
                p3_fin, k3 = process_short_clean(p3)
                c3 = list(zip(p3_fin, k3))
                random.shuffle(c3)
                p3_fin, k3 = zip(*c3) if c3 else ([],[])
                row_key.extend(k3)
                
                csv_data.append(row_key)
                
                # Rebuild Document Body
                for n in list(body.childNodes):
                    if n.localName in ['p', 'tbl']: body.removeChild(n)
                
                # Add Header
                body.appendChild(create_header_p(dom, "TRƯỜNG THPT MINH ĐỨC", "center", True))
                body.appendChild(create_header_p(dom, "ĐỀ KIỂM TRA ĐỊNH KỲ", "center", True))
                body.appendChild(create_header_p(dom, f"MÃ ĐỀ: {exam_code}", "right", True))
                body.appendChild(create_header_p(dom, "Họ tên thí sinh:............................................ Lớp:..........", "left"))
                body.appendChild(create_header_p(dom, "", "left"))
                
                # Add P1
                body.appendChild(create_header_p(dom, "PHẦN I. Trắc nghiệm (18 câu)", "left", True))
                for ix, q in enumerate(p1_fin):
                    # Renumber & Colorize Question Label
                    p_q = q[0] # Paragraph chứa "Câu X"
                    txt = get_full_text_safe(p_q)
                    
                    # Tách "Câu X." và nội dung còn lại
                    # Xử lý trường hợp "Câu 1. Nội dung"
                    new_num = f"Câu {ix+1}."
                    # Remove old num (Câu \d+.)
                    clean_txt = re.sub(r'^Câu\s*\d+[\.\:]\s*', '', txt)
                    
                    clear_node_content(p_q)
                    p_q.appendChild(create_styled_run(dom, new_num + " ", "0070C0", True))
                    p_q.appendChild(create_styled_run(dom, clean_txt, "000000", False))
                    
                    for n in q: body.appendChild(n)
                    
                # Add P2
                body.appendChild(create_header_p(dom, "PHẦN II. Đúng Sai (4 câu)", "left", True))
                for ix, q in enumerate(p2_fin):
                    p_q = q[0]
                    txt = get_full_text_safe(p_q)
                    clean_txt = re.sub(r'^Câu\s*\d+[\.\:]\s*', '', txt)
                    clear_node_content(p_q)
                    p_q.appendChild(create_styled_run(dom, f"Câu {ix+1}. ", "0070C0", True))
                    p_q.appendChild(create_styled_run(dom, clean_txt, "000000", False))
                    for n in q: body.appendChild(n)
                    
                # Add P3
                body.appendChild(create_header_p(dom, "PHẦN III. Trả lời ngắn (6 câu)", "left", True))
                for ix, q in enumerate(p3_fin):
                    p_q = q[0]
                    txt = get_full_text_safe(p_q)
                    clean_txt = re.sub(r'^Câu\s*\d+[\.\:]\s*', '', txt)
                    clear_node_content(p_q)
                    p_q.appendChild(create_styled_run(dom, f"Câu {ix+1}. ", "0070C0", True))
                    p_q.appendChild(create_styled_run(dom, clean_txt, "000000", False))
                    for n in q: body.appendChild(n)
                
                # Write Doc
                docx_io = io.BytesIO()
                with zipfile.ZipFile(docx_io, 'w', zipfile.ZIP_DEFLATED) as z:
                    for it in z_in.infolist():
                        if it.filename == "word/document.xml":
                            z.writestr(it.filename, dom.toxml().encode('utf-8'))
                        else:
                            z.writestr(it.filename, z_in.read(it.filename))
                z_out.writestr(f"De_Thi/De_{exam_code}.docx", docx_io.getvalue())
        
        # Write CSV
        csv_io = io.StringIO()
        w = csv.writer(csv_io)
        w.writerow(["Mã đề"] + [str(i) for i in range(1,19)] + [f"II_{i}" for i in range(1,5)] + [f"III_{i}" for i in range(1,7)])
        w.writerows(csv_data)
        z_out.writestr("Dap_An.csv", csv_io.getvalue().encode('utf-8-sig'))

    return outer_zip.getvalue()

# ==================== 6. UI ====================
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)
st.markdown(HEADER_HTML, unsafe_allow_html=True)

st.markdown('<div class="main-card">', unsafe_allow_html=True)
uploaded_file = st.file_uploader("Kéo thả file .docx vào đây", type=['docx'])

if uploaded_file:
    st.success(f"✅ Đã nhận: {uploaded_file.name}")
    st.info("💡 Lưu ý: Hệ thống sẽ tự động tách dòng và sửa lỗi hiển thị.")

st.write("")
num = st.number_input("Số lượng đề cần tạo", 1, 50, 4)

if st.button("🚀 BẮT ĐẦU TRỘN"):
    if not uploaded_file:
        st.warning("Vui lòng chọn file!")
    else:
        try:
            with st.spinner("Đang xử lý (Re-Engineered Core)..."):
                final_zip = generate_mix(uploaded_file.read(), num)
                st.success("✅ Thành công! Tải xuống bên dưới.")
                st.download_button(
                    "📥 Tải về (ZIP)",
                    final_zip,
                    "KetQua_TNMic_V4.zip",
                    "application/zip"
                )
        except Exception as e:
            st.error(f"Lỗi: {e}")

st.markdown('</div>', unsafe_allow_html=True)
st.markdown('<div class="footer">© 2024 TNMic • Ultimate Edition</div>', unsafe_allow_html=True)
