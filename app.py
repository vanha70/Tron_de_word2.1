"""
TNMic PRO - PHẦN MỀM TRỘN ĐỀ (RE-ENGINEERED)
------------------------------------------------
Cơ chế: Tách nội dung -> Lọc sạch nhãn cũ -> Xây lại đáp án mới.
Khắc phục: Lỗi trùng lặp, lỗi sót định dạng, lỗi layout.
"""

import streamlit as st
import re
import random
import zipfile
import io
import csv
from xml.dom import minidom

# ==================== 1. CẤU HÌNH & GIAO DIỆN (TEAL THEME) ====================
st.set_page_config(
    page_title="TNMic - Trộn Đề Thông Minh",
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

def get_text(node):
    return "".join([t.firstChild.nodeValue for t in node.getElementsByTagNameNS(W_NS, "t") if t.firstChild])

def clear_node_content(node):
    """Xóa sạch nội dung bên trong một node"""
    while node.hasChildNodes():
        node.removeChild(node.firstChild)

def create_styled_run(doc, text, color_hex="0070C0", is_bold=True):
    """Tạo Run mới với định dạng chuẩn (Xanh + Đậm + Times)"""
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

def check_correct_in_text(full_text, paragraph_node):
    """
    Check đúng sai dựa trên định dạng của paragraph gốc.
    Nếu paragraph gốc có gạch chân/đỏ -> True.
    """
    runs = paragraph_node.getElementsByTagNameNS(W_NS, "r")
    for r in runs:
        rPr = r.getElementsByTagNameNS(W_NS, "rPr")
        if rPr:
            if rPr[0].getElementsByTagNameNS(W_NS, "u"): return True
            c = rPr[0].getElementsByTagNameNS(W_NS, "color")
            if c and c[0].getAttributeNS(W_NS, "val") in ["FF0000", "RED"]: return True
    return False

# ==================== 3. XỬ LÝ NỘI DUNG (CONTENT PARSING) ====================

def extract_and_clean_options(q_block, mode="mcq"):
    """
    Hút toàn bộ nội dung đáp án, lọc bỏ nhãn cũ (A., B...).
    Trả về danh sách các đáp án (Text) và trạng thái đúng sai.
    """
    pat = r'(?:^|\s)([A-D][\.\)])' if mode == "mcq" else r'(?:^|\s)([a-d][\.\)])'
    
    # 1. Tìm các paragraph chứa đáp án
    opt_paragraphs = []
    full_text_buffer = ""
    
    for p in q_block:
        txt = get_text(p)
        # Nếu dòng chứa pattern đáp án
        if re.search(pat, txt):
            opt_paragraphs.append(p)
            full_text_buffer += " " + txt # Gộp text lại để xử lý trường hợp 1 dòng nhiều đáp án
    
    if not opt_paragraphs: return [], [], []

    # 2. Tách các đáp án từ text gộp
    # Split bằng regex, giữ lại delimiter (A., B...)
    parts = re.split(pat, full_text_buffer)
    # parts sẽ có dạng: ['', 'A.', 'Nội dung A', 'B.', 'Nội dung B'...]
    
    clean_options = []
    
    # Duyệt qua các phần đã tách
    for i in range(1, len(parts), 2):
        label = parts[i].strip()
        content = parts[i+1].strip()
        
        # Xác định đúng sai (Tương đối: Check xem paragraph chứa nội dung này có gạch chân ko)
        # Để chính xác tuyệt đối cần map lại vị trí text với paragraph.
        # Ở đây ta dùng cách đơn giản: Nếu paragraph gốc có gạch chân, ta đánh dấu.
        # Tuy nhiên, khi gộp dòng, việc này khó.
        # GIẢI PHÁP AN TOÀN: Ta check xem trong content có ký tự nào bị gạch chân không (nếu ta parse sâu).
        # Tạm thời: Ta dùng check_correct_in_text trên các paragraph gốc.
        
        is_correct = False
        # Quét lại các paragraph để xem paragraph nào chứa nội dung này và có gạch chân
        for p in opt_paragraphs:
            if content in get_text(p) and check_correct_in_text(get_text(p), p):
                is_correct = True
                break
        
        clean_options.append({"text": content, "correct": is_correct})
        
    return opt_paragraphs, clean_options

# ==================== 4. LOGIC TRỘN VÀ TÁI TẠO (REBUILD) ====================

def rebuild_paragraph_options(doc, paragraphs, options, labels):
    """
    Xây dựng lại các đoạn văn chứa đáp án.
    paragraphs: Các node paragraph cũ (để ghi đè).
    options: Danh sách nội dung đáp án đã trộn.
    labels: Nhãn mới (A., B...)
    """
    # 1. Xóa sạch nội dung các paragraph cũ
    for p in paragraphs:
        clear_node_content(p)
        
    # 2. Tính toán cách phân bố (Layout)
    # Nếu số lượng paragraph gốc == số options -> Mỗi option 1 dòng (Chuẩn nhất)
    # Nếu ít hơn -> Gộp dòng.
    # Ưu tiên: Xuất ra mỗi option 1 dòng để đảm bảo đẹp và không lỗi.
    
    # Sử dụng paragraph đầu tiên làm mẫu, các paragraph thừa có thể xóa hoặc để trống
    target_p = paragraphs[0]
    parent = target_p.parentNode
    
    # Xóa các paragraph thừa (nếu có), chỉ giữ 1 cái làm neo, sau đó insert thêm
    for p in paragraphs[1:]:
        parent.removeChild(p)
        
    # Tạo các paragraph mới cho từng option
    for i, opt in enumerate(options):
        # Tạo P mới
        new_p = doc.createElementNS(W_NS, "w:p")
        # Copy thuộc tính pPr của dòng đầu (để giữ lề, font...)
        if target_p.getElementsByTagNameNS(W_NS, "pPr"):
            new_p.appendChild(target_p.getElementsByTagNameNS(W_NS, "pPr")[0].cloneNode(True))
            
        # 1. Chèn Nhãn (Xanh + Đậm)
        new_p.appendChild(create_styled_run(doc, labels[i] + " ", "0070C0", True))
        
        # 2. Chèn Nội dung (Đen + Thường)
        new_p.appendChild(create_styled_run(doc, opt['text'], "000000", False))
        
        # Chèn vào trước target_p (hoặc vị trí cũ)
        parent.insertBefore(new_p, target_p)
        
    # Xóa cái neo cuối cùng
    parent.removeChild(target_p)

def process_mcq_rebuild(questions, doc):
    processed_qs = []
    keys = []
    labels = ["A.", "B.", "C.", "D."]
    
    for q_block in questions:
        # 1. Hút & Làm sạch
        paragraphs, options_data = extract_and_clean_options(q_block, "mcq")[:2]
        
        # Nếu không tìm thấy đáp án hoặc không đủ 2 đáp án -> Bỏ qua
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
        
        # 4. Tái tạo (Rebuild)
        # Chỉ gửi các paragraph chứa đáp án để thay thế
        rebuild_paragraph_options(doc, paragraphs, options_data, labels)
        
        # Cập nhật q_block (thực ra XML đã đổi, list q_block chỉ để tham chiếu)
        processed_qs.append(q_block)
        
    return processed_qs, keys

def process_tf_rebuild(questions, doc):
    processed_qs = []
    keys = []
    labels = ["a)", "b)", "c)", "d)"]
    
    for q_block in questions:
        paragraphs, options_data = extract_and_clean_options(q_block, "tf")[:2]
        
        if len(options_data) < 2:
            processed_qs.append(q_block)
            keys.append("")
            continue
            
        random.shuffle(options_data)
        
        # Key string
        res_str = []
        for i, opt in enumerate(options_data):
            status = "Đ" if opt['correct'] else "S"
            res_str.append(f"{labels[i][:-1]}{status}")
            
        keys.append(" - ".join(res_str))
        
        rebuild_paragraph_options(doc, paragraphs, options_data, labels)
        processed_qs.append(q_block)
        
    return processed_qs, keys

def process_short_clean(questions):
    keys = []
    for q_block in questions:
        full_text = "".join([get_text(p) for p in q_block])
        m = re.search(r'<\s*key\s*=\s*(.*?)\s*>', full_text, re.IGNORECASE)
        k = m.group(1).strip() if m else ""
        keys.append(k)
        
        # Xóa thẻ key
        for p in q_block:
            txt = get_text(p)
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
                
                # Parse
                blocks = [n for n in body.childNodes if n.localName in ['p', 'tbl']]
                questions = []
                idx = 0
                while idx < len(blocks):
                    if re.match(r'^Câu\s*\d+', get_text(blocks[idx])):
                        grp = [blocks[idx]]
                        idx += 1
                        while idx < len(blocks):
                            if re.match(r'^Câu\s*\d+', get_text(blocks[idx])) or "PHẦN" in get_text(blocks[idx]).upper(): break
                            grp.append(blocks[idx])
                            idx += 1
                        questions.append(grp)
                    else: idx += 1
                
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
                
                # REBUILD BODY
                for n in list(body.childNodes):
                    if n.localName in ['p', 'tbl']: body.removeChild(n)
                
                # Header
                body.appendChild(create_header_p(dom, "TRƯỜNG THPT MINH ĐỨC", "center", True))
                body.appendChild(create_header_p(dom, "ĐỀ KIỂM TRA ĐỊNH KỲ", "center", True))
                body.appendChild(create_header_p(dom, f"MÃ ĐỀ: {exam_code}", "right", True))
                body.appendChild(create_header_p(dom, "Họ tên thí sinh:............................................ Lớp:..........", "left"))
                body.appendChild(create_header_p(dom, "", "left"))
                
                # P1
                body.appendChild(create_header_p(dom, "PHẦN I. Trắc nghiệm (18 câu)", "left", True))
                for ix, q in enumerate(p1_fin):
                    # Đánh lại số câu (Clean & Rebuild Label)
                    txt = get_text(q[0])
                    clean_q = re.sub(r'^Câu\s*\d+[\.\:]\s*', '', txt) # Bỏ "Câu X." cũ
                    clear_node_content(q[0])
                    # Chèn "Câu X." (Xanh)
                    q[0].appendChild(create_styled_run(dom, f"Câu {ix+1}. ", "0070C0", True))
                    # Chèn nội dung (Đen)
                    q[0].appendChild(create_styled_run(dom, clean_q, "000000", False))
                    for n in q: body.appendChild(n)
                    
                # P2
                body.appendChild(create_header_p(dom, "PHẦN II. Đúng Sai (4 câu)", "left", True))
                for ix, q in enumerate(p2_fin):
                    txt = get_text(q[0])
                    clean_q = re.sub(r'^Câu\s*\d+[\.\:]\s*', '', txt)
                    clear_node_content(q[0])
                    q[0].appendChild(create_styled_run(dom, f"Câu {ix+1}. ", "0070C0", True))
                    q[0].appendChild(create_styled_run(dom, clean_q, "000000", False))
                    for n in q: body.appendChild(n)
                    
                # P3
                body.appendChild(create_header_p(dom, "PHẦN III. Trả lời ngắn (6 câu)", "left", True))
                for ix, q in enumerate(p3_fin):
                    txt = get_text(q[0])
                    clean_q = re.sub(r'^Câu\s*\d+[\.\:]\s*', '', txt)
                    clear_node_content(q[0])
                    q[0].appendChild(create_styled_run(dom, f"Câu {ix+1}. ", "0070C0", True))
                    q[0].appendChild(create_styled_run(dom, clean_q, "000000", False))
                    for n in q: body.appendChild(n)
                
                # Write
                docx_io = io.BytesIO()
                with zipfile.ZipFile(docx_io, 'w', zipfile.ZIP_DEFLATED) as z:
                    for it in z_in.infolist():
                        if it.filename == "word/document.xml":
                            z.writestr(it.filename, dom.toxml().encode('utf-8'))
                        else:
                            z.writestr(it.filename, z_in.read(it.filename))
                z_out.writestr(f"De_Thi/De_{exam_code}.docx", docx_io.getvalue())
        
        # CSV
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
    st.info("💡 Code mới: Tự động xây lại đáp án để sửa lỗi trùng lặp và màu sắc.")

st.write("")
num = st.number_input("Số lượng đề cần tạo", 1, 50, 4)

if st.button("🚀 BẮT ĐẦU TRỘN"):
    if not uploaded_file:
        st.warning("Vui lòng chọn file!")
    else:
        try:
            with st.spinner("Đang tái cấu trúc đề thi..."):
                final_zip = generate_mix(uploaded_file.read(), num)
                st.success("✅ Thành công! Đã sửa lỗi hiển thị.")
                st.download_button(
                    "📥 Tải về (ZIP)",
                    final_zip,
                    "KetQua_TNMic_Pro.zip",
                    "application/zip"
                )
        except Exception as e:
            st.error(f"Lỗi: {e}")

st.markdown('</div>', unsafe_allow_html=True)
st.markdown('<div class="footer">© 2024 TNMic • Re-Engineered Core</div>', unsafe_allow_html=True)
