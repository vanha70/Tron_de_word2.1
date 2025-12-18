"""
TNMic PRO - PHẦN MỀM TRỘN ĐỀ (FINAL STABLE)
1. Giao diện: Xanh Ngọc (Teal) + Chấm bi.
2. Core Logic: "Ăn mòn" nhãn cũ -> Chèn nhãn mới (In Đậm + Xanh).
3. Fix lỗi: Trùng lặp đáp án, sót màu.
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
    page_title="TNMic - Trộn Đề Minh Đức",
    page_icon="🧬",
    layout="centered",
    initial_sidebar_state="collapsed"
)

CUSTOM_CSS = """
<style>
    /* Ẩn header mặc định */
    header {visibility: hidden;}
    .stApp > header {display: none;}
    .block-container {padding-top: 1rem; padding-bottom: 5rem;}
    
    /* NỀN TRANG: Họa tiết chấm bi khoa học */
    [data-testid="stAppViewContainer"] {
        background-color: #f0fdfa; /* Teal rất nhạt */
        background-image: radial-gradient(#99f6e4 1px, transparent 1px);
        background-size: 20px 20px;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }

    /* HEADER */
    .header-wrapper {
        background: #ffffff;
        padding: 2rem;
        border-radius: 20px;
        text-align: center;
        box-shadow: 0 4px 20px rgba(0, 150, 136, 0.1);
        border-top: 5px solid #0d9488; /* Teal đậm */
        margin-bottom: 20px;
    }

    .school-name {
        color: #115e59;
        font-family: 'Times New Roman', serif;
        font-size: 1.8rem;
        font-weight: 900;
        text-transform: uppercase;
        margin-bottom: 10px;
    }

    .app-badge {
        display: inline-block;
        background: linear-gradient(135deg, #0d9488 0%, #0f766e 100%);
        color: white;
        padding: 8px 30px;
        border-radius: 50px;
        font-weight: 800;
        font-size: 1.4rem;
        text-transform: uppercase;
        margin: 10px 0;
        box-shadow: 0 4px 10px rgba(13, 148, 136, 0.3);
    }

    .teacher-info {
        font-size: 1rem;
        font-weight: 600;
        color: #0f766e;
        background-color: #ccfbf1;
        padding: 8px 20px;
        border-radius: 12px;
        display: inline-block;
        margin-top: 15px;
        border: 1px solid #99f6e4;
    }

    /* MAIN CARD */
    .main-card {
        background: rgba(255, 255, 255, 0.95);
        backdrop-filter: blur(10px);
        padding: 25px;
        border-radius: 20px;
        box-shadow: 0 4px 20px rgba(0,0,0,0.05);
        border: 1px solid #e2e8f0;
    }

    /* BUTTON */
    .stButton > button {
        background: linear-gradient(90deg, #0d9488, #14b8a6);
        color: white;
        border: none;
        padding: 12px 24px;
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
        background: linear-gradient(90deg, #0f766e, #0d9488);
    }

    .footer { text-align: center; color: #64748b; font-size: 0.85rem; margin-top: 40px; }
</style>
"""

HEADER_HTML = """
<div class="header-wrapper">
    <div class="school-name">TRƯỜNG THPT MINH ĐỨC</div>
    <div class="app-badge">PHẦN MỀM TRỘN ĐỀ</div><br>
    <div class="teacher-info">GV: Nguyễn Văn Hà • Zalo: 0913968302</div>
</div>
"""

# ==================== 2. HÀM XỬ LÝ WORD (CORE) ====================
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

def get_text(node):
    """Lấy toàn bộ text trong node"""
    texts = []
    for t in node.getElementsByTagNameNS(W_NS, "t"):
        if t.firstChild: texts.append(t.firstChild.nodeValue)
    return "".join(texts)

def create_run_label(doc, text, color_hex="0070C0"):
    """Tạo Run chứa nhãn với Style: Xanh Dương + In Đậm"""
    r = doc.createElementNS(W_NS, "w:r")
    rPr = doc.createElementNS(W_NS, "w:rPr")
    
    # 1. Màu sắc (Blue)
    color = doc.createElementNS(W_NS, "w:color")
    color.setAttributeNS(W_NS, "w:val", color_hex)
    rPr.appendChild(color)
    
    # 2. In đậm (Bold)
    b = doc.createElementNS(W_NS, "w:b")
    rPr.appendChild(b)
    
    # 3. Font (Times New Roman)
    rFonts = doc.createElementNS(W_NS, "w:rFonts")
    rFonts.setAttributeNS(W_NS, "w:ascii", "Times New Roman")
    rFonts.setAttributeNS(W_NS, "w:hAnsi", "Times New Roman")
    rPr.appendChild(rFonts)
    
    r.appendChild(rPr)
    
    # 4. Text
    t = doc.createElementNS(W_NS, "w:t")
    t.setAttribute("xml:space", "preserve")
    t.appendChild(doc.createTextNode(text))
    r.appendChild(t)
    
    return r

def replace_label_robust(paragraph, new_label, doc, pattern_regex):
    """
    Hàm thay thế nhãn an toàn:
    1. Tìm nhãn cũ (VD: "A. ") bằng Regex.
    2. Xóa chính xác số lượng ký tự của nhãn cũ khỏi các node văn bản đầu tiên.
    3. Chèn nhãn mới (có Style) vào đầu đoạn văn.
    """
    full_text = get_text(paragraph)
    match = re.match(pattern_regex, full_text)
    
    if match:
        chars_to_remove = len(match.group(0)) # Độ dài chuỗi cần xóa
        
        # Duyệt qua các thẻ <w:t> để xóa dần ký tự
        t_nodes = paragraph.getElementsByTagNameNS(W_NS, "t")
        nodes_to_remove = []
        
        for t in t_nodes:
            if not t.firstChild: continue
            
            val = t.firstChild.nodeValue
            len_val = len(val)
            
            if chars_to_remove > 0:
                if len_val <= chars_to_remove:
                    # Node này nằm trọn trong phần cần xóa -> Đánh dấu xóa
                    t.firstChild.nodeValue = "" 
                    chars_to_remove -= len_val
                else:
                    # Node này chỉ chứa một phần -> Cắt phần đầu
                    t.firstChild.nodeValue = val[chars_to_remove:]
                    chars_to_remove = 0
            
            if chars_to_remove == 0:
                break
    
    # Tạo Run mới chứa nhãn mới (Style: Xanh + Đậm)
    # Thêm khoảng trắng sau nhãn (VD: "A. ")
    new_run = create_run_label(doc, new_label + " ")
    
    # Chèn vào vị trí đầu tiên của paragraph
    if paragraph.hasChildNodes():
        paragraph.insertBefore(new_run, paragraph.firstChild)
    else:
        paragraph.appendChild(new_run)

# --- CÁC HÀM XỬ LÝ LOGIC TRỘN ---

def check_correct_node(paragraph):
    """Kiểm tra xem đoạn văn này có chứa đáp án đúng (Gạch chân/Đỏ) hay không"""
    runs = paragraph.getElementsByTagNameNS(W_NS, "r")
    for r in runs:
        rPr = r.getElementsByTagNameNS(W_NS, "rPr")
        if rPr:
            # Check Gạch chân
            if rPr[0].getElementsByTagNameNS(W_NS, "u"): return True
            # Check Màu đỏ
            color = rPr[0].getElementsByTagNameNS(W_NS, "color")
            if color and color[0].getAttributeNS(W_NS, "val") in ["FF0000", "RED"]: return True
    return False

def clean_formatting(paragraph):
    """Xóa bỏ gạch chân và màu đỏ (để ẩn đáp án đúng)"""
    runs = paragraph.getElementsByTagNameNS(W_NS, "r")
    for r in runs:
        rPr = r.getElementsByTagNameNS(W_NS, "rPr")
        if rPr:
            for tag in ["u", "color", "b"]: # Xóa cả Bold cũ để đồng bộ
                for node in rPr[0].getElementsByTagNameNS(W_NS, tag):
                    rPr[0].removeChild(node)

def process_mcq(questions, doc):
    """Xử lý Phần 1 (A,B,C,D)"""
    processed_qs = []
    keys = []
    
    for q_block in questions:
        # Tìm các dòng chứa A., B., ...
        pat = r'^\s*[A-D][\.\)]'
        opt_indices = [i for i, n in enumerate(q_block) if re.match(pat, get_text(n))]
        
        correct_char = "X"
        
        if len(opt_indices) >= 2:
            opts = [q_block[i] for i in opt_indices]
            
            # Xác định đáp án đúng
            target_opt = None
            for opt in opts:
                if check_correct_node(opt):
                    target_opt = opt
                clean_formatting(opt) # Xóa dấu hiệu
            
            # Trộn
            random.shuffle(opts)
            labels = ["A.", "B.", "C.", "D."]
            
            # Gán lại và Thay nhãn
            for i, idx in enumerate(opt_indices):
                q_block[idx] = opts[i]
                if opts[i] == target_opt: correct_char = labels[i][0]
                
                # THAY NHÃN MỚI (IN ĐẬM + XANH)
                replace_label_robust(opts[i], labels[i], doc, pat)
        
        keys.append(correct_char)
        processed_qs.append(q_block)
        
    return processed_qs, keys

def process_tf(questions, doc):
    """Xử lý Phần 2 (Đúng/Sai)"""
    processed_qs = []
    keys = []
    
    for q_block in questions:
        pat = r'^\s*[a-d][\.\)]'
        opt_indices = [i for i, n in enumerate(q_block) if re.match(pat, get_text(n))]
        res_str = []
        
        if len(opt_indices) >= 2:
            opts = [q_block[i] for i in opt_indices]
            
            status_map = {}
            for opt in opts:
                is_true = check_correct_node(opt)
                clean_formatting(opt)
                status_map[opt] = "Đ" if is_true else "S"
            
            random.shuffle(opts)
            labels = ["a)", "b)", "c)", "d)"]
            
            for i, idx in enumerate(opt_indices):
                q_block[idx] = opts[i]
                res_str.append(f"{labels[i][:-1]}{status_map[opts[i]]}")
                
                # THAY NHÃN MỚI
                replace_label_robust(opts[i], labels[i], doc, pat)
                
        keys.append(" - ".join(res_str))
        processed_qs.append(q_block)
        
    return processed_qs, keys

def process_short(questions):
    """Xử lý Phần 3 (Key)"""
    processed_qs = []
    keys = []
    
    for q_block in questions:
        key_val = ""
        full_text = "".join([get_text(n) for n in q_block])
        m = re.search(r'<\s*key\s*=\s*(.*?)\s*>', full_text, re.IGNORECASE)
        
        if m:
            key_val = m.group(1).strip()
            # Xóa thẻ key trong XML
            for node in q_block:
                t_nodes = node.getElementsByTagNameNS(W_NS, "t")
                for t in t_nodes:
                    if t.firstChild and '<' in t.firstChild.nodeValue:
                        val = t.firstChild.nodeValue
                        val = re.sub(r'<\s*key\s*=\s*.*?>', '', val, flags=re.IGNORECASE)
                        t.firstChild.nodeValue = val
                        
        keys.append(key_val)
        processed_qs.append(q_block)
        
    return processed_qs, keys

# ==================== 4. MAIN LOGIC ====================

def parse_docx(dom):
    body = dom.getElementsByTagNameNS(W_NS, "body")[0]
    blocks = [n for n in body.childNodes if n.localName in ['p', 'tbl']]
    intro, questions = [], []
    
    i = 0
    # Intro
    while i < len(blocks):
        if re.match(r'^Câu\s*\d+', get_text(blocks[i])): break
        intro.append(blocks[i])
        i += 1
    # Questions
    while i < len(blocks):
        if re.match(r'^Câu\s*\d+', get_text(blocks[i])):
            grp = [blocks[i]]
            i += 1
            while i < len(blocks):
                txt = get_text(blocks[i])
                if re.match(r'^Câu\s*\d+', txt) or "PHẦN" in txt.upper(): break
                grp.append(blocks[i])
                i += 1
            questions.append(grp)
        else: i += 1
    return intro, questions, body

def create_header(doc, text, align="left", bold=False):
    p = doc.createElementNS(W_NS, "w:p")
    pPr = doc.createElementNS(W_NS, "w:pPr")
    jc = doc.createElementNS(W_NS, "w:jc")
    jc.setAttributeNS(W_NS, "w:val", align)
    pPr.appendChild(jc)
    p.appendChild(pPr)
    
    # Run
    r = doc.createElementNS(W_NS, "w:r")
    rPr = doc.createElementNS(W_NS, "w:rPr")
    if bold: rPr.appendChild(doc.createElementNS(W_NS, "w:b"))
    
    # Font
    rFonts = doc.createElementNS(W_NS, "w:rFonts")
    rFonts.setAttributeNS(W_NS, "w:ascii", "Times New Roman")
    rFonts.setAttributeNS(W_NS, "w:hAnsi", "Times New Roman")
    rPr.appendChild(rFonts)
    
    r.appendChild(rPr)
    t = doc.createElementNS(W_NS, "w:t")
    t.appendChild(doc.createTextNode(text))
    r.appendChild(t)
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
                
                # Chia phần
                p1 = all_qs[0:18]
                p2 = all_qs[18:22]
                p3 = all_qs[22:]
                
                row_key = [exam_code]
                
                # Xử lý
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
                
                # REBUILD XML BODY
                # Xóa hết nội dung cũ
                for n in list(body.childNodes):
                    if n.localName in ['p', 'tbl']: body.removeChild(n)
                
                # Thêm Header Trường
                body.appendChild(create_header(dom, "TRƯỜNG THPT MINH ĐỨC", "center", True))
                body.appendChild(create_header(dom, "ĐỀ KIỂM TRA ĐỊNH KỲ", "center", True))
                body.appendChild(create_header(dom, f"MÃ ĐỀ: {exam_code}", "right", True))
                body.appendChild(create_header(dom, "Họ tên thí sinh:............................................ Lớp:..........", "left"))
                body.appendChild(create_header(dom, "", "left"))
                
                # Thêm P1
                body.appendChild(create_header(dom, "PHẦN I. Trắc nghiệm (18 câu)", "left", True))
                for idx, q in enumerate(p1_fin):
                    replace_label_robust(q[0], f"Câu {idx+1}.", dom, r'^Câu\s*\d+[\.\:]')
                    for n in q: body.appendChild(n)
                    
                # Thêm P2
                body.appendChild(create_header(dom, "PHẦN II. Đúng Sai (4 câu)", "left", True))
                for idx, q in enumerate(p2_fin):
                    replace_label_robust(q[0], f"Câu {idx+1}.", dom, r'^Câu\s*\d+[\.\:]')
                    for n in q: body.appendChild(n)
                    
                # Thêm P3
                body.appendChild(create_header(dom, "PHẦN III. Trả lời ngắn (6 câu)", "left", True))
                for idx, q in enumerate(p3_fin):
                    replace_label_robust(q[0], f"Câu {idx+1}.", dom, r'^Câu\s*\d+[\.\:]')
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
        
        # Write CSV
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
    st.info("Quy tắc: P1, P2 gạch chân đáp án. P3 dùng thẻ <key=...>")

st.write("")
num = st.number_input("Số lượng đề cần tạo", 1, 50, 4)

if st.button("🚀 BẮT ĐẦU TRỘN"):
    if not uploaded_file:
        st.warning("Vui lòng chọn file!")
    else:
        try:
            with st.spinner("Đang xử lý..."):
                final_zip = generate_mix(uploaded_file.read(), num)
                st.success("✅ Thành công! Tải xuống bên dưới.")
                st.download_button(
                    label="📥 Tải về (Đề thi + Đáp án)",
                    data=final_zip,
                    file_name="KetQua_TNMic.zip",
                    mime="application/zip"
                )
        except Exception as e:
            st.error(f"Lỗi: {e}")

st.markdown('</div>', unsafe_allow_html=True)
st.markdown('<div class="footer">© 2024 Phần mềm Trộn Đề [TNMic]</div>', unsafe_allow_html=True)
