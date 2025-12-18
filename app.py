"""
PHẦN MỀM TRỘN ĐỀ - TNMic (ULTIMATE FIX)
1. Fix lỗi nghiêm trọng: Trùng lặp đáp án (Do không xóa sạch nhãn cũ).
2. Fix lỗi màu sắc: Tạo nhãn mới hoàn toàn (Xanh Dương + Đậm + Times New Roman).
3. Giao diện: Scientific Teal (Xanh Ngọc).
"""

import streamlit as st
import re
import random
import zipfile
import io
import csv
from xml.dom import minidom

# ==================== 1. CẤU HÌNH TRANG ====================
st.set_page_config(
    page_title="TNMic - Smart Exam Mixer",
    page_icon="🧬",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# ==================== 2. CSS GIAO DIỆN (SCIENTIFIC TEAL) ====================
CUSTOM_CSS = """
<style>
    /* Ẩn header mặc định */
    header {visibility: hidden;}
    .stApp > header {display: none;}
    .block-container {
        padding-top: 1rem !important;
        padding-bottom: 5rem !important;
    }
    
    /* NỀN TRANG: Họa tiết chấm bi khoa học */
    [data-testid="stAppViewContainer"] {
        background-color: #f0fdfa; /* Teal nhạt */
        background-image: radial-gradient(#99f6e4 1px, transparent 1px);
        background-size: 20px 20px;
        font-family: 'Segoe UI', sans-serif;
    }

    /* HEADER WRAPPER */
    .header-wrapper {
        background: #ffffff;
        padding: 2.5rem;
        border-radius: 20px;
        text-align: center;
        box-shadow: 0 10px 30px -5px rgba(0, 150, 136, 0.15);
        border: 1px solid #ccfbf1;
        border-top: 6px solid #0d9488; /* Teal đậm */
        margin-bottom: 30px;
    }

    /* LOGO & TITLE */
    .school-name {
        color: #115e59;
        font-family: 'Times New Roman', serif;
        font-size: 1.8rem;
        font-weight: 900;
        text-transform: uppercase;
        margin-bottom: 5px;
        letter-spacing: 1px;
    }

    .software-badge {
        display: inline-block;
        background: linear-gradient(135deg, #0d9488 0%, #0f766e 100%);
        color: white;
        padding: 8px 30px;
        border-radius: 50px;
        font-weight: 800;
        font-size: 1.5rem;
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

    /* CARD UPLOAD */
    .main-card {
        background: rgba(255, 255, 255, 0.95);
        backdrop-filter: blur(10px);
        padding: 30px;
        border-radius: 20px;
        box-shadow: 0 4px 20px rgba(0,0,0,0.05);
        border: 1px solid #e2e8f0;
    }

    /* INPUT & BUTTON */
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

    .footer {
        text-align: center;
        color: #64748b;
        font-size: 0.85rem;
        margin-top: 40px;
    }
</style>
"""

HEADER_HTML = """
<div class="header-wrapper">
    <div class="school-name">TRƯỜNG THPT MINH ĐỨC</div>
    <div class="software-badge">PHẦN MỀM TRỘN ĐỀ</div>
    <br>
    <div class="teacher-info">GV: Nguyễn Văn Hà • Zalo: 0913968302</div>
</div>
"""

# ==================== 3. XỬ LÝ XML (CORE LOGIC) ====================
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

def create_element(doc, tag):
    return doc.createElementNS(W_NS, tag)

def create_paragraph(doc, text, align="left", bold=False):
    p = create_element(doc, "w:p")
    pPr = create_element(doc, "w:pPr")
    jc = create_element(doc, "w:jc")
    jc.setAttributeNS(W_NS, "w:val", align)
    pPr.appendChild(jc)
    p.appendChild(pPr)
    
    r = create_element(doc, "w:r")
    rPr = create_element(doc, "w:rPr")
    # Font Times New Roman
    rFonts = create_element(doc, "w:rFonts")
    rFonts.setAttributeNS(W_NS, "w:ascii", "Times New Roman")
    rFonts.setAttributeNS(W_NS, "w:hAnsi", "Times New Roman")
    rPr.appendChild(rFonts)
    if bold: rPr.appendChild(create_element(doc, "w:b"))
    r.appendChild(rPr)
    
    t = create_element(doc, "w:t")
    t.appendChild(doc.createTextNode(text))
    r.appendChild(t)
    p.appendChild(r)
    return p

def add_header_to_doc(doc, body, exam_code):
    nodes = []
    nodes.append(create_paragraph(doc, "TRƯỜNG THPT MINH ĐỨC", "center", bold=True))
    nodes.append(create_paragraph(doc, "ĐỀ KIỂM TRA ĐỊNH KỲ", "center", bold=True))
    nodes.append(create_paragraph(doc, f"MÃ ĐỀ: {exam_code}", "right", bold=True))
    nodes.append(create_paragraph(doc, "Họ tên thí sinh:...................................................... Lớp:..........", "left"))
    nodes.append(create_paragraph(doc, "", "left"))
    
    if body.hasChildNodes():
        fc = body.firstChild
        for n in reversed(nodes):
            body.insertBefore(n, fc)
    else:
        for n in nodes: body.appendChild(n)

def get_text(block):
    texts = []
    for t in block.getElementsByTagNameNS(W_NS, "t"):
        if t.firstChild: texts.append(t.firstChild.nodeValue)
    return "".join(texts).strip()

# --- XỬ LÝ ĐÁP ÁN GỐC ---
def check_is_correct(run_node):
    rPr_list = run_node.getElementsByTagNameNS(W_NS, "rPr")
    if not rPr_list: return False
    rPr = rPr_list[0]
    if rPr.getElementsByTagNameNS(W_NS, "u"): return True
    color = rPr.getElementsByTagNameNS(W_NS, "color")
    if color:
        val = color[0].getAttributeNS(W_NS, "val")
        if val and val.upper() in ["FF0000", "RED"]: return True
    return False

def remove_answer_signal(run_node):
    """Xóa gạch chân/đỏ trong run"""
    rPr_list = run_node.getElementsByTagNameNS(W_NS, "rPr")
    if not rPr_list: return
    rPr = rPr_list[0]
    for u in rPr.getElementsByTagNameNS(W_NS, "u"): rPr.removeChild(u)
    for c in rPr.getElementsByTagNameNS(W_NS, "color"): rPr.removeChild(c)

# --- THUẬT TOÁN THAY THẾ NHÃN AN TOÀN (ROBUST REPLACEMENT) ---
def replace_label_and_style(paragraph, new_label, doc, mode="mcq"):
    """
    1. Tìm chuỗi bắt đầu (A. B. hoặc a) b)...).
    2. Xóa chuỗi cũ khỏi các text node (xử lý trường hợp bị chia nhỏ).
    3. Chèn Run mới chứa nhãn mới + Style Xanh Đậm vào đầu.
    """
    pattern_str = r'^\s*[A-D][\.\)]\s*' if mode == "mcq" else r'^\s*[a-d][\.\)]\s*'
    
    # 1. Lấy toàn bộ text để xác định độ dài cần cắt
    full_text = get_text(paragraph)
    match = re.match(pattern_str, full_text)
    
    if match:
        chars_to_remove = len(match.group(0))
        
        # 2. Xóa text cũ khỏi các node (Loop qua các Run/Text node)
        t_nodes = paragraph.getElementsByTagNameNS(W_NS, "t")
        for t in t_nodes:
            if not t.firstChild: continue
            val = t.firstChild.nodeValue
            if len(val) >= chars_to_remove:
                # Cắt phần đầu (nhãn cũ), giữ phần sau
                t.firstChild.nodeValue = val[chars_to_remove:]
                chars_to_remove = 0
                break
            else:
                # Node này chứa 1 phần của nhãn, xóa hết
                chars_to_remove -= len(val)
                t.firstChild.nodeValue = ""
    
    # 3. Tạo Run mới cho Nhãn (Style chuẩn: Xanh + Đậm + Times)
    new_run = doc.createElementNS(W_NS, "w:r")
    rPr = doc.createElementNS(W_NS, "w:rPr")
    
    # Color: Blue (#0070C0)
    color = doc.createElementNS(W_NS, "w:color")
    color.setAttributeNS(W_NS, "w:val", "0070C0")
    rPr.appendChild(color)
    
    # Bold
    rPr.appendChild(doc.createElementNS(W_NS, "w:b"))
    
    # Font
    rFonts = doc.createElementNS(W_NS, "w:rFonts")
    rFonts.setAttributeNS(W_NS, "w:ascii", "Times New Roman")
    rFonts.setAttributeNS(W_NS, "w:hAnsi", "Times New Roman")
    rPr.appendChild(rFonts)
    
    new_run.appendChild(rPr)
    
    # Text
    t = doc.createElementNS(W_NS, "w:t")
    # Thêm khoảng trắng sau nhãn để đẹp (VD: "A. ")
    t.setAttribute("xml:space", "preserve")
    t.appendChild(doc.createTextNode(new_label + " "))
    new_run.appendChild(t)
    
    # Chèn vào đầu đoạn văn
    if paragraph.hasChildNodes():
        paragraph.insertBefore(new_run, paragraph.firstChild)
    else:
        paragraph.appendChild(new_run)

# --- PROCESSORS ---

def process_mcq(q_blocks, doc):
    pat = r'^\s*[A-D][\.\)]'
    indices = [k for k, b in enumerate(q_blocks) if re.match(pat, get_text(b))]
    correct_char = "X"
    
    if len(indices) >= 2:
        opts = [q_blocks[k] for k in indices]
        
        target_opt = None
        for opt in opts:
            runs = opt.getElementsByTagNameNS(W_NS, "r")
            is_cor = False
            for r in runs:
                if check_is_correct(r): is_cor = True
                remove_answer_signal(r) # Xóa dấu hiệu gốc
            if is_cor: target_opt = opt
            
        random.shuffle(opts)
        
        lbls = ["A.", "B.", "C.", "D."]
        for idx, opt in enumerate(opts):
            real_idx = indices[idx]
            q_blocks[real_idx] = opt
            
            if target_opt and opt == target_opt:
                correct_char = lbls[idx][0]
            
            # SỬ DỤNG HÀM THAY THẾ AN TOÀN
            replace_label_and_style(opt, lbls[idx], doc, "mcq")
            
    return q_blocks, correct_char

def process_tf(q_blocks, doc):
    pat = r'^\s*[a-d][\.\)]'
    indices = [k for k, b in enumerate(q_blocks) if re.match(pat, get_text(b))]
    res_str = []
    
    if len(indices) >= 2:
        opts = [q_blocks[k] for k in indices]
        
        status_map = {}
        for opt in opts:
            is_true = False
            runs = opt.getElementsByTagNameNS(W_NS, "r")
            for r in runs:
                if check_is_correct(r): is_true = True
                remove_answer_signal(r)
            status_map[opt] = "Đ" if is_true else "S"
            
        random.shuffle(opts)
        
        lbls = ["a)", "b)", "c)", "d)"]
        for idx, opt in enumerate(opts):
            real_idx = indices[idx]
            q_blocks[real_idx] = opt
            
            curr_lbl = lbls[idx]
            res_str.append(f"{curr_lbl[:-1]}{status_map[opt]}")
            
            # SỬ DỤNG HÀM THAY THẾ AN TOÀN
            replace_label_and_style(opt, curr_lbl, doc, "tf")
            
    return q_blocks, " - ".join(res_str)

def process_short(q_blocks):
    """Tìm key trong toàn bộ text, xóa key khỏi đề"""
    key_val = ""
    full_text = ""
    for b in q_blocks:
        t_nodes = b.getElementsByTagNameNS(W_NS, "t")
        for t in t_nodes:
            if t.firstChild: full_text += t.firstChild.nodeValue

    m = re.search(r'<\s*key\s*=\s*(.*?)\s*>', full_text, re.IGNORECASE)
    if m:
        key_val = m.group(1).strip()
        # Xóa thẻ key
        for b in q_blocks:
            t_nodes = b.getElementsByTagNameNS(W_NS, "t")
            for t in t_nodes:
                if t.firstChild and '<' in t.firstChild.nodeValue:
                    val = t.firstChild.nodeValue
                    val = re.sub(r'<\s*key\s*=\s*.*?>', '', val, flags=re.IGNORECASE)
                    t.firstChild.nodeValue = val
    return q_blocks, key_val

# --- PARSER ---
def parse_blocks(blocks):
    intro, questions = [], []
    i = 0
    while i < len(blocks):
        txt = get_text(blocks[i])
        if re.match(r'^Câu\s*\d+', txt): break
        intro.append(blocks[i])
        i += 1
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
    return intro, questions

# --- MAIN ENGINE ---

def generate_mix(file_bytes, num_copies):
    outer_zip_buffer = io.BytesIO()
    csv_data = []
    
    with zipfile.ZipFile(outer_zip_buffer, 'w', zipfile.ZIP_DEFLATED) as outer_zip:
        input_io = io.BytesIO(file_bytes)
        with zipfile.ZipFile(input_io, 'r') as z_in:
            xml_content = z_in.read("word/document.xml")
            
            for _ in range(num_copies):
                exam_code = str(random.randint(1001, 9999))
                dom = minidom.parseString(xml_content)
                doc = dom.documentElement
                body = dom.getElementsByTagNameNS(W_NS, "body")[0]
                blocks = [n for n in body.childNodes if n.localName in ['p', 'tbl']]
                intro, all_qs = parse_blocks(blocks)
                
                # Cắt phần
                p1_qs = all_qs[0:18]
                p2_qs = all_qs[18:22]
                p3_qs = all_qs[22:]
                
                row_key = [exam_code]
                
                # P1
                p1_fin, k1 = [], []
                for q in p1_qs:
                    q_new, k = process_mcq(q, dom)
                    p1_fin.append(q_new)
                    k1.append(k)
                c1 = list(zip(p1_fin, k1))
                random.shuffle(c1)
                if c1: p1_fin, k1 = zip(*c1)
                row_key.extend(k1)
                
                # P2
                p2_fin, k2 = [], []
                for q in p2_qs:
                    q_new, k = process_tf(q, dom)
                    p2_fin.append(q_new)
                    k2.append(k)
                c2 = list(zip(p2_fin, k2))
                random.shuffle(c2)
                if c2: p2_fin, k2 = zip(*c2)
                row_key.extend(k2)
                
                # P3
                p3_fin, k3 = [], []
                for q in p3_qs:
                    q_new, k = process_short(q)
                    p3_fin.append(q_new)
                    k3.append(k)
                c3 = list(zip(p3_fin, k3))
                random.shuffle(c3)
                if c3: p3_fin, k3 = zip(*c3)
                row_key.extend(k3)
                
                csv_data.append(row_key)
                
                # Rebuild DOC
                final_blocks = []
                final_blocks.append(create_paragraph(dom, "PHẦN I. Trắc nghiệm nhiều lựa chọn (18 câu)", bold=True))
                for i, q in enumerate(p1_fin):
                    t_list = q[0].getElementsByTagNameNS(W_NS, "t")
                    for t in t_list:
                        # Đánh lại số câu (Câu 1...)
                        if t.firstChild and re.match(r'^Câu\s*\d+', t.firstChild.nodeValue):
                            t.firstChild.nodeValue = re.sub(r'^Câu\s*\d+', f"Câu {i+1}", t.firstChild.nodeValue)
                            # Tô màu xanh số câu (Câu 1, Câu 2...)
                            replace_label_and_style(q[0], f"Câu {i+1}.", dom, "mcq") 
                            break
                    final_blocks.extend(q)
                    
                final_blocks.append(create_paragraph(dom, "PHẦN II. Trắc nghiệm đúng sai (4 câu)", bold=True))
                for i, q in enumerate(p2_fin):
                    t_list = q[0].getElementsByTagNameNS(W_NS, "t")
                    for t in t_list:
                        if t.firstChild and re.match(r'^Câu\s*\d+', t.firstChild.nodeValue):
                            t.firstChild.nodeValue = re.sub(r'^Câu\s*\d+', f"Câu {i+1}", t.firstChild.nodeValue)
                            replace_label_and_style(q[0], f"Câu {i+1}.", dom, "mcq")
                            break
                    final_blocks.extend(q)
                
                final_blocks.append(create_paragraph(dom, "PHẦN III. Trả lời ngắn (6 câu)", bold=True))
                for i, q in enumerate(p3_fin):
                    t_list = q[0].getElementsByTagNameNS(W_NS, "t")
                    for t in t_list:
                        if t.firstChild and re.match(r'^Câu\s*\d+', t.firstChild.nodeValue):
                            t.firstChild.nodeValue = re.sub(r'^Câu\s*\d+', f"Câu {i+1}", t.firstChild.nodeValue)
                            replace_label_and_style(q[0], f"Câu {i+1}.", dom, "mcq")
                            break
                    final_blocks.extend(q)
                
                for n in list(body.childNodes):
                    if n.localName in ['p', 'tbl']: body.removeChild(n)
                for b in final_blocks: body.appendChild(b)
                
                add_header_to_doc(dom, body, exam_code)
                
                new_xml = dom.toxml().encode('utf-8')
                doc_io = io.BytesIO()
                with zipfile.ZipFile(doc_io, 'w', zipfile.ZIP_DEFLATED) as z_d:
                    for item in z_in.infolist():
                        if item.filename == "word/document.xml":
                            z_d.writestr(item.filename, new_xml)
                        else:
                            z_d.writestr(item.filename, z_in.read(item.filename))
                outer_zip.writestr(f"De_Thi/De_{exam_code}.docx", doc_io.getvalue())
        
        csv_io = io.StringIO()
        writer = csv.writer(csv_io)
        head = ["Mã đề"] + [str(i) for i in range(1,19)] + [f"II_C{i}" for i in range(1,5)] + [f"III_C{i}" for i in range(1,7)]
        writer.writerow(head)
        writer.writerows(csv_data)
        outer_zip.writestr("Dap_An.csv", csv_io.getvalue().encode('utf-8-sig'))

    return outer_zip_buffer.getvalue()

# ==================== 4. UI LOGIC ====================
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
