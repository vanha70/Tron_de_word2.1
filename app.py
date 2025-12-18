"""
PHẦN MỀM TRỘN ĐỀ - TNMic (FINAL FIX: SYNTAX & COLOR)
1. Sửa lỗi SyntaxError (Lỗi cú pháp dòng lệnh).
2. Sửa lỗi Màu sắc: Ép buộc toàn bộ A. B. C. D. thành Xanh Dương + Đậm.
3. Giao diện: Xanh Ngọc (Teal) khoa học.
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
    page_title="TNMic - Trộn Đề Minh Đức",
    page_icon="📘",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# ==================== 2. CSS GIAO DIỆN (TEAL THEME) ====================
CUSTOM_CSS = """
<style>
    /* Ẩn header mặc định */
    header {visibility: hidden;}
    .stApp > header {display: none;}
    .block-container {
        padding-top: 2rem !important;
        padding-bottom: 5rem !important;
    }
    
    /* NỀN TRANG: Gradient Xanh Ngọc */
    [data-testid="stAppViewContainer"] {
        background: linear-gradient(135deg, #e0f2f1 0%, #b2dfdb 100%);
        font-family: 'Segoe UI', sans-serif;
    }

    /* HEADER */
    .header-wrapper {
        background: white;
        padding: 2.5rem;
        border-radius: 20px;
        text-align: center;
        box-shadow: 0 10px 30px rgba(0, 150, 136, 0.2);
        border-bottom: 5px solid #009688;
        margin-bottom: 30px;
    }

    /* LOGO TNMic */
    .tnmic-logo {
        background: linear-gradient(135deg, #26a69a 0%, #00897b 100%);
        color: white;
        font-family: 'Arial', sans-serif;
        font-size: 2.5rem;
        font-weight: 900;
        text-transform: uppercase;
        padding: 10px 40px;
        border-radius: 50px;
        display: inline-block;
        margin-bottom: 15px;
        box-shadow: 0 5px 15px rgba(0, 137, 123, 0.4);
        text-shadow: 1px 1px 2px rgba(0,0,0,0.2);
        letter-spacing: 2px;
    }

    .software-name {
        color: #00796b;
        font-size: 1.6rem;
        font-weight: 800;
        margin: 5px 0;
        text-transform: uppercase;
    }

    .teacher-info {
        font-size: 1.1rem;
        font-weight: 700;
        color: #004d40;
        background-color: #e0f2f1;
        padding: 10px 25px;
        border-radius: 50px;
        display: inline-block;
        margin-top: 15px;
        border: 2px solid #80cbc4;
    }

    /* CARD UPLOAD */
    .main-card {
        background: white;
        padding: 30px;
        border-radius: 20px;
        box-shadow: 0 8px 20px rgba(0,0,0,0.05);
    }

    /* NÚT BẤM (TEAL) */
    .stButton > button {
        background: linear-gradient(90deg, #26a69a, #00897b);
        color: white;
        font-weight: bold;
        border: none;
        padding: 12px 0;
        border-radius: 10px;
        width: 100%;
        font-size: 1.2rem;
        text-transform: uppercase;
        box-shadow: 0 4px 10px rgba(0, 150, 136, 0.3);
        transition: all 0.3s;
    }
    .stButton > button:hover {
        background: linear-gradient(90deg, #00897b, #004d40);
        transform: translateY(-2px);
        box-shadow: 0 6px 15px rgba(0, 150, 136, 0.5);
    }
    
    .footer {
        text-align: center;
        margin-top: 40px;
        color: #546e7a;
        font-size: 0.9rem;
    }
</style>
"""

HEADER_HTML = """
<div class="header-wrapper">
    <div class="tnmic-logo">TNMic</div>
    <div class="software-name">PHẦN MỀM TRỘN ĐỀ</div>
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
    nodes.append(create_paragraph(doc, "TNMic - TRƯỜNG THPT MINH ĐỨC", "center", bold=True))
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
    rPr_list = run_node.getElementsByTagNameNS(W_NS, "rPr")
    if not rPr_list: return
    rPr = rPr_list[0]
    for u in rPr.getElementsByTagNameNS(W_NS, "u"): rPr.removeChild(u)
    for c in rPr.getElementsByTagNameNS(W_NS, "color"): rPr.removeChild(c)

def style_label_force_blue(run_node, doc):
    """
    FIX COLOR: Ép màu Xanh Dương (#0070C0) và In Đậm
    """
    rPr_list = run_node.getElementsByTagNameNS(W_NS, "rPr")
    if rPr_list: 
        rPr = rPr_list[0]
    else:
        rPr = doc.createElementNS(W_NS, "w:rPr")
        run_node.insertBefore(rPr, run_node.firstChild)
    
    # 1. Xóa màu cũ
    for c in rPr.getElementsByTagNameNS(W_NS, "color"): rPr.removeChild(c)
    # 2. Xóa bold cũ
    for b in rPr.getElementsByTagNameNS(W_NS, "b"): rPr.removeChild(b)

    # 3. Thêm màu Xanh Dương
    color_node = doc.createElementNS(W_NS, "w:color")
    color_node.setAttributeNS(W_NS, "w:val", "0070C0")
    rPr.appendChild(color_node)

    # 4. Thêm In Đậm
    rPr.appendChild(doc.createElementNS(W_NS, "w:b"))

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

# --- PROCESSORS ---

def process_mcq(q_blocks, doc):
    # Regex tìm A. B. C. D. (để thay thế)
    pat = r'^\s*[A-D][\.\)]'
    indices = [k for k, b in enumerate(q_blocks) if re.match(pat, get_text(b))]
    correct_char = ""
    
    if len(indices) >= 2:
        opts = [q_blocks[k] for k in indices]
        
        target_opt = None
        for opt in opts:
            runs = opt.getElementsByTagNameNS(W_NS, "r")
            is_cor = False
            for r in runs:
                if check_is_correct(r): is_cor = True
                remove_answer_signal(r)
            if is_cor: target_opt = opt
            
        random.shuffle(opts)
        
        lbls = ["A.", "B.", "C.", "D."]
        for idx, opt in enumerate(opts):
            real_idx = indices[idx]
            q_blocks[real_idx] = opt
            
            if target_opt and opt == target_opt:
                correct_char = lbls[idx][0]
            
            # Thay thế nhãn và ÉP MÀU
            t_nodes = opt.getElementsByTagNameNS(W_NS, "t")
            for t in t_nodes:
                if t.firstChild:
                    val = t.firstChild.nodeValue
                    # Thay thế ký tự đầu tiên
                    new_val = re.sub(pat, lbls[idx], val, 1)
                    t.firstChild.nodeValue = new_val
                    
                    # QUAN TRỌNG: Gọi hàm ép màu cho node cha của text này
                    style_label_force_blue(t.parentNode, doc)
                    break
                    
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
            
            t_nodes = opt.getElementsByTagNameNS(W_NS, "t")
            for t in t_nodes:
                if t.firstChild:
                    t.firstChild.nodeValue = re.sub(pat, curr_lbl, t.firstChild.nodeValue, 1)
                    style_label_force_blue(t.parentNode, doc)
                    break
                    
    return q_blocks, " - ".join(res_str)

def process_short(q_blocks):
    key_val = ""
    full_text = ""
    for b in q_blocks:
        t_nodes = b.getElementsByTagNameNS(W_NS, "t")
        for t in t_nodes:
            if t.firstChild: full_text += t.firstChild.nodeValue

    m = re.search(r'<\s*key\s*=\s*(.*?)\s*>', full_text, re.IGNORECASE)
    if m:
        key_val = m.group(1).strip()
        for b in q_blocks:
            t_nodes = b.getElementsByTagNameNS(W_NS, "t")
            for t in t_nodes:
                if t.firstChild and '<' in t.firstChild.nodeValue:
                    val = t.firstChild.nodeValue
                    val = re.sub(r'<\s*key\s*=\s*.*?>', '', val, flags=re.IGNORECASE)
                    t.firstChild.nodeValue = val
    return q_blocks, key_val

# --- MAIN GENERATOR ---

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
                
                # CHIA 3 PHẦN
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
                        if t.firstChild and re.match(r'^Câu\s*\d+', t.firstChild.nodeValue):
                            t.firstChild.nodeValue = re.sub(r'^Câu\s*\d+', f"Câu {i+1}", t.firstChild.nodeValue)
                            style_label_force_blue(t.parentNode, dom)
                            break
                    final_blocks.extend(q)
                    
                final_blocks.append(create_paragraph(dom, "PHẦN II. Trắc nghiệm đúng sai (4 câu)", bold=True))
                for i, q in enumerate(p2_fin):
                    t_list = q[0].getElementsByTagNameNS(W_NS, "t")
                    for t in t_list:
                        if t.firstChild and re.match(r'^Câu\s*\d+', t.firstChild.nodeValue):
                            t.firstChild.nodeValue = re.sub(r'^Câu\s*\d+', f"Câu {i+1}", t.firstChild.nodeValue)
                            style_label_force_blue(t.parentNode, dom)
                            break
                    final_blocks.extend(q)
                    
                final_blocks.append(create_paragraph(dom, "PHẦN III. Trả lời ngắn (6 câu)", bold=True))
                for i, q in enumerate(p3_fin):
                    t_list = q[0].getElementsByTagNameNS(W_NS, "t")
                    for t in t_list:
                        if t.firstChild and re.match(r'^Câu\s*\d+', t.firstChild.nodeValue):
                            t.firstChild.nodeValue = re.sub(r'^Câu\s*\d+', f"Câu {i+1}", t.firstChild.nodeValue)
                            style_label_force_blue(t.parentNode, dom)
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
