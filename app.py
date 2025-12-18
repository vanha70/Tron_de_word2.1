"""
PHẦN MỀM TRỘN ĐỀ - TNMic (SCIENTIFIC UI + FIX COLOR)
1. Giao diện: Phong cách Khoa học, Hiện đại (Dot Matrix Background).
2. Logic: Cưỡng chế tô màu Xanh Dương (#0070C0) + In Đậm cho mọi đáp án A.B.C.D.
3. Hệ thống: Xuất 1 file Zip chứa Đề + Đáp án.
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

# ==================== 2. CSS GIAO DIỆN (SCIENTIFIC THEME) ====================
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
        background-color: #f8fafc;
        background-image: radial-gradient(#cbd5e1 1px, transparent 1px);
        background-size: 20px 20px;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }

    /* HEADER CARD */
    .header-wrapper {
        background: #ffffff;
        padding: 2.5rem;
        border-radius: 24px;
        text-align: center;
        box-shadow: 0 10px 40px -10px rgba(0,0,0,0.1);
        border: 1px solid #e2e8f0;
        margin-bottom: 30px;
        position: relative;
        overflow: hidden;
    }
    
    /* Trang trí Header */
    .header-wrapper::before {
        content: "";
        position: absolute;
        top: 0; left: 0; right: 0;
        height: 6px;
        background: linear-gradient(90deg, #3b82f6, #f97316);
    }

    /* LOGO TNMic */
    .tnmic-logo {
        display: inline-block;
        background: linear-gradient(135deg, #f97316 0%, #ea580c 100%);
        color: white;
        font-family: 'Arial', sans-serif;
        font-weight: 900;
        font-size: 2.2rem;
        padding: 8px 35px;
        border-radius: 12px;
        letter-spacing: 1px;
        margin-bottom: 15px;
        box-shadow: 0 4px 15px rgba(249, 115, 22, 0.4);
    }

    .app-title {
        color: #1e293b;
        font-size: 1.5rem;
        font-weight: 800;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        margin: 10px 0;
    }

    .teacher-badge {
        display: inline-flex;
        align-items: center;
        background: #eff6ff;
        color: #1d4ed8;
        padding: 8px 20px;
        border-radius: 50px;
        font-weight: 600;
        font-size: 0.95rem;
        border: 1px solid #bfdbfe;
        margin-top: 10px;
    }

    /* MAIN CARD */
    .main-card {
        background: rgba(255, 255, 255, 0.9);
        backdrop-filter: blur(10px);
        padding: 30px;
        border-radius: 24px;
        border: 1px solid #e2e8f0;
        box-shadow: 0 4px 20px -5px rgba(0,0,0,0.05);
    }

    /* INPUT & UPLOAD */
    [data-testid="stFileUploader"] section {
        background-color: #f1f5f9;
        border: 2px dashed #94a3b8;
        border-radius: 15px;
    }

    /* BUTTON TRỘN ĐỀ */
    .stButton > button {
        background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%); /* Xanh Khoa học */
        color: white;
        border: none;
        padding: 14px 24px;
        border-radius: 12px;
        font-weight: 700;
        font-size: 1.1rem;
        text-transform: uppercase;
        width: 100%;
        box-shadow: 0 4px 12px rgba(37, 99, 235, 0.3);
        transition: all 0.2s ease;
    }
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 8px 20px rgba(37, 99, 235, 0.4);
    }

    .footer {
        text-align: center;
        color: #64748b;
        font-size: 0.85rem;
        margin-top: 40px;
        padding-bottom: 20px;
    }
</style>
"""

HEADER_HTML = """
<div class="header-wrapper">
    <div class="tnmic-logo">TNMic</div>
    <div class="app-title">PHẦN MỀM TRỘN ĐỀ TRẮC NGHIỆM</div>
    <div class="teacher-badge">
        GV: Nguyễn Văn Hà &nbsp; • &nbsp; Zalo: 0913968302
    </div>
</div>
"""

# ==================== 3. XỬ LÝ WORD XML (CORE) ====================
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

# --- HELPER STYLE ---

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
    FIX COLOR: Cưỡng chế tô màu Xanh Dương (#0070C0) và In Đậm
    Bất chấp định dạng cũ là gì.
    """
    # 1. Lấy hoặc tạo rPr
    rPr_list = run_node.getElementsByTagNameNS(W_NS, "rPr")
    if rPr_list: 
        rPr = rPr_list[0]
    else:
        rPr = doc.createElementNS(W_NS, "w:rPr")
        run_node.insertBefore(rPr, run_node.firstChild)
    
    # 2. Xóa sạch màu cũ và thẻ bold cũ (để tránh trùng lặp)
    for old_c in rPr.getElementsByTagNameNS(W_NS, "color"): rPr.removeChild(old_c)
    for old_b in rPr.getElementsByTagNameNS(W_NS, "b"): rPr.removeChild(old_b)
    
    # 3. Thêm màu Xanh mới (Scientific Blue)
    color_node = doc.createElementNS(W_NS, "w:color")
    color_node.setAttributeNS(W_NS, "w:val", "0070C0") 
    rPr.appendChild(color_node)

    # 4. Thêm In Đậm mới
    rPr.appendChild(doc.createElementNS(W_NS, "w:b"))

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

# --- PROCESSORS ---

def process_mcq(q_blocks, doc):
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
            if target_opt and opt == target_opt: correct_char = lbls[idx][0]
            
            # Thay nhãn và tô màu
            t_nodes = opt.getElementsByTagNameNS(W_NS, "t")
            for t in t_nodes:
                if t.firstChild:
                    t.firstChild.nodeValue = re.sub(pat, lbls[idx], t.firstChild.nodeValue, 1)
                    # Gọi hàm Force Blue
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
                    # Gọi hàm Force Blue
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
                            # Tô xanh số câu (Câu 1, Câu 2...)
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
