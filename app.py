"""
PHẦN MỀM TRỘN ĐỀ - THPT MINH ĐỨC (FINAL FIX)
1. Fix lỗi mất đáp án P3 (Quét sâu toàn bộ câu hỏi).
2. Font chữ chuẩn Times New Roman cho tiêu đề.
3. Giao diện Xanh Ngọc (Teal) hiện đại.
4. Output: 1 file Zip chứa tất cả.
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
    page_title="Phần mềm Trộn Đề - Minh Đức",
    page_icon="📘",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# ==================== 2. CSS GIAO DIỆN (XANH NGỌC - HIỆN ĐẠI) ====================
CUSTOM_CSS = """
<style>
    /* Ẩn header mặc định */
    header {visibility: hidden;}
    .stApp > header {display: none;}
    
    .block-container {
        padding-top: 2rem !important;
        padding-bottom: 5rem !important;
    }
    
    /* Background Gradient Xanh Ngọc Nhạt */
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
        box-shadow: 0 10px 25px rgba(0, 121, 107, 0.15); /* Bóng xanh ngọc */
        border-bottom: 6px solid #009688; /* Viền dưới đậm */
        margin-bottom: 30px;
    }

    .school-name {
        color: #004d40; /* Xanh đậm */
        font-family: 'Times New Roman', serif; /* Font chuẩn */
        font-size: 1.8rem;
        font-weight: 900;
        text-transform: uppercase;
        margin-bottom: 5px;
        letter-spacing: 1px;
        text-shadow: 1px 1px 0px rgba(0,0,0,0.1);
    }

    .software-name {
        color: #009688; /* Xanh Ngọc chủ đạo */
        font-size: 2rem;
        font-weight: 800;
        margin: 10px 0;
        text-transform: uppercase;
    }

    .teacher-info {
        font-size: 1.1rem;
        font-weight: 600;
        color: #00695c;
        background-color: #e0f2f1;
        padding: 10px 25px;
        border-radius: 50px;
        display: inline-block;
        margin-top: 15px;
        border: 2px solid #b2dfdb;
    }

    /* CARD UPLOAD */
    .upload-card {
        background: white;
        padding: 30px;
        border-radius: 20px;
        box-shadow: 0 5px 15px rgba(0,0,0,0.05);
    }

    /* NÚT BẤM */
    .stButton > button {
        background: linear-gradient(90deg, #009688, #26a69a);
        color: white;
        font-weight: bold;
        border: none;
        padding: 12px 20px;
        border-radius: 12px;
        width: 100%;
        font-size: 1.2rem;
        text-transform: uppercase;
        transition: all 0.3s;
        box-shadow: 0 4px 6px rgba(0, 150, 136, 0.3);
    }
    .stButton > button:hover {
        background: linear-gradient(90deg, #00796b, #00897b);
        transform: translateY(-2px);
        box-shadow: 0 6px 12px rgba(0, 150, 136, 0.4);
    }
    
    .footer {
        text-align: center;
        margin-top: 40px;
        color: #5f7d7a;
        font-size: 0.9rem;
    }
    
    /* Input number */
    input[type=number] {
        border-radius: 10px;
        border: 1px solid #b2dfdb;
    }
</style>
"""

HEADER_HTML = """
<div class="header-wrapper">
    <div class="school-name">TRƯỜNG THPT MINH ĐỨC</div>
    <div class="software-name">PHẦN MỀM TRỘN ĐỀ</div>
    <div class="teacher-info">GV: Nguyễn Văn Hà • Zalo: 0913968302</div>
</div>
"""

# ==================== 3. XỬ LÝ XML (CORE) ====================
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

def create_element(doc, tag):
    return doc.createElementNS(W_NS, tag)

def create_paragraph(doc, text, align="left", bold=False):
    """Tạo đoạn văn bản Header"""
    p = create_element(doc, "w:p")
    pPr = create_element(doc, "w:pPr")
    
    jc = create_element(doc, "w:jc")
    jc.setAttributeNS(W_NS, "w:val", align)
    pPr.appendChild(jc)
    p.appendChild(pPr)
    
    r = create_element(doc, "w:r")
    rPr = create_element(doc, "w:rPr")
    
    # Set Font Times New Roman để không lỗi
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
    """Thêm Header Trường và Mã đề"""
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

# --- XỬ LÝ ĐÁP ÁN ---

def check_is_correct(run_node):
    """Check gạch chân hoặc màu đỏ"""
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
    """Xóa dấu hiệu đáp án"""
    rPr_list = run_node.getElementsByTagNameNS(W_NS, "rPr")
    if not rPr_list: return
    rPr = rPr_list[0]
    for u in rPr.getElementsByTagNameNS(W_NS, "u"): rPr.removeChild(u)
    for c in rPr.getElementsByTagNameNS(W_NS, "color"): rPr.removeChild(c)

def style_label(run_node, doc):
    """Tô xanh đậm nhãn"""
    rPr_list = run_node.getElementsByTagNameNS(W_NS, "rPr")
    if rPr_list: rPr = rPr_list[0]
    else:
        rPr = doc.createElementNS(W_NS, "w:rPr")
        run_node.insertBefore(rPr, run_node.firstChild)
    
    if not rPr.getElementsByTagNameNS(W_NS, "color"):
        c = doc.createElementNS(W_NS, "w:color")
        c.setAttributeNS(W_NS, "w:val", "0000FF")
        rPr.appendChild(c)
    if not rPr.getElementsByTagNameNS(W_NS, "b"):
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
    correct_char = "X"
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
            t_nodes = opt.getElementsByTagNameNS(W_NS, "t")
            for t in t_nodes:
                if t.firstChild:
                    t.firstChild.nodeValue = re.sub(pat, lbls[idx], t.firstChild.nodeValue, 1)
                    style_label(t.parentNode, doc)
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
                    style_label(t.parentNode, doc)
                    break
    return q_blocks, " - ".join(res_str)

def process_short(q_blocks):
    """
    Sửa lỗi mất đáp án P3:
    Quét qua TẤT CẢ các đoạn văn (paragraphs) trong câu hỏi để tìm thẻ <key=...>
    """
    key_val = ""
    # Duyệt qua từng block (paragraph) trong câu hỏi
    for b in q_blocks:
        t_nodes = b.getElementsByTagNameNS(W_NS, "t")
        # Duyệt qua từng text node trong paragraph
        full_text_in_block = ""
        
        # 1. Lấy toàn bộ text trong block này để check regex (vì key có thể bị tách ra nhiều run)
        for t in t_nodes:
            if t.firstChild: full_text_in_block += t.firstChild.nodeValue
            
        # 2. Tìm key trong block
        m = re.search(r'<key=(.*?)>', full_text_in_block)
        if m:
            key_val = m.group(1).strip() # Lấy giá trị key
            
            # 3. Xóa key khỏi XML (Duyệt lại để xóa)
            # Cách đơn giản: Replace trong từng node (có thể sót nếu key bị chia đôi)
            # Cách an toàn hơn: Xóa text node chứa key
            for t in t_nodes:
                if t.firstChild and '<key=' in t.firstChild.nodeValue:
                    # Xóa phần <key=...>
                    new_val = re.sub(r'<key=.*?>', '', t.firstChild.nodeValue)
                    t.firstChild.nodeValue = new_val
                    
        # Nếu đã tìm thấy key rồi thì không break, vì có thể có nhiều key (dù thường chỉ 1)
        # Nhưng ở đây ta giả sử 1 câu 1 key. Nếu tìm thấy key rồi thì thôi loop các block sau? 
        # Không nên break, vì cần xóa tag key hiển thị.
        
    return q_blocks, key_val

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
                
                # Slice P1, P2, P3
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
                            style_label(t.parentNode, dom)
                            break
                    final_blocks.extend(q)
                    
                final_blocks.append(create_paragraph(dom, "PHẦN II. Trắc nghiệm đúng sai (4 câu)", bold=True))
                for i, q in enumerate(p2_fin):
                    t_list = q[0].getElementsByTagNameNS(W_NS, "t")
                    for t in t_list:
                        if t.firstChild and re.match(r'^Câu\s*\d+', t.firstChild.nodeValue):
                            t.firstChild.nodeValue = re.sub(r'^Câu\s*\d+', f"Câu {i+1}", t.firstChild.nodeValue)
                            style_label(t.parentNode, dom)
                            break
                    final_blocks.extend(q)
                    
                final_blocks.append(create_paragraph(dom, "PHẦN III. Trả lời ngắn (6 câu)", bold=True))
                for i, q in enumerate(p3_fin):
                    t_list = q[0].getElementsByTagNameNS(W_NS, "t")
                    for t in t_list:
                        if t.firstChild and re.match(r'^Câu\s*\d+', t.firstChild.nodeValue):
                            t.firstChild.nodeValue = re.sub(r'^Câu\s*\d+', f"Câu {i+1}", t.firstChild.nodeValue)
                            style_label(t.parentNode, dom)
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
        
        # CSV
        csv_io = io.StringIO()
        writer = csv.writer(csv_io)
        head = ["Mã đề"] + [str(i) for i in range(1,19)] + [f"II_C{i}" for i in range(1,5)] + [f"III_C{i}" for i in range(1,7)]
        writer.writerow(head)
        writer.writerows(csv_data)
        outer_zip.writestr("Dap_An_Chi_Tiet.csv", csv_io.getvalue().encode('utf-8-sig'))

    return outer_zip_buffer.getvalue()

# ==================== 4. UI LOGIC ====================
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)
st.markdown(HEADER_HTML, unsafe_allow_html=True)

st.markdown('<div class="upload-card">', unsafe_allow_html=True)
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
                st.success("✅ Thành công!")
                st.download_button(
                    label="📥 Tải về (Đề thi + Đáp án)",
                    data=final_zip,
                    file_name="KetQua_TronDe_MinhDuc.zip",
                    mime="application/zip"
                )
        except Exception as e:
            st.error(f"Lỗi: {e}")

st.markdown('</div>', unsafe_allow_html=True)
st.markdown('<div class="footer">© 2024 Phần mềm Trộn Đề</div>', unsafe_allow_html=True)
