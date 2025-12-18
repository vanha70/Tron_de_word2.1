"""
TRỘN ĐỀ WORD - THPT MINH ĐỨC (FINAL VERSION - CLEAN ANSWER)
Tính năng:
1. Thêm Header (Trường, Tổ, Tên thi, SBD) vào đầu đề.
2. Tự động XÓA dấu hiệu đáp án (Gạch chân/Đỏ/Key) trong file đề tạo ra.
3. Xuất file Excel đáp án riêng.
"""

import streamlit as st
import re
import random
import zipfile
import io
import csv
from xml.dom import minidom

# ==================== 1. CẤU HÌNH ====================
st.set_page_config(page_title="TNMix - Trộn Đề", page_icon="📝", layout="centered")

# ==================== 2. CSS GIAO DIỆN ====================
CUSTOM_CSS = """
<style>
    header, footer {visibility: hidden;}
    .stApp > header {display: none;}
    .block-container {padding-top: 2rem; max-width: 800px;}
    
    .header-wrapper {
        background: linear-gradient(135deg, #00b09b, #96c93d);
        padding: 2rem; border-radius: 20px; color: white; text-align: center;
        margin-bottom: 20px; box-shadow: 0 4px 15px rgba(0,0,0,0.1);
    }
    .main-title {font-size: 2rem; font-weight: bold; margin: 0;}
    .stButton>button {width: 100%; border-radius: 10px; font-weight: bold;}
    .success-box {padding: 15px; background: #d4edda; color: #155724; border-radius: 10px; margin-top: 10px;}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

HEADER_HTML = """
<div class="header-wrapper">
    <div style="font-size: 1.2rem; font-weight: bold;">TRƯỜNG THPT NGUYỄN HUỆ - TỔ TOÁN</div>
    <div class="main-title">TRỘN ĐỀ KIỂM TRA</div>
    <div>Chuẩn cấu trúc 2025: 3 Phần (TN - Đ/S - TLN)</div>
</div>
"""
st.markdown(HEADER_HTML, unsafe_allow_html=True)

# ==================== 3. XỬ LÝ XML (CORE) ====================
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

def create_element(doc, tag):
    return doc.createElementNS(W_NS, tag)

def create_paragraph(doc, text, align="left", bold=False, italic=False):
    """Tạo một đoạn văn bản (dùng cho Header đề thi)"""
    p = create_element(doc, "w:p")
    pPr = create_element(doc, "w:pPr")
    
    # Căn lề
    jc = create_element(doc, "w:jc")
    jc.setAttributeNS(W_NS, "w:val", align)
    pPr.appendChild(jc)
    p.appendChild(pPr)
    
    r = create_element(doc, "w:r")
    rPr = create_element(doc, "w:rPr")
    if bold: rPr.appendChild(create_element(doc, "w:b"))
    if italic: rPr.appendChild(create_element(doc, "w:i"))
    r.appendChild(rPr)
    
    t = create_element(doc, "w:t")
    t.appendChild(doc.createTextNode(text))
    r.appendChild(t)
    p.appendChild(r)
    return p

def add_school_header(doc, body, exam_code):
    """Chèn Header Trường/Lớp/SBD vào đầu file"""
    # Tạo các dòng Header
    header_nodes = []
    header_nodes.append(create_paragraph(doc, "TRƯỜNG THPT NGUYỄN HUỆ", "center", bold=True))
    header_nodes.append(create_paragraph(doc, "TỔ TOÁN", "center", bold=True))
    header_nodes.append(create_paragraph(doc, "ĐỀ KIỂM TRA GIỮA KỲ I", "center", bold=True))
    header_nodes.append(create_paragraph(doc, f"MÃ ĐỀ: {exam_code}", "right", bold=True))
    header_nodes.append(create_paragraph(doc, "Họ tên thí sinh: ........................................................... SBD: ....................", "left"))
    header_nodes.append(create_paragraph(doc, "", "left")) # Dòng trống
    
    # Chèn vào đầu body (Lưu ý thứ tự ngược để đẩy xuống)
    if body.hasChildNodes():
        first_child = body.firstChild
        for node in reversed(header_nodes):
            body.insertBefore(node, first_child)
    else:
        for node in header_nodes:
            body.appendChild(node)

def get_text(block):
    texts = []
    for t in block.getElementsByTagNameNS(W_NS, "t"):
        if t.firstChild: texts.append(t.firstChild.nodeValue)
    return "".join(texts).strip()

def check_is_correct(run_node):
    """Kiểm tra đáp án đúng (Gạch chân hoặc Đỏ)"""
    rPr_list = run_node.getElementsByTagNameNS(W_NS, "rPr")
    if not rPr_list: return False
    rPr = rPr_list[0]
    
    # Check Underline
    if rPr.getElementsByTagNameNS(W_NS, "u"): return True
    
    # Check Color (Red)
    color = rPr.getElementsByTagNameNS(W_NS, "color")
    if color:
        val = color[0].getAttributeNS(W_NS, "val")
        if val and val.upper() in ["FF0000", "RED"]: return True
    return False

def remove_answer_signal(run_node):
    """XÓA dấu hiệu đáp án (gạch chân/màu) trong node"""
    rPr_list = run_node.getElementsByTagNameNS(W_NS, "rPr")
    if not rPr_list: return
    rPr = rPr_list[0]
    
    # Xóa thẻ u (gạch chân)
    for u in rPr.getElementsByTagNameNS(W_NS, "u"):
        rPr.removeChild(u)
    
    # Xóa thẻ color (màu sắc)
    for c in rPr.getElementsByTagNameNS(W_NS, "color"):
        rPr.removeChild(c)

def style_label(run_node, doc):
    """Tô xanh đậm cho nhãn A. B. C."""
    rPr_list = run_node.getElementsByTagNameNS(W_NS, "rPr")
    if rPr_list: rPr = rPr_list[0]
    else:
        rPr = doc.createElementNS(W_NS, "w:rPr")
        run_node.insertBefore(rPr, run_node.firstChild)
    
    # Màu xanh
    if not rPr.getElementsByTagNameNS(W_NS, "color"):
        c = doc.createElementNS(W_NS, "w:color")
        c.setAttributeNS(W_NS, "w:val", "0000FF")
        rPr.appendChild(c)
    # In đậm
    if not rPr.getElementsByTagNameNS(W_NS, "b"):
        rPr.appendChild(doc.createElementNS(W_NS, "w:b"))

def parse_questions(blocks):
    intro, questions = [], []
    i = 0
    # Intro
    while i < len(blocks):
        txt = get_text(blocks[i])
        if re.match(r'^Câu\s*\d+', txt): break
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
    return intro, questions

# --- TRỘN PHẦN 1 (MCQ) ---
def process_mcq(q_blocks, doc):
    pat = r'^\s*[A-D][\.\)]'
    indices = [k for k, b in enumerate(q_blocks) if re.match(pat, get_text(b))]
    correct_char = "X" # Mặc định nếu ko tìm thấy
    
    if len(indices) >= 2:
        opts = [q_blocks[k] for k in indices]
        
        # 1. Tìm đáp án đúng trong options gốc
        target_opt = None
        for opt in opts:
            runs = opt.getElementsByTagNameNS(W_NS, "r")
            is_correct = False
            for r in runs:
                if check_is_correct(r):
                    is_correct = True
                # QUAN TRỌNG: Luôn xóa định dạng đáp án (gạch chân/đỏ)
                remove_answer_signal(r)
            if is_correct: target_opt = opt
            
        # 2. Trộn
        random.shuffle(opts)
        
        # 3. Gán lại và đánh nhãn
        lbls = ["A.", "B.", "C.", "D."]
        for idx_shuffled, opt in enumerate(opts):
            real_idx = indices[idx_shuffled]
            q_blocks[real_idx] = opt
            
            # Nếu option này là đáp án đúng lúc nãy -> Lưu lại ký tự mới (A,B..)
            if target_opt and opt == target_opt:
                correct_char = lbls[idx_shuffled][0]
            
            # Thay đổi text nhãn (A. -> B. ...)
            t_nodes = opt.getElementsByTagNameNS(W_NS, "t")
            for t in t_nodes:
                if t.firstChild:
                    # Thay thế ký tự đầu tiên khớp pattern
                    t.firstChild.nodeValue = re.sub(pat, lbls[idx_shuffled], t.firstChild.nodeValue, 1)
                    style_label(t.parentNode, doc)
                    break
                    
    return q_blocks, correct_char

# --- TRỘN PHẦN 2 (TRUE/FALSE) ---
def process_tf(q_blocks, doc):
    pat = r'^\s*[a-d][\.\)]'
    indices = [k for k, b in enumerate(q_blocks) if re.match(pat, get_text(b))]
    # Chuỗi kết quả: a)Đ - b)S ...
    res_list = []
    
    if len(indices) >= 2:
        opts = [q_blocks[k] for k in indices]
        
        # Map trạng thái Đ/S
        status_map = {} 
        for opt in opts:
            is_true = False
            runs = opt.getElementsByTagNameNS(W_NS, "r")
            for r in runs:
                if check_is_correct(r): is_true = True
                remove_answer_signal(r) # Xóa dấu hiệu
            status_map[opt] = "Đ" if is_true else "S"
            
        random.shuffle(opts)
        
        lbls = ["a)", "b)", "c)", "d)"]
        for idx_shuffled, opt in enumerate(opts):
            real_idx = indices[idx_shuffled]
            q_blocks[real_idx] = opt
            
            # Ghi lại đáp án theo nhãn mới
            curr_lbl = lbls[idx_shuffled]
            res_list.append(f"{curr_lbl[:-1]}{status_map[opt]}") # a) -> aĐ
            
            # Thay đổi text
            t_nodes = opt.getElementsByTagNameNS(W_NS, "t")
            for t in t_nodes:
                if t.firstChild:
                    t.firstChild.nodeValue = re.sub(pat, curr_lbl, t.firstChild.nodeValue, 1)
                    style_label(t.parentNode, doc)
                    break
    
    return q_blocks, " - ".join(res_list)

# --- TRỘN PHẦN 3 (SHORT ANSWER) ---
def process_short(q_blocks):
    key_val = ""
    for b in q_blocks:
        t_nodes = b.getElementsByTagNameNS(W_NS, "t")
        for t in t_nodes:
            if t.firstChild:
                txt = t.firstChild.nodeValue
                # Tìm thẻ <key=...>
                m = re.search(r'<key=(.*?)>', txt)
                if m:
                    key_val = m.group(1)
                    # XÓA KEY KHỎI ĐỀ THI
                    t.firstChild.nodeValue = txt.replace(m.group(0), "") 
    return q_blocks, key_val

# --- MAIN PROCESS ---
def run_processing(file_bytes, num_copies):
    out_zip = io.BytesIO()
    csv_data = []
    
    with zipfile.ZipFile(out_zip, 'w', zipfile.ZIP_DEFLATED) as z_out:
        in_io = io.BytesIO(file_bytes)
        with zipfile.ZipFile(in_io, 'r') as z_in:
            xml = z_in.read("word/document.xml")
            
            for _ in range(num_copies):
                exam_code = str(random.randint(1001, 9999))
                dom = minidom.parseString(xml)
                doc = dom.documentElement
                body = dom.getElementsByTagNameNS(W_NS, "body")[0]
                
                blocks = [n for n in body.childNodes if n.localName in ['p', 'tbl']]
                intro, all_qs = parse_questions(blocks)
                
                # Chia 3 phần
                p1_qs = all_qs[0:18]
                p2_qs = all_qs[18:22]
                p3_qs = all_qs[22:]
                
                key_row = [exam_code]
                
                # Xử lý P1
                p1_fin, k1 = [], []
                for q in p1_qs:
                    q_new, k = process_mcq(q, dom)
                    p1_fin.append(q_new)
                    k1.append(k)
                # Shuffle câu hỏi P1
                c1 = list(zip(p1_fin, k1))
                random.shuffle(c1)
                if c1: p1_fin, k1 = zip(*c1)
                key_row.extend(k1)
                
                # Xử lý P2
                p2_fin, k2 = [], []
                for q in p2_qs:
                    q_new, k = process_tf(q, dom)
                    p2_fin.append(q_new)
                    k2.append(k)
                c2 = list(zip(p2_fin, k2))
                random.shuffle(c2)
                if c2: p2_fin, k2 = zip(*c2)
                key_row.extend(k2)
                
                # Xử lý P3
                p3_fin, k3 = [], []
                for q in p3_qs:
                    q_new, k = process_short(q)
                    p3_fin.append(q_new)
                    k3.append(k)
                c3 = list(zip(p3_fin, k3))
                random.shuffle(c3)
                if c3: p3_fin, k3 = zip(*c3)
                key_row.extend(k3)
                
                csv_data.append(key_row)
                
                # Rebuild DOC
                final_blocks = [] # Intro sẽ được xử lý trong Header function
                
                # Header các phần
                final_blocks.append(create_paragraph(dom, "PHẦN I. Trắc nghiệm nhiều lựa chọn (18 câu)", bold=True))
                for i, q in enumerate(p1_fin):
                    # Renumber
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
                
                # Clear body & append new blocks
                for n in list(body.childNodes):
                    if n.localName in ['p', 'tbl']: body.removeChild(n)
                for b in final_blocks: body.appendChild(b)
                
                # THÊM HEADER TRƯỜNG VÀO ĐẦU
                add_school_header(dom, body, exam_code)
                
                # Save
                new_xml = dom.toxml().encode('utf-8')
                doc_io = io.BytesIO()
                with zipfile.ZipFile(doc_io, 'w', zipfile.ZIP_DEFLATED) as z_d:
                    for item in z_in.infolist():
                        if item.filename == "word/document.xml":
                            z_d.writestr(item.filename, new_xml)
                        else:
                            z_d.writestr(item.filename, z_in.read(item.filename))
                z_out.writestr(f"De_{exam_code}.docx", doc_io.getvalue())
                
    # CSV
    csv_io = io.StringIO()
    writer = csv.writer(csv_io)
    head = ["Mã đề"] + [str(i) for i in range(1,19)] + [f"II_C{i}" for i in range(1,5)] + [f"III_C{i}" for i in range(1,7)]
    writer.writerow(head)
    writer.writerows(csv_data)
    
    return out_zip.getvalue(), csv_io.getvalue()

# ==================== 4. UI LOGIC ====================
st.markdown('<div class="main-card">', unsafe_allow_html=True)
uploaded_file = st.file_uploader("Chọn file .docx", type=['docx'])
num = st.number_input("Số lượng đề", 1, 50, 4)

if st.button("🚀 TRỘN ĐỀ NGAY"):
    if not uploaded_file:
        st.warning("Vui lòng chọn file!")
    else:
        try:
            with st.spinner("Đang xử lý..."):
                z_data, c_data = run_processing(uploaded_file.read(), num)
                st.markdown('<div class="success-box">✅ Trộn thành công!</div>', unsafe_allow_html=True)
                st.download_button("📥 Tải Bộ Đề (ZIP)", z_data, "BoDe_MinhDuc.zip", "application/zip")
                st.download_button("📊 Tải Đáp Án (CSV)", c_data.encode('utf-8-sig'), "DapAn.csv", "text/csv")
        except Exception as e:
            st.error(f"Lỗi: {e}")
st.markdown('</div>', unsafe_allow_html=True)
