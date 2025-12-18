"""
TRỘN ĐỀ WORD - THPT MINH ĐỨC (FIXED: createElementNS)
Sửa lỗi: 'Element' object has no attribute 'createElementNS'
Tính năng:
- Chia 3 phần: P1(18), P2(4), P3(6).
- Nhận diện đáp án: Gạch chân/Đỏ (P1, P2) & Thẻ <key> (P3).
- Xuất file CSV đáp án chuẩn mẫu.
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
    page_title="TNMix - Trộn Đề Minh Đức",
    page_icon="📝",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# ==================== 2. CSS GIAO DIỆN ====================
CUSTOM_CSS = """
<style>
    header {visibility: hidden;}
    .stApp > header {display: none;}
    .block-container {
        padding-top: 0rem !important;
        padding-bottom: 5rem !important;
        max-width: 100% !important;
    }
    [data-testid="stAppViewContainer"] {
        background: linear-gradient(to bottom right, #e0eafc, #cfdef3);
    }
    .header-wrapper {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding-top: 3rem;
        padding-bottom: 6rem;
        text-align: center;
        color: white;
        border-bottom-left-radius: 50px;
        border-bottom-right-radius: 50px;
        box-shadow: 0 10px 20px rgba(118, 75, 162, 0.3);
        margin-bottom: -80px;
    }
    .school-tag {
        background: rgba(255, 255, 255, 0.2);
        border: 1px solid rgba(255,255,255,0.5);
        padding: 5px 20px;
        border-radius: 30px;
        font-weight: 900;
        font-size: 1.1rem;
        text-transform: uppercase;
        display: inline-block;
        margin-bottom: 10px;
    }
    .main-h1 {
        font-size: 2.5rem;
        font-weight: 800;
        margin: 0;
        text-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    .sub-text {
        font-size: 0.95rem;
        opacity: 0.95;
        margin-top: 5px;
    }
    .info-gv {
        margin-top: 20px;
        background: white;
        color: #764ba2;
        padding: 8px 20px;
        border-radius: 15px;
        display: inline-block;
        font-weight: 700;
        box-shadow: 0 4px 10px rgba(0,0,0,0.1);
    }
    .main-card {
        background: white;
        border-radius: 30px;
        width: 90%;
        max-width: 600px;
        margin: 0 auto;
        padding: 30px;
        box-shadow: 0 20px 50px rgba(0,0,0,0.1);
        position: relative;
        z-index: 99;
    }
    .stButton > button {
        background: linear-gradient(90deg, #667eea, #764ba2);
        color: white;
        border: none;
        height: 50px;
        border-radius: 12px;
        font-weight: bold;
        width: 100%;
        margin-top: 15px;
        font-size: 1.1rem;
    }
    .footer {
        text-align: center;
        color: #888;
        margin-top: 40px;
        font-size: 0.8rem;
    }
</style>
"""

HEADER_HTML = """
<div class="header-wrapper">
<div class="school-tag">TRƯỜNG THPT MINH ĐỨC</div>
<div class="main-h1">TNMIX 2025</div>
<div class="sub-text">Cấu trúc 3 Phần: TN (18) - Đ/S (4) - TLN (6)</div>
<div class="info-gv">GV: Nguyễn Văn Hà • Zalo: 0913968302</div>
</div>
"""

# ==================== 3. LOGIC XỬ LÝ (CORE) ====================
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# --- HÀM TẠO NODE XML AN TOÀN ---
def create_header_xml(text, dom_document):
    """
    Tạo đoạn văn bản đậm.
    QUAN TRỌNG: Phải dùng dom_document (là Document) chứ không phải Element.
    """
    p = dom_document.createElementNS(W_NS, "w:p")
    r = dom_document.createElementNS(W_NS, "w:r")
    rPr = dom_document.createElementNS(W_NS, "w:rPr")
    b = dom_document.createElementNS(W_NS, "w:b")
    
    rPr.appendChild(b)
    r.appendChild(rPr)
    
    t = dom_document.createElementNS(W_NS, "w:t")
    t.appendChild(dom_document.createTextNode(text))
    
    r.appendChild(t)
    p.appendChild(r)
    return p

def get_text(block):
    texts = []
    for t in block.getElementsByTagNameNS(W_NS, "t"):
        if t.firstChild: texts.append(t.firstChild.nodeValue)
    return "".join(texts).strip()

def check_is_correct(run_node):
    """Kiểm tra gạch chân hoặc màu đỏ"""
    rPr_list = run_node.getElementsByTagNameNS(W_NS, "rPr")
    if not rPr_list: return False
    rPr = rPr_list[0]
    
    # Check Underline
    u = rPr.getElementsByTagNameNS(W_NS, "u")
    if u: return True
    
    # Check Color
    color = rPr.getElementsByTagNameNS(W_NS, "color")
    if color:
        val = color[0].getAttributeNS(W_NS, "val")
        if val and val.upper() in ["FF0000", "RED"]: return True
    return False

def style_run_blue_bold(run):
    """Tô xanh và in đậm label (A. B. C...)"""
    doc = run.ownerDocument # Lấy Document chủ sở hữu của node này
    
    rPr_list = run.getElementsByTagNameNS(W_NS, "rPr")
    if rPr_list: rPr = rPr_list[0]
    else:
        rPr = doc.createElementNS(W_NS, "w:rPr")
        run.insertBefore(rPr, run.firstChild)
    
    # Màu xanh
    color_list = rPr.getElementsByTagNameNS(W_NS, "color")
    if not color_list:
        color_el = doc.createElementNS(W_NS, "w:color")
        rPr.appendChild(color_el)
        color_el.setAttributeNS(W_NS, "w:val", "0000FF")
    
    # In đậm
    b_list = rPr.getElementsByTagNameNS(W_NS, "b")
    if not b_list:
        b_el = doc.createElementNS(W_NS, "w:b")
        rPr.appendChild(b_el)

def remove_formatting(run):
    """Xóa gạch chân/màu đỏ của đáp án"""
    rPr_list = run.getElementsByTagNameNS(W_NS, "rPr")
    if not rPr_list: return
    rPr = rPr_list[0]
    
    u_nodes = rPr.getElementsByTagNameNS(W_NS, "u")
    for u in u_nodes: rPr.removeChild(u)
    
    c_nodes = rPr.getElementsByTagNameNS(W_NS, "color")
    for c in c_nodes: rPr.removeChild(c)

def parse_all_questions(blocks):
    intro = []
    questions = []
    i = 0
    # Intro
    while i < len(blocks):
        txt = get_text(blocks[i])
        if re.match(r'^Câu\s*\d+', txt): break
        intro.append(blocks[i])
        i += 1
    # Questions
    while i < len(blocks):
        txt = get_text(blocks[i])
        if re.match(r'^Câu\s*\d+', txt):
            group = [blocks[i]]
            i += 1
            while i < len(blocks):
                t2 = get_text(blocks[i])
                if re.match(r'^Câu\s*\d+', t2) or "PHẦN" in t2.upper(): break
                group.append(blocks[i])
                i += 1
            questions.append(group)
        else:
            i += 1
    return intro, questions

# --- XỬ LÝ PHẦN 1 (MCQ) ---
def shuffle_options_mcq(q_block):
    pat = r'^\s*[A-D][\.\)]'
    indices = [x for x, b in enumerate(q_block) if re.match(pat, get_text(b))]
    correct_char = "X"
    
    if len(indices) >= 2:
        opts = [q_block[x] for x in indices]
        
        # Tìm đáp án đúng gốc
        target_opt = None
        for opt in opts:
            runs = opt.getElementsByTagNameNS(W_NS, "r")
            is_ans = False
            for r in runs:
                if check_is_correct(r):
                    is_ans = True
                    remove_formatting(r)
            if is_ans: target_opt = opt

        random.shuffle(opts)
        
        lbls = ["A.", "B.", "C.", "D."]
        for x_idx, opt in zip(indices, opts):
            q_block[x_idx] = opt
            t_nodes = opt.getElementsByTagNameNS(W_NS, "t")
            for t in t_nodes:
                if t.firstChild:
                    cur_lbl = lbls[indices.index(x_idx)] if indices.index(x_idx) < 4 else ""
                    val = t.firstChild.nodeValue
                    new_val = re.sub(pat, cur_lbl, val, 1)
                    t.firstChild.nodeValue = new_val
                    style_run_blue_bold(t.parentNode)
                    break
            
            if target_opt and opt == target_opt:
                correct_char = lbls[indices.index(x_idx)][0]

    return q_block, correct_char

# --- XỬ LÝ PHẦN 2 (Đúng/Sai) ---
def shuffle_options_tf(q_block):
    pat = r'^\s*[a-d][\.\)]'
    indices = [x for x, b in enumerate(q_block) if re.match(pat, get_text(b))]
    result_str_parts = []
    
    if len(indices) >= 2:
        opts = [q_block[x] for x in indices]
        
        status_map = {}
        for opt in opts:
            is_true = False
            runs = opt.getElementsByTagNameNS(W_NS, "r")
            for r in runs:
                if check_is_correct(r):
                    is_true = True
                    remove_formatting(r)
            status_map[opt] = "Đ" if is_true else "S"

        random.shuffle(opts)
        
        lbls = ["a)", "b)", "c)", "d)"]
        for x_idx, opt in zip(indices, opts):
            q_block[x_idx] = opt
            cur_lbl = lbls[indices.index(x_idx)] if indices.index(x_idx) < 4 else ""
            
            t_nodes = opt.getElementsByTagNameNS(W_NS, "t")
            for t in t_nodes:
                if t.firstChild:
                    t.firstChild.nodeValue = re.sub(pat, cur_lbl, t.firstChild.nodeValue, 1)
                    style_run_blue_bold(t.parentNode)
                    break
            
            result_str_parts.append(f"{cur_lbl[:-1]}{status_map[opt]}")
            
    return q_block, " - ".join(result_str_parts)

# --- XỬ LÝ PHẦN 3 (KEY) ---
def process_part_3(q_block):
    key_val = ""
    for b in q_block:
        t_nodes = b.getElementsByTagNameNS(W_NS, "t")
        for t in t_nodes:
            if t.firstChild:
                txt = t.firstChild.nodeValue
                m = re.search(r'<key=(.*?)>', txt)
                if m:
                    key_val = m.group(1)
                    new_txt = txt.replace(m.group(0), "")
                    t.firstChild.nodeValue = new_txt
    return q_block, key_val

def process_and_zip(file_bytes, num_copies):
    output_buffer = io.BytesIO()
    answer_key_data = []
    
    with zipfile.ZipFile(output_buffer, 'w', zipfile.ZIP_DEFLATED) as z_out:
        input_io = io.BytesIO(file_bytes)
        with zipfile.ZipFile(input_io, 'r') as z_in:
            xml_content = z_in.read("word/document.xml")
            
            for _ in range(num_copies):
                exam_code = str(random.randint(1001, 9999))
                # Tạo Document Object Model (DOM)
                dom = minidom.parseString(xml_content)
                # QUAN TRỌNG: Đây là đối tượng Document dùng để tạo thẻ mới
                
                body = dom.getElementsByTagNameNS(W_NS, "body")[0]
                
                blocks = [n for n in body.childNodes if n.localName in ['p', 'tbl']]
                intro, all_qs = parse_all_questions(blocks)
                
                # Chia phần
                p1_qs = all_qs[0:18]
                p2_qs = all_qs[18:22]
                p3_qs = all_qs[22:]
                
                current_key = [exam_code]
                
                # P1
                p1_final, p1_keys = [], []
                for q in p1_qs:
                    q_new, k = shuffle_options_mcq(q)
                    p1_final.append(q_new)
                    p1_keys.append(k)
                
                c_p1 = list(zip(p1_final, p1_keys))
                random.shuffle(c_p1)
                if c_p1: p1_final, p1_keys = zip(*c_p1)
                current_key.extend(p1_keys)
                
                # P2
                p2_final, p2_keys = [], []
                for q in p2_qs:
                    q_new, k = shuffle_options_tf(q)
                    p2_final.append(q_new)
                    p2_keys.append(k)
                
                c_p2 = list(zip(p2_final, p2_keys))
                random.shuffle(c_p2)
                if c_p2: p2_final, p2_keys = zip(*c_p2)
                current_key.extend(p2_keys)
                
                # P3
                p3_final, p3_keys = [], []
                for q in p3_qs:
                    q_new, k = process_part_3(q)
                    p3_final.append(q_new)
                    p3_keys.append(k)
                
                c_p3 = list(zip(p3_final, p3_keys))
                random.shuffle(c_p3)
                if c_p3: p3_final, p3_keys = zip(*c_p3)
                current_key.extend(p3_keys)
                
                answer_key_data.append(current_key)
                
                # GỘP DOC
                final_blocks = intro[:]
                
                # Header P1
                h1 = "PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn. Thí sinh trả lời từ câu 1 đến câu 18. Mỗi câu hỏi thí sinh chỉ chọn một phương án."
                # TRUYỀN 'dom' (Document) vào hàm tạo thẻ, KHÔNG truyền element
                final_blocks.append(create_header_xml(h1, dom)) 
                
                for i, q in enumerate(p1_final):
                    t_nodes = q[0].getElementsByTagNameNS(W_NS, "t")
                    for t in t_nodes:
                        if t.firstChild and re.match(r'^Câu\s*\d+', t.firstChild.nodeValue):
                            t.firstChild.nodeValue = re.sub(r'^Câu\s*\d+', f"Câu {i+1}", t.firstChild.nodeValue)
                            style_run_blue_bold(t.parentNode)
                            break
                    final_blocks.extend(q)
                    
                # Header P2
                h2 = "PHẦN II. Câu trắc nghiệm đúng sai. Thí sinh trả lời từ câu 1 đến câu 4. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai."
                final_blocks.append(create_header_xml(h2, dom))
                
                for i, q in enumerate(p2_final):
                    t_nodes = q[0].getElementsByTagNameNS(W_NS, "t")
                    for t in t_nodes:
                        if t.firstChild and re.match(r'^Câu\s*\d+', t.firstChild.nodeValue):
                            t.firstChild.nodeValue = re.sub(r'^Câu\s*\d+', f"Câu {i+1}", t.firstChild.nodeValue)
                            style_run_blue_bold(t.parentNode)
                            break
                    final_blocks.extend(q)
                    
                # Header P3
                h3 = "PHẦN III. Câu trắc nghiệm yêu cầu trả lời ngắn. Thí sinh trả lời từ câu 1 đến câu 6."
                final_blocks.append(create_header_xml(h3, dom))
                
                for i, q in enumerate(p3_final):
                    t_nodes = q[0].getElementsByTagNameNS(W_NS, "t")
                    for t in t_nodes:
                        if t.firstChild and re.match(r'^Câu\s*\d+', t.firstChild.nodeValue):
                            t.firstChild.nodeValue = re.sub(r'^Câu\s*\d+', f"Câu {i+1}", t.firstChild.nodeValue)
                            style_run_blue_bold(t.parentNode)
                            break
                    final_blocks.extend(q)
                
                # Mã đề
                p_code = create_header_xml(f"MÃ ĐỀ: {exam_code}", dom)
                final_blocks.insert(len(intro), p_code)

                # Rebuild Body
                for n in list(body.childNodes):
                    if n.localName in ['p', 'tbl']: body.removeChild(n)
                for b in final_blocks: body.appendChild(b)
                
                # Write
                new_xml = dom.toxml().encode('utf-8')
                doc_io = io.BytesIO()
                with zipfile.ZipFile(doc_io, 'w', zipfile.ZIP_DEFLATED) as z_d:
                    for item in z_in.infolist():
                        if item.filename == "word/document.xml":
                            z_d.writestr(item.filename, new_xml)
                        else:
                            z_d.writestr(item.filename, z_in.read(item.filename))
                z_out.writestr(f"De_MinhDuc_{exam_code}.docx", doc_io.getvalue())
    
    csv_io = io.StringIO()
    writer = csv.writer(csv_io)
    header = ["Mã đề"] + [str(i) for i in range(1,19)] + [f"Câu {i}" for i in range(1,5)] + [f"Câu {i}" for i in range(1,7)]
    writer.writerow(header)
    writer.writerows(answer_key_data)
    
    return output_buffer.getvalue(), csv_io.getvalue()

# ==================== 4. GIAO DIỆN CHÍNH ====================

st.markdown(CUSTOM_CSS, unsafe_allow_html=True)
st.markdown(HEADER_HTML, unsafe_allow_html=True)

st.markdown('<div class="main-card">', unsafe_allow_html=True)

cols = st.columns(2)
with cols[0]:
    st.markdown('<div style="text-align:center; color:#764ba2; font-weight:bold; border-bottom:3px solid #764ba2; padding-bottom:5px;">⚡ Trộn đề</div>', unsafe_allow_html=True)
with cols[1]:
    st.markdown('<div style="text-align:center; color:#ccc; font-weight:500;">📷 QR Code</div>', unsafe_allow_html=True)

st.write("")
uploaded_file = st.file_uploader("Kéo thả file .docx vào đây", type=['docx'])

if uploaded_file:
    st.success(f"✅ Đã nhận file: {uploaded_file.name}")
    st.info("Lưu ý: P1, P2 gạch chân đáp án đúng. P3 dùng thẻ <key=...>")

st.write("")
num = st.number_input("Số lượng đề cần tạo", 1, 50, 4)

if st.button("🚀 BẮT ĐẦU TRỘN"):
    if not uploaded_file:
        st.warning("Vui lòng chọn file!")
    else:
        try:
            with st.spinner("Đang xử lý..."):
                zip_data, csv_data = process_and_zip(uploaded_file.read(), num)
                
                st.success("Xong!")
                
                st.download_button(
                    "📥 Tải bộ đề (ZIP)",
                    zip_data,
                    "Bo_De_MinhDuc.zip",
                    "application/zip"
                )
                
                st.download_button(
                    "📊 Tải đáp án (CSV)",
                    csv_data.encode('utf-8-sig'),
                    "Dap_An.csv",
                    "text/csv"
                )
                st.balloons()
        except Exception as e:
            st.error(f"Lỗi: {e}")

st.markdown('</div>', unsafe_allow_html=True)
st.markdown('<div class="footer">© 2024 THPT Minh Đức</div>', unsafe_allow_html=True)
