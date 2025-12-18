"""
PHẦN MỀM TRỘN ĐỀ - TNMic (REWRITE CORE)
Mục tiêu: Khắc phục triệt để lỗi trùng lặp nội dung và đồng bộ màu sắc.
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
    page_title="TNMic - Trộn Đề Chuyên Nghiệp",
    page_icon="🧬",
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
        background-size: 20px 20px;
        font-family: 'Segoe UI', sans-serif;
    }

    .header-wrapper {
        background: #ffffff;
        padding: 2rem;
        border-radius: 15px;
        text-align: center;
        box-shadow: 0 4px 20px rgba(0, 150, 136, 0.15);
        border-top: 5px solid #0d9488;
        margin-bottom: 25px;
    }

    .school-name {
        color: #115e59;
        font-family: 'Times New Roman', serif;
        font-size: 1.8rem;
        font-weight: 900;
        text-transform: uppercase;
        margin-bottom: 10px;
    }

    .app-name {
        background: linear-gradient(90deg, #0d9488, #14b8a6);
        color: white;
        padding: 5px 25px;
        border-radius: 20px;
        font-weight: 700;
        font-size: 1.2rem;
        display: inline-block;
        box-shadow: 0 2px 10px rgba(13, 148, 136, 0.3);
    }

    .info-tag {
        margin-top: 15px;
        font-weight: 600;
        color: #0f766e;
        background: #ccfbf1;
        padding: 8px 15px;
        border-radius: 8px;
        display: inline-block;
    }

    .main-card {
        background: white;
        padding: 25px;
        border-radius: 15px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.05);
    }

    .stButton > button {
        background: #0d9488;
        color: white;
        font-weight: bold;
        border-radius: 8px;
        height: 50px;
        font-size: 1.1rem;
        width: 100%;
        border: none;
        transition: 0.3s;
    }
    .stButton > button:hover {
        background: #0f766e;
        box-shadow: 0 4px 12px rgba(13, 148, 136, 0.4);
    }
    
    .footer { text-align: center; color: #888; margin-top: 30px; font-size: 0.8rem; }
</style>
"""

HEADER_HTML = """
<div class="header-wrapper">
    <div class="school-name">TRƯỜNG THPT MINH ĐỨC</div>
    <div class="app-name">PHẦN MỀM TRỘN ĐỀ</div><br>
    <div class="info-tag">GV: Nguyễn Văn Hà • Zalo: 0913968302</div>
</div>
"""

# ==================== 2. HÀM XỬ LÝ XML CỐT LÕI (CORE) ====================
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

def get_text(node):
    """Lấy toàn bộ text trong một node (gộp tất cả các thẻ con)"""
    texts = []
    for t in node.getElementsByTagNameNS(W_NS, "t"):
        if t.firstChild: texts.append(t.firstChild.nodeValue)
    return "".join(texts)

def create_run_with_style(doc, text, color_hex="0070C0", bold=True):
    """Tạo một thẻ Run (<w:r>) mới với định dạng chuẩn"""
    r = doc.createElementNS(W_NS, "w:r")
    rPr = doc.createElementNS(W_NS, "w:rPr")
    
    # 1. Màu sắc
    color = doc.createElementNS(W_NS, "w:color")
    color.setAttributeNS(W_NS, "w:val", color_hex)
    rPr.appendChild(color)
    
    # 2. In đậm
    if bold:
        rPr.appendChild(doc.createElementNS(W_NS, "w:b"))
        
    # 3. Font (Times New Roman)
    rFonts = doc.createElementNS(W_NS, "w:rFonts")
    rFonts.setAttributeNS(W_NS, "w:ascii", "Times New Roman")
    rFonts.setAttributeNS(W_NS, "w:hAnsi", "Times New Roman")
    rPr.appendChild(rFonts)
    
    r.appendChild(rPr)
    
    # 4. Nội dung text
    t = doc.createElementNS(W_NS, "w:t")
    t.setAttribute("xml:space", "preserve") # Giữ khoảng trắng
    t.appendChild(doc.createTextNode(text))
    r.appendChild(t)
    
    return r

def aggressive_replace_label(paragraph, new_label_text, doc, pattern_regex):
    """
    Hàm thay thế 'ăn mòn' (Consuming Replace):
    1. Tìm chuỗi khớp pattern (VD: 'A.')
    2. Xóa chính xác số ký tự đó khỏi các thẻ <w:t> đầu tiên.
    3. Chèn thẻ <w:r> mới vào đầu.
    """
    full_text = get_text(paragraph)
    match = re.match(pattern_regex, full_text)
    
    if match:
        chars_to_remove = len(match.group(0)) # Số ký tự cần xóa (VD: 'A.' là 2)
        
        # Duyệt qua các thẻ <w:t> để xóa dần ký tự
        t_nodes = paragraph.getElementsByTagNameNS(W_NS, "t")
        for t in t_nodes:
            if not t.firstChild: continue
            
            current_val = t.firstChild.nodeValue
            len_val = len(current_val)
            
            if chars_to_remove > 0:
                if len_val <= chars_to_remove:
                    # Nếu node này ngắn hơn hoặc bằng số ký tự cần xóa -> Xóa sạch node này
                    t.firstChild.nodeValue = ""
                    chars_to_remove -= len_val
                else:
                    # Nếu node này dài hơn -> Cắt phần đầu, giữ phần đuôi
                    t.firstChild.nodeValue = current_val[chars_to_remove:]
                    chars_to_remove = 0
            
            if chars_to_remove == 0:
                break
    
    # Tạo Run mới chứa nhãn mới (VD: "B. ")
    new_run = create_run_with_style(doc, new_label_text + " ", "0070C0", True)
    
    # Chèn vào đầu đoạn văn
    if paragraph.hasChildNodes():
        paragraph.insertBefore(new_run, paragraph.firstChild)
    else:
        paragraph.appendChild(new_run)

# ==================== 3. LOGIC XỬ LÝ ĐỀ THI ====================

def parse_docx_xml(dom):
    """Tách câu hỏi từ DOM"""
    body = dom.getElementsByTagNameNS(W_NS, "body")[0]
    # Lấy tất cả paragraph và table
    all_nodes = [n for n in body.childNodes if n.localName in ['p', 'tbl']]
    
    intro = []
    questions = []
    
    # 1. Tách Intro
    idx = 0
    while idx < len(all_nodes):
        txt = get_text(all_nodes[idx])
        if re.match(r'^Câu\s*\d+', txt): # Bắt đầu câu 1
            break
        intro.append(all_nodes[idx])
        idx += 1
        
    # 2. Tách Câu hỏi
    while idx < len(all_nodes):
        txt = get_text(all_nodes[idx])
        if re.match(r'^Câu\s*\d+', txt):
            current_q = [all_nodes[idx]] # Dòng chứa "Câu X..."
            idx += 1
            # Lấy các dòng tiếp theo cho đến khi gặp câu mới hoặc hết
            while idx < len(all_nodes):
                sub_txt = get_text(all_nodes[idx])
                if re.match(r'^Câu\s*\d+', sub_txt) or "PHẦN" in sub_txt.upper():
                    break
                current_q.append(all_nodes[idx])
                idx += 1
            questions.append(current_q)
        else:
            idx += 1
            
    return intro, questions, body

def process_part1(questions, doc):
    """Xử lý Phần 1: Trộn A,B,C,D"""
    processed_qs = []
    keys = []
    
    for q_block in questions:
        # Tìm các dòng chứa đáp án A., B., ...
        opt_indices = [i for i, node in enumerate(q_block) if re.match(r'^\s*[A-D][\.\)]', get_text(node))]
        
        correct_char = "X" # Mặc định
        
        if len(opt_indices) >= 2:
            # Tách riêng các dòng đáp án để trộn
            opts = [q_block[i] for i in opt_indices]
            
            # Tìm đáp án đúng gốc (Gạch chân/Đỏ)
            target_opt_node = None
            for opt in opts:
                runs = opt.getElementsByTagNameNS(W_NS, "r")
                is_correct = False
                for r in runs:
                    rPr = r.getElementsByTagNameNS(W_NS, "rPr")
                    if rPr:
                        # Check Underline or Red
                        if rPr[0].getElementsByTagNameNS(W_NS, "u") or \
                           (rPr[0].getElementsByTagNameNS(W_NS, "color") and \
                            rPr[0].getElementsByTagNameNS(W_NS, "color")[0].getAttributeNS(W_NS, "val") in ["FF0000", "RED"]):
                            is_correct = True
                        
                        # XÓA ĐỊNH DẠNG ĐÁP ÁN (QUAN TRỌNG)
                        for tag in ["u", "color", "b"]: # Xóa cả bold cũ cho sạch
                            for n in rPr[0].getElementsByTagNameNS(W_NS, tag):
                                rPr[0].removeChild(n)
                if is_correct:
                    target_opt_node = opt

            # Trộn
            random.shuffle(opts)
            
            labels = ["A.", "B.", "C.", "D."]
            
            # Gán lại vào vị trí cũ
            for i, original_idx in enumerate(opt_indices):
                q_block[original_idx] = opts[i] # Đặt dòng đã trộn vào vị trí
                
                # Check đáp án
                if target_opt_node and opts[i] == target_opt_node:
                    correct_char = labels[i][0]
                
                # THAY THẾ NHÃN CŨ BẰNG NHÃN MỚI (FIX LỖI TRÙNG)
                aggressive_replace_label(opts[i], labels[i], doc, r'^\s*[A-D][\.\)]')

        keys.append(correct_char)
        processed_qs.append(q_block)
        
    return processed_qs, keys

def process_part2(questions, doc):
    """Xử lý Phần 2: Trộn a), b), c), d)"""
    processed_qs = []
    keys = [] # Lưu chuỗi a)Đ...
    
    for q_block in questions:
        opt_indices = [i for i, node in enumerate(q_block) if re.match(r'^\s*[a-d][\.\)]', get_text(node))]
        result_str = []
        
        if len(opt_indices) >= 2:
            opts = [q_block[i] for i in opt_indices]
            
            # Map trạng thái Đ/S
            status_map = {}
            for opt in opts:
                is_true = False
                runs = opt.getElementsByTagNameNS(W_NS, "r")
                for r in runs:
                    rPr = r.getElementsByTagNameNS(W_NS, "rPr")
                    if rPr:
                        if rPr[0].getElementsByTagNameNS(W_NS, "u") or \
                           (rPr[0].getElementsByTagNameNS(W_NS, "color") and \
                            rPr[0].getElementsByTagNameNS(W_NS, "color")[0].getAttributeNS(W_NS, "val") in ["FF0000", "RED"]):
                            is_true = True
                        # Xóa định dạng
                        for tag in ["u", "color", "b"]:
                            for n in rPr[0].getElementsByTagNameNS(W_NS, tag): rPr[0].removeChild(n)
                status_map[opt] = "Đ" if is_true else "S"
            
            random.shuffle(opts)
            labels = ["a)", "b)", "c)", "d)"]
            
            for i, original_idx in enumerate(opt_indices):
                q_block[original_idx] = opts[i]
                result_str.append(f"{labels[i][:-1]}{status_map[opts[i]]}") # a)Đ
                
                # Thay nhãn
                aggressive_replace_label(opts[i], labels[i], doc, r'^\s*[a-d][\.\)]')
        
        keys.append(" - ".join(result_str))
        processed_qs.append(q_block)
        
    return processed_qs, keys

def process_part3(questions):
    """Xử lý Phần 3: Lấy Key"""
    processed_qs = []
    keys = []
    
    for q_block in questions:
        key_val = ""
        full_text = "".join([get_text(n) for n in q_block])
        
        # Tìm key
        m = re.search(r'<\s*key\s*=\s*(.*?)\s*>', full_text, re.IGNORECASE)
        if m:
            key_val = m.group(1).strip()
            # Xóa key khỏi text hiển thị
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

def create_header_paragraph(doc, text, align="center", bold=True):
    """Tạo Header cho file Word"""
    p = doc.createElementNS(W_NS, "w:p")
    pPr = doc.createElementNS(W_NS, "w:pPr")
    jc = doc.createElementNS(W_NS, "w:jc")
    jc.setAttributeNS(W_NS, "w:val", align)
    pPr.appendChild(jc)
    p.appendChild(pPr)
    
    r = create_run_with_style(doc, text, "000000", bold) # Màu đen cho header
    p.appendChild(r)
    return p

# ==================== 4. HÀM CHÍNH (GENERATE) ====================

def generate_mix(file_bytes, num_copies):
    outer_zip = io.BytesIO()
    csv_data = []
    
    with zipfile.ZipFile(outer_zip, 'w', zipfile.ZIP_DEFLATED) as z_out:
        # Đọc nội dung file gốc 1 lần
        in_io = io.BytesIO(file_bytes)
        with zipfile.ZipFile(in_io, 'r') as z_in:
            xml_content = z_in.read("word/document.xml")
            
            for i in range(num_copies):
                exam_code = str(random.randint(1001, 9999))
                
                # Parse XML mới cho mỗi đề
                dom = minidom.parseString(xml_content)
                intro, all_qs, body = parse_docx_xml(dom)
                
                # Chia 3 phần
                p1_qs = all_qs[0:18]
                p2_qs = all_qs[18:22]
                p3_qs = all_qs[22:]
                
                row_key = [exam_code]
                
                # Xử lý P1
                p1_done, k1 = process_part1(p1_qs, dom)
                c1 = list(zip(p1_done, k1))
                random.shuffle(c1)
                p1_final, k1_final = zip(*c1) if c1 else ([], [])
                row_key.extend(k1_final)
                
                # Xử lý P2
                p2_done, k2 = process_part2(p2_qs, dom)
                c2 = list(zip(p2_done, k2))
                random.shuffle(c2)
                p2_final, k2_final = zip(*c2) if c2 else ([], [])
                row_key.extend(k2_final)
                
                # Xử lý P3
                p3_done, k3 = process_part3(p3_qs)
                c3 = list(zip(p3_done, k3))
                random.shuffle(c3)
                p3_final, k3_final = zip(*c3) if c3 else ([], [])
                row_key.extend(k3_final)
                
                csv_data.append(row_key)
                
                # --- TÁI CẤU TRÚC FILE WORD (FIX LỖI TRÙNG CÂU HỎI) ---
                # 1. Xóa sạch các node cũ trong body (trừ sectPr nếu có)
                # Lưu ý: Không xóa sectPr ở cuối để giữ lề trang
                childNodes = list(body.childNodes)
                for node in childNodes:
                    if node.localName in ['p', 'tbl']:
                        body.removeChild(node)
                
                # 2. Chèn Header Trường
                header_nodes = [
                    create_header_paragraph(dom, "TRƯỜNG THPT MINH ĐỨC", "center", True),
                    create_header_paragraph(dom, "ĐỀ KIỂM TRA ĐỊNH KỲ", "center", True),
                    create_header_paragraph(dom, f"MÃ ĐỀ: {exam_code}", "right", True),
                    create_header_paragraph(dom, "Họ tên:.......................................................... Lớp:..........", "left", False),
                    create_header_paragraph(dom, "", "left", False)
                ]
                for n in header_nodes: body.appendChild(n)
                
                # 3. Chèn P1
                body.appendChild(create_header_paragraph(dom, "PHẦN I. Trắc nghiệm (18 câu)", "left", True))
                for idx, q in enumerate(p1_final):
                    # Renumber và tô màu xanh chữ "Câu X"
                    aggressive_replace_label(q[0], f"Câu {idx+1}.", dom, r'^Câu\s*\d+[\.\:]')
                    for node in q: body.appendChild(node)
                    
                # 4. Chèn P2
                body.appendChild(create_header_paragraph(dom, "PHẦN II. Đúng Sai (4 câu)", "left", True))
                for idx, q in enumerate(p2_final):
                    aggressive_replace_label(q[0], f"Câu {idx+1}.", dom, r'^Câu\s*\d+[\.\:]')
                    for node in q: body.appendChild(node)
                    
                # 5. Chèn P3
                body.appendChild(create_header_paragraph(dom, "PHẦN III. Trả lời ngắn (6 câu)", "left", True))
                for idx, q in enumerate(p3_final):
                    aggressive_replace_label(q[0], f"Câu {idx+1}.", dom, r'^Câu\s*\d+[\.\:]')
                    for node in q: body.appendChild(node)
                
                # Ghi file
                new_xml = dom.toxml().encode('utf-8')
                docx_io = io.BytesIO()
                with zipfile.ZipFile(docx_io, 'w', zipfile.ZIP_DEFLATED) as z_docx:
                    for item in z_in.infolist():
                        if item.filename == "word/document.xml":
                            z_docx.writestr(item.filename, new_xml)
                        else:
                            z_docx.writestr(item.filename, z_in.read(item.filename))
                
                z_out.writestr(f"De_Thi/De_{exam_code}.docx", docx_io.getvalue())
        
        # Ghi CSV
        csv_io = io.StringIO()
        writer = csv.writer(csv_io)
        head = ["Mã đề"] + [str(i) for i in range(1,19)] + [f"II_C{i}" for i in range(1,5)] + [f"III_C{i}" for i in range(1,7)]
        writer.writerow(head)
        writer.writerows(csv_data)
        z_out.writestr("Dap_An.csv", csv_io.getvalue().encode('utf-8-sig'))

    return outer_zip.getvalue()

# ==================== 5. UI LOGIC ====================
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)
st.markdown(HEADER_HTML, unsafe_allow_html=True)

st.markdown('<div class="main-card">', unsafe_allow_html=True)
uploaded_file = st.file_uploader("Kéo thả file .docx vào đây", type=['docx'])

if uploaded_file:
    st.success(f"✅ Đã nhận: {uploaded_file.name}")
    st.info("💡 Lưu ý: Đáp án đúng cần được Gạch chân hoặc Tô đỏ.")

st.write("")
num = st.number_input("Số lượng đề cần tạo", 1, 50, 4)

if st.button("🚀 BẮT ĐẦU TRỘN"):
    if not uploaded_file:
        st.warning("Vui lòng chọn file!")
    else:
        try:
            with st.spinner("Đang xử lý sâu XML..."):
                final_zip = generate_mix(uploaded_file.read(), num)
                st.success("✅ Thành công! Tải xuống bên dưới.")
                st.download_button(
                    label="📥 Tải về (Đề thi + Đáp án)",
                    data=final_zip,
                    file_name="KetQua_TNMic_Final.zip",
                    mime="application/zip"
                )
        except Exception as e:
            st.error(f"Lỗi: {e}")

st.markdown('</div>', unsafe_allow_html=True)
st.markdown('<div class="footer">© 2024 TNMic - Phần mềm Trộn đề</div>', unsafe_allow_html=True)
