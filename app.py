"""
TRỘN ĐỀ WORD - THPT MINH ĐỨC (FINAL VERSION)
Tính năng:
1. Chia 3 phần chuẩn form mới (1-18, 1-4, 1-6).
2. Tự động chèn tiêu đề hướng dẫn từng phần.
3. Mã đề 4 số ngẫu nhiên.
4. Xuất file Đáp án (Excel/CSV) dựa trên định dạng Đỏ/Gạch chân.
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

# ==================== 2. CSS GIAO DIỆN (KHÔNG THỤT DÒNG) ====================
CUSTOM_CSS = """
<style>
    /* Ẩn header mặc định */
    header {visibility: hidden;}
    .stApp > header {display: none;}
    
    .block-container {
        padding-top: 0rem !important;
        padding-bottom: 5rem !important;
        max-width: 100% !important;
        padding-left: 0 !important;
        padding-right: 0 !important;
    }
    
    [data-testid="stAppViewContainer"] {
        background: linear-gradient(to bottom right, #ffefba, #ffffff);
    }

    .header-wrapper {
        background: linear-gradient(135deg, #ff7e5f 0%, #feb47b 100%);
        padding-top: 3rem;
        padding-bottom: 6rem;
        text-align: center;
        color: white;
        border-bottom-left-radius: 50px;
        border-bottom-right-radius: 50px;
        box-shadow: 0 10px 20px rgba(255, 126, 95, 0.3);
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
        color: #ff7e5f;
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
        background: linear-gradient(90deg, #ff7e5f, #feb47b);
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
        color: #aaa;
        margin-top: 40px;
        font-size: 0.8rem;
    }
</style>
"""

HEADER_HTML = """
<div class="header-wrapper">
<div class="school-tag">TRƯỜNG THPT MINH ĐỨC</div>
<div class="main-h1">TNMIX 2025</div>
<div class="sub-text">Trộn đề trắc nghiệm chuẩn cấu trúc mới (3 Phần)</div>
<div class="info-gv">GV: Nguyễn Văn Hà • Zalo: 0913968302</div>
</div>
"""

# ==================== 3. LOGIC XỬ LÝ (CORE) ====================
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# --- MẪU TIÊU ĐỀ 3 PHẦN (XML RAW) ---
# Hàm tạo đoạn văn bản đậm (Bold Paragraph) cho tiêu đề phần
def create_header_xml(text, doc):
    p = doc.createElementNS(W_NS, "w:p")
    r = doc.createElementNS(W_NS, "w:r")
    rPr = doc.createElementNS(W_NS, "w:rPr")
    b = doc.createElementNS(W_NS, "w:b")
    rPr.appendChild(b)
    r.appendChild(rPr)
    t = doc.createElementNS(W_NS, "w:t")
    t.appendChild(doc.createTextNode(text))
    r.appendChild(t)
    p.appendChild(r)
    return p

def get_text(block):
    texts = []
    for t in block.getElementsByTagNameNS(W_NS, "t"):
        if t.firstChild: texts.append(t.firstChild.nodeValue)
    return "".join(texts).strip()

def check_is_correct(run_node):
    """Kiểm tra xem đáp án có phải là đáp án đúng không (Gạch chân hoặc Màu đỏ)"""
    rPr = run_node.getElementsByTagNameNS(W_NS, "rPr")
    if not rPr: return False
    # Check Underline
    u = rPr[0].getElementsByTagNameNS(W_NS, "u")
    if u: return True
    # Check Color (Red)
    color = rPr[0].getElementsByTagNameNS(W_NS, "color")
    if color:
        val = color[0].getAttributeNS(W_NS, "val")
        # Mã màu đỏ thường gặp: FF0000, red, hoặc tô đậm
        if val and val.upper() in ["FF0000", "RED"]: return True
    return False

def style_run_blue_bold(run, doc):
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

def parse_all_questions(blocks):
    intro = []
    questions = []
    i = 0
    # Lấy phần intro (Tiêu đề đề thi)
    while i < len(blocks):
        txt = get_text(blocks[i])
        if re.match(r'^Câu\s*\d+', txt): break
        intro.append(blocks[i])
        i += 1
    
    # Tách từng câu hỏi
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

def shuffle_options_mcq(q_block, doc):
    """Trộn Phần 1 (A,B,C,D) và trả về đáp án đúng (A/B/C/D)"""
    pat = r'^\s*[A-D][\.\)]'
    indices = [x for x, b in enumerate(q_block) if re.match(pat, get_text(b))]
    correct_char = ""
    
    if len(indices) >= 2:
        opts = [q_block[x] for x in indices]
        
        # Tìm đáp án đúng trong các options gốc
        correct_opt_idx = -1
        for idx, opt in enumerate(opts):
            # Check runs inside paragraph
            runs = opt.getElementsByTagNameNS(W_NS, "r")
            for r in runs:
                if check_is_correct(r):
                    correct_opt_idx = idx
                    break
            if correct_opt_idx != -1: break
        
        # Lưu lại nội dung node đúng để theo dõi sau khi shuffle
        target_opt = opts[correct_opt_idx] if correct_opt_idx != -1 else None

        random.shuffle(opts)
        
        lbls = ["A.", "B.", "C.", "D."]
        for x_idx, opt in zip(indices, opts):
            q_block[x_idx] = opt
            # Cập nhật nhãn A, B, C, D
            t_nodes = opt.getElementsByTagNameNS(W_NS, "t")
            for t in t_nodes:
                if t.firstChild:
                    cur_lbl = lbls[indices.index(x_idx)] if indices.index(x_idx) < 4 else ""
                    t.firstChild.nodeValue = re.sub(pat, cur_lbl, t.firstChild.nodeValue, 1)
                    style_run_blue_bold(t.parentNode, doc)
                    break
            
            # Kiểm tra xem đây có phải là đáp án đúng đã bị di chuyển không
            if target_opt and opt == target_opt:
                correct_char = lbls[indices.index(x_idx)][0] # Lấy ký tự A, B, C...

    return q_block, correct_char

def shuffle_options_tf(q_block, doc):
    """Trộn Phần 2 (a,b,c,d) và trả về chuỗi đáp án (a)Đ - b)S...)"""
    pat = r'^\s*[a-d][\.\)]'
    indices = [x for x, b in enumerate(q_block) if re.match(pat, get_text(b))]
    result_str_parts = []
    
    if len(indices) >= 2:
        opts = [q_block[x] for x in indices]
        
        # Xác định trạng thái Đ/S của từng option GỐC
        # Quy ước: Gạch chân/Đỏ = ĐÚNG, còn lại = SAI
        status_map = {} # Map object -> "Đ" or "S"
        for opt in opts:
            is_true = False
            runs = opt.getElementsByTagNameNS(W_NS, "r")
            for r in runs:
                if check_is_correct(r):
                    is_true = True
                    break
            status_map[opt] = "Đ" if is_true else "S"

        random.shuffle(opts)
        
        lbls = ["a)", "b)", "c)", "d)"]
        for x_idx, opt in zip(indices, opts):
            q_block[x_idx] = opt
            cur_lbl = lbls[indices.index(x_idx)] if indices.index(x_idx) < 4 else ""
            
            # Cập nhật nhãn a), b)...
            t_nodes = opt.getElementsByTagNameNS(W_NS, "t")
            for t in t_nodes:
                if t.firstChild:
                    t.firstChild.nodeValue = re.sub(pat, cur_lbl, t.firstChild.nodeValue, 1)
                    style_run_blue_bold(t.parentNode, doc)
                    break
            
            # Xây dựng chuỗi đáp án: a)Đ
            result_str_parts.append(f"{cur_lbl[:-1]}{status_map[opt]}") # a) -> a + Đ
            
    return q_block, " - ".join(result_str_parts)

def get_short_answer(q_block):
    """Lấy nội dung trả lời ngắn (Giả sử đáp án được tô đỏ ở cuối hoặc dòng riêng)"""
    # Logic đơn giản: Tìm text đỏ/gạch chân
    ans = ""
    for b in q_block:
        runs = b.getElementsByTagNameNS(W_NS, "r")
        for r in runs:
            if check_is_correct(r):
                # Lấy text của run này
                ts = r.getElementsByTagNameNS(W_NS, "t")
                for t in ts: 
                    if t.firstChild: ans += t.firstChild.nodeValue
    return ans.strip()

def process_and_zip(file_bytes, num_copies):
    output_buffer = io.BytesIO()
    
    # Dữ liệu cho file Excel đáp án
    # Format: [Mã đề, P1_1, P1_2..., P2_1..., P3_1...]
    answer_key_data = [] 
    
    with zipfile.ZipFile(output_buffer, 'w', zipfile.ZIP_DEFLATED) as z_out:
        input_io = io.BytesIO(file_bytes)
        with zipfile.ZipFile(input_io, 'r') as z_in:
            xml_content = z_in.read("word/document.xml")
            
            for _ in range(num_copies):
                # Tạo mã đề ngẫu nhiên 4 số
                exam_code = str(random.randint(1000, 9999))
                
                dom = minidom.parseString(xml_content)
                doc = dom.documentElement
                body = dom.getElementsByTagNameNS(W_NS, "body")[0]
                
                # Lấy blocks và parse câu hỏi
                blocks = [n for n in body.childNodes if n.localName in ['p', 'tbl']]
                intro, all_qs = parse_all_questions(blocks)
                
                # CHIA 3 PHẦN (CẮT LIST)
                # Input giả định: P1(18 câu), P2(4 câu), P3(6 câu)
                # Tự động cắt nếu đủ số lượng, nếu không thì lấy hết phần còn lại
                p1_qs = all_qs[0:18]
                p2_qs = all_qs[18:22]
                p3_qs = all_qs[22:]
                
                # List chứa đáp án của mã đề này
                current_key = [exam_code]
                
                # --- XỬ LÝ PHẦN 1 ---
                p1_final = []
                p1_keys = []
                for q in p1_qs:
                    q_new, key = shuffle_options_mcq(q, doc)
                    p1_final.append(q_new)
                    p1_keys.append(key)
                
                # Trộn thứ tự câu hỏi P1
                combined_p1 = list(zip(p1_final, p1_keys))
                random.shuffle(combined_p1)
                p1_final, p1_keys = zip(*combined_p1) if combined_p1 else ([], [])
                current_key.extend(p1_keys) # Add vào đáp án tổng
                
                # --- XỬ LÝ PHẦN 2 ---
                p2_final = []
                p2_keys = []
                for q in p2_qs:
                    q_new, key = shuffle_options_tf(q, doc)
                    p2_final.append(q_new)
                    p2_keys.append(key)
                
                # Trộn thứ tự câu hỏi P2
                combined_p2 = list(zip(p2_final, p2_keys))
                random.shuffle(combined_p2)
                p2_final, p2_keys = zip(*combined_p2) if combined_p2 else ([], [])
                current_key.extend(p2_keys)
                
                # --- XỬ LÝ PHẦN 3 ---
                p3_final = []
                p3_keys = []
                for q in p3_qs:
                    # Phần 3 không trộn đáp án, chỉ lấy text đáp án (nếu có tô đỏ)
                    key = get_short_answer(q)
                    p3_final.append(q)
                    p3_keys.append(key)
                
                # Trộn thứ tự câu hỏi P3
                combined_p3 = list(zip(p3_final, p3_keys))
                random.shuffle(combined_p3)
                p3_final, p3_keys = zip(*combined_p3) if combined_p3 else ([], [])
                current_key.extend(p3_keys)
                
                # Lưu đáp án vào danh sách tổng
                answer_key_data.append(current_key)
                
                # --- GỘP CÁC PHẦN VÀO DOC ---
                final_blocks = intro[:]
                
                # HEADER PHẦN 1
                txt_p1 = "PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn. Thí sinh trả lời từ câu 1 đến câu 18. Mỗi câu hỏi thí sinh chỉ chọn một phương án."
                final_blocks.append(create_header_xml(txt_p1, doc))
                
                # Add câu hỏi P1 (Đánh số 1 -> 18)
                for i, q in enumerate(p1_final):
                    # Renumber
                    t_nodes = q[0].getElementsByTagNameNS(W_NS, "t")
                    for t in t_nodes:
                        if t.firstChild and re.match(r'^Câu\s*\d+', t.firstChild.nodeValue):
                            t.firstChild.nodeValue = re.sub(r'^Câu\s*\d+', f"Câu {i+1}", t.firstChild.nodeValue)
                            style_run_blue_bold(t.parentNode, doc)
                            break
                    final_blocks.extend(q)
                
                # HEADER PHẦN 2
                txt_p2 = "PHẦN II. Câu trắc nghiệm đúng sai. Thí sinh trả lời từ câu 1 đến câu 4. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai."
                final_blocks.append(create_header_xml(txt_p2, doc))
                
                # Add câu hỏi P2 (Đánh số 1 -> 4)
                for i, q in enumerate(p2_final):
                    t_nodes = q[0].getElementsByTagNameNS(W_NS, "t")
                    for t in t_nodes:
                        if t.firstChild and re.match(r'^Câu\s*\d+', t.firstChild.nodeValue):
                            t.firstChild.nodeValue = re.sub(r'^Câu\s*\d+', f"Câu {i+1}", t.firstChild.nodeValue)
                            style_run_blue_bold(t.parentNode, doc)
                            break
                    final_blocks.extend(q)
                
                # HEADER PHẦN 3
                txt_p3 = "PHẦN III. Câu trắc nghiệm yêu cầu trả lời ngắn. Thí sinh trả lời từ câu 1 đến câu 6."
                final_blocks.append(create_header_xml(txt_p3, doc))
                
                # Add câu hỏi P3 (Đánh số 1 -> 6)
                for i, q in enumerate(p3_final):
                    t_nodes = q[0].getElementsByTagNameNS(W_NS, "t")
                    for t in t_nodes:
                        if t.firstChild and re.match(r'^Câu\s*\d+', t.firstChild.nodeValue):
                            t.firstChild.nodeValue = re.sub(r'^Câu\s*\d+', f"Câu {i+1}", t.firstChild.nodeValue)
                            style_run_blue_bold(t.parentNode, doc)
                            break
                    final_blocks.extend(q)
                
                # Thêm Mã đề vào đầu file (Intro)
                p_code = create_header_xml(f"MÃ ĐỀ: {exam_code}", doc)
                final_blocks.insert(len(intro), p_code)

                # Rebuild XML body
                for n in list(body.childNodes):
                    if n.localName in ['p', 'tbl']: body.removeChild(n)
                for b in final_blocks: body.appendChild(b)
                
                # Save Docx to Zip
                new_xml = dom.toxml().encode('utf-8')
                doc_io = io.BytesIO()
                with zipfile.ZipFile(doc_io, 'w', zipfile.ZIP_DEFLATED) as z_d:
                    for item in z_in.infolist():
                        if item.filename == "word/document.xml":
                            z_d.writestr(item.filename, new_xml)
                        else:
                            z_d.writestr(item.filename, z_in.read(item.filename))
                z_out.writestr(f"De_MinhDuc_{exam_code}.docx", doc_io.getvalue())
    
    # TẠO FILE ĐÁP ÁN (CSV)
    csv_io = io.StringIO()
    writer = csv.writer(csv_io)
    
    # Header CSV
    header_row = ["Mã đề"] + [f"P1_C{i}" for i in range(1,19)] + [f"P2_C{i}" for i in range(1,5)] + [f"P3_C{i}" for i in range(1,7)]
    writer.writerow(header_row)
    writer.writerows(answer_key_data)
    
    return output_buffer.getvalue(), csv_io.getvalue()

# ==================== 4. GIAO DIỆN CHÍNH ====================

st.markdown(CUSTOM_CSS, unsafe_allow_html=True)
st.markdown(HEADER_HTML, unsafe_allow_html=True)

st.markdown('<div class="main-card">', unsafe_allow_html=True)

# Tab giả
cols = st.columns(2)
with cols[0]:
    st.markdown('<div style="text-align:center; color:#ff7e5f; font-weight:bold; border-bottom:3px solid #ff7e5f; padding-bottom:5px;">⚡ Trộn đề</div>', unsafe_allow_html=True)
with cols[1]:
    st.markdown('<div style="text-align:center; color:#ccc; font-weight:500;">📷 QR Code</div>', unsafe_allow_html=True)

st.write("")
uploaded_file = st.file_uploader("Kéo thả file .docx vào đây", type=['docx'])

if uploaded_file:
    st.success(f"✅ Đã nhận file: {uploaded_file.name}")
    st.info("💡 Lưu ý: File gốc phải có đủ 28 câu (18 TN, 4 Đ/S, 6 TLN). Đáp án đúng cần được TÔ ĐỎ hoặc GẠCH CHÂN.")

st.write("")
num = st.number_input("Số lượng đề cần tạo", 1, 50, 4)

if st.button("🚀 BẮT ĐẦU TRỘN"):
    if not uploaded_file:
        st.warning("Vui lòng chọn file!")
    else:
        try:
            with st.spinner("Đang xử lý..."):
                zip_data, csv_data = process_and_zip(uploaded_file.read(), num)
                
                st.success("✅ Hoàn tất!")
                
                # Download Zip Đề
                st.download_button(
                    label="📥 Tải xuống Bộ Đề (ZIP)",
                    data=zip_data,
                    file_name="Bo_De_MinhDuc.zip",
                    mime="application/zip"
                )
                
                # Download Đáp án
                st.download_button(
                    label="📊 Tải xuống Đáp Án (CSV/Excel)",
                    data=csv_data.encode('utf-8-sig'),
                    file_name="Dap_An_Chi_Tiet.csv",
                    mime="text/csv"
                )
                
        except Exception as e:
            st.error(f"Lỗi: {e}")

st.markdown('</div>', unsafe_allow_html=True)
st.markdown('<div class="footer">© 2024 THPT Minh Đức</div>', unsafe_allow_html=True)
