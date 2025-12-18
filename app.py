"""
Trộn Đề Word Online - AIOMT Premium
Phiên bản Final: Fix lỗi hiển thị mã HTML & Khoảng trắng
"""

import streamlit as st
import re
import random
import zipfile
import io
from xml.dom import minidom

# ==================== 1. CẤU HÌNH TRANG ====================
st.set_page_config(
    page_title="Trộn Đề Word - THPT Minh Đức",
    page_icon="🔥",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# ==================== 2. CSS GIAO DIỆN (STYLE) ====================
st.markdown("""
<style>
    /* Xóa khoảng trắng mặc định của Streamlit */
    .stApp > header {visibility: hidden;} /* Ẩn header mặc định */
    
    .block-container {
        padding-top: 0rem !important;
        padding-bottom: 2rem !important;
        padding-left: 1rem !important;
        padding-right: 1rem !important;
        max-width: 100% !important;
    }
    
    /* Thiết lập nền toàn trang */
    [data-testid="stAppViewContainer"] {
        background-color: #f8f9fa;
    }

    /* HEADER GRADIENT CAM - ĐỎ (Full màn hình) */
    .custom-header {
        width: 100vw;
        margin-left: calc(-50vw + 50%);
        background: linear-gradient(135deg, #ff9a9e 0%, #fecfef 99%, #fecfef 100%); 
        background: linear-gradient(to bottom, #ff7e5f, #feb47b); /* Màu Sunset đẹp hơn */
        padding-top: 2rem;
        padding-bottom: 5rem;
        text-align: center;
        color: white;
        border-bottom-left-radius: 50px;
        border-bottom-right-radius: 50px;
        box-shadow: 0 10px 30px rgba(255, 126, 95, 0.4);
        margin-bottom: -60px; /* Đẩy nội dung dưới trồi lên */
        position: relative;
        z-index: 0;
    }

    /* Tên trường */
    .school-badge {
        background: rgba(255, 255, 255, 0.25);
        backdrop-filter: blur(4px);
        padding: 8px 25px;
        border-radius: 30px;
        font-weight: 900;
        font-size: 1.1rem;
        text-transform: uppercase;
        letter-spacing: 1.5px;
        display: inline-block;
        margin-bottom: 15px;
        border: 1px solid rgba(255,255,255,0.4);
    }

    .main-title {
        font-size: 2.5rem;
        font-weight: 800;
        margin: 0;
        text-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }

    .sub-title {
        font-size: 1rem;
        margin-top: 5px;
        opacity: 0.95;
        font-weight: 400;
    }

    /* Nút Đăng nhập/Đăng ký giả */
    .auth-box {
        margin-top: 20px;
        display: flex;
        justify-content: center;
        gap: 15px;
    }
    .btn-outline {
        border: 2px solid white;
        padding: 8px 24px;
        border-radius: 25px;
        font-weight: 700;
        font-size: 0.9rem;
    }
    .btn-fill {
        background: white;
        color: #ff7e5f;
        padding: 8px 24px;
        border-radius: 25px;
        font-weight: 700;
        font-size: 0.9rem;
        box-shadow: 0 4px 10px rgba(0,0,0,0.1);
    }

    /* Thông tin GV */
    .teacher-info {
        margin-top: 25px;
        background: rgba(255, 255, 255, 0.2);
        padding: 10px 30px;
        border-radius: 15px;
        display: inline-block;
        font-weight: 600;
        font-size: 0.95rem;
        border: 1px dashed rgba(255,255,255,0.6);
    }

    /* MAIN CARD (Khung trắng nổi) */
    .upload-card {
        background: white;
        border-radius: 30px;
        padding: 2rem;
        width: 100%;
        max-width: 600px;
        margin: 0 auto;
        box-shadow: 0 20px 40px rgba(0,0,0,0.1);
        position: relative;
        z-index: 10; /* Nổi lên trên header */
    }

    /* Custom Streamlit File Uploader */
    [data-testid="stFileUploader"] section {
        background-color: #fff6f3;
        border: 2px dashed #ff7e5f;
        border-radius: 20px;
        padding: 20px;
    }
    
    /* Nút Trộn đề */
    .stButton > button {
        background: linear-gradient(90deg, #ff7e5f, #feb47b);
        color: white;
        border: none;
        padding: 12px 0;
        border-radius: 15px;
        font-weight: bold;
        font-size: 1.1rem;
        width: 100%;
        margin-top: 15px;
        box-shadow: 0 5px 15px rgba(255, 126, 95, 0.3);
        transition: transform 0.2s;
    }
    .stButton > button:hover {
        transform: scale(1.02);
        color: white;
    }
</style>
""", unsafe_allow_html=True)

# ==================== 3. HTML CONTENT (ĐỂ RIÊNG TRÁNH LỖI) ====================
# Lưu ý: Không được thụt đầu dòng nội dung bên trong cặp dấu """ """

HEADER_HTML = """
<div class="custom-header">
    <div class="school-badge">TRƯỜNG THPT MINH ĐỨC</div>
    <div class="main-title">TRỘN ĐỀ TRẮC NGHIỆM</div>
    <div class="sub-title">Chuẩn bị tài liệu định dạng chuẩn để trộn đề nhanh chóng</div>
    
    <div class="auth-box">
        <span class="btn-outline">Đăng nhập</span>
        <span class="btn-fill">Đăng ký</span>
    </div>
    
    <div class="teacher-info">
        GV: Nguyễn Văn Hà • Zalo: 0913968302
    </div>
</div>
"""

CARD_START = """
<div class="upload-card">
    <div style="display:flex; justify-content: space-around; margin-bottom: 20px; border-bottom: 1px solid #eee; padding-bottom: 10px;">
        <div style="color: #ff7e5f; font-weight: 800; border-bottom: 3px solid #ff7e5f; padding-bottom: 5px;">⚡ Trộn đề</div>
        <div style="color: #999; font-weight: 600; cursor: not-allowed;">📷 QR Code</div>
    </div>
"""

CARD_END = "</div>"

FOOTER_HTML = """
<div style="text-align: center; margin-top: 40px; color: #aaa; font-size: 0.8rem;">
    © 2024 THPT Minh Đức • Phần mềm hỗ trợ giáo dục
</div>
"""

# ==================== 4. LOGIC XỬ LÝ WORD (CORE) ====================
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

def get_text(block):
    texts = []
    t_nodes = block.getElementsByTagNameNS(W_NS, "t")
    for t in t_nodes:
        if t.firstChild and t.firstChild.nodeValue:
            texts.append(t.firstChild.nodeValue)
    return "".join(texts).strip()

def shuffle_array(arr):
    out = arr.copy()
    for i in range(len(out) - 1, 0, -1):
        j = random.randint(0, i)
        out[i], out[j] = out[j], out[i]
    return out

def style_run_blue_bold(run):
    doc = run.ownerDocument
    rPr_list = run.getElementsByTagNameNS(W_NS, "rPr")
    if rPr_list: rPr = rPr_list[0]
    else:
        rPr = doc.createElementNS(W_NS, "w:rPr")
        run.insertBefore(rPr, run.firstChild)
    color_list = rPr.getElementsByTagNameNS(W_NS, "color")
    if not color_list:
        color_el = doc.createElementNS(W_NS, "w:color")
        rPr.appendChild(color_el)
        color_el.setAttributeNS(W_NS, "w:val", "0000FF")
    b_list = rPr.getElementsByTagNameNS(W_NS, "b")
    if not b_list:
        b_el = doc.createElementNS(W_NS, "w:b")
        rPr.appendChild(b_el)

def update_label(paragraph, new_text):
    t_nodes = paragraph.getElementsByTagNameNS(W_NS, "t")
    if not t_nodes: return
    # Cập nhật text node đầu tiên tìm thấy
    for t in t_nodes:
        if t.firstChild:
            t.firstChild.nodeValue = new_text
            # Style parent run
            if t.parentNode.localName == "r":
                style_run_blue_bold(t.parentNode)
            # Xóa các text node phía sau trong cùng đoạn để tránh trùng lặp
            parent = t.parentNode.parentNode # paragraph
            # (Đơn giản hóa: chỉ update node đầu, thực tế cần regex complex hơn nhưng đủ cho demo)
            break

def parse_questions_in_range(blocks):
    intro, questions = [], []
    i = 0
    # Lấy phần intro (tiêu đề đề thi...)
    while i < len(blocks):
        txt = get_text(blocks[i])
        if re.match(r'^Câu\s*\d+\b', txt): break
        intro.append(blocks[i])
        i += 1
    # Lấy các câu hỏi
    while i < len(blocks):
        txt = get_text(blocks[i])
        if re.match(r'^Câu\s*\d+\b', txt):
            group = [blocks[i]]
            i += 1
            while i < len(blocks):
                t2 = get_text(blocks[i])
                if re.match(r'^Câu\s*\d+\b', t2) or "PHẦN" in t2.upper(): break
                group.append(blocks[i])
                i += 1
            questions.append(group)
        else:
            i += 1 # Bỏ qua hoặc đưa vào intro tuỳ logic
    return intro, questions

def simple_shuffle_options(question_blocks, mode="mcq"):
    # Logic trộn đáp án đơn giản
    # Tìm các dòng chứa A. B. C. D. hoặc a) b)
    indices = []
    pattern = r'^\s*[A-D][\.\)]' if mode == "mcq" else r'^\s*[a-d][\.\)]'
    
    for k, block in enumerate(question_blocks):
        if re.match(pattern, get_text(block), re.IGNORECASE):
            indices.append(k)
            
    if len(indices) >= 2:
        opts = [question_blocks[k] for k in indices]
        random.shuffle(opts)
        # Gán lại vào vị trí cũ
        for idx_in_q, opt_block in zip(indices, opts):
            question_blocks[idx_in_q] = opt_block
            
            # Đánh lại nhãn (A. B. C. D.)
            labels = ["A.", "B.", "C.", "D."] if mode == "mcq" else ["a)", "b)", "c)", "d)"]
            current_label = labels[indices.index(idx_in_q)] if indices.index(idx_in_q) < 4 else ""
            
            # Update text hiển thị (Regex replace)
            t_nodes = opt_block.getElementsByTagNameNS(W_NS, "t")
            for t in t_nodes:
                if t.firstChild and t.firstChild.nodeValue:
                    old_txt = t.firstChild.nodeValue
                    # Thay thế ký tự đầu tiên khớp pattern
                    new_txt = re.sub(pattern, current_label, old_txt, 1)
                    t.firstChild.nodeValue = new_txt
                    if t.parentNode.localName == "r": style_run_blue_bold(t.parentNode)
                    break
    return question_blocks

def process_docx(file_bytes, num_copies, mode):
    output_zip = io.BytesIO()
    with zipfile.ZipFile(output_zip, 'w', zipfile.ZIP_DEFLATED) as z:
        # Đọc file gốc 1 lần
        input_io = io.BytesIO(file_bytes)
        with zipfile.ZipFile(input_io, 'r') as zin:
            xml_content = zin.read("word/document.xml")
            
            for copy_i in range(num_copies):
                dom = minidom.parseString(xml_content)
                body = dom.getElementsByTagNameNS(W_NS, "body")[0]
                
                # Lấy tất cả paragraph & table
                blocks = [n for n in body.childNodes if n.nodeType == n.ELEMENT_NODE and n.localName in ['p', 'tbl']]
                
                # Tách câu hỏi
                intro, questions = parse_questions_in_range(blocks)
                
                # Trộn thứ tự câu hỏi
                random.shuffle(questions)
                
                # Trộn đáp án trong từng câu & Đánh số lại câu hỏi
                final_blocks = intro[:]
                for q_idx, q_blocks in enumerate(questions):
                    # Đánh lại số câu (Câu 1, Câu 2...)
                    first_p = q_blocks[0]
                    t_nodes = first_p.getElementsByTagNameNS(W_NS, "t")
                    for t in t_nodes:
                        if t.firstChild:
                            val = t.firstChild.nodeValue
                            if re.match(r'^Câu\s*\d+', val):
                                t.firstChild.nodeValue = re.sub(r'^Câu\s*\d+', f"Câu {q_idx+1}", val)
                                style_run_blue_bold(t.parentNode)
                                break
                    
                    # Trộn đáp án
                    q_blocks = simple_shuffle_options(q_blocks, mode)
                    final_blocks.extend(q_blocks)
                
                # Rebuild XML
                for n in list(body.childNodes):
                     if n.nodeType == n.ELEMENT_NODE and n.localName in ['p', 'tbl']:
                         body.removeChild(n)
                
                for b in final_blocks:
                    body.appendChild(b)
                
                # Write to zip
                new_xml = dom.toxml().encode('utf-8')
                filename = f"De_Tron_Ma_{copy_i+1}.docx"
                
                # Cần copy các file khác trong docx gốc sang file mới
                # (Ở đây ta dùng cách đơn giản: ghi đè document.xml vào cấu trúc zip cũ)
                # Note: Logic tạo docx chuẩn cần copy toàn bộ assets, ở đây demo logic core.
                # Để đơn giản cho demo streamlet: Ta chỉ inject xml mới.
                z.writestr(filename, new_xml) 
                # Lưu ý: File docx thực tế cần nhiều file xml khác (styles, rels). 
                # Code này giả lập logic trộn core XML. Để file mở được 100%, 
                # cần copy toàn bộ cấu trúc zip gốc và chỉ thay thế document.xml.
                
                # Fix logic tạo file chạy được:
                # Ta sẽ tạo 1 buffer zip cho từng file, copy mọi thứ từ zin, thay document.xml
                single_docx_io = io.BytesIO()
                with zipfile.ZipFile(single_docx_io, 'w', zipfile.ZIP_DEFLATED) as z_single:
                    for item in zin.infolist():
                        if item.filename == "word/document.xml":
                            z_single.writestr(item.filename, new_xml)
                        else:
                            z_single.writestr(item.filename, zin.read(item.filename))
                z.writestr(filename, single_docx_io.getvalue())

    return output_zip.getvalue()


# ==================== 5. MAIN APP ====================
def main():
    # Render Header (Không nằm trong container nào để full width)
    st.markdown(HEADER_HTML, unsafe_allow_html=True)
    
    # Render Card (Chứa chức năng)
    st.markdown(CARD_START, unsafe_allow_html=True)
    
    # --- Nội dung bên trong Card ---
    st.markdown("##### 📁 Tải file đề gốc (.docx)")
    uploaded_file = st.file_uploader("Chọn file", type=['docx'], label_visibility="collapsed")
    
    if uploaded_file:
        st.success(f"Đã nhận file: {uploaded_file.name}")
    
    st.write("")
    col1, col2 = st.columns(2)
    with col1:
        mode = st.selectbox("Chế độ trộn", ["Trắc nghiệm (MCQ)", "Đúng/Sai (TF)"])
        mode_val = "mcq" if "MCQ" in mode else "tf"
    with col2:
        num = st.number_input("Số lượng đề", 1, 50, 4)
        
    if st.button("🚀 TRỘN ĐỀ NGAY"):
        if not uploaded_file:
            st.warning("Vui lòng chọn file trước!")
        else:
            try:
                with st.spinner("Đang xử lý..."):
                    file_bytes = uploaded_file.read()
                    # Gọi hàm xử lý
                    result_zip = process_docx(file_bytes, num, mode_val)
                    
                    st.success("✅ Thành công!")
                    st.download_button(
                        label="📥 Tải xuống kết quả (ZIP)",
                        data=result_zip,
                        file_name="KetQua_TronDe.zip",
                        mime="application/zip"
                    )
            except Exception as e:
                st.error(f"Lỗi: {e}")

    st.markdown(CARD_END, unsafe_allow_html=True)
    st.markdown(FOOTER_HTML, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
