"""
PHẦN MỀM TRỘN ĐỀ - TNMic (FINAL FIX: BLUE LABELS)
1. Giao diện: Xanh Ngọc (Teal) hiện đại.
2. Logic Word: Cưỡng chế A. B. C. D. -> TOÀN BỘ MÀU XANH DƯƠNG + IN ĐẬM.
3. Output: 1 File Zip.
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

# --- XỬ LÝ ĐÁP ÁN (GẠCH CHÂN / ĐỎ / KEY) ---

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
    """Xóa gạch chân/đỏ (để ẩn đáp án)"""
    rPr_list = run_node.getElementsByTagNameNS(W_NS, "rPr")
    if not rPr_list: return
    rPr = rPr_list[0]
    for u in rPr.getElementsByTagNameNS(W_NS, "u"): rPr.removeChild(u)
    for c in rPr.getElementsByTagNameNS(W_NS, "color"): rPr.removeChild(c)

def style_label_force_blue(run_node, doc):
    """
    HÀM QUAN TRỌNG: CƯỠNG CHẾ TÔ MÀU XANH DƯƠNG (#0000FF)
    """
    rPr_list = run_node.getElementsByTagNameNS(W_NS, "rPr")
    if rPr_list: 
        rPr = rPr_list[0]
    else:
        rPr = doc.createElementNS(W_NS, "w:rPr")
        run_node.insertBefore(rPr, run_node.firstChild)
    
    # 1. XÓA MỌI ĐỊNH DẠNG MÀU CŨ (Để tránh bị đè)
    old_colors = rPr.getElementsByTagNameNS(W_NS, "color")
    for oc in old_colors: rPr.removeChild(oc)
    
    # 2. XÓA ĐỊNH DẠNG BOLD CŨ
    old_bolds = rPr.getElementsByTagNameNS(W_NS, "b")
    for ob in old_bolds: rPr.removeChild(ob)

    # 3. THÊM MÀU XANH DƯƠNG (Blue)
    color_node = doc.createElementNS(W_NS, "w:color")
    color_node.setAttributeNS(W_NS, "w:val", "0000FF") # Màu xanh dương chuẩn
    rPr.appendChild(color_node)

    # 4. THÊM IN ĐẬM
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
                remove_answer_signal(r) # Xóa dấu hiệu đáp án gốc
            if is_cor: target_opt = opt
            
        random.shuffle(opts)
        
        lbls = ["A.", "B.", "C.", "D."]
        for idx, opt in enumerate(opts):
            real_idx = indices[idx]
            q_blocks[real_idx] = opt
            
            if target_opt and opt == target_opt:
                correct_char = lbls[idx][0]
                
            t_nodes = opt.getElementsByTagNameNS(W_NS, "t")
            for t in t_nodes:
                if t.firstChild:
                    # Thay thế A. B. C. D.
                    t.firstChild.nodeValue = re.sub(pat, lbls[idx], t.firstChild.nodeValue, 1)
                    # CƯỠNG CHẾ TÔ MÀU XANH CHO LABEL
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
                    # CƯỠNG CHẾ TÔ MÀU XANH CHO LABEL
                    style_label_force_blue(t.parentNode, doc)
                    break
                    
    return q_blocks, " - ".join(res_str)

def process_short(q_blocks):
    key_val = ""
    full_text = ""
    for b in q_blocks:
        t_nodes = b.getElementsByTagNameNS(W_NS, "t")
        for t in t_nodes:
            if
