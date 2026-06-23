from pydash import tap
import streamlit as st
import docx
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO
import re
import json
import requests
import os
import pypandoc
import streamlit as st


# Kéo xuống dưới các lệnh import cũ của bạn, dán thêm 2 dòng này:
from langchain_community.embeddings import HuggingFaceEmbeddings
from langchain_community.vectorstores import Pinecone

#Chuẩn hóa latex cho Pandoc
def chuan_hoa_latex_cho_pandoc(text):
    """
    Hàm này dọn dẹp các lỗi cú pháp Toán học do AI sinh ra 
    để đảm bảo Pandoc không bị crash khi xuất file Word.
    """
    if not text:
        return ""
        
    # 1. Chuyển đổi Block Math: \[ ... \] thành $$ ... $$
    # Dùng re.DOTALL để quét được công thức Toán dài xuống nhiều dòng
    text = re.sub(r'\\\[(.*?)\\\]', r'$$\1$$', text, flags=re.DOTALL)
    
    # 2. Chuyển đổi Inline Math: \( ... \) thành $ ... $
    text = re.sub(r'\\\((.*?)\\\)', r'$\1$', text)
    
    # 3. AI hay dính lỗi bôi đậm thẻ Toán (ví dụ: **$x$**), Pandoc rất ghét cái này.
    # Ta sẽ gỡ bỏ dấu sao bôi đậm quanh công thức:
    text = re.sub(r'\*\*\$(.*?)\$\*\*', r'$\1$', text)
    
    # 4. AI hay để khoảng trắng sát dấu $ (ví dụ: $ x + 1 $), Pandoc cần viết sát vào ($x+1$)
    text = re.sub(r'\$\s+(.*?)\s+\$', r'$\1$', text)
    
    return text
# --- 1. CẤU HÌNH API ---
# Lấy API Key từ két sắt bí mật của Streamlit
API_KEY = st.secrets["API_KEY"]
st.set_page_config(page_title="EduPHB Pro - Toán học", page_icon="📐", layout="wide")
# --- HÀM TẢI BỘ NÃO RAG ---
@st.cache_resource # Lệnh này giúp web không bị load lại não nhiều lần
def load_vector_db():
    embeddings = HuggingFaceEmbeddings(model_name="sentence-transformers/paraphrase-multilingual-MiniLM-L12-v2")
    # Trỏ vào thư mục não bộ đã tạo ở Giai đoạn 1
    db = Pinecone(persist_directory="./pinecone_db_giao_an", embedding_function=embeddings)
    return db

def lay_context_tu_db(query, db):
    # Tìm 3 đoạn văn bản giống nhất với yêu cầu
    docs = db.similarity_search(query, k=3)
    context = "\n\n---\n\n".join([doc.page_content for doc in docs])
    return context

# Khởi động bộ não khi bật web
vector_db = load_vector_db()

# --- HÀM GỌI API ---
def lay_giao_an_tu_ai(prompt_text, api_key):
    response = requests.post(
        "https://openrouter.ai/api/v1/chat/completions",
        headers={"Authorization": f"Bearer {api_key}"},
        json={
            "model": "nex-agi/nex-n2-pro:free",
            "messages": [{"role": "user", "content": prompt_text}]
        }
    )
    data = response.json()
    if "choices" not in data:
        raise Exception(data.get("error", {}).get("message", str(data)))
    return data["choices"][0]["message"]["content"]

# --- GIAO DIỆN CSS ---
st.markdown("""
<style>
.stApp { background-color: #f8fafc; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
.stButton>button { background: linear-gradient(135deg, #0ea5e9, #2563eb); color: white; font-weight: bold; border-radius: 8px; padding: 10px 24px; border: none;
transition: all 0.3s ease; width: 100%; }
.stButton>button:hover { transform: translateY(-2px); box-shadow: 0 10px 15px -3px rgba(37, 99, 235, 0.3); }
.main-header { text-align: center; color: #1e293b; font-size: 2.5rem; font-weight: 800; margin-bottom: 0.5rem; }
.info-box { background-color: #e0f2fe; border-left: 5px solid #0ea5e9; padding: 15px; border-radius: 4px; color: #0369a1; margin-bottom: 20px;}
</style>
""", unsafe_allow_html=True)
st.markdown('<div class="main-header">📐  EduPHB Pro Vietnam - Môn Toán</div>', unsafe_allow_html=True)

# --- 2. CƠ SỞ DỮ LIỆU ĐỘNG ---
CAP_HOC = {
"Tiểu học (Lớp 1 - 5)": ["Toán 1", "Toán 2", "Toán 3", "Toán 4", "Toán 5"],
"Trung học Cơ sở (Lớp 6 - 9)": ["Toán 6", "Toán 7", "Toán 8", "Toán 9"],
"Trung học Phổ thông (Lớp 10 - 12)": ["Toán 10", "Toán 11", "Toán 12"]
}

try:
    with open('toan_kntt.json', 'r', encoding='utf-8-sig') as f:
        BAI_HOC_DB = json.load(f)
except FileNotFoundError:
    st.warning("⚠️ Chưa tìm thấy file toan_kntt.json. Hệ thống sẽ cho phép nhập tay.")
    BAI_HOC_DB = {}
except UnicodeDecodeError:
    with open('toan_kntt.json', 'r', encoding='latin-1') as f:
        BAI_HOC_DB = json.load(f)

# --- 3. LAYOUT GIAO DIỆN ---
col_left, col_right = st.columns([1, 1], gap="large")

with col_left:
    st.markdown("### 📝 Thiết lập thông tin")
    cap_hoc_chon = st.selectbox("1. Cấp học", list(CAP_HOC.keys()))
    lop_hoc_chon = st.selectbox("2. Khối Lớp", CAP_HOC[cap_hoc_chon])
    tap_chon = st.selectbox("3. Tập", ["Tập 1", "Tập 2"])
    
    danh_sach_bai = BAI_HOC_DB.get(lop_hoc_chon, {}).get(tap_chon, [])
    if danh_sach_bai:
        bai_chon = st.selectbox("4. Chọn Bài học (KNTT)", danh_sach_bai)
        ten_bai = bai_chon
    else:
        bai_chon = st.selectbox("4. Chọn Bài học", ["-- Dữ liệu đang cập nhật, vui lòng nhập tay --"])
        ten_bai = st.text_input("✍️ Nhập tên bài học trong SGK:")

with col_right:
    st.markdown("### 🎯 ĐỊNH HƯỚNG VÀ LỚP HỌC")
    st.markdown('<div class="info-box"><strong>SGK: Toán - Kết nối tri thức với cuộc sống</strong></div>', unsafe_allow_html=True)
    
    # 1. GỢI Ý TÌNH TRẠNG LỚP HỌC
    DS_TINH_TRANG = [
        "Học sinh có học lực đồng đều, nắm bắt kiến thức ở mức độ cơ bản.",
        "Lớp có nhiều học sinh yếu/mất gốc Toán, cần giảng dạy chậm, chi tiết và ôn lại kiến thức cũ.",
        "Lớp chọn/Khá giỏi, tiếp thu nhanh, cần nhiều bài tập nâng cao, phát triển tư duy sâu.",
        "Học sinh hiếu động, dễ mất tập trung, cần ứng dụng nhiều hoạt động trải nghiệm thực tế.",
        "Lớp ghép/Có học sinh khuyết tật học hoà nhập (Cần phương pháp phân hóa cao).",
        "Khác (Tự nhập...)"
    ]
    tinh_trang_chon = st.selectbox("1. Tình trạng lớp học hiện tại:", DS_TINH_TRANG)
    if tinh_trang_chon == "Khác (Tự nhập...)":
        tinh_trang_thuc_te = st.text_input("Nhập tình trạng lớp học của bạn:")
    else:
        tinh_trang_thuc_te = tinh_trang_chon

    # 2. GỢI Ý PHƯƠNG PHÁP GIẢNG DẠY (Chuẩn LL&PPDH Toán & GDPT 2018)
    DS_PHUONG_PHAP = [
        "Dạy học phát hiện và giải quyết vấn đề (Gợi mở, Vấn đáp).",
        "Dạy học hợp tác (Thảo luận nhóm, Kỹ thuật khăn trải bàn, Mảnh ghép).",
        "Dạy học trực quan (Sử dụng công cụ thực hành, phần mềm GeoGebra, thiết bị CNTT).",
        "Phương pháp Trò chơi hoá (Gamification) tạo hứng thú học tập.",
        "Dạy học dự án / STEM (Phù hợp cho các bài thực hành, vận dụng thực tế).",
        "Lớp học đảo ngược (Flipped Classroom - Yêu cầu HS nghiên cứu trước ở nhà).",
        "Khác (Tự nhập...)"
    ]
    phuong_phap_chon = st.selectbox("2. Phương pháp & Kỹ thuật dạy học chủ đạo:", DS_PHUONG_PHAP)
    if phuong_phap_chon == "Khác (Tự nhập...)":
        phuong_phap_thuc_te = st.text_input("Nhập phương pháp giáo viên muốn áp dụng:")
    else:
        phuong_phap_thuc_te = phuong_phap_chon

    # 3. YÊU CẦU THÊM (Optional)
    ghi_chu_them = st.text_area("3. Ghi chú thêm cho AI (Tùy chọn):", placeholder="VD: Tập trung nhiều vào phần giải toán có lời văn...", height=68)

    # ĐÓNG GÓI CHUỖI 'DIEM_NHAN' ĐỂ GỬI CHO AI PROMPT
    diem_nhan = f"- Tình trạng lớp: {tinh_trang_thuc_te}\n- Phương pháp áp dụng: {phuong_phap_thuc_te}"
    if ghi_chu_them:
        diem_nhan += f"\n- Ghi chú thêm: {ghi_chu_them}"

st.markdown("---")
col1, col2, col3 = st.columns(3)
with col1:
    submit_btn = st.button("📝 TẠO KẾ HOẠCH BÀI DẠY", type="primary", use_container_width=True)
with col2:
    game_btn = st.button("🎲 TẠO TRÒ CHƠI KHỞI ĐỘNG", use_container_width=True)
with col3:
    exam_btn = st.button("💯 TẠO MA TRẬN & ĐỀ THI", use_container_width=True)

# --- 4. XỬ LÝ LOGIC (MEGA VIP PRO PROMPT) ---
if submit_btn:
    
    if not ten_bai or not tinh_trang_thuc_te or not phuong_phap_thuc_te:
        st.error("⚠️ Vui lòng điền/chọn đủ thông tin bài học và tình trạng lớp!")
    else:
        with st.spinner("🧠 Đang phân tích dữ liệu SGK Toán KNTT và thiết kế tiến trình..."):
            # TÌM KIẾM TRONG 44 FILE GIÁO ÁN
            # Câu lệnh để AI đi tìm tài liệu
            query_tim_kiem = f"Giáo án môn Toán {lop_hoc_chon} bài {ten_bai} phương pháp {phuong_phap_thuc_te}"
            # Lấy ra các đoạn giáo án mẫu
            extracted_context = lay_context_tu_db(query_tim_kiem, vector_db)
            # Phân tách cấu trúc dựa trên cấp học (Tiểu học dùng CV 2345, Trung học dùng CV 5512)
            is_primary = "Tiểu học" in cap_hoc_chon
            
            legal_doc = "Official Dispatch 2345/BGDĐT-GDTH" if is_primary else "Official Dispatch 5512/BGDĐT-GDTrH"
            
            if is_primary:
                structure_template = """
## I. YÊU CẦU CẦN ĐẠT
1. Kiến thức, kĩ năng: [Cụ thể cho bài học]
2. Năng lực: [Năng lực chung và Năng lực đặc thù Toán học]
3. Phẩm chất: [Phẩm chất liên quan]
## II. ĐỒ DÙNG DẠY HỌC
- Giáo viên: [Thiết bị, phần mềm, học liệu]
- Học sinh: [Đồ dùng học tập]
## III. CÁC HOẠT ĐỘNG DẠY HỌC CHỦ YẾU
### 1. Hoạt động Khởi động (Mục tiêu, Thời gian, Cách tiến hành)
### 2. Hoạt động Khám phá/Hình thành kiến thức mới (Mục tiêu, Thời gian, Cách tiến hành)
### 3. Hoạt động Thực hành/Luyện tập (Mục tiêu, Thời gian, Cách tiến hành)
### 4. Hoạt động Vận dụng (Mục tiêu, Thời gian, Cách tiến hành)
## IV. ĐIỀU CHỈNH SAU BÀI DẠY (Nếu có)
"""
            else:
                structure_template = """
## I. MỤC TIÊU
1. Kiến thức: [Mục tiêu kiến thức cụ thể]
2. Năng lực:
| Năng lực | Yêu cầu cần đạt |
| :--- | :--- |
| **Năng lực đặc thù** | (Tư duy và lập luận toán học; Mô hình hóa toán học; Giải quyết vấn đề; Giao tiếp toán học; Sử dụng công cụ). [Phân tích chi tiết theo bài học] |
| **Năng lực chung** | (Tự chủ và tự học; Giao tiếp và hợp tác; Giải quyết vấn đề và sáng tạo). [Phân tích chi tiết] |
3. Phẩm chất:
| Phẩm chất | Yêu cầu cần đạt |
| :--- | :--- |
| (Trách nhiệm, Chăm chỉ, Trung thực, Nhân ái) | [Phân tích chi tiết] |

## II. THIẾT BỊ DẠY HỌC VÀ HỌC LIỆU
- Máy chiếu, phiếu học tập, bảng phụ, phần mềm (Geogebra nếu cần)...

## III. TIẾN TRÌNH DẠY HỌC
[For EACH of the 4 activities below, STRICTLY use this exact sub-structure: a) Mục tiêu, b) Nội dung, c) Sản phẩm, d) Tổ chức thực hiện (Bước 1: Giao nhiệm vụ, Bước 2: Thực hiện nhiệm vụ, Bước 3: Báo cáo thảo luận, Bước 4: Kết luận nhận định)]

### Hoạt động 1: Mở đầu / Xác định vấn đề
[Insert a real-world mathematical context/problem here to spark curiosity]

### Hoạt động 2: Hình thành kiến thức mới
[Break down into sub-activities if needed (e.g., Hoạt động 2.1, 2.2). Include a Markdown table for 'Bảng kiểm đánh giá năng lực' in Bước 4 as seen in standard Vietnamese elite lesson plans].

### Hoạt động 3: Luyện tập
[Provide specific math exercises ranging from basic to advanced. Provide full solutions in the 'Sản phẩm' section].

### Hoạt động 4: Vận dụng
[Provide a practical, real-world application problem based on the newly acquired knowledge].
"""

            # Lắp ráp MEGA PROMPT
            prompt = f"""
[SYSTEM ROLE]
You are an elite, world-class Mathematics Educator, Curriculum Developer, and Educational Technologist in Vietnam. You have absolute mastery over the Vietnamese General Education Curriculum 2018 (Chương trình GDPT 2018), Modern Pedagogical Psychology, and the "Kết nối tri thức với cuộc sống" (KNTT) textbook series.

[YOUR TASK]
Generate a highly professional, scientifically accurate, and practical Lesson Plan (Kế hoạch bài dạy) for the requested math lesson.
The output MUST be written entirely in STANDARD ACADEMIC VIETNAMESE (Tiếng Việt chuẩn mực).
Do NOT include any conversational filler, greetings, or meta-commentary. Output ONLY the lesson plan.

[INPUT CONTEXT]
- Education Level: {cap_hoc_chon}
- Grade: {lop_hoc_chon}
- Volume: {tap_chon}
- Lesson Name: {ten_bai}
- Textbook Series: Toán - Kết nối tri thức với cuộc sống.
- Raw Input (Class Context & Pedagogy): "{diem_nhan}"

[PRE-PROCESSING & SANITIZATION ENGINE FOR RAW INPUT]
Before generating the lesson plan, you MUST strictly evaluate the "Raw Input" through the following 4-step pipeline to ensure high-quality, culturally and scientifically appropriate pedagogy:
1. GIBBERISH/SPAM FILTER: If the Raw Input contains random keystrokes, profanity, or lacks logical semantic meaning, IGNORE IT COMPLETELY. Do not let it influence the lesson plan.
2. NOISE REDUCTION: If the input contains a mix of relevant pedagogy and weird/spam words, filter out the noise and extract ONLY the educational intent.
3. CULTURAL TRANSLATION: If the input uses Vietnamese regional dialects (North, Central, South), local slang, or informal phrasing, TRANSLATE its core meaning into standard, formal educational terminology used by the Vietnamese Ministry of Education and Training (BGDĐT).
4. PEDAGOGICAL VALIDATION: Cross-reference the extracted intent with recognized educational sciences in Vietnam. If the requested method is scientifically valid and applicable to Mathematics (e.g., Gamification, STEM, Flipped Classroom, Group Work), seamlessly adapt the lesson's activities to fit this method. If it is pseudo-scientific or inappropriate, silently discard it and default to standard "Active Learning" and "Constructivism".

[STRICT FORMATTING & MATHEMATICAL NOTATION RULES]
- Regulatory Compliance: You MUST strictly follow the structure defined in {legal_doc}.
- Markdown Excellence: Use ## for main headings, ### for sub-headings. Use Markdown tables strictly for Competencies, Qualities, and Evaluation Rubrics (Bảng kiểm).
- Mathematical Typography (CRITICAL FOR EXPORT): You MUST use strict LaTeX formatting for ALL mathematical symbols, numbers, variables, and formulas. 
  + Use `$x$` for inline math (e.g., `$f(x) = ax^2 + bx + c$`).
  + Use `$$...$$` for block math.
  + Do NOT use plain text for math variables (e.g., write `$A$` instead of A, `$x$` instead of x).
  + Ensure no syntax errors so the docx compiler (Pandoc) does not crash.
- Content Quality: Draw inspiration from top-tier Vietnamese lesson plans. Use real-world contextualization (Mô hình hóa toán học) exactly like the KNTT curriculum philosophy.

[OUTPUT STRUCTURE TEMPLATE]
# KẾ HOẠCH BÀI DẠY: {ten_bai.upper()}
**Môn học:** Toán | **Lớp:** {lop_hoc_chon} | **{tap_chon}**
**Bộ sách:** Kết nối tri thức với cuộc sống

{structure_template}

[EXECUTE TASK NOW]
"""
            try:
                # Gọi API
                noi_dung_giao_an = lay_giao_an_tu_ai(prompt, API_KEY)
                st.success("✅ Đã tạo xong giáo án siêu VIP PRO!")
                
                tab1, tab2 = st.tabs([" 📄  Xem trước", " 📥  Tải xuống (Word)"])
                with tab1:
                    st.markdown(f'<div style="background-color: white; padding: 30px; border-radius: 8px; color: black;">\n\n{noi_dung_giao_an}\n\n</div>', unsafe_allow_html=True)
                
                with tab2:
                    temp_docx = "temp_giao_an.docx"
                    try:
                        # 1. DỌN DẸP LỖI LATEX TRƯỚC BẰNG HÀM VỪA TẠO
                        noi_dung_da_don_dep = chuan_hoa_latex_cho_pandoc(noi_dung_giao_an)
                        
                        # 2. CHO NỘI DUNG "ĐÃ SẠCH" VÀO PANDOC
                        pypandoc.convert_text(noi_dung_da_don_dep, 'docx', format='md', outputfile=temp_docx)
                        
                        with open(temp_docx, "rb") as f:
                            docx_bytes = f.read()
                            
                        st.download_button(
                            label="⬇️ Tải file Giáo án chuẩn (.docx)", 
                            data=docx_bytes, 
                            file_name=f"GiaoAn_{lop_hoc_chon}_{ten_bai.replace(' ', '_')}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )
                        
                        if os.path.exists(temp_docx):
                            os.remove(temp_docx)
                            
                    except Exception as e:
                        st.error(f"Lỗi hệ thống khi biên dịch file Word. Chi tiết lỗi: {e}")
            except Exception as e:
                st.error(f"Lỗi kết nối AI: {e}")

# --- TÍNH NĂNG 3: TẠO NGÂN HÀNG TRÒ CHƠI KHỞI ĐỘNG ---
if game_btn:
    if not ten_bai:
        st.warning("⚠️ Vui lòng nhập Tên bài học / Chuyên đề trước khi tạo trò chơi!")
    else:
        with st.spinner("🎲 Đang thiết kế 3 ý tưởng trò chơi siêu hấp dẫn..."):
            
            prompt_game = f"""
[SYSTEM ROLE]
You are a highly creative EdTech Expert and Mathematics Teacher in Vietnam, specializing in Gamification and Active Learning for Gen Z students.

[YOUR TASK]
Generate EXACTLY 3 highly engaging, practical, and fun Warm-up Games (Hoạt động khởi động) for the following math lesson.
The output MUST be written in STANDARD ACADEMIC VIETNAMESE (Tiếng Việt).

[INPUT CONTEXT]
- Cấp học: {cap_hoc_chon}
- Lớp: {lop_hoc_chon}
- Tên bài: {ten_bai}
- Phương pháp / Yêu cầu thêm từ giáo viên: {tinh_trang_thuc_te} / {phuong_phap_thuc_te}

[GAME DESIGN REQUIREMENTS]
Each game must follow this strict structure:
1. Tên trò chơi: (Catchy, fun, modern name. E.g., "Ai là triệu phú", "Giải cứu công chúa", "Mật mã bảo vật", "Kahoot Racing").
2. Thời gian: (Strictly 5 - 7 minutes).
3. Đồ dùng / Công cụ chuẩn bị: (Keep it simple - e.g., Plickers, Quizizz, Bảng phụ, Giấy A4, or just verbal).
4. Luật chơi chi tiết: (Step-by-step instructions on how the teacher runs it).
5. Kết nối vào bài mới: (CRITICAL: Explain exactly how the outcome of this game naturally transitions into the theoretical concept of the lesson "{ten_bai}").

[DIVERSITY RULE]
- Game 1: Must be a Technology-based game (using Quizizz, Kahoot, Plickers, Blooket, Wordwall...).
- Game 2: Must be a Physical/Team-based game (using whiteboard, running, passing objects, relay).
- Game 3: Must be a Real-world Mystery/Puzzle (using a real-life scenario, story-telling, or a trick question related to the math concept).

[EXECUTE TASK NOW]
"""
            try:
                # Gọi API với prompt thiết kế game
                noi_dung_game = lay_giao_an_tu_ai(prompt_game, API_KEY)
                
                # Hiển thị kết quả ra màn hình trong một cái khung nổi bật
                st.success("🎉 Đã tạo xong! Dưới đây là 3 ý tưởng dành riêng cho tiết học này:")
                st.markdown(f"""
                <div style="background-color: #f8f9fa; border-left: 5px solid #ff4b4b; padding: 20px; border-radius: 5px; color: #1f1f1f;">
                {noi_dung_game}
                </div>
                """, unsafe_allow_html=True)
                
                # Nút copy nhanh (Tuỳ chọn thêm cho xịn)
                st.download_button(
                    label="📥 Tải ý tưởng này về (.txt)",
                    data=noi_dung_game,
                    file_name=f"TroChoi_{ten_bai}.txt",
                    mime="text/plain"
                )

            except Exception as e:
                st.error(f"Lỗi kết nối AI: {e}")

# --- TÍNH NĂNG 4: TẠO MA TRẬN VÀ ĐỀ KIỂM TRA ---
if exam_btn:
    if not ten_bai:
        st.warning("⚠️ Vui lòng nhập Tên bài học / Chuyên đề trước khi tạo đề!")
    else:
        with st.spinner("💯 Đang xây dựng ma trận, biên soạn 10 câu trắc nghiệm và 2 câu tự luận..."):
            
            prompt_exam = f"""
[SYSTEM ROLE]
You are an elite Mathematics Assessment Expert and Curriculum Developer in Vietnam. You deeply understand the 2018 General Education Curriculum (Chương trình GDPT 2018) and the standard testing matrix guidelines of the Ministry of Education and Training (BGDĐT).

[YOUR TASK]
Create a rigorous, scientifically accurate 15-minute or 45-minute Mini-Test (Đề kiểm tra) for the requested math lesson.
The output MUST be written entirely in STANDARD ACADEMIC VIETNAMESE (Tiếng Việt).

[INPUT CONTEXT]
- Cấp học: {cap_hoc_chon}
- Khối lớp: {lop_hoc_chon}
- Tên bài/Chuyên đề: {ten_bai}

[STRICT OUTPUT STRUCTURE]
You must format your response EXACTLY like this:

## I. MA TRẬN ĐỀ KIỂM TRA
[Create a Markdown table showing the cognitive levels: Nhận biết (NB), Thông hiểu (TH), Vận dụng (VD), Vận dụng cao (VDC).
Distribute exactly:
- 10 Multiple Choice questions (0.5 pts each): 4 NB, 4 TH, 2 VD.
- 2 Essay questions: 1 VD (3.0 pts) and 1 VDC (2.0 pts).]

## II. ĐỀ KIỂM TRA
### A. PHẦN TRẮC NGHIỆM (5.0 điểm)
[Write 10 multiple-choice questions based on the matrix. Each question must have 4 distinct options: A, B, C, D. Ensure distractors are plausible student errors.]

### B. PHẦN TỰ LUẬN (5.0 điểm)
**Câu 1 (3.0 điểm):** [Practical, real-world application problem at the 'Vận dụng' level related to the lesson].
**Câu 2 (2.0 điểm):** [Advanced theoretical or complex problem at the 'Vận dụng cao' level].

## III. ĐÁP ÁN VÀ HƯỚNG DẪN CHẤM
### A. ĐÁP ÁN TRẮC NGHIỆM
[Provide a simple Markdown table with the correct letters for Q1-Q10].

### B. HƯỚNG DẪN CHẤM TỰ LUẬN
[Provide step-by-step mathematical solutions for Câu 1 and Câu 2. EXPLICITLY state the points awarded for each logical step (e.g., "+0.5 điểm", "+1.0 điểm") matching the total points of the question].

[FORMATTING & MATH RULES]
- You MUST use strict LaTeX formatting for ALL mathematical symbols, equations, and numbers.
- Use `$x$` for inline math.
- Use `$$x^2$$` for block math.
- DO NOT wrap math tags in bold/italic markdown (e.g., NO **$x$**).

[EXECUTE TASK NOW]
"""
            try:
                # Gọi API
                noi_dung_de_thi = lay_giao_an_tu_ai(prompt_exam, API_KEY)
                
                st.success("🎉 Đã tạo xong Ma trận và Đề kiểm tra chuẩn Bộ GD&ĐT!")
                
                tab_view, tab_download = st.tabs(["📄 Xem trước Đề thi", "📥 Tải xuống (Word)"])
                
                with tab_view:
                    st.markdown(f'<div style="background-color: white; padding: 30px; border-radius: 8px; color: black; box-shadow: 0 2px 4px rgba(0,0,0,0.1);">\n\n{noi_dung_de_thi}\n\n</div>', unsafe_allow_html=True)
                
                with tab_download:
                    temp_exam_docx = "temp_de_thi.docx"
                    try:
                        # Gọi hàm dọn dẹp LaTeX từ Bước 2 để chống crash
                        de_thi_da_don_dep = chuan_hoa_latex_cho_pandoc(noi_dung_de_thi)
                        
                        # Dùng pypandoc xuất Word
                        pypandoc.convert_text(de_thi_da_don_dep, 'docx', format='md', outputfile=temp_exam_docx)
                        
                        with open(temp_exam_docx, "rb") as f:
                            docx_bytes_exam = f.read()
                            
                        st.download_button(
                            label="⬇️ Tải file Đề Kiểm Tra (.docx)", 
                            data=docx_bytes_exam, 
                            file_name=f"DeKiemTra_{lop_hoc_chon}_{ten_bai.replace(' ', '_')}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )
                        
                        if os.path.exists(temp_exam_docx):
                            os.remove(temp_exam_docx)
                            
                    except Exception as e:
                        st.error(f"Lỗi biên dịch file Word. Vui lòng kiểm tra lại cấu trúc: {e}")

            except Exception as e:
                st.error(f"Lỗi kết nối AI: {e}")