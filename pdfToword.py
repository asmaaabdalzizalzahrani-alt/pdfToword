import streamlit as st
import fitz  # مكتبة PyMuPDF لقراءة PDF
import pytesseract
from PIL import Image
import io
import re
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# 1. إعدادات الصفحة والتجاوب مع جميع الأجهزة
st.set_page_config(
    page_title="محول PDF إلى Word الذكي",
    page_icon="📄",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# 2. تصميم الواجهة باللون البنفسجي وتنسيق النصوص العربية (RTL)
st.markdown("""
    <style>
    /* خلفية الصفحة بالكامل باللون البنفسجي */
    .stApp {
        background: linear-gradient(135deg, #4A148C 0%, #7B1FA2 50%, #9C27B0 100%);
        color: #FFFFFF;
        direction: rtl;
        text-align: right;
    }
    
    /* تنسيق العناوين */
    h1, h2, h3, p, label {
        color: #FFFFFF !important;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        text-align: right;
    }
    
    /* تنسيق صندوق رفع الملفات */
    [data-testid="stFileUploader"] {
        background-color: rgba(255, 255, 255, 0.12);
        border: 2px dashed #E1BEE7;
        border-radius: 15px;
        padding: 20px;
    }
    
    /* تنسيق الأزرار */
    .stButton > button, .stDownloadButton > button {
        background-color: #FFFFFF !important;
        color: #4A148C !important;
        font-weight: bold !important;
        font-size: 18px !important;
        border-radius: 10px !important;
        border: none !important;
        padding: 12px 24px !important;
        width: 100%;
        box-shadow: 0 4px 6px rgba(0,0,0,0.2);
        transition: all 0.3s ease;
    }
    
    .stButton > button:hover, .stDownloadButton > button:hover {
        background-color: #F3E5F5 !important;
        transform: translateY(-2px);
    }
    </style>
""", unsafe_allow_html=True)

# 3. شاشة الدخول الترحيبية (بدون رقم سري)
st.markdown("<h1 style='text-align: center; margin-bottom: 10px;'>✨ أهلاً بك لتحويل ملف PDF إلى Word بكل سهولة ✨</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; font-size: 18px; color: #E1BEE7;'>أداة ذكية لاستخراج النصوص بدقة عالية وتنظيف الملفات تلقائياً</p>", unsafe_allow_html=True)
st.write("---")

# دوال الذكاء والتنظيف (تجاهل الفهارس والمصادر)
def is_ignored_section(text):
    """التحقق مما إذا كانت الصفحة عبارة عن فهرس أو قائمة مصادر"""
    clean_text = text.strip()
    # كلمات دلالية للفهارس والمصادر في بداية الصفحة
    ignored_keywords = [
        r'^(الفهرس|فهرس المحتويات|المحتويات|قائمة المحتويات)',
        r'^(المصادر|المراجع|قائمة المصادر|قائمة المراجع|المصادر والمراجع)',
        r'^(Table of Contents|Contents|Index|References|Bibliography)'
    ]
    first_lines = "\n".join(clean_text.split("\n")[:4]) # فحص أول 4 أسطر فقط
    for pattern in ignored_keywords:
        if re.search(pattern, first_lines, re.IGNORECASE | re.MULTILINE):
            return True
            
    # اكتشاف الفهارس عن طريق كثرة النقاط المتتالية (مثل: المقدمة ........... 5)
    dotted_lines = len(re.findall(r'\.{4,}\s*\d+', clean_text))
    if dotted_lines >= 4:
        return True
        
    return False

def clean_extracted_text(text):
    """تنظيف النص من الرموز العشوائية الناتجة عن الرسومات"""
    lines = text.split('\n')
    cleaned_lines = []
    for line in lines:
        line = line.strip()
        # تجاهل الأسطر التي تحتوي على رموز غريبة فقط (بقايا رسومات)
        if len(line) > 0 and not re.match(r'^[\W_]+$', line):
            cleaned_lines.append(line)
    return "\n".join(cleaned_lines)

def process_pdf_to_word(pdf_bytes):
    """قراءة الـ PDF وتحويله إلى Word مع تجاهل الرسومات والفهارس والمصادر"""
    doc_pdf = fitz.open(stream=pdf_bytes, filetype="pdf")
    doc_word = Document()
    
    # إعداد اتجاه ملف الوورد ليدعم العربية بشكل سليم
    style = doc_word.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(12)

    total_pages = len(doc_pdf)
    progress_bar = st.progress(0)
    status_text = st.empty()

    for page_num in range(total_pages):
        status_text.text(f"جاري معالجة وقراءة الصفحة {page_num + 1} من {total_pages}...")
        page = doc_pdf.load_page(page_num)
        
        # 1. محاولة استخراج النص المباشر (تجاهل الصور والرسومات تلقائياً)
        text = page.get_text("text").strip()
        
        # 2. إذا كانت الصفحة عبارة عن صورة ممسوحة ضوئياً (لا يوجد بها نص قابل للنسخ)
        if len(text) < 30:
            # تحويل الصفحة لصورة عالية الدقة (300 DPI) لقراءة النصوص وأرقام الصفحات بوضوح
            pix = page.get_pixmap(dpi=300)
            img = Image.open(io.BytesIO(pix.tobytes("png")))
            # استخراج النص باللغتين العربية والإنجليزية
            text = pytesseract.image_to_string(img, lang='ara+eng')
            
        # 3. التحقق من الفهارس والمصادر وتجاهلها
        if is_ignored_section(text):
            progress_bar.progress((page_num + 1) / total_pages)
            continue
            
        # 4. تنظيف النص وإضافته لملف الوورد
        cleaned_text = clean_extracted_text(text)
        if cleaned_text:
            for paragraph_text in cleaned_text.split('\n'):
                p = doc_word.add_paragraph(paragraph_text)
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            
            # إضافة رقم الصفحة في نهاية كل صفحة بوضوح
            page_footer = doc_word.add_paragraph(f"--- صفحة {page_num + 1} ---")
            page_footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # فاصل صفحة (إلا في الصفحة الأخيرة)
            if page_num < total_pages - 1:
                doc_word.add_page_break()

        progress_bar.progress((page_num + 1) / total_pages)

    status_text.empty()
    progress_bar.empty()
    
    # حفظ ملف الوورد في الذاكرة لتجهيزه للتحميل
    word_io = io.BytesIO()
    doc_word.save(word_io)
    word_io.seek(0)
    return word_io

# 4. خانة إدخال ملف الـ PDF
uploaded_file = st.file_uploader("📂 اختر ملف الـ PDF من جهازك أو اسحبه هنا:", type=["pdf"])

if uploaded_file is not None:
    st.info(f"تم رفع الملف: **{uploaded_file.name}** بنجاح.")
    
    if st.button("🚀 ابدأ التحويل الآن"):
        with st.spinner("جاري تحليل الصفحات واستخراج النصوص بدقة... يرجى الانتظار"):
            try:
                pdf_bytes = uploaded_file.read()
                word_file = process_pdf_to_word(pdf_bytes)
                
                # 5. إظهار رسالة النجاح وزر التحميل
                st.success("🎉 تم تحويل الملف إلى Word بنجاح ودقة عالية!")
                
                output_filename = uploaded_file.name.rsplit('.', 1)[0] + ".docx"
                st.download_button(
                    label="📥 اضغط هنا لتحميل ملف الـ Word جاهزاً",
                    data=word_file,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as e:
                st.error(f"حدث خطأ أثناء المعالجة: {e}")
                st.warning("تأكد من تثبيت أداة Tesseract OCR لدعم قراءة الصور.")