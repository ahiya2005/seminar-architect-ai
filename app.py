import streamlit as st
import google.generativeai as genai
import io
import time
import re
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import PyPDF2

# --- 1. הגדרות דף ואימות ---
st.set_page_config(page_title="Seminar Architect PRO", page_icon="🎓", layout="wide")

api_key = st.secrets.get("GEMINI_API_KEY")
if api_key:
    genai.configure(api_key=api_key)

# CSS לעיצוב
st.markdown("""
    <style>
    header {visibility: hidden;} footer {visibility: hidden;}
    .stButton>button { width: 100%; background-color: #2c3e50; color: white; border-radius: 10px; font-weight: bold; height: 3.5em; }
    </style>
""", unsafe_allow_html=True)

# --- 2. פונקציות עזר ---
def is_valid_email(email):
    return re.match(r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z]{2,}$", email) is not None

def clean_ai_text(text):
    """ניקוי אגרסיבי של תגיות סורס ומספרים בסוגריים מרובעים"""
    text = re.sub(r"\[.*?source.*?\]", "", text, flags=re.IGNORECASE)
    text = re.sub(r"\[\d+\]", "", text)
    return text.strip()

def extract_text(file):
    if file is None: return ""
    try:
        if file.name.endswith('.pdf'):
            return "\n".join([p.extract_text() for p in PyPDF2.PdfReader(file).pages])
        return file.getvalue().decode("utf-8")
    except: return ""

def set_rtl(p):
    pPr = p._element.get_or_add_pPr()
    bidi = OxmlElement('w:bidi')
    bidi.set(qn('w:val'), '1')
    pPr.append(bidi)

def add_run(p, text, size=12, bold=False):
    run = p.add_run(text)
    run.font.name = 'David'
    run.font.size = Pt(size)
    run.bold = bold
    rPr = run._element.get_or_add_rPr()
    rtl = OxmlElement('w:rtl')
    rtl.set(qn('w:val'), '1')
    rPr.append(rtl)
    return run

# --- 3. מנוע ה-AI הרשמי (Official SDK) ---
def call_professor(title, topic, extra, notes, is_bib=False):
    model = genai.GenerativeModel('gemini-1.5-flash')
    
    if is_bib:
        prompt = f"תפקיד: ביבליוגרף אקדמי. צור רשימת מקורות APA מלאה (12 מקורות) בנושא '{topic}'. רשימה בלבד ללא הסברים."
    else:
        prompt = f"""
        תפקיד: פרופסור אקדמי בכיר. כתוב פרק עמוק מאוד (1000 מילים) תחת הכותרת '# {title}' לעבודה בנושא '{topic}'.
        חוקים: פסקאות ארוכות, ציטוטי APA בתוך הטקסט, חלק ל-3 תתי נושאים עם '##'. ללא הקדמות.
        הנחיות סטודנט: {extra}. סילבוס: {notes}.
        """

    for attempt in range(3):
        try:
            response = model.generate_content(
                prompt,
                generation_config={"temperature": 0.4, "max_output_tokens": 8000},
                safety_settings=[{"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_NONE"}]
            )
            if response.text:
                return clean_ai_text(response.text)
        except Exception as e:
            if "429" in str(e):
                st.warning("⏳ עומס בשרתי גוגל. ממתין 40 שניות לאיפוס...")
                time.sleep(40)
            else:
                time.sleep(5)
    return f"שגיאה בייצור הפרק {title}."

# --- 4. ממשק ומנגנון ייצור ---
lang = st.radio("🌐 שפה:", ["עברית", "English"], horizontal=True)
if lang == "עברית": st.markdown("<style>.block-container{direction:rtl; text-align:right;}</style>", unsafe_allow_html=True)

if 'logged_in' not in st.session_state: st.session_state['logged_in'] = False

if not st.session_state['logged_in']:
    email = st.text_input("אימייל:")
    if st.button("🚀 כניסה"):
        if is_valid_email(email): st.session_state['logged_in'] = True; st.rerun()
else:
    col1, col2 = st.columns(2)
    with col1:
        topic = st.text_input("נושא:")
        name = st.text_input("שם הסטודנט:")
        inst = st.text_input("מוסד:")
    with col2:
        extra = st.text_area("דגשים:", height=130)
    up = st.file_uploader("סילבוס:", type=['pdf', 'txt'])

    if st.button("🚀 צא לדרך!"):
        if not topic or not name: st.error("חסרים נתונים")
        else:
            st.warning("⚠️ תהליך אקדמי מורכב החל (כ-8 דקות). נא לא לסגור את הדף.")
            notes = extract_text(up)
            chapters = ["מבוא מורחב", "רקע תיאורטי", "ניתוח מערכות", "מקרי בוחן", "מסקנות", "ביבליוגרפיה"]
            
            res = []
            bar = st.progress(0)
            status = st.empty()
            
            for i, head in enumerate(chapters):
                status.info(f"⏳ כותב: **{head}**...")
                is_bib = (head == "ביבליוגרפיה")
                content = call_professor(head, topic, extra, notes, is_bib)
                res.append(content)
                bar.progress((i + 1) / (len(chapters) + 1))
                time.sleep(5)
            
            status.info("⏳ מסכם לתקציר...")
            abs_text = call_professor("תקציר", topic, "סיכום קצר של המחקר", "", False)
            res.insert(0, abs_text)
            
            bar.progress(1.0)
            status.success("🎉 הסמינריון מוכן!")
            
            # יצירת הוורד
            doc = Document()
            for s in doc.sections: s.top_margin = s.bottom_margin = s.left_margin = s.right_margin = Inches(0.98)
            
            # שער - תיקון הצמדת טקסט
            doc.add_paragraph('\n\n\n')
            p1 = doc.add_paragraph(); p1.alignment = 1; set_rtl(p1)
            add_run(p1, f"עבודה סמינריונית בנושא:\n{topic}", 24, True)
            doc.add_paragraph('\n\n')
            p2 = doc.add_paragraph(); p2.alignment = 1; set_rtl(p2)
            add_run(p2, f"מוגש על ידי: {name}", 16)
            if inst:
                doc.add_paragraph(); p3 = doc.add_paragraph(); p3.alignment = 1; set_rtl(p3)
                add_run(p3, f"מוסד אקדמי: {inst}", 16)
            doc.add_page_break()

            for t in res:
                for line in t.split('\n'):
                    line = line.strip()
                    if not line: continue
                    p = doc.add_paragraph(); p.alignment = 2; set_rtl(p)
                    if line.startswith('#'):
                        add_run(p, line.replace('#','').strip(), 16, True)
                    else:
                        p.paragraph_format.line_spacing = 1.5
                        p.paragraph_format.alignment = 3
                        add_run(p, line.replace('*',''), 12)
                doc.add_page_break()
            
            buf = io.BytesIO(); doc.save(buf); buf.seek(0)
            st.download_button("📥 הורד קובץ וורד", buf, f"Seminar_{name}.docx")
            st.balloons()
