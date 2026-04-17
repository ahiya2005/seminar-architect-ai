import streamlit as st
import requests
import json
import io
import time
import re
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import PyPDF2

# --- 1. הגדרות דף ועיצוב ---
st.set_page_config(page_title="Seminar Architect PRO", page_icon="🎓", layout="wide")

api_key = st.secrets.get("GEMINI_API_KEY")

st.markdown("""
    <style>
    header {visibility: hidden;}
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    div[data-testid="InputInstructions"] { display: none !important; }
    .stButton>button { 
        width: 100%; background-color: #2c3e50; 
        color: white; border-radius: 10px; 
        font-weight: bold; height: 3.5em; 
    }
    </style>
""", unsafe_allow_html=True)

# --- 2. פונקציות עזר (מותאמות לנייד) ---
def is_valid_email(email):
    pattern = r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z]{2,}$"
    return re.match(pattern, email) is not None

def clean_tags(text):
    # מנגנון ניקוי תגיות קצר ובטוח נגד שגיאות העתקה
    pattern = r"\"
    return re.sub(pattern, "", text)

def extract_text_from_file(uploaded_file):
    if uploaded_file is None: return ""
    try:
        if uploaded_file.name.endswith('.pdf'):
            reader = PyPDF2.PdfReader(uploaded_file)
            return "\n".join([page.extract_text() for page in reader.pages])
        return uploaded_file.getvalue().decode("utf-8")
    except Exception: return ""

# --- 3. הלב של המערכת: Master Prompt ---

def generate_dynamic_outline(topic, extra, key, lang="עברית"):
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key={key}"
    
    prompt = f"""
    תפקיד: פרופסור אקדמי בכיר.
    משימה: בנה 7 כותרות לסמינריון בנושא '{topic}'.
    הנחיות: {extra}
    
    חוקים:
    1. אין כותרות גנריות. הכל מותאם לנושא.
    2. פרק אחרון: 'ביבליוגרפיה'.
    
    החזר רק את הכותרות מופרדות בפסיק (,) ללא טקסט נוסף.
    """
    payload = {
        "contents": [{"parts": [{"text": prompt}]}], 
        "generationConfig": {"temperature": 0.3}
    }
    
    try:
        response = requests.post(url, json=payload, timeout=40)
        if response.status_code == 200:
            text = response.json()['candidates'][0]['content']['parts'][0]['text']
            return [c.strip() for c in text.split(',') if c.strip()]
    except: pass
    return ["מבוא", "רקע תיאורטי", "ניתוח מערכות", "השפעות ותהליכים", "מקרי בוחן", "מסקנות", "ביבליוגרפיה"]

def call_gemini_master_professor(title, topic, name, extra, notes, key, lang="עברית"):
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key={key}"
    
    instruction = f"""
    תפקיד: פרופסור אקדמי.
    משימה: כתוב פרק '{title}' לנושא '{topic}' עבור הסטודנט {name}.
    
    חוקי ברזל:
    1. טקסט ארוך ומעמיק (1500 מילים). פסקאות של 4 שורות לפחות.
    2. חובה ציטוטי APA פנימיים.
    3. אסור להשתמש בתגיות קוד או סוגריים מרובעים בשום אופן.
    4. חלק ל-4 תתי-נושאים בעזרת '##'.
    5. ללא הקדמות. התחל עם '# {title}'.
    
    סילבוס: {notes}
    דגשים: {extra}
    """
    
    payload = {
        "contents": [{"parts": [{"text": instruction}]}], 
        "generationConfig": {"temperature": 0.4, "maxOutputTokens": 6000}
    }
    
    for attempt in range(3):
        try:
            response = requests.post(url, json=payload, timeout=120)
            if response.status_code == 200:
                text = response.json()['candidates'][0]['content']['parts'][0]['text']
                clean_text = clean_tags(text)
                if len(clean_text) > 400: return clean_text.strip()
            time.sleep(3) 
        except: time.sleep(3)
    return f"שגיאה בייצור הפרק '{title}'."

# --- 4. עיצוב וורד תקני ---
def create_master_doc(title, author, institution, content_list, lang):
    doc = Document()
    font_name = 'David' if lang == "עברית" else 'Times New Roman'
    
    for section in doc.sections:
        section.top_margin = Inches(0.98)
        section.bottom_margin = Inches(0.98)
        section.left_margin = Inches(0.98)
        section.right_margin = Inches(0.98)

    # עמוד שער
    doc.add_paragraph('\n\n\n\n')
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(f"עבודה סמינריונית בנושא:\n{title}")
    r.bold = True
    r.font.size = Pt(24)
    r.font.name = font_name
    
    doc.add_paragraph('\n\n\n\n')
    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    inst_text = f"\nמוסד אקדמי: {institution}" if institution else ""
    p2.add_run(f"מוגש על ידי: {author}{inst_text}").font.name = font_name
    doc.add_page_break()

    for text in content_list:
        for line in text.split('\n'):
            line = line.strip()
            if not line: continue
            
            if line.startswith('# ') and not line.startswith('## '):
                h = doc.add_heading(line.replace('#', '').strip(), level=1)
                h.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            elif line.startswith('## '):
                sh = doc.add_heading(line.replace('##', '').strip(), level=2)
                sh.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            else:
                clean_line = clean_tags(line)
                p = doc.add_paragraph(clean_line.replace('*', ''))
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                p.paragraph_format.line_spacing = 1.5
                p.paragraph_format.alignment = 3 
        doc.add_page_break()
    return doc

# --- 5. ממשק משתמש ---
lang = st.radio("🌐 שפת ממשק:", ["עברית", "English"], horizontal=True)
if lang == "עברית":
    st.markdown("<style>.block-container { direction: rtl; text-align: right; }</style>", unsafe_allow_html=True)

if 'logged_in' not in st.session_state: st.session_state['logged_in'] = False

st.title("🎓 Seminar Architect PRO")

if not st.session_state['logged_in']:
    email_input = st.text_input("אימייל להתחברות:")
    if st.button("🚀 כניסה למערכת"):
        if email_input and is_valid_email(email_input):
            st.session_state['user_email'] = email_input.lower()
            st.session_state['logged_in'] = True
            st.rerun()
        else: st.error("נא להזין אימייל תקין עם סיומת מורשית.")
else:
    st.sidebar.success(f"מחובר כ: {st.session_state['user_email']}")
    col1, col2 = st.columns(2)
    with col1:
        topic = st.text_input("נושא הסמינריון:")
        name = st.text_input("שם הסטודנט:")
        institution = st.text_input("מוסד אקדמי (אופציונלי):")
    with col2:
        extra = st.text_area("דגשים ספציפיים ופרוטוקול אישי:", height=130)
    uploaded = st.file_uploader("העלאת סילבוס (PDF/TXT):", type=['pdf', 'txt'])

    if st.button("🚀 צא לדרך! הפעל פרוטוקול כתיבה אקדמי"):
        if not topic or not name: st.error("הזן נושא ושם סטודנט!")
        elif not api_key: st.error("שגיאת API KEY.")
        else:
            st.warning("⚠️ פרוטוקול כתיבה מתקדם מופעל. נא לא לסגור את החלון!")
            notes = extract_text_from_file(uploaded)
            
            status_text = st.empty()
            time_est = st.empty()
            progress_bar = st.progress(0)
            
            status_text.info("⏳ מאבחן נושא ובונה שלד פרקים...")
            chapters = generate_dynamic_outline(topic, extra, api_key, lang)
            
            generated_content = []
            total_steps = len(chapters) + 1 
            
            for i, head in enumerate(chapters):
                pct = int((i / total_steps) * 100)
                rem = (total_steps - i) * 55 
                
                status_text.info(f"⏳ ({pct}%) כותב פרק עומק: **{head}**...")
                time_est.markdown(f"⏱️ **זמן לסיום:** כ-{rem} שניות")
                
                content = call_gemini_master_professor(head, topic, name, extra, notes, api_key, lang)
                generated_content.append(content)
                progress_bar.progress((i + 1) / total_steps)
                time.sleep(4)
            
            status_text.info("⏳ מסכם את העבודה...")
            summary_prompt = f"כתוב תקציר לסמינריון בנושא '{topic}'. סיכום בלבד."
            abstract = call_gemini_master_professor("תקציר", topic, name, summary_prompt, "", api_key, lang)
            generated_content.insert(0, abstract) 
            
            progress_bar.progress(1.0)
            status_text.success("🎉 הסמינריון הושלם בהצלחה!")
            
            doc = create_master_doc(topic, name, institution, generated_content, lang)
            buf = io.BytesIO()
            doc.save(buf)
            buf.seek(0)
            st.download_button("📥 הורד סמינריון מושלם (Word)", buf, f"Seminar_{name}.docx")
            st.balloons()
