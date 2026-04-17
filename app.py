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

# --- 1. הגדרות דף ועיצוב (ממשק משתמש) ---
st.set_page_config(page_title="Seminar Architect PRO", page_icon="🎓", layout="wide")

api_key = st.secrets.get("GEMINI_API_KEY")

st.markdown("""
    <style>
    header {visibility: hidden;}
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    div[data-testid="InputInstructions"] { display: none !important; }
    .stButton>button { width: 100%; background-color: #2c3e50; color: white; border-radius: 10px; font-weight: bold; height: 3.5em; }
    .stProgress > div > div > div {
        background-image: linear-gradient(45deg, rgba(255, 255, 255, .15) 25%, transparent 25%, transparent 50%, rgba(255, 255, 255, .15) 50%, rgba(255, 255, 255, .15) 75%, transparent 75%, transparent);
        background-size: 1rem 1rem;
        animation: progress-bar-stripes 1s linear infinite;
    }
    @keyframes progress-bar-stripes { from { background-position: 1rem 0; } to { background-position: 0 0; } }
    </style>
""", unsafe_allow_html=True)

# --- 2. אימות אימייל נוקשה (תיקון הפרצה שמצאת) ---
def is_valid_email(email):
    # בודק מבנה תקין וסיומת של לפחות 2 אותיות (מונע gmail.c)
    pattern = r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z]{2,}$"
    return re.match(pattern, email) is not None

def extract_text_from_file(uploaded_file):
    if uploaded_file is None: return ""
    try:
        if uploaded_file.name.endswith('.pdf'):
            reader = PyPDF2.PdfReader(uploaded_file)
            return "\n".join([page.extract_text() for page in reader.pages])
        return uploaded_file.getvalue().decode("utf-8")
    except Exception: return ""

# --- 3. הלב של המערכת: Master Prompt Implementation ---

def generate_dynamic_outline(topic, extra, key, lang="עברית"):
    """שלב 1 של הפרופסור: אבחון ומיקוד מחקרי ובניית שלד פרקים לא גנריים"""
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key={key}"
    
    prompt = f"""
    ניתוח אקדמי עבור הנושא: {topic}
    הנחיות נוספות: {extra}
    
    תפקיד: פרופסור אקדמי בכיר.
    משימה: בצע אבחון ומיקוד מחקרי (שלב 1). קבע זווית מחקרית וסיווג עבודה (עיונית/אמפירית).
    לאחר מכן, בנה שלד של 7 פרקים מרכזיים.
    איסור מוחלט: אין להשתמש בכותרות גנריות כמו 'מבוא', 'סקירה ספרותית', 'מתודולוגיה' או 'ממצאים'.
    כל כותרת חייבת להיות כותרת תוכן ממשית וספציפית לנושא.
    הפרק האחרון חייב להיות 'ביבליוגרפיה'.
    
    החזר אך ורק את 7 הכותרות, מופרדות בפסיק (,), ללא שום הסברים נוספים.
    """
    
    payload = {"contents": [{"parts": [{"text": prompt}]}], "generationConfig": {"temperature": 0.3}}
    try:
        response = requests.post(url, json=payload, timeout=40)
        if response.status_code == 200:
            text = response.json()['candidates'][0]['content']['parts'][0]['text']
            return [c.strip() for c in text.split(',') if c.strip()]
    except: pass
    return ["מבוא ומיקוד שאלת המחקר", "רקע היסטורי ותיאורטי", "ניתוח היבטים מרכזיים א'", "ניתוח היבטים מרכזיים ב'", "מקרה בוחן ויישומיות", "דיון ומסקנות", "ביבליוגרפיה"]

def call_gemini_master_professor(chapter_title, total_context, key, lang="עברית", max_retries=3):
    """הפעלת הפרופסור האקדמי לכתיבת הפרק לפי ה-Master Prompt"""
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key={key}"
    
    # הטמעת פקודת המערכת הקשיחה שלך בתוך הקוד
    master_prompt_instruction = f"""
    תפקיד: פרופסור אקדמי בכיר.
    משימה: כתוב את הפרק הבא בעבודה סמינריונית: '{chapter_title}'.
    
    הנחיות שלב 2 (פורמט):
    - משלב אקדמי גבוה, גוף שלישי בלבד.
    
    הנחיות שלב 3 ו-5 (עומק ואיכות):
    - עומק הפסקה: לפחות 4 שורות. כל פסקה כוללת: טענה, הסבר והוכחה ממקור.
    - מקורות: חובה לשלב ציטוטים אקדמיים (APA) בתוך הטקסט (שם מחבר, שנה).
    - חוט מקשר: סיים את הפרק בשורת מעבר המקשרת לפרק הבא ולשאלת המחקר.
    - איכות: אל תמציא נתונים. השתמש בניתוח לוגי מעמיק.
    - כותרות משנה: חלק את הפרק ל-3-4 תתי-נושאים ספציפיים בעזרת '##'.
    
    התחל לכתוב ישירות את תוכן הפרק תחת הכותרת '# {chapter_title}'. ללא הקדמות.
    
    הקשר העבודה המלא: {total_context}
    """
    
    payload = {"contents": [{"parts": [{"text": master_prompt_instruction}]}], "generationConfig": {"temperature": 0.5, "maxOutputTokens": 4000}}
    
    for attempt in range(max_retries):
        try:
            response = requests.post(url, json=payload, timeout=120)
            if response.status_code == 200:
                text = response.json()['candidates'][0]['content']['parts'][0]['text']
                if len(text) > 400: return text.strip()
            time.sleep(3) # 3 שניות המתנה לפי בקשתך
        except: time.sleep(3)
    return f"שגיאה בייצור הפרק '{chapter_title}'. נא לנסות שוב."

# --- 4. עיצוב וורד אקדמי (שלב 2 של ה-Master Prompt) ---
def create_master_doc(title, author, content_list, lang):
    doc = Document()
    font_name = 'David' if lang == "עברית" else 'Times New Roman'
    
    # הגדרת שוליים (2.5 ס"מ)
    for section in doc.sections:
        section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = Inches(0.98)

    # עמוד שער
    doc.add_paragraph('\n\n\n')
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(f"עבודה סמינריונית בנושא:\n{title}")
    r.bold = True
    r.font.size = Pt(24)
    r.font.name = font_name
    
    doc.add_paragraph('\n\n\n')
    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.add_run(f"מוגש על ידי: {author}\nמוסד אקדמי: אורות").font.name = font_name
    doc.add_page_break()

    # הוספת תקציר (תקציר נכתב כפרק ראשון לאחר שער)
    
    # גוף העבודה
    for text in content_list:
        for line in text.split('\n'):
            line = line.strip()
            if not line: continue
            
            if line.startswith('# '):
                h = doc.add_heading(line.replace('#', '').strip(), level=1)
                h.alignment = WD_ALIGN_PARAGRAPH.RIGHT if lang == "עברית" else WD_ALIGN_PARAGRAPH.LEFT
            elif line.startswith('## '):
                sh = doc.add_heading(line.replace('##', '').strip(), level=2)
                sh.alignment = WD_ALIGN_PARAGRAPH.RIGHT if lang == "עברית" else WD_ALIGN_PARAGRAPH.LEFT
            else:
                p = doc.add_paragraph(line.replace('*', ''))
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if lang == "עברית" else WD_ALIGN_PARAGRAPH.LEFT
                p.paragraph_format.line_spacing = 1.5 # מרווח 1.5 כפי שביקשת
                p.paragraph_format.alignment = 3 # Justify
        doc.add_page_break()
    return doc

# --- 5. ממשק המשתמש (UI) ---
lang = st.radio("🌐 שפת ממשק:", ["עברית", "English"], horizontal=True)
if lang == "עברית":
    st.markdown("<style>.block-container { direction: rtl; text-align: right; }</style>", unsafe_allow_html=True)

if 'logged_in' not in st.session_state: st.session_state['logged_in'] = False

st.title("🎓 Seminar Architect PRO")

if not st.session_state['logged_in']:
    email_input = st.text_input("אימייל להתחברות / Email:")
    if st.button("🚀 כניסה / Login"):
        if email_input and is_valid_email(email_input):
            st.session_state['user_email'] = email_input.lower()
            st.session_state['logged_in'] = True
            st.rerun()
        else: st.error("נא להזין אימייל תקין (למשל: name@gmail.com).")
else:
    st.sidebar.success(f"מחובר כ: {st.session_state['user_email']}")
    col1, col2 = st.columns(2)
    with col1:
        topic = st.text_input("נושא הסמינריון:")
        name = st.text_input("שם הסטודנט:")
    with col2:
        extra = st.text_area("דגשים ספציפיים (שלב 1):", height=100)
    uploaded = st.file_uploader("העלאת הנחיות מרצה (PDF/TXT):", type=['pdf', 'txt'])

    if st.button("🚀 צא לדרך! תן לפרופסור לבנות לך סמינריון מנצח"):
        if not topic: st.error("הזן נושא!")
        else:
            st.warning("⚠️ **המערכת מייצרת את העבודה לפי ה-Master Prompt. נא לא לסגור את החלון!**")
            notes = extract_text_from_file(uploaded)
            
            status_text = st.empty()
            time_est = st.empty()
            progress_bar = st.progress(0)
            
            # שלב האבחון ובניית השלד הדינמי
            status_text.info("⏳ הפרופסור מאבחן את הנושא ובונה שלד פרקים ייחודי...")
            chapters = generate_dynamic_outline(topic, extra, api_key, lang)
            
            generated_content = []
            total = len(chapters) + 1 # +1 עבור התקציר שיכתב בסוף
            
            # כתיבת הפרקים הדינמיים
            for i, head in enumerate(chapters):
                pct = int((i / total) * 100)
                rem = (total - i) * 45
                status_text.info(f"⏳ ({pct}%) כותב פרק עומק: **{head}**...")
                time_est.markdown(f"⏱️ **זמן משוער לסיום:** כ-{rem} שניות")
                
                content = call_gemini_master_professor(head, f"Topic: {topic}. Extra: {extra}. Guidelines: {notes}", api_key, lang)
                generated_content.append(content)
                progress_bar.progress((i + 1) / total)
                time.sleep(3)
            
            # שלב התקציר (שלב 3 ב-Master Prompt - נכתב בסוף)
            status_text.info("⏳ (90%) מסכם את העבודה ומייצר תקציר אקדמי...")
            summary_prompt = f"Write a 1-page Abstract for the following seminar content: {str(generated_content[:2])}. Focus on research questions and main conclusions."
            abstract = call_gemini_master_professor("תקציר", summary_prompt, api_key, lang)
            generated_content.insert(0, abstract) # הכנסה לראש העבודה
            
            progress_bar.progress(1.0)
            status_text.success("🎉 (100%) הסמינריון הושלם ברמה אקדמית גבוהה!")
            
            doc = create_master_doc(topic, name, generated_content, lang)
            buf = io.BytesIO()
            doc.save(buf)
            buf.seek(0)
            st.download_button("📥 הורד סמינריון מושלם (Word)", buf, f"Seminar_{name}.docx")
            st.balloons()
