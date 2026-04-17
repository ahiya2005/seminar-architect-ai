import streamlit as st
import requests
import json
import io
import time
import re
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import PyPDF2

st.set_page_config(page_title="Seminar Architect PRO", page_icon="🎓", layout="wide")

api_key = st.secrets.get("GEMINI_API_KEY")

st.markdown("""
    <style>
    header {visibility: hidden;}
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    div[data-testid="InputInstructions"] { display: none !important; }
    .stButton>button { width: 100%; background-color: #2c3e50; color: white; border-radius: 10px; font-weight: bold; }
    .stProgress > div > div > div {
        background-image: linear-gradient(45deg, rgba(255, 255, 255, .15) 25%, transparent 25%, transparent 50%, rgba(255, 255, 255, .15) 50%, rgba(255, 255, 255, .15) 75%, transparent 75%, transparent);
        background-size: 1rem 1rem;
        animation: progress-bar-stripes 1s linear infinite;
    }
    @keyframes progress-bar-stripes { from { background-position: 1rem 0; } to { background-position: 0 0; } }
    </style>
""", unsafe_allow_html=True)

def is_valid_email(email):
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

# --- פונקציה חדשה: בניית ראשי פרקים דינמיים ---
def generate_dynamic_outline(topic, key, lang="עברית"):
    """מייצר כותרות אקדמיות שנגזרות ישירות מהנושא, במקום כותרות גנריות"""
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key={key}"
    
    if lang == "עברית":
        prompt = f"צור רשימה של 6 כותרות לפרקים עבור עבודה סמינריונית בנושא '{topic}'.\nחובה: אל תשתמש במילים גנריות כמו 'סקירת ספרות', 'מתודולוגיה' או 'ממצאים'. כל כותרת חייבת להיות ספציפית ואקדמית. הפרק הראשון צריך להיות סוג של מבוא מותאם לנושא, והפרק האחרון חייב להיות המילה 'ביבליוגרפיה'.\nהחזר אך ורק את 6 הכותרות, מופרדות בפסיק (,), ללא שום טקסט נוסף."
    else:
        prompt = f"Create 6 specific academic chapter titles for a seminar on '{topic}'.\nRule: Do NOT use generic words like 'Methodology' or 'Findings'. Make them specific to the topic. The last must be 'Bibliography'.\nReturn ONLY the titles separated by commas (,)."

    payload = {"contents": [{"parts": [{"text": prompt}]}], "generationConfig": {"temperature": 0.3, "maxOutputTokens": 200}}
    
    try:
        response = requests.post(url, json=payload, timeout=30)
        if response.status_code == 200:
            text = response.json()['candidates'][0]['content']['parts'][0]['text']
            chapters = [c.strip() for c in text.split(',') if c.strip()]
            if len(chapters) >= 3:
                return chapters
    except: pass
    
    # גיבוי למקרה שה-AI מסתבך
    if lang == "עברית": return [f"מבוא ורקע היסטורי: {topic}", f"גישות תיאורטיות מרכזיות ל{topic}", f"ניתוח עומק והשלכות של {topic}", f"מקרי בוחן ומגמות עכשוויות", f"סיכום, דיון ומסקנות", "ביבליוגרפיה"]
    else: return [f"Introduction to {topic}", f"Theoretical Frameworks of {topic}", f"In-depth Analysis of {topic}", f"Current Trends and Case Studies", f"Conclusion and Discussion", "Bibliography"]

# --- מנוע כתיבת התוכן (מעודכן ל-3 ניסיונות, 3 שניות) ---
def call_gemini_with_retry(chapter_title, topic, extra, guidelines, key, lang="עברית", max_retries=3):
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key={key}"
    
    system_instruction = f"""
    ROLE: Senior Academic Writer.
    TASK: Write the chapter titled '{chapter_title}' for a seminar on '{topic}'.
    RULES:
    1. TITLE: Start strictly with '# {chapter_title}'.
    2. SUB-TOPICS: Break the chapter into 3-4 highly specific sub-topics using '##'.
    3. DEPTH: Write very long, professional academic content under each sub-topic.
    4. SOURCES: Embed real APA academic citations in the text.
    5. GUIDELINES: {extra}. {guidelines}
    6. NO META-TEXT: Start directly with the # title.
    """
    if lang == "עברית": system_instruction += " 7. כתיבה בעברית אקדמית ותקנית בלבד."
    
    payload = {"contents": [{"parts": [{"text": system_instruction}]}], "generationConfig": {"temperature": 0.4, "maxOutputTokens": 5000}}
    
    last_error = ""
    for attempt in range(max_retries):
        try:
            response = requests.post(url, json=payload, timeout=90)
            if response.status_code == 200:
                text = response.json()['candidates'][0]['content']['parts'][0]['text']
                if len(text) > 300: return text.strip()
            else:
                last_error = f"API Error {response.status_code}"
            time.sleep(3) # שונה ל-3 שניות לפי בקשתך
        except Exception as e:
            last_error = str(e)
            time.sleep(3)
            
    return f"שגיאה בייצור התוכן. ניסינו 3 פעמים ללא הצלחה. (שגיאה אחרונה: {last_error})"

def create_pro_doc(title, author, content_list, lang):
    doc = Document()
    font_name = 'David' if lang == "עברית" else 'Times New Roman'
    
    doc.add_paragraph('\n\n\n')
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(f"עבודה סמינריונית בנושא:\n{title}" if lang == "עברית" else f"Academic Seminar:\n{title}")
    r.bold = True
    r.font.size = Pt(24)
    r.font.name = font_name
    doc.add_paragraph('\n\n')
    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.add_run(f"מגיש: {author}" if lang == "עברית" else f"By: {author}").font.name = font_name
    doc.add_page_break()

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
                p.paragraph_format.line_spacing = 1.5
        doc.add_page_break()
    return doc

lang = st.radio("🌐 שפת ממשק / System Language:", ["עברית", "English"], horizontal=True)
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
        else: st.error("נא להזין אימייל תקין עם סיומת (למשל: .com).")
else:
    st.sidebar.success(f"מחובר כ: {st.session_state['user_email']}")
    
    col1, col2 = st.columns(2)
    with col1:
        topic = st.text_input("נושא הסמינריון:")
        name = st.text_input("שם הסטודנט:")
    with col2:
        extra = st.text_area("הנחיות ספציפיות:", height=100)
    
    uploaded = st.file_uploader("העלאת הנחיות (PDF/TXT):", type=['pdf', 'txt'])

    if st.button("🚀 צא לדרך! בנה לי סמינריון מנצח"):
        if not topic: st.error("הזן נושא!")
        else:
            st.warning("⚠️ **המערכת מייצרת את העבודה. נא לא לסגור או לרענן את החלון!**")
            notes = extract_text_from_file(uploaded)
            
            # שלב 1: בניית ראשי הפרקים הדינמיים
            status_text = st.empty()
            time_est = st.empty()
            progress_bar = st.progress(0)
            
            status_text.info("⏳ מנתח את הנושא ובונה ראשי פרקים אקדמיים...")
            chapters = generate_dynamic_outline(topic, api_key, lang)
            
            # שלב 2: ייצור התוכן לכל פרק
            generated_content = []
            
            for i, head in enumerate(chapters):
                pct = int((i / len(chapters)) * 100)
                rem = (len(chapters) - i) * 35 
                
                # עכשיו התצוגה תראה לך כותרות אקדמיות אמיתיות שקשורות לנושא!
                status_text.info(f"⏳ ({pct}%) כותב כרגע את פרק: **{head}**...")
                time_est.markdown(f"⏱️ **זמן משוער לסיום:** כ-{rem} שניות")
                
                content = call_gemini_with_retry(head, topic, extra, notes, api_key, lang)
                generated_content.append(content)
                
                progress_bar.progress((i + 1) / len(chapters))
                time.sleep(3) # השהייה של 3 שניות כפי שביקשת
            
            time_est.empty()
            status_text.success("🎉 העבודה מוכנה ומעוצבת!")
            
            doc = create_pro_doc(topic, name, generated_content, lang)
            buf = io.BytesIO()
            doc.save(buf)
            buf.seek(0)
            st.download_button("📥 הורד עבודה סמינריונית (Word)", buf, f"Seminar_{name}.docx")
            st.balloons()
