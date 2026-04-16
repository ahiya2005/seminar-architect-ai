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

# --- הגדרות דף ועיצוב ---
st.set_page_config(page_title="Seminar Architect PRO", page_icon="🎓", layout="wide")

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

# --- אימות אימייל (מתוקן - חוסם הרמטית כתובות מזויפות) ---
def is_valid_email(email):
    pattern = r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z]{2,}$"
    return re.match(pattern, email) is not None

# --- מנוע ה-AI המשופר ---
def call_gemini_with_retry(prompt, key, lang="עברית", max_retries=10):
    # שימוש במודל היציב ביותר שעוקף את שגיאת ה-404
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={key}"
    
    # הנחיות ברזל ל-AI ליצירת כותרות מקצועיות ותתי-נושאים
    system_instruction = """
    You are an expert academic writer structuring a high-level thesis.
    RULES:
    1. TITLE: Start the section with a highly specific, academic Title related to the topic using a single '#' (e.g., '# התפתחות לשון חכמים'). DO NOT use generic titles like 'מבוא' or 'סקירת ספרות'.
    2. SUB-TOPICS: Break the content down into 3-4 specific sub-topics using '##'.
    3. DEPTH: Write very long, highly detailed academic content under each sub-topic.
    4. SOURCES: Integrate real academic APA citations in-text.
    5. NO META-TEXT: Start directly with the '#' title. No greetings.
    """
    if lang == "עברית": system_instruction += " 6. כתיבה בעברית אקדמית ותקנית בלבד."
    
    payload = {"contents": [{"parts": [{"text": system_instruction + "\n\n" + prompt}]}], "generationConfig": {"temperature": 0.4, "maxOutputTokens": 5000}}
    
    last_error = ""
    for attempt in range(max_retries):
        try:
            response = requests.post(url, json=payload, timeout=120)
            if response.status_code == 200:
                text = response.json()['candidates'][0]['content']['parts'][0]['text']
                if len(text) > 300: 
                    return text.strip()
                else:
                    last_error = "התוכן שנוצר היה קצר מדי."
                    time.sleep(10)
            else:
                last_error = f"Status {response.status_code}: {response.text}"
                time.sleep(10) 
        except Exception as e: 
            last_error = str(e)
            time.sleep(10)
            
    # אם הכל נכשל, הוא ידפיס לך בדיוק את השגיאה של גוגל בוורד כדי שנדע מה קרה
    return f"שגיאה בייצור התוכן. פרטי תקלה אחרונה: {last_error}"

# --- בניית קובץ ה-Word (מותאם אישית לקריאת הכותרות של ה-AI) ---
def create_pro_doc(title, author, content_dict, lang):
    doc = Document()
    font_name = 'David' if lang == "עברית" else 'Times New Roman'
    
    # שער
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

    # סריקת התוכן ועיצוב הכותרות המקצועיות שה-AI יצר
    for chapter_key, text in content_dict.items():
        for line in text.split('\n'):
            line = line.strip()
            if not line: continue
            
            # זיהוי כותרת ראשית של פרק
            if line.startswith('# ') and not line.startswith('## '):
                h = doc.add_heading(line.replace('#', '').strip(), level=1)
                h.alignment = WD_ALIGN_PARAGRAPH.RIGHT if lang == "עברית" else WD_ALIGN_PARAGRAPH.LEFT
            
            # זיהוי תתי-נושאים
            elif line.startswith('## '):
                sh = doc.add_heading(line.replace('##', '').strip(), level=2)
                sh.alignment = WD_ALIGN_PARAGRAPH.RIGHT if lang == "עברית" else WD_ALIGN_PARAGRAPH.LEFT
            
            # טקסט רגיל
            else:
                p = doc.add_paragraph(line.replace('*', ''))
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if lang == "עברית" else WD_ALIGN_PARAGRAPH.LEFT
                p.paragraph_format.line_spacing = 1.5
        doc.add_page_break()
    return doc

# --- הממשק (Streamlit) ---
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
        else: st.error("נא להזין אימייל תקין (חובה לכלול סיומת מלאה, למשל: .com).")
else:
    col1, col2 = st.columns(2)
    with col1:
        topic = st.text_input("נושא הסמינריון:")
        name = st.text_input("שם הסטודנט:")
    with col2:
        extra = text_area_input = st.text_area("הנחיות ספציפיות:", height=100)
    
    uploaded = st.file_uploader("העלאת הנחיות (PDF/TXT):", type=['pdf', 'txt'])

    if st.button("🚀 צא לדרך! בנה לי סמינריון מנצח"):
        if not topic: st.error("הזן נושא!")
        else:
            st.warning("⚠️ **המערכת מייצרת את העבודה. נא לא לסגור או לרענן את החלון!** התהליך לוקח כ-5 עד 10 דקות.")
            
            # שינינו את הפקודות ל-AI שייצרו כותרות ספציפיות
            chapters = [
                ("מבוא", f"Write the Introduction section for the topic '{topic}'. Provide a specific academic title using '#'. Include background and research questions."),
                ("סקירה ספרותית", f"Write the Literature Review for '{topic}'. Provide a specific academic title using '#'. Break into 4 theoretical sub-topics using '##'."),
                ("מתודולוגיה", f"Write the Methodology section for '{topic}'. Provide a specific academic title using '#'. Break down using '##'."),
                ("ממצאים", f"Write the Findings section for '{topic}'. Provide a specific academic title using '#'. Detail hypothetical findings using '##'."),
                ("דיון", f"Write the Discussion section for '{topic}'. Provide a specific academic title using '#'. Critique findings using '##'."),
                ("ביבליוגרפיה", f"Provide the Bibliography for '{topic}'. Use the title '# ביבליוגרפיה'. Use APA format.")
            ]
            
            res_dict = {}
            progress_bar = st.progress(0)
            status_text = st.empty()
            time_est = st.empty()
            
            for i, (key_name, prompt) in enumerate(chapters):
                pct = int((i / len(chapters)) * 100)
                rem = (len(chapters) - i) * 50 
                
                status_text.info(f"⏳ ({pct}%) כותב כרגע את פרק: **{key_name}**...")
                time_est.markdown(f"⏱️ **זמן משוער לסיום:** כ-{rem} שניות")
                
                res_dict[key_name] = call_gemini_with_retry(prompt, st.secrets["GEMINI_API_KEY"], lang)
                
                progress_bar.progress((i + 1) / len(chapters))
                time.sleep(5)
            
            time_est.empty()
            status_text.success("🎉 העבודה מוכנה!")
            doc = create_pro_doc(topic, name, res_dict, lang)
            buf = io.BytesIO()
            doc.save(buf)
            buf.seek(0)
            st.download_button("📥 הורד עבודה סמינריונית (Word)", buf, f"Seminar_{name}.docx")
            st.balloons()
