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
from streamlit_gsheets import GSheetsConnection

# --- 1. הגדרות דף ועיצוב גלובלי ---
st.set_page_config(page_title="Seminar Architect PRO", page_icon="🎓", layout="wide")

api_key = st.secrets.get("GEMINI_API_KEY")
gsheet_url = st.secrets.get("GSHEETS_URL")

# ניסיון חיבור לגוגל שיטס
try:
    conn = st.connection("gsheets", type=GSheetsConnection)
except:
    conn = None

st.markdown("""
    <style>
    /* הסתרת רכיבי מערכת של Streamlit */
    header {visibility: hidden;}
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    div[data-testid="InputInstructions"] { display: none !important; }
    
    /* עיצוב הכפתור הראשי */
    .stButton>button { width: 100%; background-color: #2c3e50; color: white; border-radius: 10px; font-weight: bold; }
    .stButton>button:hover { background-color: #34495e; }
    
    /* אנימציית סרגל ההתקדמות (Pulse) */
    .stProgress > div > div > div {
        background-image: linear-gradient(45deg, rgba(255, 255, 255, .15) 25%, transparent 25%, transparent 50%, rgba(255, 255, 255, .15) 50%, rgba(255, 255, 255, .15) 75%, transparent 75%, transparent);
        background-size: 1rem 1rem;
        animation: progress-bar-stripes 1s linear infinite;
    }
    @keyframes progress-bar-stripes { from { background-position: 1rem 0; } to { background-position: 0 0; } }
    </style>
""", unsafe_allow_html=True)


# --- 2. פונקציות בסיס נתונים ואימות ---
def is_valid_email(email):
    """בדיקה נוקשה שכתובת האימייל באמת הגיונית"""
    pattern = r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$"
    return re.match(pattern, email) is not None

def save_project_to_db(email, topic, name, extra, status="In Progress", content={}):
    """שמירת נתוני הפרויקט במסד הנתונים של גוגל"""
    if conn is None or not gsheet_url:
        return
    try:
        df = conn.read(spreadsheet=gsheet_url)
        new_row = {"email": email, "topic": topic, "name": name, "extra": extra, "status": status, "content_json": json.dumps(content)}
        df = df[~((df['email'] == email) & (df['topic'] == topic))]
        df = df.append(new_row, ignore_index=True)
        conn.update(spreadsheet=gsheet_url, data=df)
    except: pass

def get_user_projects(email):
    """שליפת פרויקטים קודמים של המשתמש"""
    if conn is None or not gsheet_url:
        return []
    try:
        df = conn.read(spreadsheet=gsheet_url)
        return df[df['email'] == email]
    except: return []


# --- 3. פונקציות ליבה (קריאת קבצים, AI, וייצור וורד) ---
def extract_text_from_file(uploaded_file):
    """חילוץ טקסט מקובץ הנחיות מרצה"""
    if uploaded_file is None: return ""
    try:
        if uploaded_file.name.endswith('.pdf'):
            reader = PyPDF2.PdfReader(uploaded_file)
            return "\n".join([page.extract_text() for page in reader.pages])
        return uploaded_file.getvalue().decode("utf-8")
    except Exception as e: return f"Error: {str(e)}"

def call_gemini_with_retry(prompt, key, lang="עברית", max_retries=10):
    """קריאה לבינה המלאכותית עם מנגנון חסינות לתקלות ובקרת איכות"""
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key={key}"
    system_instruction = "ROLE: Senior Academic Researcher. RULES: 1. STRUCTURE: Breakdown chapter into 3-4 sub-topics using '##'. 2. DEPTH: Write very long, detailed content. 3. SOURCES: Real academic sources only. 4. NO META-TEXT."
    if lang == "עברית": system_instruction += " 5. כתיבה בעברית אקדמית בלבד."
    payload = {"contents": [{"parts": [{"text": system_instruction + "\n\n" + prompt}]}], "generationConfig": {"temperature": 0.5, "maxOutputTokens": 5000}}
    
    for attempt in range(max_retries):
        try:
            response = requests.post(url, headers={'Content-Type': 'application/json'}, data=json.dumps(payload), timeout=100)
            if response.status_code == 200:
                text = response.json()['candidates'][0]['content']['parts'][0]['text']
                # בקרת איכות - אם הטקסט קצר מדי, הוא מנסה שוב
                if len(text) > 400: return text.strip()
            time.sleep(10)
        except: time.sleep(10)
    return "שגיאה בייצור התוכן." if lang == "עברית" else "Error generating content."

def create_pro_doc(title, author, content_dict, lang):
    """יצירת מסמך וורד תקני ומעוצב"""
    doc = Document()
    font_name = 'David' if lang == "עברית" else 'Times New Roman'
    
    # עמוד שער
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

    # תוכן הפרקים
    for chapter, text in content_dict.items():
        h = doc.add_heading(chapter, level=1)
        h.alignment = WD_ALIGN_PARAGRAPH.RIGHT if lang == "עברית" else WD_ALIGN_PARAGRAPH.LEFT
        for line in text.split('\n'):
            line = line.strip()
            if not line: continue
            if line.startswith('##'):
                sh = doc.add_heading(line.replace('##', '').strip(), level=2)
                sh.alignment = WD_ALIGN_PARAGRAPH.RIGHT if lang == "עברית" else WD_ALIGN_PARAGRAPH.LEFT
            else:
                p = doc.add_paragraph(line.replace('*', ''))
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if lang == "עברית" else WD_ALIGN_PARAGRAPH.LEFT
                p.paragraph_format.line_spacing = 1.5
        doc.add_page_break()
    return doc


# --- 4. הממשק הראשי (UI) ---
lang = st.radio("🌐 שפת ממשק:", ["עברית", "English"], horizontal=True)
dir_css = "rtl" if lang == "עברית" else "ltr"
st.markdown(f'<div style="direction: {dir_css}; text-align: {"right" if lang=="עברית" else "left"};">', unsafe_allow_html=True)

if 'logged_in' not in st.session_state: 
    st.session_state['logged_in'] = False

st.title("🎓 Seminar Architect PRO")

# מסך התחברות
if not st.session_state['logged_in']:
    email_input = st.text_input("הכנס אימייל להתחברות:" if lang == "עברית" else "Enter your email:")
    if st.button("🚀 כניסה ושמירת נתונים" if lang == "עברית" else "🚀 Login"):
        if email_input:
            if is_valid_email(email_input):
                st.session_state['user_email'] = email_input.lower()
                st.session_state['logged_in'] = True
                st.rerun()
            else:
                st.error("כתובת האימייל אינה תקינה. נא להזין כתובת אמיתית (לדוגמה: name@gmail.com)." if lang == "עברית" else "Invalid email format.")
        else:
            st.warning("נא להזין אימייל." if lang == "עברית" else "Please enter your email.")
            
# מסך הפרויקט (לאחר התחברות)
else:
    user_email = st.session_state['user_email']
    st.sidebar.success(f"מחובר כ: {user_email}")
    
    # תפריט פרויקטים קודמים
    projects = get_user_projects(user_email)
    if len(projects) > 0:
        with st.expander("📁 הפרויקטים הקודמים שלך" if lang == "עברית" else "📁 Your Previous Projects"):
            for i, row in projects.iterrows():
                if st.button(f"המשך פרויקט: {row['topic']}", key=f"proj_{i}"):
                    st.session_state['current_topic'] = row['topic']
                    st.session_state['current_name'] = row['name']
                    st.info("נטען בהצלחה! אפשר להמשיך להפיק את העבודה." if lang == "עברית" else "Loaded successfully!")

    st.divider()
    col1, col2 = st.columns(2)
    with col1:
        topic = st.text_input("נושא הסמינריון:" if lang == "עברית" else "Seminar Topic:", value=st.session_state.get('current_topic', ''))
        name = st.text_input("שם הסטודנט:" if lang == "עברית" else "Student Name:", value=st.session_state.get('current_name', ''))
    with col2:
        extra = st.text_area("הנחיות ספציפיות:" if lang == "עברית" else "Focus Instructions:", height=100)

    uploaded = st.file_uploader("העלאת הנחיות מרצה (PDF/TXT):" if lang == "עברית" else "Upload Guidelines:", type=['pdf', 'txt'])

    # כפתור הפעלה ממריץ
    if st.button("🚀 צא לדרך! תן לנו לבנות לך סמינריון מנצח" if lang == "עברית" else "🚀 Let's Go! Generate My Seminar"):
        if not topic: 
            st.error("חובה להזין נושא!" if lang == "עברית" else "Topic is required!")
        else:
            save_project_to_db(user_email, topic, name, extra, "Started")
            
            # אזהרה בולטת מופרדת
            st.warning("⚠️ **המערכת בונה את העבודה. תהליך זה אורך מספר דקות. נא לא לסגור או לרענן את החלון!**" if lang == "עברית" else "⚠️ **Generating... DO NOT close or refresh this window!**")
            
            notes = extract_text_from_file(uploaded)
            full_context = f"Topic: {topic}. Notes: {extra}. Guidelines: {notes}"
            
            # מבנה הפרקים
            if lang == "עברית":
                chapters = [
                    ("מבוא", "Write Intro. Breakdown: Background, Research Question, Rationale."),
                    ("סקירה ספרותית", "Write Lit Review. Breakdown into 4 theoretical topics."),
                    ("מתודולוגיה", "Write Methodology. Breakdown: Tools, Population, Method."),
                    ("ממצאים", "Write Findings. Detailed hypothetical breakdown."),
                    ("דיון ומסקנות", "Write Discussion. Critique findings vs theories."),
                    ("ביבליוגרפיה", "Final Bibliography. APA style. Real sources only.")
                ]
            else:
                chapters = [
                    ("Introduction", "Write Intro. Breakdown: Background, Research Question, Rationale."),
                    ("Literature Review", "Write Lit Review. Breakdown into 4 theoretical topics."),
                    ("Methodology", "Write Methodology. Breakdown: Tools, Population, Method."),
                    ("Findings", "Write Findings. Detailed hypothetical breakdown."),
                    ("Discussion and Conclusions", "Write Discussion. Critique findings vs theories."),
                    ("Bibliography", "Final Bibliography. APA style. Real sources only.")
                ]
            
            res_dict = {}
            progress_bar = st.progress(0)
            status_text = st.empty()
            time_estimate = st.empty()
            
            # לולאת בניית הפרקים
            for i, (head, prompt) in enumerate(chapters):
                pct = int((i / len(chapters)) * 100)
                status_text.info(f"⏳ ({pct}%) כותב כרגע: **{head}**..." if lang == "עברית" else f"⏳ ({pct}%) Writing: **{head}**...")
                time_estimate.markdown(f"⏱️ **זמן משוער לסיום:** כ-{(len(chapters)-i)*45} שניות" if lang == "עברית" else f"⏱️ **Estimated time:** ~{(len(chapters)-i)*45} seconds")
                
                res_dict[head] = call_gemini_with_retry(f"Topic: {topic}. Notes: {extra}. Task: {prompt}", api_key, lang)
                
                save_project_to_db(user_email, topic, name, extra, f"Writing {head}", res_dict)
                progress_bar.progress((i + 1) / len(chapters))
                time.sleep(10)
            
            time_estimate.empty()
            status_text.success("🎉 (100%) העבודה מוכנה ומעוצבת!" if lang == "עברית" else "🎉 (100%) Ready!")
            save_project_to_db(user_email, topic, name, extra, "Completed", res_dict)
            
            doc = create_pro_doc(topic, name, res_dict, lang)
            buf = io.BytesIO()
            doc.save(buf)
            buf.seek(0)
            
            st.download_button("📥 הורד עבודה סמינריונית (Word)" if lang == "עברית" else "📥 Download Paper (Word)", buf, f"Seminar_{name}.docx")
            st.balloons()
