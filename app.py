import streamlit as st
import requests
import json
import io
import time
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import PyPDF2

# --- הגדרות דף ---
st.set_page_config(page_title="Seminar Architect PRO", page_icon="🎓", layout="wide")

api_key = st.secrets.get("GEMINI_API_KEY")

# --- CSS גלובלי להסתרת רכיבי מערכת ועיצוב נקי ---
st.markdown("""
    <style>
    header {visibility: hidden;}
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    div[data-testid="InputInstructions"] { display: none !important; }
    .stButton>button { width: 100%; background-color: #2c3e50; color: white; border-radius: 10px; font-weight: bold; }
    .stButton>button:hover { background-color: #34495e; }
    /* אנימציה לסרגל הטעינה (Pulse) */
    .stProgress > div > div > div {
        background-image: linear-gradient(45deg, rgba(255, 255, 255, .15) 25%, transparent 25%, transparent 50%, rgba(255, 255, 255, .15) 50%, rgba(255, 255, 255, .15) 75%, transparent 75%, transparent);
        background-size: 1rem 1rem;
        animation: progress-bar-stripes 1s linear infinite;
    }
    @keyframes progress-bar-stripes {
        from { background-position: 1rem 0; }
        to { background-position: 0 0; }
    }
    </style>
""", unsafe_allow_html=True)

# --- ממשק שפות וכיוון דף ---
lang = st.radio("🌐 שפת מערכת / System Language:", ["עברית", "English"], horizontal=True)

if lang == "עברית":
    st.markdown("""
        <style>
        .block-container { direction: rtl; text-align: right; }
        p, div, span, label, h1, h2, h3 { direction: rtl; text-align: right; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
        .stTextInput input, .stTextArea textarea { direction: rtl; text-align: right; }
        </style>
        """, unsafe_allow_html=True)
else:
    st.markdown("""
        <style>
        .block-container { direction: ltr; text-align: left; }
        p, div, span, label, h1, h2, h3 { direction: ltr; text-align: left; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
        .stTextInput input, .stTextArea textarea { direction: ltr; text-align: left; }
        </style>
        """, unsafe_allow_html=True)

def extract_text_from_file(uploaded_file):
    if uploaded_file is None: return ""
    try:
        if uploaded_file.name.endswith('.pdf'):
            reader = PyPDF2.PdfReader(uploaded_file)
            return "\n".join([page.extract_text() for page in reader.pages])
        return uploaded_file.getvalue().decode("utf-8")
    except Exception as e: return f"Error: {str(e)}"

def call_gemini_with_retry(prompt, key, lang="עברית", max_retries=10):
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key={key}"
    
    system_instruction = """
    ROLE: Senior Academic Researcher.
    RULES:
    1. STRUCTURE: Breakdown every chapter into 3-4 deep sub-topics using '##'. 
    2. DEPTH: Provide extensive academic analysis. Write very long, deep and detailed content.
    3. SOURCES: Use ONLY real academic sources. APA style citations inside text.
    4. QUALITY: Use perfect grammar, punctuation, and formal academic tone. 
    5. NO META-TEXT: Start the content immediately. No "Sure" or "Here is".
    6. VALIDATION: Ensure logical flow between paragraphs.
    """
    if lang == "עברית": system_instruction += " 7. כתיבה בעברית אקדמית גבוהה ומדעית בלבד."
    else: system_instruction += " 7. Write in high-level formal academic English."

    payload = {
        "contents": [{"parts": [{"text": system_instruction + "\n\n" + prompt}]}],
        "generationConfig": {"temperature": 0.5, "maxOutputTokens": 5000}
    }
    
    for attempt in range(max_retries):
        try:
            response = requests.post(url, headers={'Content-Type': 'application/json'}, data=json.dumps(payload), timeout=100)
            if response.status_code == 200:
                result = response.json()
                if 'candidates' in result and len(result['candidates']) > 0:
                    text = result['candidates'][0]['content']['parts'][0]['text']
                    if len(text) < 400:
                        time.sleep(10)
                        continue 
                    lines = text.split('\n')
                    if lines and any(word in lines[0] for word in ["להלן", "הנה", "Here", "Sure"]):
                        lines = lines[1:]
                    return "\n".join(lines).strip()
            elif response.status_code == 429: # במקרה של עומס בגוגל נחכה יותר
                time.sleep(30)
            else:
                time.sleep(10)
        except Exception:
            time.sleep(10)
    return "שגיאה: לא הצלחנו לייצר את הפרק באיכות הנדרשת. ייתכן שיש עומס. אנא נסה שוב." if lang == "עברית" else "Error generating content."

def create_pro_doc(title, author, content_dict, lang):
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

# --- ניהול Session State (הכנה לשמירת משתמשים) ---
if 'logged_in' not in st.session_state:
    st.session_state['logged_in'] = False
if 'user_email' not in st.session_state:
    st.session_state['user_email'] = ""

st.title("🎓 Seminar Architect PRO")

if not st.session_state['logged_in']:
    email = st.text_input("אימייל להתחברות / Email:")
    if st.button("🚀 כניסה מהירה / Login"):
        if email:
            st.session_state['user_email'] = email
            st.session_state['logged_in'] = True
            # TODO: כאן נוסיף בעתיד את החיבור ל-Google Sheets לשמירת המייל במסד הנתונים
            st.rerun()
        else:
            st.warning("נא להזין אימייל." if lang == "עברית" else "Please enter your email.")
else:
    col1, col2 = st.columns(2)
    with col1:
        topic = st.text_input("נושא הסמינריון:" if lang == "עברית" else "Seminar Topic:")
        name = st.text_input("שם הסטודנט:" if lang == "עברית" else "Student Name:")
    with col2:
        extra = st.text_area("תיאור והנחיות ספציפיות (למשל: דגש על פסיכולוגיה):" if lang == "עברית" else "Focus Instructions:", height=100)
    
    uploaded = st.file_uploader("העלאת הנחיות מרצה (PDF/TXT):" if lang == "עברית" else "Upload Guidelines:", type=['pdf', 'txt'])

    btn_text = "🚀 צא לדרך! תן לנו לבנות לך סמינריון מנצח" if lang == "עברית" else "🚀 Let's Go! Generate My Seminar"
    if st.button(btn_text):
        if not topic: 
            st.error("חובה להזין נושא!" if lang == "עברית" else "Topic is required!")
        else:
            wait_msg = "⚠️ המערכת בונה כרגע את העבודה. תהליך זה אורך מספר דקות ומחייב חיבור רציף. **נא לא לסגור או לרענן את החלון!**" if lang == "עברית" else "⚠️ Generating... Please DO NOT close or refresh this window!"
            st.warning(wait_msg)
            
            notes = extract_text_from_file(uploaded)
            full_context = f"Topic: {topic}. Notes: {extra}. Guidelines: {notes}"
            
            if lang == "עברית":
                chapters = [
                    ("מבוא", "Write Intro. Breakdown: Background, Research Question, Rationale."),
                    ("סקירה ספרותית", "Write Lit Review. Breakdown into 4 theoretical topics."),
                    ("מתודולוגיה", "Write Methodology. Breakdown: Tools, Population, Method."),
                    ("ממצאים", "Write Findings. Detailed hypothetical breakdown."),
                    ("דיון ומסקנות", "Write Discussion. Critique findings vs theories."),
                    ("ביבליוגרפיה", "Final Bibliography. APA style. Real sources only. Alphabetical.")
                ]
            else:
                chapters = [
                    ("Introduction", "Write Intro. Breakdown: Background, Research Question, Rationale."),
                    ("Literature Review", "Write Lit Review. Breakdown into 4 theoretical topics."),
                    ("Methodology", "Write Methodology. Breakdown: Tools, Population, Method."),
                    ("Findings", "Write Findings. Detailed hypothetical breakdown."),
                    ("Discussion and Conclusions", "Write Discussion. Critique findings vs theories."),
                    ("Bibliography", "Final Bibliography. APA style. Real sources only. Alphabetical.")
                ]
            
            res_dict = {}
            progress_bar = st.progress(0)
            status_text = st.empty()
            time_estimate = st.empty() # רכיב השעון
            
            total_chapters = len(chapters)
            estimated_time_per_chapter = 45 # הערכה: 45 שניות לפרק כולל זמן המתנה ועיבוד
            
            for i, (head, prompt) in enumerate(chapters):
                pct = int((i / total_chapters) * 100)
                remaining_time = (total_chapters - i) * estimated_time_per_chapter
                
                status_text.info(f"⏳ ({pct}%) כותב כרגע: **{head}** (מבצע בקרת איכות)..." if lang == "עברית" else f"⏳ ({pct}%) Writing: **{head}**...")
                time_estimate.markdown(f"⏱️ **זמן משוער לסיום:** כ-{remaining_time} שניות" if lang == "עברית" else f"⏱️ **Estimated time:** ~{remaining_time} seconds")
                
                res_dict[head] = call_gemini_with_retry(f"Context: {full_context}. Task: {prompt}", api_key, lang)
                
                progress_bar.progress((i + 1) / total_chapters)
                
                if i < total_chapters - 1:
                    time.sleep(12) 
            
            time_estimate.empty() # מחיקת שעון העצר בסיום
            status_text.success("🎉 (100%) העבודה מוכנה ומעוצבת!" if lang == "עברית" else "🎉 (100%) Ready!")
            
            doc = create_pro_doc(topic, name, res_dict, lang)
            buf = io.BytesIO()
            doc.save(buf)
            buf.seek(0)
            
            st.download_button("📥 הורד עבודה סמינריונית (Word)" if lang == "עברית" else "📥 Download Paper (Word)", buf, f"Seminar_{name}.docx")
            st.balloons()
