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
    /* הסתרת תפריטי מערכת למראה אפליקציה עצמאית */
    header {visibility: hidden;}
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    
    /* הסתרת הוראות אנגלית של המערכת */
    div[data-testid="InputInstructions"] { display: none !important; }
    
    /* עיצוב כפתורים */
    .stButton>button { width: 100%; background-color: #2c3e50; color: white; border-radius: 10px; font-weight: bold; }
    .stButton>button:hover { background-color: #34495e; }
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
    
    # לולאת ניסיונות (מעודכן ל-10 ניסיונות ו-10 שניות המתנה)
    for attempt in range(max_retries):
        try:
            response = requests.post(url, headers={'Content-Type': 'application/json'}, data=json.dumps(payload), timeout=100)
            
            if response.status_code == 200:
                result = response.json()
                if 'candidates' in result and len(result['candidates']) > 0:
                    text = result['candidates'][0]['content']['parts'][0]['text']
                    
                    # בקרת איכות: דחיית תשובות קצרות מדי
                    if len(text) < 400:
                        time.sleep(10)
                        continue 
                    
                    lines = text.split('\n')
                    if lines and any(word in lines[0] for word in ["להלן", "הנה", "Here", "Sure"]):
                        lines = lines[1:]
                    return "\n".join(lines).strip()
            
            # המתנה של 10 שניות בין ניסיונות בכל מקרה של תקלה או עומס
            time.sleep(10)
                
        except Exception:
            time.sleep(10)
            
    return "שגיאה: לא הצלחנו לייצר את הפרק באיכות הנדרשת. אנא נסה שוב." if lang == "עברית" else "Error generating content."

def create_pro_doc(title, author, content_dict, lang):
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

    # תוכן
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

# --- הממשק הראשי ---
if 'logged_in' not in st.session_state:
    st.session_state['logged_in'] = False

st.title("🎓 Seminar Architect PRO")

if not st.session_state['logged_in']:
    email = st.text_input("אימייל להתחברות / Email:")
    if st.button("כניסה / Login"):
        if email:
            st.session_state['logged_in'] = True
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

    if st.button("🚀 התחל בייצור העבודה - נא לא לסגור את החלון!" if lang == "עברית" else "🚀 Start Generating!"):
        if not topic: 
            st.error("חובה להזין נושא!")
        else:
            # הודעת אזהרה קבועה (סעיף 3 בבקשתך)
            wait_msg = "המערכת מייצרת כרגע את העבודה. תהליך זה לוקח בין 5 ל-10 דקות (תלוי בעומס השרתים). נא לא לסגור את החלון." if lang == "עברית" else "Generating paper... This process takes 5-10 minutes. Please stay on this page."
            st.warning(f"⚠️ **{wait_msg}**")
            
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
            progress = st.progress(0)
            status = st.empty()
            
            for i, (head, prompt) in enumerate(chapters):
                status.info(f"✍️ כותב כרגע: **{head}** (מבצע בקרת איכות ומקורות...)" if lang == "עברית" else f"✍️ Writing: **{head}**...")
                
                # קריאה ל-API
                res_dict[head] = call_gemini_with_retry(f"Context: {full_context}. Task: {prompt}", api_key, lang)
                
                progress.progress((i + 1) / len(chapters))
                
                if i < len(chapters) - 1:
                    time.sleep(12) # המתנת ביטחון קלה בין פרקים מוצלחים
            
            status.success("✅ העבודה מוכנה ומעוצבת!" if lang == "עברית" else "✅ Ready!")
            doc = create_pro_doc(topic, name, res_dict, lang)
            buf = io.BytesIO()
            doc.save(buf)
            buf.seek(0)
            
            st.download_button("📥 הורד עבודה סמינריונית (Word)", buf, f"Seminar_{name}.docx")
            st.balloons()
