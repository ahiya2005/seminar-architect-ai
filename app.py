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

# --- ממשק שפות וכיוון דף (RTL/LTR) ---
lang = st.sidebar.radio("🌐 שפת מערכת / System Language", ["עברית", "English"])

if lang == "עברית":
    direction = "rtl"
    st.markdown("""
        <style>
        .main, .stApp, input, textarea, label { direction: rtl; text-align: right; }
        div[data-testid="stToolbar"] { direction: ltr; }
        .stButton>button { width: 100%; background-color: #2c3e50; color: white; border-radius: 10px; }
        </style>
        """, unsafe_allow_html=True)
else:
    direction = "ltr"
    st.markdown("""
        <style>
        .main, .stApp, input, textarea, label { direction: ltr; text-align: left; }
        .stButton>button { width: 100%; background-color: #2c3e50; color: white; border-radius: 10px; }
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

def call_gemini_with_retry(prompt, key, lang="עברית"):
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key={key}"
    
    # הנחיות כתיבה קשיחות: עומק, פיסוק, מקורות אמת וחלוקה לתתי-נושאים
    system_instruction = """
    ROLE: Senior Academic Researcher.
    RULES:
    1. STRUCTURE: Breakdown every chapter into 3-4 deep sub-topics using '##'. 
    2. DEPTH: Provide extensive academic analysis. No fluff. 
    3. SOURCES: Use ONLY real academic sources. Check facts. APA style only. 
    4. QUALITY: Use perfect grammar, punctuation, and formal academic tone. 
    5. NO META-TEXT: Do not say "Here is your text". Start the content immediately.
    6. VALIDATION: Ensure logical flow between paragraphs.
    """
    if lang == "עברית": system_instruction += " 7. כתיבה בעברית אקדמית גבוהה בלבד."
    else: system_instruction += " 7. Write in high-level formal English."

    payload = {
        "contents": [{"parts": [{"text": system_instruction + "\n\n" + prompt}]}],
        "generationConfig": {"temperature": 0.4, "maxOutputTokens": 4000}
    }
    
    # לולאת "אל תתקדם עד שיהיה מושלם"
    while True:
        try:
            response = requests.post(url, headers={'Content-Type': 'application/json'}, data=json.dumps(payload))
            if response.status_code == 200:
                result = response.json()
                text = result['candidates'][0]['content']['parts'][0]['text']
                if len(text) > 200: # לוודא שלא קיבלנו תשובה קצרה מדי או שגיאה מוסווית
                    return text.strip()
            elif response.status_code == 429: # עומס
                time.sleep(45) # לחכות דקה כמעט כדי שגוגל ישחררו
            else:
                time.sleep(10)
        except:
            time.sleep(10)

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
st.title("🎓 Seminar Architect PRO")

if not st.session_state.get('logged_in'):
    email = st.text_input("אימייל להתחברות / Email:")
    if st.button("כניסה / Login"):
        st.session_state['logged_in'] = True
        st.rerun()
else:
    col1, col2 = st.columns(2)
    with col1:
        topic = st.text_input("נושא הסמינריון:" if lang == "עברית" else "Seminar Topic:")
        name = st.text_input("שם הסטודנט:" if lang == "עברית" else "Student Name:")
    with col2:
        extra = st.text_area("תיאור והנחיות ספציפיות (למשל: דגש על פסיכולוגיה):" if lang == "עברית" else "Focus Instructions:", height=100)
    
    uploaded = st.file_uploader("העלאת הנחיות מרצה (PDF/TXT):" if lang == "עברית" else "Upload Guidelines:", type=['pdf', 'txt'])

    if st.button("🚀 התחל בייצור העבודה - נא לא לסגור את החלון!"):
        if not topic: st.error("חובה להזין נושא!")
        else:
            st.warning("⚠️ המערכת מייצרת את העבודה כרגע. **נא לא לסגור את החלון או לעבור דף**, אחרת התהליך יתאפס." if lang == "עברית" else "⚠️ Generating... **DO NOT close or leave this page**.")
            
            notes = extract_text_from_file(uploaded)
            full_context = f"Topic: {topic}. Notes: {extra}. Guidelines: {notes}"
            
            chapters = [
                ("מבוא", "Write the Intro. Sub-topics: Background, Research Question, Rationale."),
                ("סקירה ספרותית", "Write the Lit Review. Breakdown into 4 main theoretical sub-topics."),
                ("מתודולוגיה", "Write Methodology. Breakdown: Tools, Population, Method."),
                ("ממצאים", "Write Findings. Detailed hypothetical breakdown of results."),
                ("דיון ומסקנות", "Write Discussion. Critique findings vs theories. Summary."),
                ("ביבליוגרפיה", "Final Bibliography. APA style. Real sources only. Alphabetical.")
            ]
            
            res_dict = {}
            progress = st.progress(0)
            status = st.empty()
            
            for i, (head, prompt) in enumerate(chapters):
                status.info(f"✍️ כותב כרגע: **{head}** (בדיקת איכות ומקורות בעיצומה...)" if lang == "עברית" else f"✍️ Writing: **{head}**...")
                res_dict[head] = call_gemini_with_retry(f"Context: {full_context}. Task: {prompt}", api_key, lang)
                progress.progress((i + 1) / len(chapters))
                time.sleep(10) # השהיית ביטחון
            
            status.success("✅ העבודה מוכנה ומעוצבת!" if lang == "עברית" else "✅ Ready!")
            doc = create_pro_doc(topic, name, res_dict, lang)
            buf = io.BytesIO()
            doc.save(buf)
            buf.seek(0)
            st.download_button("📥 הורד עבודה סמינריונית (Word)", buf, f"Seminar_{name}.docx")
            st.balloons()
