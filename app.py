import streamlit as st
import requests
import json
import io
import time
import re
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
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
    .stProgress > div > div > div {
        background-image: linear-gradient(45deg, rgba(255, 255, 255, .15) 25%, transparent 25%, transparent 50%, rgba(255, 255, 255, .15) 50%, rgba(255, 255, 255, .15) 75%, transparent 75%, transparent);
        background-size: 1rem 1rem;
        animation: progress-bar-stripes 1s linear infinite;
    }
    @keyframes progress-bar-stripes { from { background-position: 1rem 0; } to { background-position: 0 0; } }
    </style>
""", unsafe_allow_html=True)

# --- 2. פונקציות עזר וניקוי (RTL ותגיות) ---
def is_valid_email(email):
    pattern = r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z]{2,}$"
    return re.match(pattern, email) is not None

def aggressive_clean_tags(text):
    # שואב אבק אגרסיבי שמחסל כל תגית סורס
    text = re.sub(r"\[.*?source.*?\]", "", text, flags=re.IGNORECASE)
    text = re.sub(r"\[\d+\]", "", text)
    return text

def extract_text_from_file(uploaded_file):
    if uploaded_file is None: return ""
    try:
        if uploaded_file.name.endswith('.pdf'):
            reader = PyPDF2.PdfReader(uploaded_file)
            return "\n".join([page.extract_text() for page in reader.pages])
        return uploaded_file.getvalue().decode("utf-8")
    except Exception: return ""

def set_rtl_paragraph(p):
    pPr = p._element.get_or_add_pPr()
    bidi = OxmlElement('w:bidi')
    bidi.set(qn('w:val'), '1')
    pPr.append(bidi)

def add_rtl_run(paragraph, text, font_name='David', font_size=12, bold=False):
    run = paragraph.add_run(text)
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.bold = bold
    
    rPr = run._element.get_or_add_rPr()
    rtl = OxmlElement('w:rtl')
    rtl.set(qn('w:val'), '1')
    rPr.append(rtl)
    
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:cs'), font_name)
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)
    return run

# --- 3. הלב של המערכת: Master Prompt מוחזר במלואו ---

def generate_dynamic_outline(topic, extra, key, lang="עברית"):
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={key}"
    prompt = f"""
    תפקיד: פרופסור אקדמי בכיר.
    משימה: בנה 6 כותרות אקדמיות לעבודה סמינריונית בנושא '{topic}'.
    הנחיות: {extra}
    חוקים:
    1. אין כותרות גנריות (כמו 'ממצאים'). הכל מותאם ספציפית לנושא.
    2. פרק אחרון חובה: 'ביבליוגרפיה'.
    החזר רק את הכותרות מופרדות בפסיק (,) ללא מספור וללא טקסט נוסף.
    """
    payload = {
        "contents": [{"parts": [{"text": prompt}]}], 
        "generationConfig": {"temperature": 0.3},
        "safetySettings": [
            {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_NONE"},
            {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_NONE"},
            {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_NONE"},
            {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_NONE"}
        ]
    }
    try:
        response = requests.post(url, json=payload, timeout=40)
        if response.status_code == 200:
            text = response.json()['candidates'][0]['content']['parts'][0]['text']
            return [c.strip() for c in text.split(',') if c.strip()]
    except: pass
    return ["מבוא מורחב", "רקע תיאורטי", "ניתוח מערכות", "מקרי בוחן", "מסקנות", "ביבליוגרפיה"]

def call_gemini_master_professor(title, topic, name, extra, notes, key, status_element):
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={key}"
    
    if "ביבליוגרפיה" in title or "מקורות" in title:
        instruction = f"""
        תפקיד: ביבליוגרף אקדמי.
        משימה: צור רשימה ביבליוגרפית (APA) מקיפה לעבודה בנושא '{topic}'.
        חוקי ברזל:
        1. רשימה בלבד! ללא פסקאות הסבר.
        2. פורמט: שם מחבר, שנה, כותרת, הוצאה.
        3. רשום לפחות 12 מקורות קשורים.
        """
        min_length = 100 
    else:
        instruction = f"""
        תפקיד: פרופסור אקדמי.
        משימה: כתוב פרק עומק אקדמי תחת הכותרת '{title}' לעבודה בנושא '{topic}' עבור הסטודנט {name}.
        חוקי ברזל נוקשים:
        1. אורך: טקסט ארוך ומעמיק של כ-800 עד 1000 מילים לפחות. פסקאות בשרניות.
        2. ציטוטים: חובה לשלב ציטוטי APA פנימיים בתוך הטקסט (מחבר, שנה).
        3. תגיות: אסור בהחלט להשתמש בתגיות קוד, סוגריים מרובעים או במילה source.
        4. מבנה: חלק את הפרק ל-3 תתי-נושאים לפחות בעזרת '##'.
        5. ללא הקדמות: התחל מיד עם הכותרת '# {title}'.
        """
        min_length = 400

    payload = {
        "contents": [{"parts": [{"text": instruction}]}], 
        "generationConfig": {"temperature": 0.4, "maxOutputTokens": 8192},
        "safetySettings": [
            {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_NONE"},
            {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_NONE"},
            {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_NONE"},
            {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_NONE"}
        ]
    }
    
    for attempt in range(3):
        try:
            response = requests.post(url, json=payload, timeout=180)
            if response.status_code == 200:
                res_data = response.json()
                if 'candidates' in res_data and len(res_data['candidates']) > 0:
                    text = res_data['candidates'][0]['content']['parts'][0]['text']
                    clean_text = aggressive_clean_tags(text)
                    if len(clean_text) >= min_length: 
                        return clean_text.strip()
            elif response.status_code == 429:
                status_element.warning(f"⏳ חריגה ממכסת מילים חינמית בגוגל. ממתין 60 שניות לאיפוס המכסה...")
                time.sleep(60)
            else:
                time.sleep(5)
        except Exception:
            time.sleep(5)
            
    return f"שגיאה בייצור הפרק '{title}' (גוגל חסמה את הבקשה עקב עומס מילים)."

# --- 4. עיצוב וורד תקני (RTL מלא ותיקון עמוד השער) ---
def create_master_doc(topic, author, institution, content_list, lang):
    doc = Document()
    font_name = 'David' if lang == "עברית" else 'Times New Roman'
    
    for section in doc.sections:
        section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = Inches(0.98)

    # עמוד שער אקדמי - תיקון הריווח בין השם למוסד
    doc.add_paragraph('\n\n\n\n')
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_rtl_paragraph(p)
    add_rtl_run(p, f"עבודה סמינריונית בנושא:\n{topic}", font_name, 24, True)
    
    doc.add_paragraph('\n\n\n\n')
    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_rtl_paragraph(p2)
    add_rtl_run(p2, f"מוגש על ידי: {author}", font_name, 16, False)
    if institution:
        # פקודת ירידת שורה קשיחה שלא מפספסת
        add_rtl_run(p2, "\n", font_name, 16, False)
        add_rtl_run(p2, f"מוסד אקדמי: {institution}", font_name, 16, False)
    doc.add_page_break()

    # יצירת תוכן
    for text in content_list:
        text = aggressive_clean_tags(text)
        
        for line in text.split('\n'):
            line = line.strip()
            if not line: continue
            
            if line.startswith('# ') and not line.startswith('## '):
                p_head = doc.add_paragraph()
                p_head.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                set_rtl_paragraph(p_head)
                add_rtl_run(p_head, line.replace('#', '').strip(), font_name, 18, True)
            
            elif line.startswith('## '):
                p_sub = doc.add_paragraph()
                p_sub.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                set_rtl_paragraph(p_sub)
                add_rtl_run(p_sub, line.replace('##', '').strip(), font_name, 14, True)
            
            else:
                p_text = doc.add_paragraph()
                p_text.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                p_text.paragraph_format.line_spacing = 1.5
                p_text.paragraph_format.alignment = 3 
                set_rtl_paragraph(p_text)
                add_rtl_run(p_text, line.replace('*', ''), font_name, 12, False)
                
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
        else: st.error("נא להזין אימייל תקין.")
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
            st.warning("⚠️ פקודות העומק הופעלו. המערכת תמתין באופן יזום בין פרק לפרק כדי למנוע חסימות גוגל (כ-10 דקות).")
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
                rem = (total_steps - i) * 65 # זמנים ארוכים יותר בגלל ההמתנה
                
                status_text.info(f"⏳ ({pct}%) מנתח וכותב פרק עומק: **{head}**...")
                time_est.markdown(f"⏱️ **זמן משוער לסיום:** כ-{rem} שניות")
                
                content = call_gemini_master_professor(head, topic, name, extra, notes, api_key, status_text)
                generated_content.append(content)
                progress_bar.progress((i + 1) / total_steps)
                
                # מנגנון קירור חובה: השהייה קשיחה של 25 שניות בין פרק לפרק לאיפוס מכסת המילים החינמית
                if i < len(chapters) - 1:
                    status_text.info("⏳ מקרר את השרת למניעת חסימת מפתח...")
                    time.sleep(25) 
            
            status_text.info("⏳ מסכם את העבודה ומייצר תקציר...")
            summary_prompt = f"כתוב תקציר אקדמי קצר לעבודה בנושא '{topic}'."
            abstract = call_gemini_master_professor("תקציר", topic, name, summary_prompt, "", api_key, status_text)
            generated_content.insert(0, abstract) 
            
            progress_bar.progress(1.0)
            status_text.success("🎉 הסמינריון הושלם בהצלחה! כללי הפורמט, ה-RTL והעומק הוחזרו.")
            
            doc = create_master_doc(topic, name, institution, generated_content, lang)
            buf = io.BytesIO()
            doc.save(buf)
            buf.seek(0)
            st.download_button("📥 הורד סמינריון מושלם (Word)", buf, f"Seminar_{name}.docx")
            st.balloons()
