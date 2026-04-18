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

# --- 2. פונקציות עזר וניקוי ---
def is_valid_email(email):
    pattern = r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z]{2,}$"
    return re.match(pattern, email) is not None

def aggressive_clean_tags(text):
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

# --- 3. מנוע ה-AI המשודרג (מודל PRO המעמיק + עקיפת מסנני בטיחות) ---

def generate_dynamic_outline(topic, extra, key, lang="עברית"):
    # שדרוג למודל הפרו המעמיק
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-pro-latest:generateContent?key={key}"
    prompt = f"""
    תפקיד: פרופסור אקדמי בכיר.
    משימה: בנה 6 כותרות אקדמיות לעבודה סמינריונית בנושא '{topic}'.
    הנחיות: {extra}
    חוקים:
    1. הוקדש זמן לחשיבה מעמיקה (Deep reasoning) על הנושא לפני כתיבת הכותרות.
    2. אין כותרות גנריות (כמו 'ממצאים'). הכל מותאם ספציפית לנושא.
    3. פרק אחרון חובה: 'ביבליוגרפיה'.
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
        response = requests.post(url, json=payload, timeout=60)
        if response.status_code == 200:
            text = response.json()['candidates'][0]['content']['parts'][0]['text']
            return [c.strip() for c in text.split(',') if c.strip()]
    except: pass
    return ["מבוא מורחב", "רקע תיאורטי", "ניתוח מערכות", "מקרי בוחן", "מסקנות", "ביבליוגרפיה"]

def call_gemini_master_professor(title, topic, name, extra, notes, key, lang="עברית"):
    # שדרוג למודל הפרו המעמיק עבור כתיבת התוכן
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-pro-latest:generateContent?key={key}"
    
    if "ביבליוגרפיה" in title or "מקורות" in title:
        instruction = f"""
        תפקיד: ביבליוגרף אקדמי.
        משימה: צור רשימה ביבליוגרפית (APA) מקיפה לעבודה אקדמית בנושא '{topic}'.
        חוקי ברזל:
        1. רשימה בלבד! ללא פסקאות הסבר וללא הקדמות. אל תכתוב שום דבר שאינו מקור אקדמי.
        2. עשה שימוש בפורמט: שם מחבר, שנת פרסום, כותרת, הוצאה/כתב עת.
        3. רשום לפחות 12 מקורות קשורים לנושא.
        """
        min_length_required = 100 
    else:
        instruction = f"""
        תפקיד: פרופסור אקדמי.
        משימה: כתוב פרק עומק אקדמי תחת הכותרת '{title}' לעבודה בנושא '{topic}'.
        חוקי ברזל נוקשים:
        1. חשיבה מעמיקה: נתח את הנושא לעומק, הבא מקרי בוחן, תיאוריות מורכבות ודיון ביקורתי.
        2. אורך: טקסט מעמיק מאוד של כ-800 עד 1000 מילים לפחות. 
        3. ציטוטים: שלב ציטוטי APA פנימיים (מחבר, שנה) כדי לגבות כל טענה.
        4. תגיות: אסור להשתמש בתגיות, בקוד או בסוגריים מרובעים.
        5. מבנה: חלק את הפרק ל-3 תתי-נושאים לפחות בעזרת '##'.
        6. ללא הקדמות: התחל מיד עם '# {title}'.
        """
        min_length_required = 400

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
            # מודל הפרו צריך זמן לחשוב (העלינו את הטיימר ל-200 שניות)
            response = requests.post(url, json=payload, timeout=200) 
            if response.status_code == 200:
                res_data = response.json()
                if 'candidates' in res_data and len(res_data['candidates']) > 0:
                    if 'parts' in res_data['candidates'][0]['content']:
                        text = res_data['candidates'][0]['content']['parts'][0]['text']
                        clean_text = aggressive_clean_tags(text)
                        if len(clean_text) >= min_length_required: 
                            return clean_text.strip()
            time.sleep(10) # מנוחה ארוכה יותר לשרת של גוגל בין ניסיונות
        except Exception as e:
            time.sleep(10)
            
    return f"שגיאה בייצור הפרק '{title}'. השרת היה עמוס בחישובים."

# --- 4. עיצוב וורד תקני ---
def create_master_doc(topic, author, institution, content_list, lang):
    doc = Document()
    font_name = 'David' if lang == "עברית" else 'Times New Roman'
    
    for section in doc.sections:
        section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = Inches(0.98)

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
        add_rtl_run(p2, "\n", font_name, 16, False)
        add_rtl_run(p2, f"מוסד אקדמי: {institution}", font_name, 16, False)
    doc.add_page_break()

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

    if st.button("🚀 צא לדרך! הפעל פרוטוקול כתיבה אקדמי (מודל מעמיק)"):
        if not topic or not name: st.error("הזן נושא ושם סטודנט!")
        elif not api_key: st.error("שגיאת API KEY.")
        else:
            st.warning("⚠️ המערכת כעת משתמשת במודל ה'מעמיק' של גוגל (Gemini 1.5 Pro). התהליך מורכב ולוקח עד 15 דקות - נא לא לסגור את החלון!")
            notes = extract_text_from_file(uploaded)
            
            status_text = st.empty()
            time_est = st.empty()
            progress_bar = st.progress(0)
            
            status_text.info("⏳ מאבחן נושא (חשיבה מעמיקה)...")
            chapters = generate_dynamic_outline(topic, extra, api_key, lang)
            
            generated_content = []
            total_steps = len(chapters) + 1 
            
            for i, head in enumerate(chapters):
                pct = int((i / total_steps) * 100)
                rem = (total_steps - i) * 75 # הארכת הזמן המשוער בתצוגה
                
                status_text.info(f"⏳ ({pct}%) מנתח וכותב פרק עומק: **{head}**...")
                time_est.markdown(f"⏱️ **זמן משוער לסיום:** כ-{rem} שניות")
                
                content = call_gemini_master_professor(head, topic, name, extra, notes, api_key, lang)
                generated_content.append(content)
                progress_bar.progress((i + 1) / total_steps)
                time.sleep(8) # מרווח נשימה משמעותי למודל הפרו
            
            status_text.info("⏳ מסכם את העבודה ומייצר תקציר...")
            summary_prompt = f"כתוב תקציר אקדמי לעבודה בנושא '{topic}'. סיכום קצר של שאלת המחקר והמסקנות בלבד."
            abstract = call_gemini_master_professor("תקציר", topic, name, summary_prompt, "", api_key, lang)
            generated_content.insert(0, abstract) 
            
            progress_bar.progress(1.0)
            status_text.success("🎉 הסמינריון המעמיק הושלם בהצלחה! הקובץ מוכן להורדה.")
            
            doc = create_master_doc(topic, name, institution, generated_content, lang)
            buf = io.BytesIO()
            doc.save(buf)
            buf.seek(0)
            st.download_button("📥 הורד סמינריון מושלם (Word)", buf, f"Seminar_{name}.docx")
            st.balloons()
