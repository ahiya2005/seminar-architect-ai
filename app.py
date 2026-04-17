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
    </style>
""", unsafe_allow_html=True)

# --- 2. פונקציות עזר וניקוי ---
def is_valid_email(email):
    pattern = r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z]{2,}$"
    return re.match(pattern, email) is not None

def aggressive_clean_tags(text):
    """שואב אבק אגרסיבי שמחסל כל תגית סורס שעשויה להופיע"""
    # מנקה כל תבנית של סוגריים מרובעים שמכילה את המילה source
    clean_text = re.sub(r"\[.*?source.*?\]", "", text, flags=re.IGNORECASE)
    return clean_text

def extract_text_from_file(uploaded_file):
    if uploaded_file is None: return ""
    try:
        if uploaded_file.name.endswith('.pdf'):
            reader = PyPDF2.PdfReader(uploaded_file)
            return "\n".join([page.extract_text() for page in reader.pages])
        return uploaded_file.getvalue().decode("utf-8")
    except Exception: return ""

# הזרקת XML לוורד כדי לתקן את הסוגריים והכיווניות (RTL)
def set_rtl(element):
    """מכריח את הוורד להבין שמדובר בטקסט מימין לשמאל"""
    pPr = element._element.get_or_add_pPr()
    bidi = OxmlElement('w:bidi')
    bidi.set(qn('w:val'), '1')
    pPr.append(bidi)

# --- 3. הלב של המערכת: Master Prompt ---

def generate_dynamic_outline(topic, extra, key, lang="עברית"):
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key={key}"
    
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
        "generationConfig": {"temperature": 0.3}
    }
    
    try:
        response = requests.post(url, json=payload, timeout=40)
        if response.status_code == 200:
            text = response.json()['candidates'][0]['content']['parts'][0]['text']
            return [c.strip() for c in text.split(',') if c.strip()]
    except: pass
    return ["מבוא מורחב", "רקע תיאורטי", "ניתוח מערכות", "מקרי בוחן", "מסקנות", "ביבליוגרפיה"]

def call_gemini_master_professor(title, topic, name, extra, notes, key, lang="עברית"):
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key={key}"
    
    instruction = f"""
    תפקיד: פרופסור אקדמי.
    משימה: כתוב פרק עומק אקדמי תחת הכותרת '{title}' לעבודה בנושא '{topic}'.
    
    חוקי ברזל נוקשים:
    1. אורך: כתוב טקסט ארוך ומעמיק מאוד. הרחב בכל תת-נושא.
    2. ציטוטים: חובה לשלב ציטוטי APA פנימיים בתוך הטקסט (מחבר, שנה).
    3. תגיות: אסור בהחלט להשתמש בתגיות קוד, סוגריים מרובעים או המילה source.
    4. מבנה: חלק את הפרק ל-3 תתי-נושאים לפחות בעזרת '##'.
    5. ללא הקדמות: התחל מיד עם הכותרת '# {title}'.
    """
    
    payload = {
        "contents": [{"parts": [{"text": instruction}]}], 
        # הורדתי מעט את כמות הטוקנים כדי למנוע את קריסת השרת של גוגל (Timeout)
        "generationConfig": {"temperature": 0.4, "maxOutputTokens": 4500}
    }
    
    for attempt in range(4): # הגדלתי ל-4 ניסיונות
        try:
            # הגדלתי את זמן ההמתנה המקסימלי לתשובה כדי למנוע שגיאות
            response = requests.post(url, json=payload, timeout=150)
            if response.status_code == 200:
                text = response.json()['candidates'][0]['content']['parts'][0]['text']
                # מפעיל את שואב האבק על כל הטקסט
                clean_text = aggressive_clean_tags(text)
                if len(clean_text) > 300: return clean_text.strip()
            time.sleep(6) # מרווח נשימה לשרת
        except: time.sleep(6)
    return f"שגיאה בייצור הפרק '{title}'. (השרת עמוס, נסה שוב מאוחר יותר)."

# --- 4. עיצוב וורד תקני (עם תיקון פיסוק וכיווניות RTL) ---
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
    set_rtl(p)
    r = p.add_run(f"עבודה סמינריונית בנושא:\n{title}")
    r.bold = True
    r.font.size = Pt(24)
    r.font.name = font_name
    
    doc.add_paragraph('\n\n\n\n')
    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_rtl(p2)
    
    inst_text = f"\nמוסד אקדמי: {institution}" if institution else ""
    p2.add_run(f"מוגש על ידי: {author}{inst_text}").font.name = font_name
    doc.add_page_break()

    # יצירת התוכן
    for text in content_list:
        # ניקוי נוסף למקרה שפספסנו משהו
        text = aggressive_clean_tags(text)
        
        for line in text.split('\n'):
            line = line.strip()
            if not line: continue
            
            if line.startswith('# ') and not line.startswith('## '):
                h = doc.add_heading(line.replace('#', '').strip(), level=1)
                h.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                set_rtl(h) # החלת חוקי ימין-לשמאל על הכותרת
            elif line.startswith('## '):
                sh = doc.add_heading(line.replace('##', '').strip(), level=2)
                sh.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                set_rtl(sh) # החלת חוקי ימין-לשמאל על תת-הכותרת
            else:
                p = doc.add_paragraph(line.replace('*', ''))
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                p.paragraph_format.line_spacing = 1.5
                p.paragraph_format.alignment = 3 # יישור דו צדדי
                set_rtl(p) # החלת חוקי ימין-לשמאל על הפסקה (מתקן את הסוגריים!)
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
            st.warning("⚠️ פרוטוקול כתיבה מופעל. נא לא לסגור את החלון! (התהליך לוקח כ-5 עד 10 דקות)")
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
                rem = (total_steps - i) * 60 # עדכון חישוב הזמן (הוספנו השהיות)
                
                status_text.info(f"⏳ ({pct}%) כותב פרק עומק: **{head}**...")
                time_est.markdown(f"⏱️ **זמן משוער לסיום:** כ-{rem} שניות")
                
                content = call_gemini_master_professor(head, topic, name, extra, notes, api_key, lang)
                generated_content.append(content)
                progress_bar.progress((i + 1) / total_steps)
                time.sleep(6) # השהייה ארוכה יותר למניעת חסימות
            
            status_text.info("⏳ מסכם את העבודה ומייצר תקציר...")
            summary_prompt = f"כתוב תקציר אקדמי לעבודה בנושא '{topic}'. סיכום של שאלת המחקר והמסקנות בלבד."
            abstract = call_gemini_master_professor("תקציר", topic, name, summary_prompt, "", api_key, lang)
            generated_content.insert(0, abstract) 
            
            progress_bar.progress(1.0)
            status_text.success("🎉 הסמינריון הושלם בהצלחה! הקובץ מוכן להורדה.")
            
            doc = create_master_doc(topic, name, institution, generated_content, lang)
            buf = io.BytesIO()
            doc.save(buf)
            buf.seek(0)
            st.download_button("📥 הורד סמינריון מושלם (Word)", buf, f"Seminar_{name}.docx")
            st.balloons()
