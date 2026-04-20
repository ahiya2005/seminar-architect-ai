import streamlit as st
import requests
import io
import time
import re
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import PyPDF2

# --- 1. System Config ---
st.set_page_config(page_title="Seminar Architect PRO", page_icon="🎓", layout="wide")
api_key = st.secrets.get("GEMINI_API_KEY")

st.markdown("""
    <style>
    header {visibility: hidden;} #MainMenu {visibility: hidden;} footer {visibility: hidden;}
    div[data-testid="InputInstructions"] { display: none !important; }
    .stButton>button { width: 100%; background-color: #2c3e50; color: white; border-radius: 10px; font-weight: bold; height: 3.5em; }
    </style>
""", unsafe_allow_html=True)

# --- 2. Safe Helpers (Translation-Proof) ---
def is_valid_email(email):
    return "@" in email and "." in email

def clean_text(text):
    # מנוקה מביטויים מורכבים שגוגל טרנסלייט שובר בטלפון
    text = re.sub("\\[.*?source.*?\\]", "", text, flags=re.IGNORECASE)
    text = re.sub("\\[\\d+\\]", "", text)
    return text.strip()

def extract_txt(f):
    if not f: return ""
    try:
        if f.name.endswith('.pdf'): return "\n".join([p.extract_text() for p in PyPDF2.PdfReader(f).pages])
        return f.getvalue().decode("utf-8")
    except Exception: return ""

def set_rtl(p):
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    pPr = p._element.get_or_add_pPr()
    bidi = OxmlElement('w:bidi')
    bidi.set(qn('w:val'), '1')
    pPr.append(bidi)

def add_r(p, text, size=12, bold=False):
    run = p.add_run(text)
    run.font.name = 'David'
    run.font.size = Pt(size)
    run.bold = bold
    rPr = run._element.get_or_add_rPr()
    rtl = OxmlElement('w:rtl')
    rtl.set(qn('w:val'), '1')
    rPr.append(rtl)
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.insert(0, rFonts)
    rFonts.set(qn('w:cs'), 'David')
    rFonts.set(qn('w:ascii'), 'David')
    rFonts.set(qn('w:hAnsi'), 'David')
    return run

# --- 3. The Core API Engine ---
def ask_gemini(prompt, max_t, key):
    # שימוש ישיר במודל שעבד לך ב-100% לפני ששינינו
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key={key}"
    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {"temperature": 0.4, "maxOutputTokens": max_t},
        "safetySettings": [
            {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_NONE"},
            {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_NONE"},
            {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_NONE"},
            {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_NONE"}
        ]
    }
    
    for attempt in range(4):
        try:
            res = requests.post(url, json=payload, timeout=90)
            if res.status_code == 200:
                data = res.json()
                if 'candidates' in data and len(data['candidates']) > 0:
                    ans = data['candidates'][0]['content']['parts'][0]['text']
                    return clean_text(ans)
            elif res.status_code == 429:
                time.sleep(25) # הגנת עומס
            else:
                time.sleep(5)
        except Exception: time.sleep(5)
    return ""

# --- 4. Content Generation (Chunking) ---
def build_chunked_chapter(title, topic, name, extra, notes, key, status):
    if "ביבליוגרפיה" in title or "מקורות" in title:
        p = f"תפקיד: ביבליוגרף. רשום 12 מקורות אקדמיים (APA) בנושא '{topic}'. רשימה בלבד ללא הסברים."
        ans = ask_gemini(p, 1500, key)
        time.sleep(10)
        return f"# {title}\n{ans}" if ans else f"# {title}\nשגיאה בייצור המקורות."

    status.info(f"⏳ מכין תתי-נושאים לפרק '{title}'...")
    sub_p = f"צור 3 כותרות של תתי-נושאים לפרק '{title}' העוסק ב-'{topic}'. החזר רק מופרד בפסיקים."
    subs_str = ask_gemini(sub_p, 200, key)
    time.sleep(5)
    
    subs = [s.strip() for s in subs_str.split(',') if s.strip()] if subs_str else ["חלק א", "חלק ב", "חלק ג"]
    full_text = f"# {title}\n\n"
    
    for sub in subs[:3]:
        status.info(f"⏳ כותב תת-נושא: **{sub}** (מתוך '{title}')...")
        cp = f"""תפקיד: פרופסור. כתוב תוכן אקדמי מעמיק (400 מילים) על תת-הנושא '{sub}' מתוך הפרק '{title}' בנושא '{topic}'. 
        חובה: פסקאות ארוכות, ציטוטי APA (מחבר, שנה). אסור להשתמש בתגיות קוד. ללא כותרות ראשיות."""
        
        part_text = ask_gemini(cp, 2000, key)
        if part_text: full_text += f"## {sub}\n{part_text}\n\n"
        
        # זה הקסם שמונע חסימות:
        time.sleep(10)
        
    return full_text

# --- 5. UI & Main Loop ---
lang = st.radio("🌐 System Language:", ["עברית", "English"], horizontal=True)
if lang == "עברית": st.markdown("<style>.block-container{direction:rtl; text-align:right;}</style>", unsafe_allow_html=True)

if 'logged_in' not in st.session_state: st.session_state['logged_in'] = False

st.title("🎓 Seminar Architect PRO")

if not st.session_state['logged_in']:
    email = st.text_input("אימייל להתחברות:")
    if st.button("🚀 כניסה"):
        if is_valid_email(email): st.session_state['logged_in'] = True; st.rerun()
else:
    col1, col2 = st.columns(2)
    with col1:
        topic = st.text_input("נושא הסמינריון:")
        name = st.text_input("שם הסטודנט:")
        inst = st.text_input("מוסד אקדמי (אופציונלי):")
    with col2:
        extra = st.text_area("דגשים ספציפיים:", height=130)
    up = st.file_uploader("העלאת סילבוס (PDF/TXT):", type=['pdf', 'txt'])

    if st.button("🚀 צא לדרך! (הפרד ומשול)"):
        if not topic or not name: st.error("חסרים נתונים!")
        elif not api_key: st.error("API KEY חסר.")
        else:
            st.warning("⚠️ המערכת בונה את העבודה חלק-אחר-חלק כדי למנוע עומס (ייקח כ-8 דקות). לא לסגור!")
            notes = extract_txt(up)
            status = st.empty()
            bar = st.progress(0)
            
            status.info("⏳ בונה שלד פרקים...")
            out_p = f"צור 6 כותרות אקדמיות לפרקים בנושא '{topic}'. פרק אחרון חובה: 'ביבליוגרפיה'. החזר מופרד בפסיקים."
            ch_str = ask_gemini(out_p, 300, api_key)
            time.sleep(5)
            chapters = [c.strip() for c in ch_str.split(',') if c.strip()][:6] if ch_str else ["מבוא", "תיאוריה", "ניתוח", "מקרי בוחן", "מסקנות", "ביבליוגרפיה"]
            
            res_content = []
            for i, head in enumerate(chapters):
                content = build_chunked_chapter(head, topic, name, extra, notes, api_key, status)
                res_content.append(content)
                bar.progress((i + 1) / (len(chapters) + 1))
            
            status.info("⏳ מסכם לתקציר...")
            abs_p = f"כתוב תקציר אקדמי (300 מילים) לעבודה בנושא '{topic}'."
            abstract = ask_gemini(abs_p, 1000, api_key)
            res_content.insert(0, f"# תקציר\n{abstract}" if abstract else "# תקציר\nשגיאה.")
            
            bar.progress(1.0)
            status.success("🎉 הסמינריון נבנה בהצלחה! מוריד את העבודה...")
            
            # עיצוב הוורד
            doc = Document()
            for s in doc.sections: s.top_margin = s.bottom_margin = s.left_margin = s.right_margin = Inches(0.98)
            
            doc.add_paragraph('\n\n\n\n')
            p1 = doc.add_paragraph(); set_rtl(p1)
            add_r(p1, f"עבודה סמינריונית בנושא:\n{topic}", 24, True)
            
            doc.add_paragraph('\n\n\n\n')
            p2 = doc.add_paragraph(); set_rtl(p2)
            add_r(p2, f"מוגש על ידי: {name}", 16, False)
            if inst:
                add_r(p2, "\n", 16, False)
                add_r(p2, f"מוסד אקדמי: {inst}", 16, False)
            doc.add_page_break()

            for text in res_content:
                for line in text.split('\n'):
                    line = line.strip()
                    if not line: continue
                    p = doc.add_paragraph(); set_rtl(p)
                    if line.startswith('# '): add_r(p, line.replace('#','').strip(), 18, True)
                    elif line.startswith('## '): add_r(p, line.replace('#','').strip(), 14, True)
                    else:
                        p.paragraph_format.line_spacing = 1.5
                        add_r(p, line.replace('*',''), 12, False)
                doc.add_page_break()
                
            buf = io.BytesIO(); doc.save(buf); buf.seek(0)
            st.download_button("📥 הורד סמינריון מושלם (Word)", buf, f"Seminar_{name}.docx")
            st.balloons()
