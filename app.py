import streamlit as st
import requests
import json
import io
import time
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import PyPDF2

# --- הגדרות עיצוב ודף ---
st.set_page_config(page_title="Seminar Architect PRO", page_icon="🎓", layout="wide")

st.markdown("""
    <style>
    .main {background-color: #f8f9fa;}
    h1 {color: #2c3e50; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;}
    .stButton>button {background-color: #3498db; color: white; border-radius: 8px; border: none; padding: 10px 24px; font-weight: bold;}
    .stButton>button:hover {background-color: #2980b9;}
    </style>
    """, unsafe_allow_html=True)

api_key = st.secrets.get("GEMINI_API_KEY")

def extract_text_from_file(uploaded_file):
    if uploaded_file is None:
        return ""
    try:
        if uploaded_file.name.endswith('.pdf'):
            reader = PyPDF2.PdfReader(uploaded_file)
            text = ""
            for page in reader.pages:
                text += page.extract_text() + "\n"
            return text
        elif uploaded_file.name.endswith('.txt'):
            return uploaded_file.getvalue().decode("utf-8")
    except Exception as e:
        return f"שגיאה בקריאת הקובץ: {str(e)}"
    return ""

def call_gemini_direct(prompt, key, lang="עברית", retries=3):
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key={key}"
    
    # פקודת מערכת קשוחה למניעת חרטוטים, הוספת מקורות אמיתיים, וללא "משפטי פתיחה"
    system_instruction = """
    You are an expert academic writer. You must strictly follow these rules:
    1. DO NOT include any conversational filler, greetings, or meta-text (e.g., "Here is the chapter", "להלן פרק המבוא"). Start writing the academic text immediately.
    2. DO NOT invent or hallucinate citations. Use ONLY real, existing academic sources.
    3. Output in clean text. Use '##' for sub-headings.
    """
    if lang == "עברית":
        system_instruction += "4. Write ONLY in formal, high-level academic Hebrew."
    else:
        system_instruction += "4. Write ONLY in formal, high-level academic English."

    payload = {
        "contents": [{"parts": [{"text": system_instruction + "\n\n" + prompt}]}],
        "generationConfig": {"temperature": 0.5, "maxOutputTokens": 4000}
    }
    
    headers = {'Content-Type': 'application/json'}
    
    # מנגנון הגנה מפני קריסת API (שגיאה 429)
    for attempt in range(retries):
        try:
            response = requests.post(url, headers=headers, data=json.dumps(payload))
            if response.status_code == 200:
                text = response.json()['candidates'][0]['content']['parts'][0]['text']
                # ניקוי שאריות של "להלן..." אם ה-AI בכל זאת התעקש
                lines = text.split('\n')
                if lines and ("להלן" in lines[0] or "הנה" in lines[0] or "Here is" in lines[0]):
                    lines = lines[1:]
                return "\n".join(lines).strip()
            elif response.status_code == 429:
                time.sleep(20) # המתנה של 20 שניות אם הגענו למגבלת הקצב של גוגל
            else:
                return f"Error ({response.status_code}): {response.text}"
        except Exception as e:
            time.sleep(5)
    return "שגיאה: המערכת עמוסה כרגע. אנא נסה שוב מאוחר יותר."

def create_word_document(title, author, parts_dict, lang):
    doc = Document()
    
    # הגדרת פונט אחיד למסמך כולו
    style = doc.styles['Normal']
    font = style.font
    font.name = 'David' if lang == "עברית" else 'Times New Roman'
    font.size = Pt(12)
    
    # --- עמוד שער אקדמי תקני ---
    doc.add_paragraph('\n\n\n\n')
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_p.add_run(f"עבודה סמינריונית בנושא:\n{title}")
    title_run.bold = True
    title_run.font.size = Pt(22)
    
    doc.add_paragraph('\n\n\n')
    author_p = doc.add_paragraph()
    author_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    author_run = author_p.add_run(f"מגיש: {author}" if lang == "עברית" else f"Submitted by: {author}")
    author_run.font.size = Pt(14)
    doc.add_page_break()

    # --- יצירת הפרקים ---
    for part_name, content in parts_dict.items():
        # כותרת ראשית של פרק (Heading 1) - מוגדרת כך שבוורד תוכלו ליצור תוכן עניינים אוטומטי
        h = doc.add_heading(part_name, level=1)
        h.alignment = WD_ALIGN_PARAGRAPH.RIGHT if lang == "עברית" else WD_ALIGN_PARAGRAPH.LEFT
        
        paragraphs = content.split('\n')
        for p_text in paragraphs:
            p_text = p_text.strip()
            if not p_text:
                continue
            
            # אם ה-AI ייצר תת-כותרת (Heading 2)
            if p_text.startswith('##'):
                sub_h = doc.add_heading(p_text.replace('##', '').strip(), level=2)
                sub_h.alignment = WD_ALIGN_PARAGRAPH.RIGHT if lang == "עברית" else WD_ALIGN_PARAGRAPH.LEFT
            else:
                # פסקה רגילה
                p = doc.add_paragraph(p_text.replace('*', ''))
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if lang == "עברית" else WD_ALIGN_PARAGRAPH.LEFT
                p.paragraph_format.line_spacing = 1.5
                p.paragraph_format.space_after = Pt(10)
        
        doc.add_page_break()
        
    return doc

# --- ניהול מצב (Session State) ---
if 'logged_in' not in st.session_state:
    st.session_state['logged_in'] = False
if 'email' not in st.session_state:
    st.session_state['email'] = ""

# --- ממשק ---
st.title("🎓 Seminar Architect PRO")
lang = st.radio("שפת הממשק / Language", ["עברית", "English"], horizontal=True)

if not st.session_state['logged_in']:
    st.subheader("התחברות למערכת" if lang == "עברית" else "Login")
    email_input = st.text_input("אימייל / Email:")
    if st.button("התחבר / Login"):
        if email_input:
            st.session_state['email'] = email_input
            st.session_state['logged_in'] = True
            st.rerun()
        else:
            st.warning("נא להזין אימייל.")
else:
    st.success(f"מחובר כ: {st.session_state['email']}" if lang == "עברית" else f"Logged in as: {st.session_state['email']}")
    
    col1, col2 = st.columns(2)
    with col1:
        title = st.text_input("נושא העבודה (לדוגמה: השפעת לחץ חברתי על השתתפות בקמפיינים ויראליים):")
        author = st.text_input("שם הסטודנט:")
    with col2:
        description = st.text_area("מיקוד והנחיות (מה תרצה שיודגש בעבודה? כיוון ספציפי?):", height=130)
    
    st.markdown("### 📎 העלאת הנחיות מרצה/סילבוס (אופציונלי)")
    uploaded_file = st.file_uploader("העלה קובץ PDF או TXT עם דרישות העבודה", type=['pdf', 'txt'])

    if st.button("🚀 צור עבודה סמינריונית מושלמת"):
        if not title:
            st.error("חובה להזין נושא!")
        else:
            lecturer_notes = extract_text_from_file(uploaded_file)
            context = f"Topic: {title}\n"
            if description: context += f"Student's Specific Focus: {description}\n"
            if lecturer_notes: context += f"Lecturer Strict Guidelines: {lecturer_notes}\n"

            progress = st.progress(0)
            status = st.empty()
            
            # שלבי עבודה - השמות כאן הם מה שיודפס ככותרות עליונות במסמך, לכן אין "חלק א/ב"
            parts_plan = [
                ("מבוא", "Write the Introduction. Must include background, modern context, research question, and rationale."),
                ("סקירת ספרות: רקע תיאורטי", "Write the first part of the Literature Review focusing purely on theoretical background and classic theories."),
                ("סקירת ספרות: מחקרים עדכניים", "Write the second part of the Literature Review focusing on empirical studies from the last 5 years. Identify the research gap."),
                ("מתודולוגיית המחקר", "Write the Methodology chapter. Describe the research method, tools, population, and limitations."),
                ("ממצאים", "Write the Findings chapter. Present deep, detailed hypothetical findings categorized logically."),
                ("דיון ומסקנות", "Write the Discussion and Conclusion chapter. Critically analyze the findings against the literature review. Suggest future research."),
                ("ביבליוגרפיה", "List all references used. MUST be formatted perfectly in APA style. Only real academic sources. Sort alphabetically. No extra blank lines.")
            ]

            generated_parts = {}
            
            for i, (doc_header, prompt_instruction) in enumerate(parts_plan):
                status.info(f"⏳ מנתח וכותב כרגע את: **{doc_header}** (התהליך מבוקר למניעת קריסות, נא להמתין)...")
                
                full_prompt = f"Context:\n{context}\n\nTask: {prompt_instruction}"
                content = call_gemini_direct(full_prompt, api_key, lang)
                
                generated_parts[doc_header] = content
                progress.progress((i + 1) / len(parts_plan))
            
            status.success("🎉 העבודה מוכנה ועוצבה כראוי!")
            
            doc = create_word_document(title, author, generated_parts, lang)
            buffer = io.BytesIO()
            doc.save(buffer)
            buffer.seek(0)
            
            st.download_button(
                label="📥 הורד עבודה סמינריונית למחשב (Word)",
                data=buffer,
                file_name=f"Seminar_{author}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            st.balloons()
