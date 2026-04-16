import streamlit as st
import requests
import json
import io
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import PyPDF2

# --- הגדרות עיצוב ודף ---
st.set_page_config(page_title="Seminar Architect PRO", page_icon="🎓", layout="wide")

# עיצוב מודרני ובהיר (CSS מותאם)
st.markdown("""
    <style>
    .main {background-color: #f8f9fa;}
    h1 {color: #2c3e50; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;}
    .stButton>button {background-color: #3498db; color: white; border-radius: 8px; border: none; padding: 10px 24px; font-weight: bold;}
    .stButton>button:hover {background-color: #2980b9;}
    .stTextInput>div>div>input, .stTextArea>div>div>textarea {border-radius: 8px; border: 1px solid #bdc3c7;}
    </style>
    """, unsafe_allow_html=True)

# --- פונקציות עזר ---
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

def call_gemini_direct(prompt, key, lang="עברית"):
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key={key}"
    
    # פקודת בסיס שמונעת המצאת מקורות ומכריחה כתיבה אקדמית
    system_instruction = "You are an expert academic writer. You must write at the highest academic standard. CRITICAL: DO NOT invent or hallucinate citations. Use only real, peer-reviewed academic sources. If you do not know a real source for a claim, state the claim generally without a fake citation. Use proper APA formatting for all references. "
    if lang == "עברית":
        system_instruction += "כתוב בשפה עברית אקדמית, תקינה ועשירה."
    else:
        system_instruction += "Write in high-level, formal academic English."

    payload = {
        "contents": [{"parts": [{"text": system_instruction + "\n\n" + prompt}]}],
        "generationConfig": {"temperature": 0.6, "maxOutputTokens": 3500}
    }
    
    try:
        response = requests.post(url, headers={'Content-Type': 'application/json'}, data=json.dumps(payload))
        if response.status_code == 200:
            return response.json()['candidates'][0]['content']['parts'][0]['text']
        else:
            return f"Error ({response.status_code}): {response.text}"
    except Exception as e:
        return f"Connection Error: {str(e)}"

def create_word_document(title, author, parts_dict, lang):
    doc = Document()
    
    # הגדרת פונט אחיד למסמך (David לעברית)
    style = doc.styles['Normal']
    font = style.font
    font.name = 'David' if lang == "עברית" else 'Times New Roman'
    font.size = Pt(12)
    
    # עמוד שער
    doc.add_paragraph('\n\n\n')
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_p.add_run(title)
    title_run.bold = True
    title_run.font.size = Pt(24)
    
    author_p = doc.add_paragraph()
    author_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    author_p.add_run(f"\nמגיש: {author}" if lang == "עברית" else f"\nSubmitted by: {author}")
    doc.add_page_break()

    # תוכן
    for part_name, content in parts_dict.items():
        h = doc.add_heading(part_name, level=1)
        h.alignment = WD_ALIGN_PARAGRAPH.RIGHT if lang == "עברית" else WD_ALIGN_PARAGRAPH.LEFT
        
        # ניקוי טקסט ופסקאות
        paragraphs = content.split('\n')
        for p_text in paragraphs:
            if p_text.strip():
                p = doc.add_paragraph(p_text.replace('#', '').replace('*', ''))
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if lang == "עברית" else WD_ALIGN_PARAGRAPH.LEFT
                p.paragraph_format.line_spacing = 1.5 # רווח 1.5 סטנדרטי
        doc.add_page_break()
        
    return doc

# --- ניהול מצב (Session State) ---
if 'logged_in' not in st.session_state:
    st.session_state['logged_in'] = False
if 'email' not in st.session_state:
    st.session_state['email'] = ""

# --- ממשק משתמש ---
st.title("🎓 Seminar Architect PRO")

# בחירת שפה
lang = st.radio("בחר שפת ממשק / Choose Language", ["עברית", "English"], horizontal=True)

if not st.session_state['logged_in']:
    st.subheader("התחברות למערכת / Login" if lang == "עברית" else "Login to System")
    email_input = st.text_input("אימייל / Email:")
    if st.button("התחבר / Login"):
        if email_input:
            st.session_state['email'] = email_input
            st.session_state['logged_in'] = True
            st.rerun()
        else:
            st.warning("נא להזין אימייל." if lang == "עברית" else "Please enter an email.")

else:
    st.success(f"מחובר כ: {st.session_state['email']}" if lang == "עברית" else f"Logged in as: {st.session_state['email']}")
    
    col1, col2 = st.columns(2)
    with col1:
        title = st.text_input("נושא העבודה:" if lang == "עברית" else "Seminar Topic:")
        author = st.text_input("שם הסטודנט:" if lang == "עברית" else "Student Name:")
    with col2:
        description = st.text_area("מיקוד והנחיות (על מה לשים דגש?):" if lang == "עברית" else "Specific focus/instructions:", height=130)
    
    st.markdown("### 📎 העלאת הנחיות מרצה (אופציונלי)" if lang == "עברית" else "### 📎 Upload Lecturer Guidelines (Optional)")
    uploaded_file = st.file_uploader("העלה קובץ PDF או TXT" if lang == "עברית" else "Upload PDF or TXT", type=['pdf', 'txt'])

    if st.button("🚀 צור עבודה סמינריונית (17 עמודים)" if lang == "עברית" else "🚀 Generate Seminar Paper"):
        if not title:
            st.error("חובה להזין נושא!" if lang == "עברית" else "Topic is required!")
        elif not api_key:
            st.error("API Key is missing in secrets.")
        else:
            # קריאת הנחיות קובץ
            lecturer_notes = extract_text_from_file(uploaded_file)
            
            # בניית הקשר מלא לבינה המלאכותית
            context = f"Topic: {title}\n"
            if description: context += f"Specific Focus: {description}\n"
            if lecturer_notes: context += f"Lecturer Strict Guidelines: {lecturer_notes}\n"

            progress = st.progress(0)
            status = st.empty()
            
            # מבנה העבודה
            if lang == "עברית":
                parts = [
                    ("מבוא", "כתוב מבוא אקדמי מורחב מאוד (לפחות 3 עמודים). הצג את שאלת המחקר והרציונל."),
                    ("סקירת ספרות - תיאוריות", "סקור תיאוריות מרכזיות בנושא. כתוב בעומק עם מושגים אקדמיים."),
                    ("סקירת ספרות - מחקרים", "סקור מחקרים אמפיריים מהשנים האחרונות וזהה פער מחקרי."),
                    ("מתודולוגיה", "תאר בפירוט את שיטת המחקר, אוכלוסיית המחקר, וכלי המחקר."),
                    ("ממצאים", "הצג ממצאים היפותטיים מפורטים בחלוקה לנושאים."),
                    ("דיון", "נתח את הממצאים מול סקירת הספרות. זהו הפרק החשוב ביותר, הרחב מאוד."),
                    ("סיכום", "סכם את העבודה, ציין מגבלות והצע מחקרי המשך."),
                    ("ביבליוגרפיה", "רשימת מקורות אקדמיים אמיתיים בלבד בפורמט APA (לפחות 15). אסור להמציא!")
                ]
            else:
                parts = [
                    ("Introduction", "Write a very extensive academic introduction. Include background and research question."),
                    ("Literature Review - Theories", "Review core theories related to the topic in depth."),
                    ("Literature Review - Studies", "Review empirical studies from recent years. Identify the research gap."),
                    ("Methodology", "Detail the research method, population, and tools."),
                    ("Findings", "Present detailed hypothetical findings categorized by themes."),
                    ("Discussion", "Critically analyze the findings against the literature review. Expand significantly."),
                    ("Conclusion", "Summarize, state limitations, and suggest future research."),
                    ("References", "List only REAL academic sources in APA format. Do not hallucinate sources!")
                ]

            generated_parts = {}
            
            # אנימציות ואימוג'י רצים
            for i, (part_name, specific_prompt) in enumerate(parts):
                status.info(f"⏳ מנתח וכותב כרגע את פרק: **{part_name}**... נא להמתין" if lang == "עברית" else f"⏳ Writing chapter: **{part_name}**... Please wait")
                
                full_prompt = f"Based on this context:\n{context}\n\nTask: {specific_prompt}"
                content = call_gemini_direct(full_prompt, api_key, lang)
                
                generated_parts[part_name] = content
                progress.progress((i + 1) / len(parts))
            
            status.success("🎉 העבודה מוכנה ועוצבה בהצלחה!" if lang == "עברית" else "🎉 The paper is ready and formatted!")
            
            # יצירת קובץ Word אמיתי ומעוצב
            doc = create_word_document(title, author, generated_parts, lang)
            
            # שמירה לזיכרון להורדה
            buffer = io.BytesIO()
            doc.save(buffer)
            buffer.seek(0)
            
            st.download_button(
                label="📥 הורד עבודה סמינריונית (Word)" if lang == "עברית" else "📥 Download Seminar (Word)",
                data=buffer,
                file_name=f"Seminar_{author}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            st.balloons()
