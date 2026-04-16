import streamlit as st
import google.generativeai as genai
import base64

# הגדרת דף
st.set_page_config(page_title="Seminar Architect AI", page_icon="🎓")

# פונקציה לייצור חלק מהעבודה
def generate_part(model, prompt):
    response = model.generate_content(prompt)
    return response.text

st.title("🎓 Seminar Architect AI")
st.write("צור עבודה סמינריונית מלאה בתוך דקות")

# הגדרות צד
with st.sidebar:
    st.header("הגדרות")
    api_key = st.text_input("הכנס Gemini API Key:", type="password")
    st.info("ניתן להשיג מפתח בחינם ב-Google AI Studio")

# שדות קלט
title = st.text_input("נושא העבודה (לדוגמה: השפעת הבינה המלאכותית על שוק העבודה)")
author = st.text_input("שם הסטודנט")

if st.button("התחל בייצור העבודה"):
    if not api_key:
        st.error("חובה להזין מפתח API")
    elif not title:
        st.error("חובה להזין נושא")
    else:
        try:
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel('gemini-1.5-flash')
            
            progress = st.progress(0)
            status = st.empty()
            
            # שלבים לייצור 17 עמודים (מבנה מודולרי)
            parts = [
                ("מבוא ורקע", "כתוב מבוא ארוך ומפורט מאוד (לפחות 3 עמודים ב-Word) על: " + title),
                ("סקירת ספרות א'", "כתוב סקירה תיאורטית רחבה מאוד על המושגים המרכזיים ב: " + title),
                ("סקירת ספרות ב'", "תאר מחקרים קודמים וגישות שונות בנושא של: " + title),
                ("מתודולוגיה", "תאר את שיטת המחקר, הרציונל וכלי המחקר."),
                ("דיון וסיכום", "כתוב פרק דיון מעמיק, מסקנות והמלצות."),
                ("ביבליוגרפיה", "צור רשימת מקורות אקדמיים אמיתיים בפורמט APA.")
            ]
            
            full_content = f"# {title}\n\nמגיש: {author}\n\n---\n\n"
            
            for i, (name, prompt) in enumerate(parts):
                status.write(f"מייצר כרגע: {name}...")
                content = generate_part(model, prompt)
                full_content += f"## {name}\n\n{content}\n\n"
                progress.progress((i + 1) / len(parts))
            
            status.success("העבודה מוכנה!")
            st.download_button("הורד עבודה (Markdown/Word)", full_content, file_name="seminar.doc")
            st.markdown(full_content)
            
        except Exception as e:
            st.error(f"שגיאה: {str(e)}")

