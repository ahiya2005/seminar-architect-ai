import streamlit as st
import google.generativeai as genai

# הגדרת דף
st.set_page_config(page_title="Seminar Architect AI", page_icon="🎓")

# ניסיון למשוך את המפתח מה-Secrets, ואם אין - מהתיבה באתר
api_key = st.secrets.get("GEMINI_API_KEY") or st.sidebar.text_input("הכנס Gemini API Key:", type="password")

def generate_part(model, prompt):
    try:
        response = model.generate_content(prompt)
        return response.text
    except Exception as e:
        return f"שגיאה בייצור חלק: {str(e)}"

st.title("🎓 Seminar Architect AI")
st.write("צור עבודה סמינריונית מלאה בתוך דקות")

title = st.text_input("נושא העבודה (לדוגמה: היתר עיסוק בעולם המודרני)")
author = st.text_input("שם הסטודנט")

if st.button("התחל בייצור העבודה"):
    if not api_key:
        st.error("חובה להזין מפתח API ב-Secrets או בתפריט הצד")
    elif not title:
        st.error("חובה להזין נושא")
    else:
        try:
            genai.configure(api_key=api_key)
            # שינוי המודל לשם המעודכן ביותר שגוגל תומכת בו כרגע
            model = genai.GenerativeModel('gemini-1.5-flash')
            
            progress = st.progress(0)
            status = st.empty()
            
            parts = [
                ("מבוא ורקע", f"כתוב מבוא ארוך ואקדמי על {title}. כלול שאלת מחקר וחשיבות הנושא."),
                ("סקירת ספרות", f"כתוב סקירה תיאורטית רחבה מאוד על {title} עם מקורות."),
                ("דיון ומסקנות", f"כתוב פרק דיון מעמיק המנתח את {title} וסיכום סופי."),
                ("ביבליוגרפיה", f"צור רשימת מקורות אקדמיים בפורמט APA עבור {title}.")
            ]
            
            full_content = f"# {title}\n\nמגיש: {author}\n\n---\n\n"
            
            for i, (name, prompt) in enumerate(parts):
                status.write(f"מייצר כרגע: {name}...")
                content = generate_part(model, prompt)
                full_content += f"## {name}\n\n{content}\n\n"
                progress.progress((i + 1) / len(parts))
            
            st.success("העבודה מוכנה!")
            st.download_button("הורד עבודה (Word/Text)", full_content, file_name="seminar.doc")
            st.markdown(full_content)
            
        except Exception as e:
            st.error(f"שגיאת מערכת: {str(e)}")

