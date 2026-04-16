import streamlit as st
import google.generativeai as genai

st.set_page_config(page_title="Seminar Architect AI", page_icon="🎓")

# משיכת מפתח מה-Secrets
api_key = st.secrets.get("GEMINI_API_KEY")

def get_working_model(api_key):
    genai.configure(api_key=api_key)
    # רשימת שמות מודלים אפשריים לפי סדר עדיפות
    for model_name in ['gemini-1.5-flash', 'gemini-1.5-pro', 'gemini-pro']:
        try:
            model = genai.GenerativeModel(model_name)
            # בדיקה קצרה אם המודל מגיב
            model.generate_content("hi", generation_config={"max_output_tokens": 1})
            return model
        except:
            continue
    return None

def generate_part(model, prompt):
    try:
        response = model.generate_content(prompt)
        return response.text
    except Exception as e:
        return f"\n(שגיאה בייצור חלק זה: {str(e)})\n"

st.title("🎓 Seminar Architect PRO")
st.write("מערכת ליצירת עבודות סמינריוניות (15-20 עמודים)")

if not api_key:
    st.error("מפתח API לא נמצא ב-Secrets של Streamlit.")
else:
    title = st.text_input("נושא העבודה:")
    author = st.text_input("שם הסטודנט:")

    if st.button("צור עבודה מלאה"):
        if not title:
            st.error("נא להזין נושא.")
        else:
            model = get_working_model(api_key)
            if not model:
                st.error("שגיאה: לא נמצא מודל Gemini זמין. בדוק את תקינות המפתח שלך.")
            else:
                progress = st.progress(0)
                status = st.empty()
                
                parts = [
                    ("מבוא", f"כתוב מבוא אקדמי ארוך ומעמיק (לפחות 3 עמודים) על {title}. כלול רקע, שאלת מחקר וחשיבות."),
                    ("סקירת ספרות א", f"סקור תיאוריות מרכזיות ומאמרים קלאסיים בנושא {title}. לפחות 1000 מילים."),
                    ("סקירת ספרות ב", f"סקור מחקרים אמפיריים מהשנים האחרונות (2020-2026) על {title}."),
                    ("מתודולוגיה", f"תאר בפירוט את שיטת המחקר, הרציונל והכלים עבור {title}."),
                    ("ממצאים ודיון", f"תאר ממצאים מרכזיים ונתח אותם באופן ביקורתי לאור הספרות."),
                    ("סיכום והמלצות", f"כתוב סיכום סופי, מסקנות מעשיות והמלצות למחקרי המשך."),
                    ("ביבליוגרפיה", f"צור רשימת מקורות אקדמיים מלאה (20 מקורות) בפורמט APA עבור {title}.")
                ]
                
                full_text = f"# עבודה סמינריונית: {title}\n## מגיש: {author}\n\n---\n\n"
                
                for i, (name, prompt) in enumerate(parts):
                    status.write(f"✍️ כותב כרגע: **{name}**...")
                    content = generate_part(model, prompt)
                    full_text += f"\n\n## {name}\n\n{content}\n"
                    progress.progress((i + 1) / len(parts))
                
                status.success("✅ העבודה מוכנה!")
                # הורדה כקובץ .md - וורד יודע לפתוח אותו מושלם עם כותרות
                st.download_button("📥 הורד עבודה לוורד", full_text, file_name=f"seminar_{author}.md")
                st.markdown("### תצוגה מקדימה:")
                st.markdown(full_text)
