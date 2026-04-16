import streamlit as st
import google.generativeai as genai

st.set_page_config(page_title="Seminar Architect AI", page_icon="🎓")

# משיכת מפתח מה-Secrets
api_key = st.secrets.get("GEMINI_API_KEY")

def generate_part(model, prompt):
    try:
        response = model.generate_content(prompt)
        return response.text
    except Exception as e:
        return f"\n(שגיאה זמנית בייצור חלק זה: {str(e)})\n"

st.title("🎓 Seminar Architect PRO")
st.write("מערכת ליצירת עבודות סמינריוניות עמוקות (15-20 עמודים)")

if not api_key:
    st.error("חסר מפתח API ב-Secrets")
else:
    genai.configure(api_key=api_key)
    # שימוש במודל היציב ביותר
    model = genai.GenerativeModel('gemini-1.5-flash')

    title = st.text_input("נושא העבודה:")
    author = st.text_input("שם הסטודנט:")

    if st.button("צור עבודה מלאה"):
        if not title:
            st.error("נא להזין נושא.")
        else:
            progress = st.progress(0)
            status = st.empty()
            
            # פיצול ל-12 חלקים כדי להבטיח אורך של 17 עמודים
            parts = [
                ("שער ומבוא", f"כתוב מבוא אקדמי ארוך מאוד (3 עמודים) על {title}. כלול רקע, שאלת מחקר וחשיבות."),
                ("סקירת ספרות א", f"סקור תיאוריות מרכזיות בנושא {title}. לפחות 1000 מילים."),
                ("סקירת ספרות ב", f"סקור מחקרים אמפיריים מהשנים האחרונות על {title}."),
                ("סקירת ספרות ג", f"דון בהיבטים חברתיים וכלכליים של {title}."),
                ("סקירת ספרות ד", f"הצג ביקורת וגישות מנוגדות בנושא {title}."),
                ("מתודולוגיה", f"תאר בפירוט את שיטת המחקר עבור {title}."),
                ("ממצאים", f"תאר ממצאים היפותטיים מפורטים מאוד עבור {title}."),
                ("דיון א", f"נתח את הממצאים לעומק (חלק 1)."),
                ("דיון ב", f"נתח את הממצאים לאור הספרות (חלק 2)."),
                ("סיכום", f"כתוב סיכום ומסקנות סופיות."),
                ("המלצות", f"כתוב המלצות ליישום ומחקרי המשך."),
                ("ביבליוגרפיה", f"צור רשימת מקורות APA מלאה (20 מקורות) עבור {title}.")
            ]
            
            full_text = f"עבודה סמינריונית בנושא: {title}\nמגיש: {author}\n\n"
            
            for i, (name, prompt) in enumerate(parts):
                status.write(f"✍️ כותב כרגע: {name}...")
                content = generate_part(model, prompt)
                full_text += f"\n\n--- {name} ---\n\n{content}\n"
                progress.progress((i + 1) / len(parts))
            
            status.success("✅ העבודה מוכנה!")
            
            # הורדה כקובץ טקסט פשוט כדי שוורד לא יסתבך
            st.download_button("📥 הורד עבודה לוורד", full_text, file_name=f"seminar_{author}.txt")
            st.text_area("הטקסט המוכן (ניתן להעתיק מכאן):", full_text, height=400)
