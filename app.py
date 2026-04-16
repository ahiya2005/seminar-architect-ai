import streamlit as st
import google.generativeai as genai

st.set_page_config(page_title="Seminar Architect AI", page_icon="🎓", layout="wide")

# משיכת מפתח מה-Secrets
api_key = st.secrets.get("GEMINI_API_KEY")

def generate_part(model, prompt):
    try:
        config = genai.types.GenerationConfig(
            temperature=0.7,
            top_p=0.9,
            max_output_tokens=2048,
        )
        response = model.generate_content(prompt, generation_config=config)
        return response.text
    except Exception as e:
        return f"Error: {str(e)}"

st.title("🎓 Seminar Architect PRO")
st.write("מערכת ליצירת עבודות סמינריוניות (15-20 עמודים)")

if not api_key:
    st.error("שגיאה: מפתח API לא הוגדר ב-Secrets.")
else:
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('gemini-1.5-flash-latest')

    title = st.text_input("נושא העבודה:")
    author = st.text_input("שם הסטודנט:")

    if st.button("צור עבודה מלאה"):
        if not title:
            st.error("נא להזין נושא.")
        else:
            progress = st.progress(0)
            status = st.empty()
            
            parts = [
                ("מבוא", f"כתוב מבוא אקדמי ארוך על {title}."),
                ("סקירת ספרות א", f"כתוב סקירה תיאורטית על {title}."),
                ("סקירת ספרות ב", f"כתוב על מחקרים קודמים בנושא {title}."),
                ("מתודולוגיה", f"תאר את שיטת המחקר עבור {title}."),
                ("דיון וסיכום", f"נתח את הנושא {title} וסכם."),
                ("ביבליוגרפיה", f"צור רשימת מקורות APA עבור {title}.")
            ]
            
            full_content = f"# {title}\n\nמגיש: {author}\n\n---\n\n"
            
            for i, (name, prompt) in enumerate(parts):
                status.write(f"✍️ כותב כרגע: {name}...")
                content = generate_part(model, prompt)
                full_content += f"\n\n## {name}\n\n{content}\n"
                progress.progress((i + 1) / len(parts))
            
            status.success("✅ העבודה מוכנה!")
            st.download_button("📥 הורד עבודה", full_content, file_name="seminar.doc")
            st.markdown(full_content)

    if st.button("צור עבודה מלאה (תהליך של כ-5 דקות)"):
        if not title:
            st.error("נא להזין נושא.")
        else:
            progress = st.progress(0)
            status = st.empty()
            
            # פיצול ל-10 חלקים כדי להגיע ל-17 עמודים
            parts = [
                ("עמוד שער ומבוא", f"כתוב מבוא אקדמי מורחב מאוד (3 עמודים) על {title}. כלול רקע היסטורי, רציונל וחשיבות המחקר."),
                ("סקירת ספרות - חלק א'", f"כתוב סקירה תיאורטית מעמיקה על המושגים המרכזיים ב{title}. לפחות 800 מילים."),
                ("סקירת ספרות - חלק ב'", f"סקור מחקרים מהעשור האחרון שנעשו בנושא {title}. הצג גישות שונות."),
                ("סקירת ספרות - חלק ג'", f"נתח את הפער המחקרי הקיים בנושא {title} וכיצד עבודה זו עונה עליו."),
                ("מתודולוגיה", f"תאר בפירוט את שיטת המחקר, כלי המחקר, אוכלוסיית המחקר והליך איסוף הנתונים."),
                ("פרק הממצאים", f"צור פרק ממצאים היפותטי (או מבוסס נתונים) מפורט מאוד עם תתי-פרקים."),
                ("דיון וניתוח", f"נתח את הממצאים לאור סקירת הספרות. כתוב לפחות 3 עמודים של ניתוח ביקורתי."),
                ("סיכום ומסקנות", f"סכם את עיקרי העבודה, תאר את המגבלות והצע מחקרי המשך."),
                ("ביבליוגרפיה", f"צור רשימת מקורות אקדמיים (לפחות 15 מקורות) בפורמט APA תקני.")
            ]
            
            full_content = f"# {title}\n\nמגיש: {author}\n\nתאריך: אפריל 2026\n\n---\n\n"
            
            for i, (name, prompt) in enumerate(parts):
                status.write(f"✍️ כותב כרגע: **{name}** (אל תסגור את הדף)...")
                content = generate_part(model, prompt)
                full_content += f"\n\n## {name}\n\n{content}\n"
                progress.progress((i + 1) / len(parts))
            
            status.success("✅ העבודה מוכנה בשלמותה!")
            
            # יצירת קובץ להורדה (בפורמט שוורד יקרא בקלות)
            st.download_button("📥 הורד עבודה מוכנה לוורד", full_content, file_name=f"seminar_{author}.doc")
            st.markdown("---")
            st.markdown(full_content)

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

