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
