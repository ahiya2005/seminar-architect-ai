import streamlit as st
import requests
import json

st.set_page_config(page_title="Seminar Architect AI", page_icon="🎓", layout="wide")

# משיכת המפתח מה-Secrets
api_key = st.secrets.get("GEMINI_API_KEY")

def call_gemini_direct(prompt, key):
    # הכתובת המדויקת מה-CURL שעבד לך ב-AI Studio
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key={key}"
    
    payload = {
        "contents": [{
            "parts": [{"text": prompt}]
        }],
        "generationConfig": {
            "temperature": 0.8,
            "maxOutputTokens": 3000, # העליתי את הכמות כדי שיהיה יותר טקסט
        }
    }
    
    headers = {'Content-Type': 'application/json'}
    
    try:
        response = requests.post(url, headers=headers, data=json.dumps(payload))
        if response.status_code == 200:
            result = response.json()
            return result['candidates'][0]['content']['parts'][0]['text']
        else:
            return f"שגיאה (קוד {response.status_code}): {response.text}"
    except Exception as e:
        return f"שגיאת התחברות: {str(e)}"

st.title("🎓 Seminar Architect PRO")
st.write("מערכת ייצור עבודות סמינריוניות - גרסה יציבה")

if not api_key:
    st.error("אנא הגדר GEMINI_API_KEY ב-Secrets של Streamlit.")
else:
    title = st.text_input("נושא העבודה:")
    author = st.text_input("שם הסטודנט:")

    if st.button("צור עבודה של 17 עמודים"):
        if not title:
            st.error("הכנס נושא.")
        else:
            progress = st.progress(0)
            status = st.empty()
            
            # פיצול ל-8 חלקים רחבים מאוד
            parts = [
                ("שער ומבוא", f"כתוב מבוא אקדמי מפורט וארוך מאוד (לפחות 1200 מילים) על {title}. כלול רקע, שאלת מחקר וחשיבות."),
                ("סקירת ספרות - חלק א'", f"כתוב סקירה תיאורטית מעמיקה על המושגים המרכזיים ב{title}. השתמש בשפה אקדמית גבוהה."),
                ("סקירת ספרות - חלק ב'", f"סקור מחקרים קודמים וגישות שונות בנושא {title} מהשנים האחרונות."),
                ("מתודולוגיה", f"תאר בפירוט את שיטת המחקר, כלי המחקר והליך איסוף הנתונים עבור {title}."),
                ("פרק הממצאים", f"הצג ממצאים היפותטיים מפורטים מאוד עם תתי-כותרות עבור {title}."),
                ("דיון וניתוח", f"נתח את הממצאים באופן ביקורתי ומעמיק. כתוב לפחות 4 עמודים של ניתוח."),
                ("סיכום והמלצות", f"סכם את העבודה, הצג מסקנות מעשיות והמלצות למחקרי המשך."),
                ("ביבליוגרפיה", f"צור רשימת מקורות אקדמיים מלאה (20 מקורות) בפורמט APA עבור {title}.")
            ]
            
            full_text = f"# {title}\nמגיש: {author}\nתאריך: אפריל 2026\n\n"
            full_text += "="*30 + "\n\n"
            
            for i, (name, p_prompt) in enumerate(parts):
                status.write(f"✍️ כותב כרגע: **{name}** (זה לוקח זמן כי הטקסט ארוך)...")
                content = call_gemini_direct(p_prompt, api_key)
                full_text += f"\n\n## {name}\n\n{content}\n"
                full_text += "\n" + "-"*50 + "\n" # קו מפריד בין פרקים
                progress.progress((i + 1) / len(parts))
            
            status.success("✅ העבודה מוכנה בשלמותה!")
            # ייצוא כקובץ טקסט פשוט כדי שוורד יקרא אותו בלי בעיות
            st.download_button("📥 הורד עבודה לוורד", full_text, file_name=f"seminar_{author}.txt")
            st.markdown(full_text)
