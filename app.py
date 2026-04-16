import streamlit as st
import requests
import json

st.set_page_config(page_title="Seminar Architect AI", page_icon="🎓", layout="wide")

# משיכת המפתח מה-Secrets
api_key = st.secrets.get("GEMINI_API_KEY")

def call_gemini_direct(prompt, key):
    # הכתובת הישירה מה-CURL שלך
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={key}"
    
    payload = {
        "contents": [{
            "parts": [{"text": prompt}]
        }],
        "generationConfig": {
            "temperature": 0.7,
            "maxOutputTokens": 2048,
        }
    }
    
    headers = {'Content-Type': 'application/json'}
    
    response = requests.post(url, headers=headers, data=json.dumps(payload))
    
    if response.status_code == 200:
        result = response.json()
        return result['candidates'][0]['content']['parts'][0]['text']
    else:
        return f"שגיאה מהשרת ({response.status_code}): {response.text}"

st.title("🎓 Seminar Architect PRO")
st.write("ייצור עבודה אקדמית מלאה באמצעות גישה ישירה ל-API")

if not api_key:
    st.error("אנא הגדר GEMINI_API_KEY ב-Secrets.")
else:
    title = st.text_input("נושא העבודה:")
    author = st.text_input("שם הסטודנט:")

    if st.button("צור עבודה של 17 עמודים"):
        if not title:
            st.error("הכנס נושא.")
        else:
            progress = st.progress(0)
            status = st.empty()
            
            # רשימה מורחבת מאוד כדי להגיע לנפח דפים
            parts = [
                ("שער ומבוא", f"כתוב מבוא אקדמי מורחב מאוד על {title}. כלול רקע היסטורי, רציונל וחשיבות המחקר (לפחות 1000 מילים)."),
                ("סקירת ספרות א'", f"סקור תיאוריות קלאסיות בנושא {title}. כתוב בעומק רב עם מושגים מקצועיים."),
                ("סקירת ספרות ב'", f"נתח מחקרים עכשוויים (2020-2026) בנושא {title} והצג את הפער המחקרי."),
                ("מתודולוגיה", f"תאר בפירוט את שיטת המחקר, כלי המחקר והליך איסוף הנתונים עבור {title}."),
                ("ממצאים ודיון", f"הצג ממצאים היפותטיים ונתח אותם באופן ביקורתי מול הספרות שסקרת."),
                ("סיכום והמלצות", f"סכם את המסקנות העיקריות והצע המלצות למדיניות ולמחקר עתידי."),
                ("ביבליוגרפיה", f"צור רשימת מקורות אקדמיים מלאה (20 מקורות) בפורמט APA תקני עבור {title}.")
            ]
            
            full_text = f"# {title}\nמגיש: {author}\nתאריך: אפריל 2026\n\n"
            
            for i, (name, p_prompt) in enumerate(parts):
                status.write(f"✍️ כותב כרגע: {name}...")
                content = call_gemini_direct(p_prompt, api_key)
                full_text += f"\n\n## {name}\n\n{content}\n"
                progress.progress((i + 1) / len(parts))
            
            status.success("✅ העבודה מוכנה!")
            st.download_button("📥 הורד עבודה לוורד", full_text, file_name=f"seminar_{author}.txt")
            st.markdown(full_text)
