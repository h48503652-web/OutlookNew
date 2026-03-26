## 📧 OutlookDraft-Bridge 
### Seamless Web-to-Desktop Integration for Automated Email Drafting

![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)
![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)
![Flask](https://img.shields.io/badge/flask-%23000.svg?style=flat&logo=flask&logoColor=white)

**OutlookDraft-Bridge** הוא פתרון חכם המגשר על הפער שבין אפליקציית Web לבין תוכנת Outlook המקומית. הפרויקט מאפשר לייצר טיוטות מייל מרובות (אחת לכל נמען) הכוללות **קבצים מצורפים מהמחשב המקומי** — פעולה שהיא בלתי אפשרית לביצוע באמצעות פרוטוקול `mailto:` הסטנדרטי.

---

## 🚀 למה הפרויקט הזה נולד? (Problem vs. Solution)

פיתוח אפליקציות Web נתקל לעיתים קרובות בחומת האבטחה של הדפדפן (Sandbox). כשמנסים לפתוח את Outlook עם קובץ מצורף, נתקלים במגבלות הבאות:

* **הבעיה ב-`mailto:`:** אינו תומך בצירוף קבצים באופן אמין, מוגבל מאוד באורך התווים, ואינו מסוגל לייצר מספר טיוטות נפרדות בלחיצה אחת.
* **מגבלות הדפדפן:** מטעמי אבטחה, לדפדפן אין גישה למערכת הקבצים (FileSystem) או ליכולת להפעיל תוכנות חיצוניות דרך COM API.

**הפתרון:** שימוש ב-Helper מקומי מבוסס Python המשמש כ"גשר". ה-Web שולח בקשה לשרת מקומי, שמבצע את הפקודות ישירות מול Outlook על המחשב.

---

## 🛠️ ארכיטקטורת המערכת

המערכת בנויה משני רכיבים עיקריים שעובדים בסנכרון:

1.  **Frontend (Vanilla JS/HTML):** ממשק משתמש נקי המאפשר הזנת נתוני נמענים ובחירת קובץ.
2.  **Local Helper (Flask + PyWin32):** שרת Micro-service מקומי המאזין לכתובת ה-Loopback (`127.0.0.1`) ומפעיל את Outlook באמצעות אובייקטי COM של Windows.

---

## ✨ תכונות עיקריות (Features)

* **Multi-Draft Generation:** יצירת טיוטות נפרדות למספר נמענים בו-זמנית בלחיצת כפתור אחת.
* **Local Attachment Support:** יכולת "להזריק" קבצים פיזיים מהמחשב ישירות לתוך הטיוטה ב-Outlook.
* **User in the Loop:** המערכת אינה שולחת את המייל אוטומטית (ללא ידיעת המשתמש), אלא פותחת אותו לבדיקה ועריכה סופית — אידיאלי למניעת טעויות.
* **Zero Cloud Dependency:** כל המידע נשאר מקומי על המכונה של המשתמש, ללא שליחת נתונים לשרתים חיצוניים.

---

## ⚙️ התקנה והרצה (Quick Start)

### דרישות קדם
* מערכת הפעלה **Windows** (נדרש עבור קישוריות ל-Outlook).
* **Microsoft Outlook** מותקן ומחובר לחשבון פעיל.
* **Python 3.10** ומעלה.

### 1. הגדרת ה-Helper המקומי
```powershell
# ניווט לתיקיית ה-helper
cd helper

# יצירת סביבה וירטואלית והפעלתה
python -m venv .venv
.venv\Scripts\Activate.ps1

# התקנת ספריות נדרשות
pip install -r requirements.txt

# הרצת השרת המקומי
python app.py
