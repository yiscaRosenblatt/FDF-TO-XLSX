import pdfplumber
import pandas as pd
import arabic_reshaper
from bidi.algorithm import get_display
import re

class pdf_to_xlsx:
    def __init__(self):
        self.conversion()

    def conversion(self):
        pdf_file_path = r"C:\Users\yisca\Desktop\BetaYeda\מאגר מסכם לשאלות בטיחות ועזרה ראשונה_220531_195438.pdf"
        excel_file_path = r"C:\Users\yisca\Desktop\BetaYeda\questionsInELSX.xlsx"

        QuestionID = []
        Description = []
        Option1 = []
        Option2 = []
        Option3 = []
        Option4 = []
        current_answer = ""
        inside_answer = False


        with pdfplumber.open(pdf_file_path) as pdf: #פתיחת קובץ
            for page in pdf.pages: #מעבר על כל העמודים
                text = page.extract_text() #מזקקת את כל הטקבט מהעמוד הנוכחי
                if text: #בדיקה שאכן יש טקסט בדף
                    lines = text.split("\n")
                    for line in lines:
                        cleaned_line = line.strip()
                        if cleaned_line:
                            options = re.split(r'(\s*[א-ד]\s*\.\s*)', cleaned_line) #מחזיר רשימה
                            current_option = ""
                            for option in options:
                                if re.match(r'(\s*[א-ד]\s*\.\s*)', option):
                                    if current_option.strip():
                                        reshaped = arabic_reshaper.reshape(current_option.strip()) #מתקן שהאותיות יופיעו נכון
                                        fixed_line = get_display(reshaped) #3יצייג את כיוון הטקסט מימין לשמאול
                                        QuestionID.append([fixed_line]) #מויסף את התשובה
                                    current_option = option
                                else:
                                    current_option += " " + cleaned_line  # המשך של אותה תשובה
                            # if current_option.strip():
                            #     reshaped = arabic_reshaper.reshape(current_option.strip())
                            #     fixed_line = get_display(reshaped)
                            #     QuestionID.append([fixed_line])
                                # current_option = ""

        df = pd.DataFrame(QuestionID, columns=["Description"])
        df.to_excel(excel_file_path, index=False)

        print("עובד")



