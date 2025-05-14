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
        excel_file_path = r"C:\Users\yisca\Desktop\BetaYeda\Option.xlsx"

        QuestionID = []
        Description = []
        Option1 = []
        Option2 = []
        Option3 = []
        Option4 = []
        current_answer = ""
        inside_answer = False
        current_option = ""

        with pdfplumber.open(pdf_file_path) as pdf: #פתיחת קובץ
            for page in pdf.pages: #מעבר על כל העמודים
                text = page.extract_text() #מזקקת את כל הטקבט מהעמוד הנוכחי
                if text: #בדיקה שאכן יש טקסט בדף
                    lines = text.split("\n")
                    for line in lines:
                        cleaned_line = line.strip()
                        if cleaned_line:
                            # if re.match(r'([א-ד]\s*\.\s*)', cleaned_line):
                            #     current_option = ""
                            #     current_option = cleaned_line
                            # else:
                            #     current_option += cleaned_line
                            #
                            # if current_option.strip():
                            #     reshaped = arabic_reshaper.reshape(current_option.strip())
                            #     fixed_line = get_display(reshaped)
                            #     QuestionID.append([fixed_line])
                            # options = re.split(r'([א-ד]\s*\.\s*)', cleaned_line) #מחזיר רשימה
                            options = re.split(r'\s\.', cleaned_line)
                            current_option = ""
                            if len(options) == 2 and options[0] != "":
                                current_answer = options[0]
                            elif len(options) == 1:
                                current_answer = cleaned_line + current_answer

                            if re.match(r'\.', current_answer):
                                reshaped = arabic_reshaper.reshape(current_answer.strip()) #מתקן שהאותיות יופיעו נכון
                                fixed_line = get_display(reshaped) #3יצייג את כיוון הטקסט מימין לשמאול
                                QuestionID.append([fixed_line]) #מויסף את התשובה

                            # if current_option.strip():
                            #     reshaped = arabic_reshaper.reshape(current_option.strip())
                            #     fixed_line = get_display(reshaped)
                            #     QuestionID.append([fixed_line])

        df = pd.DataFrame(QuestionID, columns=["Description"])
        df.to_excel(excel_file_path, index=False)


        print("עובד")



