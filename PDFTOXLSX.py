import pdfplumber
import pandas as pd
import arabic_reshaper
from bidi.algorithm import get_display
import re
import fitz  # PyMuPDF

class pdf_to_xlsx:
    def __init__(self):
        self.conversion()

    def conversion(self):
        pdf_file_path = r"C:\Users\yisca\Desktop\BetaYeda\מאגר מסכם לשאלות בטיחות ועזרה ראשונה_220531_195438.pdf"
        excel_file_path = r"C:\Users\yisca\Desktop\BetaYeda\Option15.xlsx"

        data = []
        columns = {
            "Step Name": "",
            "Widget Type": "",
            "Description": "",
            "Required": "",
            "Graded": "",
            "Instant Response": "",
            "Option 1": "",
            "Option 2": "",
            "Option 3": "",
            "Option 4": "",
            "Correct Answer": ""
        }
        min_answer_font_size = 10

        current_description = ""
        current_answer = ""
        inside_answer = False
        current_option = ""
        num_answer = 1
        allowed_fonts = {"Admin"}

        # frame_rects = []
        # with fitz.open(pdf_file_path) as doc:
        #     for page in doc:
        #         for block in page.get_text("dict")["blocks"]:
        #             if block.get("type") == 0:  # רק בלוקים טקסט
        #                 x0, y0, x1, y1 = block["bbox"]
        #                 # שמירת כל תיבות המסגרת (כולל מלבנים, קווים ואובייקטים גרפיים)
        #                 if "Rectangle" in block.get("lines", [{}])[0].get("spans", [{}])[0].get("font", ""):
        #                     frame_rects.append((x0, y0, x1, y1))

        with pdfplumber.open(pdf_file_path) as pdf: #פתיחת קובץ
            for page in pdf.pages: #מעבר על כל העמודים
                text = page.extract_text() #מזקקת את כל הטקבט מהעמוד הנוכחי
                if text: #בדיקה שאכן יש טקסט בדף
                    lines = text.split("\n")
                    for line in lines:
                        cleaned_line = line.strip()
                        # for line in block["lines"]:
                        #     for span in line["spans"]:
                        #         text = span["text"].strip()
                        #         font_name = span["font"].split(",")[0]  # רק שם הפונט בלי סגנון
                        #         font_size = span["size"]
                        #
                        #         # דילוג על טקסטים בפונטים שונים מהפונט הרגיל
                        #         if font_name not in allowed_fonts:
                        #             continue

                        if cleaned_line:
                            if line == 5:
                                current_description = cleaned_line
                            if re.match(r'\?', cleaned_line) or re.match(':', cleaned_line):
                                current_description = cleaned_line + current_description
                                columns["Description"] = current_description
                            # for char in lines:
                            #     font_size = char["size"]
                            options = re.split(r'\s\.', cleaned_line)
                            if len(options) == 2 and options[0] != "":
                                current_answer = options[0]
                            elif len(options) == 1:
                                current_answer = cleaned_line + current_answer
                            elif len(options) == 3:
                                columns[f"Option {num_answer}"] = current_answer
                                num_answer += 1

                            if re.match(r'\.', current_answer):
                                columns[f"Option {num_answer}"] = current_answer
                                current_description = ""

                                num_answer += 1
                                if num_answer == 5:
                                    for key in columns:
                                        reshaped = arabic_reshaper.reshape(columns[key])
                                        fixed_line = get_display(reshaped)
                                        columns[key] = fixed_line
                                    data.append([
                                        columns["Step Name"],
                                        columns["Widget Type"],
                                        columns["Description"],
                                        columns["Required"],
                                        columns["Graded"],
                                        columns["Instant Response"],
                                        columns["Option 1"],
                                        columns["Option 2"],
                                        columns["Option 3"],
                                        columns["Option 4"],
                                        columns["Correct Answer"],
                                    ])
                                    num_answer = 1


                                    # reshaped = arabic_reshaper.reshape(columns) #מתקן שהאותיות יופיעו נכון
                                    # fixed_line = get_display(reshaped) #3יצייג את כיוון הטקסט מימין לשמאול
                                    # QuestionID.append([fixed_line]) #מויסף את התשובה
                                    # num_answer = 0

        df = pd.DataFrame(data, columns=["Step Name", "Widget Type", "Description", "Required", "Graded",
                                         "Instant Response", "Option 1", "Option 2", "Option 3", "Option 4", "Correct Answer"])
        df.to_excel(excel_file_path, index=False)


        print("עובד")



