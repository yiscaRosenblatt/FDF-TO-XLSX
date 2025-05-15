import pandas as pd
from docx import Document
import arabic_reshaper
from bidi.algorithm import get_display
import re


class docx_to_xlsx:
    def __init__(self):
        self.conversion()

    def conversion(self):
        docx_file_path = r"C:\Users\yisca\Desktop\BetaYeda\מאגר_מסכם_לשאלות_בטיחות_ועזרה_ראשונה_220531_195438[1].docx"
        excel_file_path = r"C:\Users\yisca\Desktop\BetaYeda\outputdox3.xlsx"

        data = []
        REQUIRED_HEADERS = [
            "Step Name", "Widget Type", "Description", "Required", "Graded",
            "Instant Response", "Option 1", "Option 2", "Option 3", "Option 4", "Correct Answer"
        ]

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
        num_answer = 1
        # פתיחת קובץ DOCX
        doc = Document(docx_file_path)

        for paragraph in doc.paragraphs:
            cleaned_line = paragraph.text.strip()
            # זיהוי תיאור (Description)
            if re.search(r'\?', cleaned_line) or re.search(r':', cleaned_line):
                num_answer = self.nextq(data, columns)
                columns["Description"] = cleaned_line.strip()

            elif re.compile(r"[א-ד]+\.").findall(cleaned_line):

                # זיהוי אפשרויות (Options)
                options = re.compile(r"[א-ד]+\.").split(cleaned_line)

                for word in options:
                    if word != '':
                        columns[f"Option {num_answer}"] = word
                        num_answer += 1

        # המרה לאקסל
        df = pd.DataFrame(data, columns=["Step Name", "Widget Type", "Description", "Required", "Graded",
                                         "Instant Response", "Option 1", "Option 2", "Option 3", "Option 4",
                                         "Correct Answer"])
        df.to_excel(excel_file_path, index=False)
        print("work")

    def nextq(self, data, columns):
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
        return 1



