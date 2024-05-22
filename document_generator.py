import json
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import html2text
import re
import logging

class DocumentGenerator:
    def __init__(self, data_file):
        self.data_file = data_file
        self.document = Document()
        self.h = html2text.HTML2Text()
        self.h.body_width = 0
        self.setup_logging()
        
    def setup_logging(self):
        logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

    def clean_text(self, text):
        pattern = r'\n[ \t]*\n[ \t]*\n'
        text = self.h.handle(text).strip()
        return re.sub(pattern, '\n\n', text)

    def set_table_border(self, table, border_color="000000", border_size="4"):
        tbl = table._element
        tblBorders = OxmlElement('w:tblBorders')
        for border_name in ["top", "left", "bottom", "right", "insideH", "insideV"]:
            border = OxmlElement(f'w:{border_name}')
            border.set(qn('w:val'), 'single')
            border.set(qn('w:sz'), border_size)
            border.set(qn('w:color'), border_color)
            tblBorders.append(border)
        tbl.tblPr.append(tblBorders)

    def set_cell_background(self, cell, color):
        cell_properties = cell._element.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:fill'), color)
        cell_properties.append(shd)

    def set_cell_margins(self, cell, top=0.08, start=0.16, bottom=0.08, end=0.16):
        tcPr = cell._element.get_or_add_tcPr()
        tcMar = OxmlElement('w:tcMar')
        for margin_type, margin_size in [('top', top), ('start', start), ('bottom', bottom), ('end', end)]:
            mar = OxmlElement(f'w:{margin_type}')
            mar.set(qn('w:w'), str(int(margin_size * 1440)))
            mar.set(qn('w:type'), 'dxa')
            tcMar.append(mar)
        tcPr.append(tcMar)

    def set_cell_width(self, cell, width):
        cell.width = Inches(width)

    def build_question(self, quest):
        def add_row(table, name, value):
            value = value if value else "N/A"
            row_cells = table.add_row().cells
            row_cells[0].text = self.clean_text(name)
            row_cells[1].text = self.clean_text(value.strip())
            self.set_cell_background(row_cells[0], "F5F5F5")
            self.set_cell_margins(row_cells[0], .04, .08, .04, .08)
            self.set_cell_width(row_cells[0], .75)
            self.set_cell_background(row_cells[1], "F6FAFF")
            self.set_cell_margins(row_cells[1], .04, .08, .04, .08)
            return table

        question_table = self.document.add_table(rows=0, cols=2, style='Table Grid')
        self.set_table_border(question_table, border_color="CCCCCC", border_size="4")

        row_cells = question_table.add_row().cells
        row_cells[0].merge(row_cells[1])
        row_cells[0].text = self.clean_text(quest["elementNumber"]) + " " + self.clean_text(quest["questionText"])
        self.set_cell_background(row_cells[0], "F5F5F5")
        self.set_cell_margins(row_cells[0], .04, .08, .04, .08)

        if quest["triggerValue"]:
            followup_text = self.clean_text(quest["triggerValue"]) + " on " + self.clean_text(quest["parentElementNumber"])
            question_table = add_row(question_table, "Follow-up", followup_text)

        question_table = add_row(question_table, "Answer", quest["answerValue"])
        question_table = add_row(question_table, "Comment", quest["answerComments"])
        documents = [doc for docList in quest["docListAnswer"] for doc in docList["attachedFiles"]]
        documentAsList = ", ".join(documents)
        question_table = add_row(question_table, "Documents", documentAsList)

    def generate_document(self):
        with open(self.data_file, "r") as f:
            data = json.load(f)

        meta_data = {
            "Assessment": data["pqiBasicInfo"]["questionnaireBasicInfo"]["qStructureBasicInfo"]["name"],
            "Report run on": data["reportGeneratedOn"],
            "Published by": data["pqiBasicInfo"]["publishedByUser"]["companyName"] + " at " + data["pqiBasicInfo"]["publishedByUser"]["userProfile"]["fullName"],
            "Published on": data["pqiBasicInfo"]["publishedDate"]
        }

        evaluation_data = {
            "Completed on": data["pqiBasicInfo"]["evaluationCompleteDate"],
            "Grade": data["pqiBasicInfo"]["gradeDetails"]["name"] if data["pqiBasicInfo"]["gradeDetails"] else "",
            "Score": data["pqiBasicInfo"]["evaluationScore"]
        }

        is_evaluated = bool(data["pqiBasicInfo"]["evaluationCompleteDate"])

        self.document.add_heading('Assessment Export', level=1)
        self.document.add_heading('Details', level=2)
        table = self.document.add_table(rows=0, cols=2, style='Table Grid')
        self.set_table_border(table, border_color="CCCCCC", border_size="4")

        for entry, value in meta_data.items():
            row_cells = table.add_row().cells
            row_cells[0].text = entry
            self.set_cell_background(row_cells[0], "DFEFFD")
            self.set_cell_margins(row_cells[0])
            row_cells[1].text = value
            self.set_cell_background(row_cells[1], "F5F5F5")
            self.set_cell_margins(row_cells[1])

        if is_evaluated:
            for entry, value in evaluation_data.items():
                row_cells = table.add_row().cells
                row_cells[0].text = entry
                row_cells[1].text = value

        self.document.add_page_break()

        for section in data["sections"]:
            section_heading = section["elementNumber"] + " " + section["sectionName"]
            self.document.add_heading(section_heading, level=2)
            for quest in section["questionDetails"]:
                self.build_question(quest)
                self.document.add_paragraph()

        self.document.save('demo.docx')
        logging.info("Document generated successfully.")

if __name__ == "__main__":
    doc_gen = DocumentGenerator("data.json")
    try:
        doc_gen.generate_document()
    except Exception as e:
        logging.error(f"Error generating document: {e}")