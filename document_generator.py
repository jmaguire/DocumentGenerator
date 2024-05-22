import json
from typing import Any, Dict
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import html2text
import re
import logging


class DocumentGenerator:
    BORDER_COLOR_DEFAULT = "000000"
    BORDER_SIZE_DEFAULT = "4"
    CELL_BG_COLOR_WHITE = "FFFFFF"
    CELL_BG_COLOR_HEADER = "F5F5F5"
    CELL_BG_COLOR_BODY = "F6FAFF"
    TABLE_STYLE = 'Table Grid'
    BORDER_COLOR_GREY = "CCCCCC"

    def __init__(self, data_file):
        self.data_file = data_file
        self.document = Document()
        self.h = html2text.HTML2Text()
        self.h.body_width = 0
        self.setup_logging()

    def setup_logging(self):
        logging.basicConfig(level=logging.INFO,
                            format='%(asctime)s - %(levelname)s - %(message)s')

    def clean_text(self, text):
        """
        Clean and format text from HTML to plain text.

        :param text: The HTML text to clean.
        :return: The cleaned plain text.
        """
        pattern = r'\n[ \t]*\n[ \t]*\n'
        text = self.h.handle(text).strip()
        return re.sub(pattern, '\n\n', text)

    def parse_markdown_table(self, markdown):
        lines = markdown.strip().split('\n')
        headers = re.split(r'\s*\|\s*', lines[0].strip())
        rows = [re.split(r'\s*\|\s*', line.strip()) for line in lines[2:]]
        return headers, rows

    def handle_answer_type(self, answerValue, answerType):
        if answerType == "Currency":
            return answerValue.replace(" ", "")
        if answerType == "Percentage":
            return answerValue.strip() + "%"
        if answerType == "Multiple Choice(Select all that apply)":
            return answerValue.strip()
        return self.clean_text(answerValue)

    def set_table_border(self, table, border_color=BORDER_COLOR_DEFAULT, border_size=BORDER_SIZE_DEFAULT):
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
        def add_question_row(table, name, value):
            value = value if value else "N/A"
            row_cells = table.add_row().cells
            row_cells[0].text = name
            row_cells[1].text = value
            self.set_cell_background(row_cells[0], self.CELL_BG_COLOR_HEADER)
            self.set_cell_margins(row_cells[0], .04, .08, .04, .08)
            self.set_cell_width(row_cells[0], .75)
            self.set_cell_background(row_cells[1], self.CELL_BG_COLOR_BODY)
            self.set_cell_margins(row_cells[1], .04, .08, .04, .08)
            return table
        
        def add_question_table_row(table, name, value):
            table_data = self.parse_markdown_table(value)
            cols = len(table_data[0])
            rows = len(table_data[1])
            row_cells = table.add_row().cells
            row_cells[0].text = name
            table_answer = row_cells[1].add_table(rows=0, cols=cols)
            self.set_table_border(table_answer, border_color=self.BORDER_COLOR_GREY, border_size=self.BORDER_SIZE_DEFAULT)
            hdr_cells = table_answer.add_row().cells
            for index, value in enumerate(table_data[0]):
                hdr_cells[index].text = value
                self.set_cell_background(hdr_cells[index], self.CELL_BG_COLOR_WHITE)
            for row in table_data[1]:
                data_cells = table_answer.add_row().cells
                for index, value in enumerate(row):
                    data_cells[index].text = value
                    self.set_cell_background(data_cells[index], self.CELL_BG_COLOR_WHITE)
            self.set_cell_background(row_cells[0], self.CELL_BG_COLOR_HEADER)
            self.set_cell_margins(row_cells[0], .04, .08, .04, .08)
            self.set_cell_width(row_cells[0], .75)
            self.set_cell_background(row_cells[1], self.CELL_BG_COLOR_HEADER)
            self.set_cell_margins(row_cells[1], .04, .08, .04, .08)
            self.set_cell_background(row_cells[0], self.CELL_BG_COLOR_HEADER)
            self.set_cell_margins(row_cells[0], .04, .08, .04, .08)
            return table

        question_table = self.document.add_table(
            rows=0, cols=2, style=self.TABLE_STYLE)
        self.set_table_border(
            question_table, border_color=self.BORDER_COLOR_GREY, border_size=self.BORDER_SIZE_DEFAULT)

        # Question header row
        row_cells = question_table.add_row().cells
        row_cells[0].merge(row_cells[1])
        row_cells[0].text = self.clean_text(
            quest["elementNumber"]) + " " + self.clean_text(quest["questionText"])
        self.set_cell_background(row_cells[0], self.CELL_BG_COLOR_HEADER)
        self.set_cell_margins(row_cells[0], .04, .08, .04, .08)

        # Handle followup
        if quest["triggerValue"]:
            followup_text = self.clean_text(
                quest["triggerValue"]) + " on " + self.clean_text(quest["parentElementNumber"])
            question_table = add_question_row(
                question_table, "Follow-up", followup_text)

        # Add answer
        answer_value = self.handle_answer_type(
                quest["answerValue"], quest["answerType"])
        if quest["answerType"] != "Table" and quest["answerType"] != "JSON":
            question_table = add_question_row(question_table, "Answer", answer_value)
        else: 
            question_table = add_question_table_row(question_table, "Answer", answer_value)

        # Comments and documents
        question_table = add_question_row(
            question_table, "Comment", self.clean_text(quest["answerComments"]))
        documents = [doc for docList in quest["docListAnswer"]
                     for doc in docList["attachedFiles"]]
        documentAsList = ", ".join(documents)
        question_table = add_question_row(question_table, "Documents", documentAsList)

    def get_nested_value(self, data, keys, default=None):
        for key in keys:
            if isinstance(data, dict):
                data = data.get(key, default)
            else:
                return default
            if data is default:
                break
        return data

    def add_meta_data_row(self, table, name, value):
        if not value:
            return table
        row_cells = table.add_row().cells
        row_cells[0].text = name
        self.set_cell_background(row_cells[0], self.CELL_BG_COLOR_HEADER)
        self.set_cell_margins(row_cells[0])
        row_cells[1].text = value
        self.set_cell_background(row_cells[1], self.CELL_BG_COLOR_BODY)
        self.set_cell_margins(row_cells[1])
        return table
    
    def add_assessment_metadata(self, data: Dict[str, Any]):
        """Add the section details to the document."""
        table = self.document.add_table(rows=0, cols=2, style=self.TABLE_STYLE)
        self.set_table_border(table, border_color=self.BORDER_COLOR_GREY, border_size=self.BORDER_SIZE_DEFAULT)

        assessment_details = [
            ("Assessment", ["pqiBasicInfo", "questionnaireBasicInfo", "qStructureBasicInfo", "name"]),
            ("Report Generated", ["reportGeneratedOn"]),
            ("Publisher", ["pqiBasicInfo", "publishedByUser", "companyName"]),
            ("Published By", ["pqiBasicInfo", "publishedByUser", "userProfile", "fullName"]),
            ("Published On", ["pqiBasicInfo", "publishedDate"]),
            ("Recipient", ["partner", "name"]),
            ("Received By", ["pqiBasicInfo", "publishedToContact", "contactProfile", "fullName"])
        ]

        for detail_name, detail_keys in assessment_details:
            detail_value = self.get_nested_value(data, detail_keys, "")
            self.add_meta_data_row(table, detail_name, detail_value)
    
    def generate_document(self):

        try:
            with open(self.data_file, "r") as f:
                data = json.load(f)
        except FileNotFoundError:
            logging.error("Data file not found.")
            return
        except json.JSONDecodeError:
            logging.error("Error decoding JSON from the data file.")
            return

        self.document.add_heading('Assessment Export', level=1)
        self.document.add_heading('Details', level=2)
        table = self.document.add_table(rows=0, cols=2, style=self.TABLE_STYLE)
        self.set_table_border(table, border_color=self.BORDER_COLOR_GREY, border_size=self.BORDER_SIZE_DEFAULT)
        
        ## Add assessment metadata
        self.add_assessment_metadata(data)
        self.document.add_page_break()

        ## Add questions
        for section in data["sections"]:
            section_heading = section["elementNumber"] + \
                " " + section["sectionName"]
            self.document.add_heading(section_heading, level=2)
            for quest in section["questionDetails"]:
                self.build_question(quest)
                self.document.add_paragraph()
            for subSection in section["subSections"]:
                sub_section_heading = subSection["elementNumber"] + \
                " " + subSection["sectionName"]
                self.document.add_heading(sub_section_heading, level=2)
                for quest in subSection["questionDetails"]:
                    self.build_question(quest)
                    self.document.add_paragraph()


        self.document.save('demo.docx')
        logging.info("Document generated successfully.")


if __name__ == "__main__":
    doc_gen = DocumentGenerator("mediumdata.json")
    try:
        doc_gen.generate_document()
    except Exception as e:
        logging.error(f"Error generating document: {e}")
