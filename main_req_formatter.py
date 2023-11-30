import re

from docx import Document
from docx.oxml import OxmlElement, ns
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.text import WD_LINE_SPACING
from docx.shared import RGBColor
from docx.shared import Cm


class MainRequirementsFormatter:

    @staticmethod
    def number_pages(doc):
        for section in doc.sections:
            section.footer.is_linked_to_previous = True

        new_paragraph = doc.sections[0].footer.add_paragraph()
        new_run = new_paragraph.add_run()

        new_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        doc.sections[0].different_first_page_header_footer = True

        fld_char1 = OxmlElement('w:fldChar')
        fld_char1.set(ns.qn('w:fldCharType'), 'begin')

        instr_text = OxmlElement('w:instrText')
        instr_text.set(ns.qn('xml:space'), 'preserve')
        instr_text.text = "PAGE"

        fld_char2 = OxmlElement('w:fldChar')
        fld_char2.set(ns.qn('w:fldCharType'), 'end')

        new_run._r.append(fld_char1)
        new_run._r.append(instr_text)
        new_run._r.append(fld_char2)

    @staticmethod
    def change_font(font_name: str, paragraph):
        for run in paragraph.runs:
            run.font.name = font_name

    @staticmethod
    def change_font_size(font_size: int, paragraph):
        for run in paragraph.runs:
            run.font.size = Pt(font_size)


    @staticmethod
    def change_font_color(color: RGBColor, paragraph):
        for run in paragraph.runs:
            run.font.color.rgb = color

    @staticmethod
    def change_line_spacing(spacing: WD_LINE_SPACING, paragraph):
        paragraph.paragraph_format.line_spacing_rule = spacing

    @staticmethod
    def change_margins(left_cm, right_cm, top_cm, bottom_cm, doc: Document):
        for section in doc.sections:
            section.left_margin = Cm(left_cm)
            section.right_margin = Cm(right_cm)
            section.top_margin = Cm(top_cm)
            section.bottom_margin = Cm(bottom_cm)

    @staticmethod
    def change_left_paragraph_indentation(indentation_cm, paragraph):
        paragraph.paragraph_format.left_indent = Cm(indentation_cm)

    @staticmethod
    def change_title_page_year(doc: Document, year: str):
        regex = re.compile(r"20[0-9][0-9]")

        for p in doc.paragraphs:
            for r in p.runs:
                if regex.search(r.text):
                    r.text = f'\n{year}'
                    return

    @staticmethod
    def format_document(doc: Document):
        for paragraph in doc.paragraphs:
            MainRequirementsFormatter.change_font('Times New Roman', paragraph)
            MainRequirementsFormatter.change_font_size(14, paragraph)
            MainRequirementsFormatter.change_font_color(RGBColor(0, 0, 0), paragraph)
            MainRequirementsFormatter.change_line_spacing(WD_LINE_SPACING.ONE_POINT_FIVE, paragraph)
            MainRequirementsFormatter.change_left_paragraph_indentation(1.25, paragraph)
        MainRequirementsFormatter.number_pages(doc)
        MainRequirementsFormatter.change_margins(3, 1.5, 2, 2, doc)