import re

from docx import Document


class SourceLinksFormatter:

    @staticmethod
    def get_number_of_ref_subscripts(doc: Document) -> int:
        counter = 0
        for p in doc.paragraphs:
            for r in p.runs:
                if r.font.superscript and re.match(r"\[[0-9]+]", r.text):
                    counter += 1
        return counter

    @staticmethod
    def get_numbered_list_id(paragraph):
        return paragraph._element.xpath('./w:pPr/w:numPr/w:numId/@w:val')

    @staticmethod
    def get_number_of_source_links(doc: Document, section_header: str) -> int:
        section_found = False
        first_list_id = None
        counter = 0
        for p in doc.paragraphs:
            if p.text.casefold() == section_header.casefold():
                section_found = True
            if section_found:
                list_id = SourceLinksFormatter.get_numbered_list_id(p)
                if list_id and not first_list_id:
                    first_list_id = list_id
                if first_list_id and list_id == first_list_id:
                    counter += 1

        return counter

    @staticmethod
    def check_for_links_presence(doc: Document):
        ref_subscripts_number = SourceLinksFormatter.get_number_of_ref_subscripts(doc)
        source_links_number = SourceLinksFormatter.get_number_of_source_links(doc, 'список источников и литературы')
        if ref_subscripts_number != source_links_number:
            print('Количество выделенных ссылок не соответствует количеству источников литературы.'
                  ' Возможно Ваш список источников не является цельным списком.')