from docx import Document
from os import path

from main_req_formatter import MainRequirementsFormatter
from source_links_formatter import SourceLinksFormatter
import argparse


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('path_to_docx', type=str, help='Path to docx file')
    args = parser.parse_args()

    if not path.exists(args.path_to_docx) or not path.isfile(args.path_to_docx):
        print('Введен неверный путь до файла')
        exit(1)

    try:
        doc = Document(args.path_to_docx)
    except ValueError:
        print('Документ должен быть типа docx')
        exit(1)

    changes =[]
    MainRequirementsFormatter.format_document(doc, changes)
    SourceLinksFormatter.extract_bibliography(doc, 'библиография.txt')
    MainRequirementsFormatter.change_title_page_year(doc, '2024', changes)
    SourceLinksFormatter.check_for_links_presence(doc, changes)

    doc.save("edited.docx")
    changes_file = open("changelog.txt", "w+")
    for item in changes:
        changes_file.write("%s\n" % item)
    changes_file.close()


if __name__ == '__main__':
    main()