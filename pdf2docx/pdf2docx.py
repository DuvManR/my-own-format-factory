
from pdf2docx import Converter
from docx import Document
import os


# Constants:
ERROR_LOG = '\n~~~~~~~~~~~~~~~~~~~~~~~~~~\nError with the given path!\n~~~~~~~~~~~~~~~~~~~~~~~~~~\n'
OUTPUT_LOG = 'Enter the path of the output DOCX file: '
INPUT_LOG = 'Enter the path of the input PDF file: '
COLOR_GREEN = 'color a'
PAUSE = 'pause'


# Gets a PDF file and converts it to DOCX
def pdf_to_docx(input_pdf_file, output_docx_file):
    cv = Converter(input_pdf_file)
    cv.convert(output_docx_file, start=0, end=None, keep_layout=True, parse_lattice_table=False)
    cv.close()

    # Open the converted DOCX file
    doc = Document(output_docx_file)

    # Set alignment to justify for all paragraphs
    for paragraph in doc.paragraphs:
        paragraph.alignment = 3  # 3 corresponds to "justified" alignment

    # Save the modified DOCX file
    doc.save(output_docx_file)


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    try:
        os.system(COLOR_GREEN)
        pdf_file = (input(INPUT_LOG)).strip('"')
        docx_file = (input(OUTPUT_LOG)).strip('"')

        pdf_to_docx(pdf_file, docx_file)
    except:
        print(ERROR_LOG)
    finally:
        os.system(PAUSE)
