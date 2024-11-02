from docx2pdf import convert
import os


# Constants:
ERROR_LOG = '\n~~~~~~~~~~~~~~~~~~~~~~~~~~\nError with the given path!\n~~~~~~~~~~~~~~~~~~~~~~~~~~\n'
OUTPUT_LOG = 'Enter the path of the input DOCX file: '
INPUT_LOG = 'Enter the path of the output PDF file: '
COLOR_GREEN = 'color a'
PAUSE = 'pause'


# Function to convert DOCX to PDF
def docx_to_pdf(input_docx_file, output_pdf_file):
    convert(input_docx_file, output_pdf_file)


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    try:
        os.system(COLOR_GREEN)
        input_docx_file = (input(OUTPUT_LOG)).strip('"')
        output_pdf_file = (input(INPUT_LOG)).strip('"')

        docx_to_pdf(input_docx_file, output_pdf_file)

    except:
        print(ERROR_LOG)
    finally:
        os.system(PAUSE)
