from pdf2docx import Converter

def convert_pdf_to_word(pdf_file,docx_file):
    cv = Converter(pdf_file)
    cv.convert(docx_file)
    cv.close()


