import pythoncom
from docx2pdf import convert

def convert_word_to_pdf(input_file, output_file):
    pythoncom.CoInitialize()   # Initialize COM for this thread
    try:
        convert(input_file, output_file)
    finally:
        pythoncom.CoUninitialize()  # Always clean up
