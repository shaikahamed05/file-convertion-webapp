from .html_to_pdf import html_to_pdf_converter
import sys
import os

# Ensure we can import from the modules folder
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '../..')))

from modules.pdf.pdf_to_word import convert_pdf_to_word

def convert_html_to_word(input_file,output_file):
    # Step 1: HTML → PDF
    html_to_pdf_converter(input_file, 'output.pdf')

    # Step 2: PDF → Word
    convert_pdf_to_word('output.pdf', output_file)

    # Step 3: Delete the intermediate PDF
    try:
        os.remove('output.pdf')
        print("🗑️ Removed temporary output.pdf file")
    except FileNotFoundError:
        print("⚠ output.pdf was already removed or not found")
    except Exception as e:
        print(f"❌ Could not delete output.pdf: {e}")
