import os
from xhtml2pdf import pisa
from io import BytesIO
from pathlib import Path
 
def convert_html_to_pdf(html_content, pdf_path):
    """Convert HTML content to PDF using xhtml2pdf"""
    try:
        with open(pdf_path, 'wb') as output_file:
            # Create PDF
            result = pisa.CreatePDF(
                html_content,
                dest=output_file,
                encoding='utf-8',
                link_callback=None,
                show_error_as_pdf=True
            )
           
        if not result.err:
            return True, "Conversion successful"
        else:
            return False, f"Error during PDF generation: {result.err}"
           
    except Exception as e:
        return False, f"Error in PDF conversion: {str(e)}"
 
def html_to_pdf_converter(html_path, pdf_path):
    try:
        print(f"Starting conversion from {html_path} to {pdf_path}")
       
        # Validate input file
        if not os.path.exists(html_path):
            error_msg = f"HTML file not found: {html_path}"
            print(error_msg)
            return False, error_msg
           
        if not html_path.lower().endswith(('.html', '.htm')):
            error_msg = "Input file must be an HTML file (.html or .htm)"
            print(error_msg)
            return False, error_msg
           
        if not pdf_path.lower().endswith('.pdf'):
            error_msg = "Output file must be a PDF file (.pdf)"
            print(error_msg)
            return False, error_msg
       
        # Read HTML content
        try:
            with open(html_path, 'r', encoding='utf-8') as f:
                html_content = f.read()
                if not html_content.strip():
                    error_msg = "HTML file is empty"
                    print(error_msg)
                    return False, error_msg
        except Exception as e:
            error_msg = f"Error reading HTML file: {str(e)}"
            print(error_msg)
            return False, error_msg
           
        # Convert HTML to PDF
        success, message = convert_html_to_pdf(html_content, pdf_path)
       
        if success:
            # Verify the PDF was created
            if not os.path.exists(pdf_path):
                error_msg = f"PDF file was not created at {pdf_path}"
                print(error_msg)
                return False, error_msg
               
            if os.path.getsize(pdf_path) == 0:
                error_msg = "PDF file was created but is empty"
                print(error_msg)
                return False, error_msg
               
            print("Conversion successful!")
            return True, "Conversion successful"
        else:
            print(f"Conversion failed: {message}")
            return False, message
           
    except Exception as e:
        import traceback
        error_msg = f"Unexpected error: {str(e)}\n\n{traceback.format_exc()}"
        print(error_msg)
        return False, error_msg