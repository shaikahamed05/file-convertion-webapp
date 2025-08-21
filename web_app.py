from flask import render_template, redirect, url_for, session
import mysql.connector
from mysql.connector import Error
from flask import Flask, request, send_file
import tempfile
import os

from modules.excel.excel_to_pdf import excel_to_pdf_converter
from modules.excel.excel_to_html import excel_to_html_enhanced
from modules.excel.excel_to_word import excel_to_word_converter

# from modules.html.html_to_xml import html_to_xml_converter
from modules.html.html_to_excel import convert_html_to_excel
from modules.html.html_to_word import convert_html_to_word
from modules.html.html_to_pdf import html_to_pdf_converter

from modules.pdf.pdf_to_html import convert_pdf_to_html
from modules.pdf.pdf_to_word import convert_pdf_to_word
from modules.pdf.pdf_to_excel import convert_pdf_to_xlsx
# from modules.pdf.pdf_to_xml import convert_pdf_to_xml

from modules.word.word_to_pdf import convert_word_to_pdf
from modules.word.word_to_excel import convert_word_to_excel
# from modules.word.word_to_xml import convert_word_to_xml
from modules.word.word_to_html import convert_word_to_html


app = Flask(__name__)
app.secret_key = 'docscraper'

# MySQL connection config
db_config = {
    'host': 'localhost',
    'user': 'root',
    'password': '',
    'database': 'docscraper'
}

def get_db_connection():
    try:
        conn = mysql.connector.connect(**db_config)
        return conn
    except Error as e:
        print(f"Error connecting to MySQL: {e}")
        raise



@app.route('/')
def index():
    # if 'username' in session:
    if True:
        return render_template("converter.html")
    return render_template('login.html')

# 🔐 Login Handler
@app.route('/login', methods=['POST'])
def login():
    username = request.form.get('username')
    password = request.form.get('password')

    try:
        conn = get_db_connection()
        cursor = conn.cursor(dictionary=True)
        cursor.execute("SELECT * FROM users WHERE username=%s AND password=%s", (username, password))
        user = cursor.fetchone()
        cursor.close()
        conn.close()

        if user:
            session['username'] = user['username']
            return redirect(url_for('index'))  # ✅ redirect instead of render
        else:
            return render_template('login.html', error='Invalid username or password')

    except Exception as e:
        print(f"DB error: {e}")
        return render_template('login.html', error="Database connection error")

@app.route('/logout')
def logout():
    session.pop('username', None)
    return redirect(url_for('index'))

# html routes
@app.route('/html_to_excel')
def html_to_excel():
    if 'username' in session:
        return render_template("converter.html",converter_type=" HTML to Excel", from_file=".html", file_type="HTML", to_file="Excel",to_converted_file=".xlsx",convertion_root="html_to_excel_convert")
    return render_template("login.html")

@app.route('/html_to_pdf')
def html_to_pdf():
    if 'username' in session:
        return render_template("converter.html",converter_type=" HTML to PDF", from_file=".html", file_type="HTML", to_file="PDF",to_converted_file=".pdf",convertion_root="html_to_pdf_convert")
    return render_template("login.html")

@app.route('/html_to_word')
def html_to_word():
    if 'username' in session:
        return render_template("converter.html",converter_type=" HTML to Word", from_file=".html", file_type="HTML", to_file="Word",to_converted_file=".docx",convertion_root="html_to_word_convert")
    return render_template("login.html")

# excel routes
@app.route('/excel_to_word')
def excel_to_word():
    if 'username' in session:
        return render_template("converter.html",converter_type=" Excel to Word", from_file=".xlsx", file_type="Excel", to_file="Word",to_converted_file=".docx",convertion_root="excel_to_word_convert")
    return render_template("login.html")

@app.route('/excel_to_pdf')
def excel_to_pdf():
    if 'username' in session:
        return render_template("converter.html",converter_type="Excel to PDF", from_file=".xlsx", file_type="Excel", to_file="PDF", to_converted_file=".pdf",convertion_root="excel_to_pdf_convert")
    return render_template("login.html")

@app.route('/excel_to_html')
def excel_to_html():
    if 'username' in session:
        return render_template("converter.html",converter_type=" Excel to HTML", from_file=".xlsx", file_type="Excel", to_file="HTML",to_converted_file=".html",convertion_root="excel_to_html_convert")
    return render_template("login.html")

# word routes
@app.route('/word_to_excel')
def word_to_excel():
    if 'username' in session:
        return render_template("converter.html",converter_type=" Word to Excel", from_file=".docx", file_type="Word", to_file="Excel",to_converted_file=".xlsx",convertion_root="word_to_excel_convert")
    return render_template("login.html")

@app.route('/word_to_pdf')
def word_to_pdf():
    if 'username' in session:
        return render_template("converter.html",converter_type=" Word to PDF", from_file=".docx", file_type="Word", to_file="PDF",to_converted_file=".pdf",convertion_root="word_to_pdf_convert")
    return render_template("login.html")

@app.route('/word_to_html')
def word_to_html():
    if 'username' in session:
        return render_template("converter.html",converter_type=" Word to HTML", from_file=".docx", file_type="Word", to_file="HTML",to_converted_file=".html",convertion_root="word_to_html_convert")
    return render_template("login.html")

# pdf routes
@app.route('/pdf_to_html')
def pdf_to_html():
    if 'username' in session:
        return render_template("converter.html",converter_type=" PDF to HTML", from_file=".pdf", file_type="PDF", to_file="HTML",to_converted_file=".html",convertion_root="pdf_to_html_convert")
    return render_template("login.html")

@app.route('/pdf_to_excel')
def pdf_to_excel():
    if 'username' in session:
        return render_template("converter.html",converter_type=" PDF to Excel", from_file=".pdf", file_type="PDF", to_file="Excel",to_converted_file=".xlsx",convertion_root="pdf_to_excel_convert")
    return render_template("login.html")

@app.route('/pdf_to_word')
def pdf_to_word():
    if 'username' in session:
        return render_template("converter.html",converter_type=" PDF to Word", from_file=".pdf", file_type="PDF", to_file="Word",to_converted_file=".docx",convertion_root="pdf_to_word_convert")
    return render_template("login.html")


# excel convertion routes
@app.route('/excel_to_word_convert', methods=['POST'])
def excel_to_word_convert():
    uploaded_file = request.files.get('file')
    if not uploaded_file:
        return "No file uploaded", 400
    original_filename = os.path.splitext(uploaded_file.filename)[0]

    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_xlsx:
        uploaded_file.save(temp_xlsx.name)
        temp_xlsx_path = temp_xlsx.name

    # Create temp PDF first
    temp_pdf_path = temp_xlsx_path.replace('.xlsx', '.pdf')
    excel_to_pdf_converter(temp_xlsx_path, temp_pdf_path)

    # Convert PDF to Word
    temp_docx_path = temp_xlsx_path.replace('.xlsx', '.docx')
    convert_pdf_to_word(temp_pdf_path, temp_docx_path)
    
    # Serve the DOCX
    return send_file(temp_docx_path, as_attachment=True, download_name=f'{original_filename}.docx')

@app.route('/excel_to_pdf_convert', methods=['POST'])
def excel_to_pdf_convert():
    uploaded_file = request.files.get('file')
    if not uploaded_file:
        return "No file uploaded", 400

    # Save uploaded file temporarily
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_excel:
        uploaded_file.save(temp_excel.name)
        temp_excel_path = temp_excel.name

    # Output PDF path
    temp_pdf_path = temp_excel_path.replace('.xlsx', '.pdf')

    # Convert Excel to PDF
    excel_to_pdf_converter(temp_excel_path, temp_pdf_path)
    
    # Serve the PDF
    return send_file(temp_pdf_path, as_attachment=True, download_name=f'{original_filename}.pdf')

@app.route('/excel_to_html_convert', methods=['POST'])
def excel_to_html_convert():
    uploaded_file = request.files.get('file')
    if not uploaded_file:
        return "No file uploaded", 400
    original_filename = os.path.splitext(uploaded_file.filename)[0]

    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_xlsx:
        uploaded_file.save(temp_xlsx.name)
        temp_xlsx_path = temp_xlsx.name

    # Create temp PDF first
    temp_pdf_path = temp_xlsx_path.replace('.xlsx', '.html')
    excel_to_html_enhanced(temp_xlsx_path, temp_pdf_path)
    
    # Serve the DOCX
    return send_file(temp_pdf_path, as_attachment=True, download_name=f'{original_filename}.html')


# html convertion routes
@app.route('/html_to_excel_convert', methods=['POST'])
def html_to_excel_convert():
    uploaded_file = request.files.get('file')
    if not uploaded_file:
        return "No file uploaded", 400
    original_filename = os.path.splitext(uploaded_file.filename)[0]
    # Save uploaded file temporarily
    with tempfile.NamedTemporaryFile(delete=False, suffix='.html') as temp_html:
        uploaded_file.save(temp_html.name)
        temp_html_path = temp_html.name

    # Output PDF path
    temp_excel_path = temp_html_path.replace('.html', '.xlsx')

    # Convert Excel to PDF
    convert_html_to_excel(temp_html_path, temp_excel_path,columns_to_merge=None)
    
    # Serve the PDF
    return send_file(temp_excel_path, as_attachment=True, download_name=f'{original_filename}.xlsx')

@app.route('/html_to_pdf_convert', methods=['POST'])
def html_to_pdf_convert():
    uploaded_file = request.files.get('file')
    if not uploaded_file:
        return "No file uploaded", 400
    original_filename = os.path.splitext(uploaded_file.filename)[0]

    # Save uploaded file temporarily
    with tempfile.NamedTemporaryFile(delete=False, suffix='.html') as temp_html:
        uploaded_file.save(temp_html.name)
        temp_html_path = temp_html.name

    # Output PDF path
    temp_excel_path = temp_html_path.replace('.html', '.pdf')

    # Convert Excel to PDF
    html_to_pdf_converter(temp_html_path, temp_excel_path)
    
    # Serve the PDF
    return send_file(temp_excel_path, as_attachment=True, download_name=f'{original_filename}.pdf')


# pdf convertion routes
@app.route('/pdf_to_word_convert', methods=['POST'])
def pdf_to_word_convert():
    uploaded_file = request.files.get('file')
    if not uploaded_file:
        return "No file uploaded", 400
    original_filename = os.path.splitext(uploaded_file.filename)[0]

    # Save uploaded file temporarily
    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as temp_html:
        uploaded_file.save(temp_html.name)
        temp_html_path = temp_html.name
    original_filename = os.path.splitext(uploaded_file.filename)[0]

    # Output PDF path
    temp_excel_path = temp_html_path.replace('.pdf', '.docx')

    # Convert Excel to PDF
    convert_pdf_to_word(temp_html_path, temp_excel_path)
    
    # Serve the PDF
    return send_file(temp_excel_path, as_attachment=True, download_name=f'{original_filename}.docx')

@app.route('/pdf_to_excel_convert', methods=['POST'])
def pdf_to_excel_convert():
    uploaded_file = request.files.get('file')
    if not uploaded_file:
        return "No file uploaded", 400
    original_filename = os.path.splitext(uploaded_file.filename)[0]

    # Save uploaded PDF temporarily
    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as temp_pdf:
        uploaded_file.save(temp_pdf.name)
        temp_pdf_path = temp_pdf.name

    # Convert PDF → Excel (watermark removed)
    cleaned_excel_path = convert_pdf_to_xlsx(temp_pdf_path)
    
    # Serve the Excel file
    return send_file(cleaned_excel_path, as_attachment=True, download_name=f'{original_filename}.xlsx')

@app.route('/pdf_to_html_convert', methods=['POST'])
def pdf_to_html_convert():
    uploaded_file = request.files.get('file')
    if not uploaded_file:
        return "No file uploaded", 400
    original_filename = os.path.splitext(uploaded_file.filename)[0]

    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as temp_xlsx:
        uploaded_file.save(temp_xlsx.name)
        temp_xlsx_path = temp_xlsx.name

    # Create temp PDF first
    temp_pdf_path = temp_xlsx_path.replace('.pdf', '.html')
    convert_pdf_to_html(temp_xlsx_path, temp_pdf_path)
    
    # Serve the DOCX
    return send_file(temp_pdf_path, as_attachment=True, download_name=f'{original_filename}.html')


#word convertion routes
@app.route('/word_to_pdf_convert', methods=['POST'])
def word_to_pdf_convert():
    uploaded_file = request.files.get('file')
    if not uploaded_file:
        return "No file uploaded", 400
    original_filename = os.path.splitext(uploaded_file.filename)[0]

    # Save uploaded file temporarily
    with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as temp_html:
        uploaded_file.save(temp_html.name)
        temp_html_path = temp_html.name

    # Output PDF path
    temp_excel_path = temp_html_path.replace('.docx', '.pdf')

    # Convert Excel to PDF
    convert_word_to_pdf(temp_html_path, temp_excel_path)
    
    # Serve the PDF
    return send_file(temp_excel_path, as_attachment=True, download_name=f'{original_filename}.pdf')

@app.route('/word_to_excel_convert', methods=['POST'])
def word_to_excel_convert():
    uploaded_file = request.files.get('file')
    if not uploaded_file:
        return "No file uploaded", 400
    original_filename = os.path.splitext(uploaded_file.filename)[0]

    with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as temp_xlsx:
        uploaded_file.save(temp_xlsx.name)
        temp_xlsx_path = temp_xlsx.name

    # Create temp PDF first
    temp_pdf_path = temp_xlsx_path.replace('.docx', '.xlsx')
    convert_word_to_excel(temp_xlsx_path, temp_pdf_path)
    
    # Serve the DOCX
    return send_file(temp_pdf_path, as_attachment=True, download_name=f'{original_filename}.docx')

@app.route('/word_to_html_convert', methods=['POST'])
def word_to_html_convert():
    uploaded_file = request.files.get('file')
    if not uploaded_file:
        return "No file uploaded", 400
    original_filename = os.path.splitext(uploaded_file.filename)[0]

    with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as temp_pdf:
        uploaded_file.save(temp_pdf.name)
        temp_pdf_path = temp_pdf.name

    temp_xml_path = temp_pdf_path.replace('.docx', '.html')
    convert_word_to_html(temp_pdf_path, temp_xml_path)
    
    return send_file(temp_xml_path, as_attachment=True, download_name=f'{original_filename}.html')


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0',port=5000)
