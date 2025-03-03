from flask import Flask, request, jsonify, send_from_directory, render_template
from pdf2docx import Converter
from PyPDF2 import PdfReader
from openpyxl import Workbook
from pdf2image import convert_from_path
import os

app = Flask(__name__)

# Directories
UPLOAD_FOLDER = 'static/uploads'
DOWNLOAD_FOLDER = 'static/downloads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['DOWNLOAD_FOLDER'] = DOWNLOAD_FOLDER

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert_pdf():
    if 'pdf' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400

    file = request.files['pdf']
    format_type = request.form.get('format', 'word')
    filename = file.filename
    pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(pdf_path)

    output_filename = filename.rsplit('.', 1)[0]
    output_path = os.path.join(app.config['DOWNLOAD_FOLDER'], f"{output_filename}.{format_type}")

    try:
        if format_type == 'word':
            cv = Converter(pdf_path)
            cv.convert(output_path)
            cv.close()
            output_extension = 'docx'

        elif format_type == 'excel':
            pdf = PdfReader(pdf_path)
            wb = Workbook()
            ws = wb.active
            for i, page in enumerate(pdf.pages):
                text = page.extract_text()
                ws[f'A{i+1}'] = text
            output_path = output_path.replace('.excel', '.xlsx')
            wb.save(output_path)
            output_extension = 'xlsx'

        elif format_type == 'jpg':
            images = convert_from_path(pdf_path)
            output_path = output_path.replace('.jpg', '_page1.jpg')
            images[0].save(output_path, 'JPEG')
            output_extension = 'jpg'

        else:
            return jsonify({'error': 'Invalid format'}), 400

        download_url = f"/downloads/{os.path.basename(output_path)}"
        return jsonify({'download_url': download_url})

    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/downloads/<filename>')
def download_file(filename):
    return send_from_directory(app.config['DOWNLOAD_FOLDER'], filename)

if __name__ == '__main__':
    app.run(debug=True)