from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import os
import tempfile
from werkzeug.utils import secure_filename

class PdfConverter:
    @staticmethod
    def convert_pdf_table_to_excel(pdf_path, output_file):
        import pdfplumber
        import pandas as pd
        
        all_tables = []
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                table = page.extract_table()
                if table:
                    all_tables.extend(table)
        if not all_tables:
            return None, "No tables found in PDF"
        df = pd.DataFrame(all_tables[1:], columns=all_tables[0])
        df.to_excel(output_file, index=False)
        return output_file, "Excel file generated successfully"

    @staticmethod
    def extract_tables_from_pdf(pdf_path):
        import fitz
        
        pdf_document = fitz.open(pdf_path)
        tables = []
        for page_num in range(len(pdf_document)):
            page = pdf_document[page_num]
            tables_in_page = page.find_tables()
            for table in tables_in_page:
                table_data = table.extract()
                tables.append(table_data)
        pdf_document.close()
        return tables

    @staticmethod
    def export_tables_to_excel(tables, output_excel_path):
        import pandas as pd
        
        with pd.ExcelWriter(output_excel_path, engine='xlsxwriter') as writer:
            for i, table in enumerate(tables):
                df = pd.DataFrame(table[1:], columns=table[0])
                df.to_excel(writer, sheet_name=f"Table_{i + 1}", index=False)

    @staticmethod
    def convert_pdf_text_to_excel(pdf_path, output_file):
        import fitz
        import pandas as pd
        
        doc = fitz.open(pdf_path)
        extracted_text = [page.get_text("text") for page in doc]
        doc.close()
        df = pd.DataFrame({"Texte extrait": extracted_text})
        df.to_excel(output_file, index=False)
        return output_file, "Excel file generated successfully"

    @staticmethod
    def process_ocr(pdf_path, output_file):
        from pdf2image import convert_from_path
        import easyocr
        import cv2
        import numpy as np
        
        images = convert_from_path(pdf_path)
        detections = []
        reader = easyocr.Reader(['fr'])
        for i, image in enumerate(images):
            image_cv = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
            result = reader.readtext(image_cv)
            detections.append(result)
        PdfConverter.export_to_excel(detections, output_file)
        return output_file, "Excel file generated successfully"

    @staticmethod
    def export_to_excel(detections, output_file):
        import pandas as pd
        
        data = {}
        for page_num, page_detections in enumerate(detections):
            for (bbox, text, prob) in page_detections:
                x0, y0 = bbox[0]
                x1, y1 = bbox[2]
                row = int(y0 // 20)
                col = int(x0 // 20)
                if (row, col) not in data:
                    data[(row, col)] = text
                else:
                    data[(row, col)] += " " + text
        df = pd.DataFrame.from_dict(data, orient="index", columns=["Text"])
        df.index = pd.MultiIndex.from_tuples(df.index, names=["Row", "Column"])
        df = df.unstack().fillna("")
        df.to_excel(output_file)

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = tempfile.gettempdir()
ALLOWED_EXTENSIONS = {'pdf'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/api/convert', methods=['POST'])
def convert_pdf():
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400
    
    file = request.files['file']
    conversion_type = request.form.get('conversion_type', 'table')
    
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400
    
    if not allowed_file(file.filename):
        return jsonify({'error': 'Only PDF files are allowed'}), 400
    
    try:
        filename = secure_filename(file.filename)
        pdf_path = os.path.join(UPLOAD_FOLDER, filename)
        file.save(pdf_path)
        
        output_file = os.path.join(UPLOAD_FOLDER, f"output_{os.path.splitext(filename)[0]}.xlsx")
        
        if conversion_type == 'table':
            result, message = PdfConverter.convert_pdf_table_to_excel(pdf_path, output_file)
        elif conversion_type == 'sheets':
            tables = PdfConverter.extract_tables_from_pdf(pdf_path)
            PdfConverter.export_tables_to_excel(tables, output_file)
            result, message = output_file, "Excel file generated successfully"
        elif conversion_type == 'text':
            result, message = PdfConverter.convert_pdf_text_to_excel(pdf_path, output_file)
        elif conversion_type == 'ocr':
            result, message = PdfConverter.process_ocr(pdf_path, output_file)
        else:
            return jsonify({'error': 'Invalid conversion type'}), 400
        
        if result:
            return send_file(result, as_attachment=True, download_name=os.path.basename(result))
        else:
            return jsonify({'error': message}), 400
            
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        if os.path.exists(pdf_path):
            os.remove(pdf_path)

@app.route('/api/health', methods=['GET'])
def health_check():
    return jsonify({'status': 'ok'})

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
