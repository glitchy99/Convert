import os
import re
import pdfplumber
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font
from flask import Flask, request, render_template, send_file, jsonify
from werkzeug.utils import secure_filename
from functools import lru_cache
import io
import gc

app = Flask(__name__)

# PythonAnywhere specific configuration
app.config['UPLOAD_FOLDER'] = '/home/Glitchy99/mysite/uploads'  # Adjust username if different
app.config['MAX_CONTENT_LENGTH'] = 5 * 1024 * 1024  # 5MB max file size
app.config['SECRET_KEY'] = 'your-secret-key-here'

# Ensure upload folder exists
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# Cache regex patterns
ACHAT_REMISE_PATTERN = re.compile(r'ACHAT REMISE TPE N°\s*:\s*(\d+)')
DATE_PATTERN = re.compile(r'DU\s*:\s*(\d{2}/\d{2}/\d{2})')
VALUE_PATTERN = re.compile(r'([\d,]+\.\d{2})\s*$')

@lru_cache(maxsize=32)
def extract_value(line):
    """Cache frequently used value extractions"""
    match = VALUE_PATTERN.search(line.strip())
    return match.group(1) if match else None

def process_block(block):
    facture_match = ACHAT_REMISE_PATTERN.search(block)
    date_match = DATE_PATTERN.search(block)
    
    if not (facture_match and date_match):
        return None
    
    lines = block.split('\n')
    remise = commissions = tva = solde = None
    
    for i, line in enumerate(lines):
        if "TOTAL REMISE (DH)" in line:
            value = extract_value(line)
            if not value and i+1 < len(lines):
                value = extract_value(lines[i+1])
            remise = value
        elif "TOTAL COMMISSIONS HT" in line:
            value = extract_value(line)
            if not value and i+1 < len(lines):
                value = extract_value(lines[i+1])
            commissions = value
        elif "TOTAL TVA SUR COMMISSIONS" in line:
            value = extract_value(line)
            if not value and i+1 < len(lines):
                value = extract_value(lines[i+1])
            tva = value
        elif "SOLDE NET REMISE" in line:
            value = extract_value(line)
            if not value and i+1 < len(lines):
                value = extract_value(lines[i+1])
            solde = value
    
    if all([remise, commissions, tva, solde]):
        try:
            return {
                'TPE': facture_match.group(1),
                'Date': date_match.group(1),
                'Total Remise (DH)': float(remise.replace(',', '')),
                'Total Commissions HT': float(commissions.replace(',', '')),
                'Total TVA': float(tva.replace(',', '')),
                'Solde Net Remise': float(solde.replace(',', ''))
            }
        except ValueError:
            return None
    return None

def process_pdf(pdf_path):
    transactions = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if not text:
                    continue
                
                current_block = ""
                for line in text.split('\n'):
                    line = line.strip()
                    
                    if line.startswith("ACHAT REMISE"):
                        if current_block:
                            result = process_block(current_block)
                            if result:
                                transactions.append(result)
                        current_block = line + "\n"
                    elif current_block:
                        current_block += line + "\n"
                
                if current_block:
                    result = process_block(current_block)
                    if result:
                        transactions.append(result)
                
                # Force garbage collection after each page
                gc.collect()
        
        return transactions
    except Exception as e:
        raise Exception(f"Error processing PDF: {str(e)}")

def create_excel(transactions):
    wb = Workbook()
    ws = wb.active
    ws.title = "Transactions"
    
    headers = ['N° TPE', 'Date', 'TOTAL REMISE (DH)', 
              'TOTAL COMMISSIONS HT', 'TOTAL TVA SUR COMMISSIONS', 'SOLDE NET REMISE']
    ws.append(headers)
    
    for cell in ws[1]:
        cell.font = Font(bold=True)
    
    for transaction in transactions:
        ws.append([
            transaction['TPE'],
            transaction['Date'],
            transaction['Total Remise (DH)'],
            transaction['Total Commissions HT'],
            transaction['Total TVA'],
            transaction['Solde Net Remise']
        ])
    
    # Save to memory instead of disk
    excel_buffer = io.BytesIO()
    wb.save(excel_buffer)
    excel_buffer.seek(0)
    return excel_buffer

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    
    if not file.filename.endswith('.pdf'):
        return jsonify({'error': 'File must be a PDF'}), 400
    
    try:
        # Save PDF to memory instead of disk
        pdf_buffer = io.BytesIO()
        file.save(pdf_buffer)
        pdf_buffer.seek(0)
        
        transactions = process_pdf(pdf_buffer)
        if not transactions:
            return jsonify({'error': 'No transactions found in the PDF'}), 400
        
        # Generate Excel file in memory
        now = datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_filename = f"transaction_{now}.xlsx"
        excel_buffer = create_excel(transactions)
        
        # Store the buffer in the session for download
        app.config[excel_filename] = excel_buffer
        
        return jsonify({
            'success': True,
            'message': f'Successfully processed {len(transactions)} transactions',
            'filename': excel_filename
        })
        
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/download/<filename>')
def download_file(filename):
    try:
        if filename not in app.config:
            return jsonify({'error': 'File not found'}), 404
        
        excel_buffer = app.config[filename]
        excel_buffer.seek(0)
        
        def generate():
            yield from excel_buffer
            # Clean up after sending
            del app.config[filename]
            excel_buffer.close()
        
        return app.response_class(
            generate(),
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            headers={'Content-Disposition': f'attachment; filename={filename}'}
        )
        
    except Exception as e:
        return jsonify({'error': str(e)}), 404

# Add a cleanup route that can be called periodically
@app.route('/cleanup')
def cleanup():
    try:
        # Clean up any files older than 1 hour in the uploads folder
        current_time = datetime.now()
        files_cleaned = 0
        
        for filename in os.listdir(app.config['UPLOAD_FOLDER']):
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            if os.path.isfile(file_path):
                # Check if file is older than 1 hour
                file_time = datetime.fromtimestamp(os.path.getctime(file_path))
                if (current_time - file_time).seconds > 3600:
                    os.remove(file_path)
                    files_cleaned += 1
                    
        return jsonify({'message': f'Cleaned up {files_cleaned} old files'}), 200
    except Exception as e:
        return jsonify({'error': str(e)}), 500 