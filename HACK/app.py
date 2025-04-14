import re
import pdfplumber
import sys
import os
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QLabel, QPushButton, QFileDialog, QMessageBox, 
                            QProgressBar, QLineEdit)
from PyQt5.QtCore import Qt, QThread, pyqtSignal

class PdfExtractorThread(QThread):
    progress_updated = pyqtSignal(int)
    extraction_complete = pyqtSignal(list)
    error_occurred = pyqtSignal(str)

    def __init__(self, pdf_path):
        super().__init__()
        self.pdf_path = pdf_path

    def run(self):
        try:
            transactions = []
            with pdfplumber.open(self.pdf_path) as pdf:
                total_pages = len(pdf.pages)
                
                for i, page in enumerate(pdf.pages):
                    self.progress_updated.emit(int((i / total_pages) * 80))
                    
                    text = page.extract_text()
                    if not text:
                        continue
                    
                    current_block = ""
                    for line in text.split('\n'):
                        line = line.strip()
                        
                        if line.startswith("ACHAT REMISE"):
                            if current_block:
                                result = self.process_block(current_block)
                                if result:
                                    transactions.append(result)
                            current_block = line + "\n"
                        elif current_block:
                            current_block += line + "\n"
                    
                    if current_block:
                        result = self.process_block(current_block)
                        if result:
                            transactions.append(result)
                
                self.progress_updated.emit(90)
                self.extraction_complete.emit(transactions)
                self.progress_updated.emit(100)
                
        except Exception as e:
            self.error_occurred.emit(f"Error: {str(e)}")

    def process_block(self, block):
        facture_match = re.search(r'ACHAT REMISE TPE N°\s*:\s*(\d+)', block)
        date_match = re.search(r'DU\s*:\s*(\d{2}/\d{2}/\d{2})', block)
        
        lines = block.split('\n')
        remise = commissions = tva = solde = None
        
        for i, line in enumerate(lines):
            if "TOTAL REMISE (DH)" in line:
                value = self.extract_value(line)
                if not value and i+1 < len(lines):
                    value = self.extract_value(lines[i+1])
                remise = value
            elif "TOTAL COMMISSIONS HT" in line:
                value = self.extract_value(line)
                if not value and i+1 < len(lines):
                    value = self.extract_value(lines[i+1])
                commissions = value
            elif "TOTAL TVA SUR COMMISSIONS" in line:
                value = self.extract_value(line)
                if not value and i+1 < len(lines):
                    value = self.extract_value(lines[i+1])
                tva = value
            elif "SOLDE NET REMISE" in line:
                value = self.extract_value(line)
                if not value and i+1 < len(lines):
                    value = self.extract_value(lines[i+1])
                solde = value
        
        if all([facture_match, date_match, remise, commissions, tva, solde]):
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

    def extract_value(self, line):
        match = re.search(r'([\d,]+\.\d{2})\s*$', line.strip())
        return match.group(1) if match else None

class PdfToExcelApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF to Excel Converter")
        self.setGeometry(100, 100, 500, 300)
        
        self.initUI()
        self.extractor_thread = None
        self.pdf_path = ""

    def initUI(self):
        main_widget = QWidget()
        layout = QVBoxLayout()
        
        # PDF Selection
        pdf_layout = QVBoxLayout()
        pdf_label = QLabel("Select PDF File:")
        self.pdf_path_edit = QLineEdit()
        self.pdf_path_edit.setReadOnly(True)
        browse_btn = QPushButton("Browse...")
        browse_btn.clicked.connect(self.browse_pdf)
        
        pdf_layout.addWidget(pdf_label)
        pdf_layout.addWidget(self.pdf_path_edit)
        pdf_layout.addWidget(browse_btn)
        
        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        
        # Convert button
        convert_btn = QPushButton("Convert to Excel")
        convert_btn.clicked.connect(self.start_conversion)
        
        # Add widgets to main layout
        layout.addLayout(pdf_layout)
        layout.addWidget(self.progress_bar)
        layout.addWidget(convert_btn)
        
        main_widget.setLayout(layout)
        self.setCentralWidget(main_widget)

    def browse_pdf(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select PDF File", "", "PDF Files (*.pdf)"
        )
        if file_path:
            self.pdf_path = file_path
            self.pdf_path_edit.setText(file_path)

    def start_conversion(self):
        if not self.pdf_path:
            QMessageBox.warning(self, "Warning", "Please select a PDF file first.")
            return
        
        self.progress_bar.setValue(0)
        self.extractor_thread = PdfExtractorThread(self.pdf_path)
        self.extractor_thread.progress_updated.connect(self.update_progress)
        self.extractor_thread.extraction_complete.connect(self.save_to_excel)
        self.extractor_thread.error_occurred.connect(self.show_error)
        self.extractor_thread.start()

    def update_progress(self, value):
        self.progress_bar.setValue(value)

    def save_to_excel(self, transactions):
        if not transactions:
            QMessageBox.warning(self, "Warning", "No transactions found in the PDF.")
            return
        
        # Generate filename with current datetime
        now = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f"transaction_{now}.xlsx"
        
        try:
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
            
            # Save to user's Downloads folder
            downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
            output_path = os.path.join(downloads_path, output_file)
            wb.save(output_path)
            
            QMessageBox.information(
                self, 
                "Success", 
                f"Successfully extracted {len(transactions)} transactions to:\n{output_path}"
            )
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save Excel file: {str(e)}")
        
        self.progress_bar.setValue(100)

    def show_error(self, message):
        QMessageBox.critical(self, "Error", message)
        self.progress_bar.setValue(0)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PdfToExcelApp()
    window.show()
    sys.exit(app.exec_())