# PDF to Excel Converter

A web application that converts PDF transaction files to Excel format. Built with Flask and modern web technologies.

## Features

- Drag and drop PDF file upload
- Modern, responsive user interface
- Automatic transaction extraction
- Excel file generation with formatted data
- Progress tracking
- Error handling

## Requirements

- Python 3.7 or higher
- pip (Python package installer)

## Installation

1. Clone this repository or download the source code
2. Create a virtual environment (recommended):
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```
3. Install the required packages:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. Start the Flask application:
   ```bash
   python web_app.py
   ```
2. Open your web browser and navigate to `http://localhost:5000`
3. Drag and drop your PDF file or click "Choose File" to select it
4. Wait for the processing to complete
5. Click the "Download Excel File" button to save the converted file

## File Format

The application expects PDF files containing transaction data in the following format:
- Each transaction should start with "ACHAT REMISE"
- The file should contain transaction details including:
  - TPE number
  - Date
  - Total Remise (DH)
  - Total Commissions HT
  - Total TVA
  - Solde Net Remise

## Security

- Maximum file size: 16MB
- Only PDF files are accepted
- Files are processed securely and cleaned up after conversion
- Temporary files are stored in the `uploads` directory

## License

This project is licensed under the MIT License - see the LICENSE file for details. 