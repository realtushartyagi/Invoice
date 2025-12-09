# Backend Setup Guide

## Installation

The backend dependencies have been successfully configured and installed.

### Install Dependencies

```bash
cd backend
python -m pip install -r requirements.txt
```

### Run the Backend Server

```bash
python backend.py
```

The server will start on `http://localhost:5000`

## Available Features

✅ **Working Features:**
- PDF Table to Excel conversion (using pdfplumber)
- PDF Multi-sheet Excel export (using PyMuPDF)
- PDF Text to Excel conversion

⚠️ **OCR Feature Note:**
The OCR functionality requires additional dependencies (pdf2image, easyocr, opencv-python) that need a C++ compiler to install on Windows with Python 3.14. If you need OCR functionality:
1. Install Visual Studio Build Tools or MinGW
2. Add the missing packages to requirements.txt:
   - pdf2image
   - easyocr
   - opencv-python

## API Endpoints

- `POST /api/convert` - Convert PDF to Excel
  - Form data: `file` (PDF file), `conversion_type` (table/sheets/text/ocr)
- `GET /api/health` - Health check endpoint

## Requirements

- Python 3.x
- Flask
- flask-cors
- pdfplumber
- pandas
- openpyxl
- xlsxwriter
- PyMuPDF
- Werkzeug
