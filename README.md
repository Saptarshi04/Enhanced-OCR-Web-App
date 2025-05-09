# 🧠 Enhanced OCR Web App

A powerful web application for converting scanned documents (PDFs and images) into **searchable PDFs** or **Word documents**, with advanced features like:

- 🧾 Enhanced OCR using Tesseract and OCRmyPDF
- 📊 Smart table detection with PyMuPDF, Camelot, and Tabula
- 📝 Custom Word output with styled tables and headers
- 🌐 Flask-powered asynchronous processing and API endpoints

---

## 🚀 Features

- Upload image (`.jpg`, `.jpeg`, `.png`, `.tif`, `.tiff`) or PDF files
- Choose output format: `searchable PDF` or `DOCX`
- Language selection for OCR
- Options for deskewing, image cleaning, and table detection
- Real-time status updates via background job threads
- Clean and organized output with styled DOCX tables

---

## 🛠 Setup & Installation

### 1. Clone the repository

```bash
git clone https://github.com/Saptarshi04/Enhanced-OCR-Web-App.git
cd Enhanced-OCR-Web-App

python -m venv venv
source venv/bin/activate  # or venv\Scripts\activate on Windows
pip install -r requirements.txt

python app.py
Then go to http://127.0.0.1:5000 in your browser.
