# Local-pdf-pro
A professional, offline desktop application for PDF management. Built with Python and PyQt6, this tool processes all data locally on your machine, ensuring maximum privacy and security for your documents. It mimics the functionality of premium online PDF tools but without file upload limits or data privacy concerns.

## Project Screenshot
<img width="1919" height="1002" alt="image" src="https://github.com/user-attachments/assets/d0d64012-1cb7-4e57-9936-e783b6308117" />

## Features

### Dashboard & Navigation
- **Grid Dashboard:** Quick access to all 16 tools upon launch.
- **Sidebar Navigation:** Persistent access to tools with drag-and-drop support.
- **Theme Toggle:** Switch between Dark Mode and Light Mode.

### Core PDF Tools
- **Merge PDF:** Combine multiple PDF files into a single document.
- **Split PDF:** Advanced splitting options:
    - Split all pages into separate files.
    - Split by specific page ranges (e.g., 1-5, 8).
    - Extract selected ranges into a new single PDF.
- **Visual Organiser:** View page thumbnails to reorder, rotate (90 degrees left/right), or delete specific pages.
- **Compress PDF:** Reduce file size with three levels of compression:
    - Low (Lossless)
    - Medium (Optimized)
    - Extreme (Rasterize text to images for maximum reduction)

### Conversion Tools
**To PDF:**
- **Images to PDF:** Convert JPG/PNG files to PDF. Includes a "Smart Scan" feature with automatic edge detection and a manual corner adjustment tool for perspective correction.
- **Word to PDF:** Convert .docx files to PDF (Requires Microsoft Word).
- **PowerPoint to PDF:** Convert .pptx slides to PDF (Requires Microsoft PowerPoint).

**From PDF:**
- **PDF to Image:** Extract all pages as high-quality JPEG images.
- **PDF to Word:** Convert PDF documents to editable .docx files.
- **PDF to PowerPoint:** Convert PDF pages into PowerPoint slides.

### Security
- **Protect PDF:** Encrypt documents with a password. Choose from multiple encryption algorithms:
    - AES-256 (Strongest)
    - AES-128 (Standard)
    - RC4-128 (Legacy compatibility)
- **Unlock PDF:** Remove passwords from protected PDFs (requires knowing the original password) to save an unlocked copy.

### Pro Features
- **OCR (Optical Character Recognition):** Convert scanned PDFs or images into searchable, text-selectable PDFs using Tesseract.
- **Watermark:** Add custom text watermarks to every page with adjustable opacity and rotation.
- **Page Numbers:** Add "Page X of Y" numbering to the bottom-centre, bottom-right, or top-right of the document.
- **Edit Metadata:** View and modify PDF properties including Title, Author, Subject, Producer, and Creator.

## Installation & Setup

### 1. Install Python Dependencies
Ensure you have Python installed. Run the following command in your terminal:
```bash
pip install -r requirements.txt
```

# Install External Tools (Crucial!)
Some features rely on external tools to work correctly:
## **Poppler**(Required for Images/Thumbnails):

Download the latest binary from Poppler for Windows.

Extract the ZIP file to a folder (e.g., C:\Program Files\poppler).

Add the bin folder (e.g., C:\Program Files\poppler\Library\bin) to your System PATH environment variable.

## Tesseract OCR (Required for OCR features):

Download and install Tesseract OCR.

During installation, ensure "Add to Path" is checked.

If the application cannot find Tesseract, update the path configuration in pdf_engine.py.

## **Microsoft Office (Optional):**

Microsoft Word is required for Word -> PDF conversion.

Microsoft PowerPoint is required for PowerPoint -> PDF conversion.

## How to Run
Ensure main.py and pdf_engine.py are in the same directory.

Run the application:

Bash
```
python main.py
```

## Technical Stack
GUI: PyQt6

PDF Manipulation: pypdf, pikepdf

Image Processing: OpenCV, Pillow, pdf2image

OCR: pytesseract (Tesseract Engine)

Document Conversion: pdf2docx, docx2pdf, comtypes

PDF Generation: reportlab, img2pdf
