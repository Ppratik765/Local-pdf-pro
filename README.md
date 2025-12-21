# Local-pdf-pro
A powerful, offline desktop application that mimics the functionality of iLovePDF. Built with Python and PyQt6, this tool processes everything locally on your machine, ensuring your data never leaves your computer.

## 🚀 Features

**Organize & Edit**
- **Merge PDF:** Combine multiple PDFs into a single file.
- **Split PDF:** Extract all pages from a PDF into separate files.
- **Visual Organize:** Grid view to drag-and-drop page reordering and deletion.

**Convert to PDF**
- **Images to PDF:** Convert JPG/PNG files to a single PDF.
- **Word to PDF:** Convert `.docx` to PDF (Requires MS Word).
- **PowerPoint to PDF:** Convert `.pptx` to PDF (Requires MS PowerPoint).

**Convert from PDF**
- **PDF to Image:** Extract pages as high-quality JPEGs.
- **PDF to Word:** Convert PDF to editable `.docx` files.
- **PDF to PowerPoint:** Convert PDF pages into PowerPoint slides.

**Security**
- **Protect PDF:** Encrypt your documents with a password.

## 🛠️ Installation & Setup

### 1. Install Python Dependencies
Run the following command in your terminal:
```bash
pip install -r requirements.txt
```

# Install External Tools (Crucial!)
Some features rely on external tools to work correctly:
**Poppler**(Required for Images/Thumbnails):

Download the latest binary from Poppler for Windows.

Extract the ZIP file to a folder (e.g., C:\Program Files\poppler).

Add the bin folder (e.g., C:\Program Files\poppler\Library\bin) to your System PATH environment variable.

**Microsoft Office (Optional):**

Microsoft Word is required for Word -> PDF conversion.

Microsoft PowerPoint is required for PowerPoint -> PDF conversion.

## 🏃‍♂️ How to Run
Ensure main.py and pdf_engine.py are in the same directory.

Run the application:

Bash
```
python main.py
```
