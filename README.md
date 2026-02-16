# Local PDF pro

Local PDF Pro is a comprehensive, secure, and offline PDF toolkit built with Python and PyQt6. This tool processes all data locally on your machine, ensuring maximum privacy and security for your documents. It mimics the functionality of premium online PDF tools but without file upload limits or data privacy concerns.

## Project Screenshot
<img width="1919" height="1007" alt="Screenshot 2026-02-16 133442" src="https://github.com/user-attachments/assets/e87086f5-b9ba-4561-a34d-5b9bc22ebc4b" />

## Features

This application includes 20+ tools organised into a dashboard for easy access:

### Most Used
* **Merge PDF:** Combine multiple PDF files into a single document.
* **Visual Organise:** Reorder, rotate, or delete specific pages visually.
* **Split PDF:** Split a PDF into separate files or extract specific pages.
* **Compress PDF:** Reduce file size with Low, Medium, or Extreme compression levels.
* **Page Numbers:** Add customizable page numbering to your documents.

### Convert To PDF
* **Images to PDF:** Convert JPG, PNG, and other image formats to PDF. Includes a "Smart Scan" feature to crop and warp document photos.
* **Word to PDF:** Convert Microsoft Word documents (.docx) to PDF.
* **PPT to PDF:** Convert PowerPoint presentations (.pptx) to PDF.
* **HTML to PDF:** Render raw HTML code or files into PDF using a modern Chromium engine.

### Convert From PDF
* **PDF to JPG:** Extract pages as high-quality images.
* **PDF to Word:** Convert PDF documents into editable Word files.
* **PDF to PPT:** Convert PDF slides into editable PowerPoint presentations.

### Security
* **Protect PDF:** Encrypt documents with a password (AES-256).
* **Unlock PDF:** Remove passwords from protected PDFs.

### Pro Features
* **OCR Searchable:** Convert scanned documents into text-searchable PDFs using Tesseract.
* **Watermark:** Add text watermarks to every page.
* **Edit Metadata:** View and modify title, author, subject, and creator tags.
* **Extract Images:** Extract all embedded images from a PDF file in their original quality.
* **Flatten PDF:** Merge form fields and annotations into the page content (useful for preventing further edits).
* **Grayscale PDF:** Convert colored documents to black and white to save ink.

## Installation

### 1. Prerequisites
Ensure you have Python 3.9 or higher installed. You also need the following external tools for specific features:

* **Tesseract OCR:** Required for the OCR feature.
    * Windows: Download the installer from UB-Mannheim/tesseract.
    * Linux: `sudo apt-get install tesseract-ocr`
    * Mac: `brew install tesseract`
* **Poppler:** Required for PDF to Image conversion.
    * Windows: Download the latest binary and add the `bin` folder to your System PATH.
    * Linux: `sudo apt-get install poppler-utils`
    * Mac: `brew install poppler`
* **Microsoft Office:** Required for native Word/PPT conversion (Windows only).

### 2. Install Dependencies
Run the following command to install the required Python packages:

```bash
pip install -r requirements.txt
```

### 3. Install Playwright Browsers
The HTML to PDF feature uses a headless browser. Run this command once after installing dependencies:
```bash
playwright install chromium
```
## Running the Application

To start the application, simply run the main script:
```bash
python main.py
```

## Building a Standalone Executable (.exe)
You can package this entire application into a single .exe file so it can be run on any Windows machine without needing Python installed.
1. Install PyInstaller:
   
   Ensure you have PyInstaller installed in your environment:
   ```bash
   pip install PyInstaller
   ```
2. Run the Build Command:
   
   Execute the following command in your terminal from the project's root directory. Ensure you have your pdf.ico icon file in the same folder.
   ```bash
   PyInstaller --noconsole --onefile --icon=pdf.ico --name="LocalPDFPro" --collect-all qtawesome main.py
   ```
   Command Breakdown:
   * --noconsole: Hides the background terminal window when the app is running.

   * --onefile: Bundles all Python dependencies, libraries, and code into a single executable file.

   * --icon=pdf.ico: Applies your custom icon to the executable.

   * --name="LocalPDFPro": Sets the output file name to LocalPDFPro.exe.

   * --collect-all qtawesome: Ensures PyInstaller gathers the font files required to render the application's icons.
3. Locate your App:
   
   Once the process finishes, navigate to the newly created dist folder. Inside, you will find your standalone LocalPDFPro.exe file ready to use or distribute.
