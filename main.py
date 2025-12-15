import sys
import os
import threading
import tempfile
import shutil
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QLabel, QListWidget, 
                             QStackedWidget, QFileDialog, QMessageBox, QFrame,
                             QListWidgetItem, QAbstractItemView, QInputDialog, 
                             QLineEdit, QScrollArea) # <--- Added QScrollArea
from PyQt6.QtCore import Qt, pyqtSignal, QObject, QSize
from PyQt6.QtGui import QDragEnterEvent, QDropEvent, QIcon, QFont, QPixmap, QKeyEvent
from pdf2image import convert_from_path

# Import backend engine
from pdf_engine import PDFEngine

# --- MODERN STYLESHEET ---
STYLESHEET = """
QMainWindow { background-color: #1e1e2e; }
QFrame#Sidebar { background-color: #252535; border-right: 1px solid #333; }
QLabel { color: #e0e0e0; font-family: 'Segoe UI'; }

/* Standard List View */
QListWidget#FileList {
    background-color: #2b2b3b; border: 2px dashed #444; border-radius: 10px;
    color: white; padding: 10px; font-size: 13px;
}
/* Grid View for Organizer */
QListWidget#PageGrid {
    background-color: #2b2b3b; border: 2px solid #444; border-radius: 10px;
    color: white; padding: 15px;
}
QListWidget::item:selected {
    background-color: #ff4757;
    border-radius: 5px;
}

/* Navigation Buttons */
QPushButton.nav-btn {
    background-color: transparent; color: #a0a0a0; text-align: left;
    padding: 12px 20px; font-size: 14px; border: none; border-radius: 5px;
}
QPushButton.nav-btn:hover { background-color: #2e2e3e; color: white; }
QPushButton.nav-btn:checked { background-color: #ff4757; color: white; font-weight: bold; }

/* Action Buttons */
QPushButton.action-btn {
    background-color: #ff4757; color: white; border-radius: 6px;
    padding: 10px 20px; font-size: 15px; font-weight: bold; border: none;
}
QPushButton.action-btn:hover { background-color: #ff6b81; }
QPushButton.action-btn:disabled { background-color: #555; color: #aaa; }

QPushButton.upload-btn {
    background-color: #3742fa; color: white; border-radius: 6px;
    padding: 8px 15px; font-size: 13px; border: none;
}
QPushButton.upload-btn:hover { background-color: #5352ed; }

/* Custom Scrollbar for Sidebar */
QScrollBar:vertical {
    background: #252535;
    width: 8px;
    margin: 0px 0px 0px 0px;
}
QScrollBar::handle:vertical {
    background: #555;
    min-height: 20px;
    border-radius: 4px;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0px;
}
"""

# --- WORKER THREAD ---
class WorkerSignals(QObject):
    finished = pyqtSignal(str)
    error = pyqtSignal(str)
    result_data = pyqtSignal(object)

class TaskWorker(threading.Thread):
    def __init__(self, func, *args):
        super().__init__()
        self.func = func
        self.args = args
        self.signals = WorkerSignals()

    def run(self):
        try:
            result = self.func(*self.args)
            self.signals.finished.emit("Task Completed Successfully!")
            if result:
                 self.signals.result_data.emit(result)
        except Exception as e:
            self.signals.error.emit(str(e))

# --- WIDGETS ---
class FileDropList(QListWidget):
    def __init__(self, allowed_exts=('.pdf',)):
        super().__init__()
        self.setObjectName("FileList")
        self.allowed_exts = allowed_exts
        self.setAcceptDrops(True)
        self.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self.setIconSize(QSize(50, 70))
        self.setSpacing(5)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls(): event.accept()
        else: event.ignore()

    def dropEvent(self, event: QDropEvent):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            self.handle_files([url.toLocalFile() for url in urls])
        else:
            super().dropEvent(event)

    def addItems(self, file_paths):
        self.handle_files(file_paths)

    def handle_files(self, file_paths):
        valid_files = [f for f in file_paths if f.lower().endswith(self.allowed_exts)]
        for file_path in valid_files:
            self.add_file_with_thumbnail(file_path)

    def add_file_with_thumbnail(self, file_path):
        item = QListWidgetItem(os.path.basename(file_path))
        item.setToolTip(file_path)
        item.setData(Qt.ItemDataRole.UserRole, file_path)
        
        if file_path.lower().endswith('.pdf'):
            try:
                images = convert_from_path(file_path, first_page=1, last_page=1, size=(100, None))
                if images:
                    from io import BytesIO
                    byte_io = BytesIO()
                    images[0].save(byte_io, 'PNG')
                    pixmap = QPixmap()
                    pixmap.loadFromData(byte_io.getvalue())
                    item.setIcon(QIcon(pixmap))
            except: pass
        
        self.addItem(item)

    def keyPressEvent(self, event: QKeyEvent):
        if event.key() == Qt.Key.Key_Delete:
            for item in self.selectedItems():
                self.takeItem(self.row(item))
        else:
            super().keyPressEvent(event)

class PageGridWidget(QListWidget):
    def __init__(self):
        super().__init__()
        self.setObjectName("PageGrid")
        self.setViewMode(QListWidget.ViewMode.IconMode)
        self.setResizeMode(QListWidget.ResizeMode.Adjust)
        self.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self.setIconSize(QSize(120, 160))
        self.setSpacing(15)
        self.setWrapping(True)

    def keyPressEvent(self, event: QKeyEvent):
        if event.key() == Qt.Key.Key_Delete:
            for item in self.selectedItems():
                self.takeItem(self.row(item))
        else:
            super().keyPressEvent(event)

# --- PAGES ---
class BaseToolPage(QWidget):
    def __init__(self, title, desc, btn_text="Process", allowed_exts=('.pdf',), use_grid=False):
        super().__init__()
        self.allowed_exts = allowed_exts
        self.worker = None
        layout = QVBoxLayout()
        layout.setContentsMargins(30, 30, 30, 30)
        
        head_lbl = QLabel(title)
        head_lbl.setFont(QFont("Segoe UI", 22, QFont.Weight.Bold))
        layout.addWidget(head_lbl)
        desc_lbl = QLabel(desc)
        desc_lbl.setStyleSheet("color: #888; font-size: 14px;")
        layout.addWidget(desc_lbl)
        layout.addSpacing(15)
        
        ctl_layout = QHBoxLayout()
        self.btn_upload = QPushButton("📂 Add Files")
        self.btn_upload.setProperty("class", "upload-btn")
        self.btn_upload.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_upload.clicked.connect(self.open_file_dialog)
        self.btn_clear = QPushButton("Clear")
        self.btn_clear.setStyleSheet("background-color: #444; color: white; border-radius: 5px; padding: 8px 15px;")
        self.btn_clear.clicked.connect(self.clear_list)
        
        ctl_layout.addWidget(self.btn_upload)
        ctl_layout.addWidget(self.btn_clear)
        ctl_layout.addStretch()
        layout.addLayout(ctl_layout)

        if use_grid: self.file_list = PageGridWidget()
        else: self.file_list = FileDropList(allowed_exts)
        layout.addWidget(self.file_list)
        
        self.btn_process = QPushButton(btn_text)
        self.btn_process.setProperty("class", "action-btn")
        self.btn_process.setCursor(Qt.CursorShape.PointingHandCursor)
        self.lbl_status = QLabel("")
        self.lbl_status.setStyleSheet("color: #00d2d3; font-weight: bold; margin-top: 10px;")
        
        layout.addWidget(self.btn_process)
        layout.addWidget(self.lbl_status)
        self.setLayout(layout)

    def clear_list(self): self.file_list.clear()

    def open_file_dialog(self):
        ext_filter = " ".join([f"*{e}" for e in self.allowed_exts])
        files, _ = QFileDialog.getOpenFileNames(self, "Select Files", "", f"Files ({ext_filter})")
        if files: self.file_list.addItems(files)

    def get_files(self):
        return [self.file_list.item(i).data(Qt.ItemDataRole.UserRole) 
                for i in range(self.file_list.count())]

    def run_worker(self, func, *args, success_callback=None):
        self.btn_process.setEnabled(False)
        self.btn_upload.setEnabled(False)
        self.btn_process.setText("Processing...")
        self.lbl_status.setText("Working... Please wait.")
        self.worker = TaskWorker(func, *args)
        self.worker.signals.finished.connect(self.on_worker_finished)
        self.worker.signals.error.connect(self.on_worker_error)
        if success_callback: self.worker.signals.result_data.connect(success_callback)
        self.worker.start()

    def on_worker_finished(self, msg):
        self.lbl_status.setText(msg)
        self.btn_process.setText("Process Completed")
        self.btn_process.setEnabled(True)
        self.btn_upload.setEnabled(True)
        QMessageBox.information(self, "Success", msg)

    def on_worker_error(self, err):
        self.lbl_status.setText("Error occurred.")
        self.btn_process.setText("Retry")
        self.btn_process.setEnabled(True)
        self.btn_upload.setEnabled(True)
        QMessageBox.critical(self, "Error", f"An error occurred:\n{err}")

# --- SPECIAL PAGES ---
class OrganizePage(BaseToolPage):
    def __init__(self):
        super().__init__("Organize PDF", "Load a PDF to view pages. Drag to reorder. Delete to remove.", "Save Ordered PDF", use_grid=True)
        self.btn_upload.setText("📂 Load PDF")
        self.btn_upload.disconnect()
        self.btn_upload.clicked.connect(self.load_pdf_dialog)
        self.btn_process.clicked.connect(self.save_pdf)
        self.temp_dir = None

    def load_pdf_dialog(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select PDF", "", "PDF (*.pdf)")
        if path:
            self.file_list.clear()
            self.cleanup_temp()
            self.temp_dir = tempfile.mkdtemp()
            self.run_worker(PDFEngine.pdf_to_images, path, self.temp_dir, 150, "jpeg", success_callback=self.populate_grid)

    def populate_grid(self, image_paths):
        for i, img_path in enumerate(image_paths):
            item = QListWidgetItem(f"Page {i+1}")
            item.setData(Qt.ItemDataRole.UserRole, img_path)
            item.setIcon(QIcon(img_path))
            self.file_list.addItem(item)
        self.lbl_status.setText(f"Loaded {len(image_paths)} pages.")

    def save_pdf(self):
        if self.file_list.count() == 0: return
        ordered_images = self.get_files()
        save_path, _ = QFileDialog.getSaveFileName(self, "Save PDF", "organized.pdf", "PDF (*.pdf)")
        if save_path: self.run_worker(PDFEngine.images_to_pdf, ordered_images, save_path)

    def cleanup_temp(self):
        if self.temp_dir and os.path.exists(self.temp_dir):
             shutil.rmtree(self.temp_dir)
             self.temp_dir = None
    
    def clear_list(self):
        super().clear_list()
        self.cleanup_temp()

class MergePage(BaseToolPage):
    def __init__(self):
        super().__init__("Merge PDF", "Combine multiple PDFs into one.")
        self.btn_process.clicked.connect(self.action)
    def action(self):
        files = self.get_files()
        if len(files) < 2: return QMessageBox.warning(self, "Info", "Select 2+ files.")
        save_path, _ = QFileDialog.getSaveFileName(self, "Save Merged", "merged.pdf", "PDF (*.pdf)")
        if save_path: self.run_worker(PDFEngine.merge_pdfs, files, save_path)

class SplitPage(BaseToolPage):
    def __init__(self):
        super().__init__("Split PDF", "Extract all pages into separate files.")
        self.file_list.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.btn_process.clicked.connect(self.action)
    def action(self):
        files = self.get_files()
        if not files: return
        folder = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if folder: self.run_worker(PDFEngine.split_pdf, files[0], folder)

class ImgToPdfPage(BaseToolPage):
    def __init__(self):
        super().__init__("Images to PDF", "Convert JPG/PNG to PDF.", "Convert", ('.jpg', '.png', '.jpeg'))
        self.btn_process.clicked.connect(self.action)
    def action(self):
        files = self.get_files()
        if not files: return
        save_path, _ = QFileDialog.getSaveFileName(self, "Save PDF", "images.pdf", "PDF (*.pdf)")
        if save_path: self.run_worker(PDFEngine.images_to_pdf, files, save_path)

class PdfToImgPage(BaseToolPage):
    def __init__(self):
        super().__init__("PDF to JPG", "Extract pages as JPEG images.", "Extract")
        self.btn_process.clicked.connect(self.action)
    def action(self):
        files = self.get_files()
        if not files: return
        folder = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if folder: self.run_worker(PDFEngine.pdf_to_images, files[0], folder)

class PdfToWordPage(BaseToolPage):
    def __init__(self):
        super().__init__("PDF to Word", "Convert PDF to Docx.", "Convert")
        self.btn_process.clicked.connect(self.action)
    def action(self):
        files = self.get_files()
        if not files: return
        save_path, _ = QFileDialog.getSaveFileName(self, "Save Word", "converted.docx", "Word (*.docx)")
        if save_path: self.run_worker(PDFEngine.pdf_to_word, files[0], save_path)

class WordToPdfPage(BaseToolPage):
    def __init__(self):
        super().__init__("Word to PDF", "Convert Docx to PDF (Req. Word).", "Convert", ('.docx', '.doc'))
        self.btn_process.clicked.connect(self.action)
    def action(self):
        files = self.get_files()
        if not files: return
        save_path, _ = QFileDialog.getSaveFileName(self, "Save PDF", "converted.pdf", "PDF (*.pdf)")
        if save_path: self.run_worker(PDFEngine.word_to_pdf, files[0], save_path)

class PptxToPdfPage(BaseToolPage):
    def __init__(self):
        super().__init__("PowerPoint to PDF", "Convert PPTX to PDF (Req. PowerPoint).", "Convert", ('.pptx', '.ppt'))
        self.btn_process.clicked.connect(self.action)
    def action(self):
        files = self.get_files()
        if not files: return
        save_path, _ = QFileDialog.getSaveFileName(self, "Save PDF", "presentation.pdf", "PDF (*.pdf)")
        if save_path: self.run_worker(PDFEngine.pptx_to_pdf, files[0], save_path)

class PdfToPptxPage(BaseToolPage):
    def __init__(self):
        super().__init__("PDF to PowerPoint", "Convert PDF pages to Slides.", "Convert")
        self.btn_process.clicked.connect(self.action)
    def action(self):
        files = self.get_files()
        if not files: return
        save_path, _ = QFileDialog.getSaveFileName(self, "Save PPTX", "converted.pptx", "PowerPoint (*.pptx)")
        if save_path: self.run_worker(PDFEngine.pdf_to_pptx, files[0], save_path)

class ProtectPage(BaseToolPage):
    def __init__(self):
        super().__init__("Protect PDF", "Encrypt PDF with password.", "Encrypt")
        self.btn_process.clicked.connect(self.action)
    def action(self):
        files = self.get_files()
        if not files: return
        pwd, ok = QInputDialog.getText(self, "Password", "Enter Password:", QLineEdit.EchoMode.Password)
        if ok and pwd:
            save_path, _ = QFileDialog.getSaveFileName(self, "Save PDF", "protected.pdf", "PDF (*.pdf)")
            if save_path: self.run_worker(PDFEngine.protect_pdf, files[0], save_path, pwd)

# --- MAIN WINDOW ---
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Local PDF Pro")
        self.resize(1150, 750)
        self.setStyleSheet(STYLESHEET)
        
        central = QWidget()
        self.setCentralWidget(central)
        layout = QHBoxLayout(central)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # 1. Sidebar Container
        self.sidebar = QFrame()
        self.sidebar.setObjectName("Sidebar")
        self.sidebar.setFixedWidth(260)
        
        # Outer layout for sidebar (Title fixed, Buttons scrollable)
        sidebar_outer_layout = QVBoxLayout(self.sidebar)
        sidebar_outer_layout.setContentsMargins(0, 0, 0, 0)
        sidebar_outer_layout.setSpacing(0)

        # 2. Title Area (Fixed at top)
        title_frame = QFrame()
        title_layout = QVBoxLayout(title_frame)
        title_layout.setContentsMargins(20, 30, 10, 20)
        title = QLabel("PDF TOOLKIT")
        title.setStyleSheet("font-size: 20px; font-weight: bold; color: white;")
        title_layout.addWidget(title)
        sidebar_outer_layout.addWidget(title_frame)

        # 3. Scroll Area for Buttons
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        # Hide horizontal bar, style vertical
        self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.scroll_area.setStyleSheet("QScrollArea { border: none; background: transparent; }")
        
        # Content Widget for Scroll Area
        self.nav_content = QWidget()
        self.nav_content.setStyleSheet("background: transparent;")
        self.nav_layout = QVBoxLayout(self.nav_content)
        self.nav_layout.setContentsMargins(10, 0, 10, 20)
        self.nav_layout.setSpacing(8)
        
        self.scroll_area.setWidget(self.nav_content)
        sidebar_outer_layout.addWidget(self.scroll_area)

        # 4. Main Stacked Widget
        self.stack = QStackedWidget()
        self.btns = []
        
        # Tools
        self.add_header("ORGANIZE")
        self.add_tool("Merge PDF", MergePage())
        self.add_tool("Split PDF", SplitPage())
        self.add_tool("Visual Organize", OrganizePage())
        
        self.add_header("CONVERT TO PDF")
        self.add_tool("Images to PDF", ImgToPdfPage())
        self.add_tool("Word to PDF", WordToPdfPage())
        self.add_tool("PowerPoint to PDF", PptxToPdfPage())

        self.add_header("CONVERT FROM PDF")
        self.add_tool("PDF to JPG", PdfToImgPage())
        self.add_tool("PDF to Word", PdfToWordPage())
        self.add_tool("PDF to PowerPoint", PdfToPptxPage())
        
        self.add_header("SECURITY")
        self.add_tool("Protect PDF", ProtectPage())

        # Push everything up in the scroll area
        self.nav_layout.addStretch()

        layout.addWidget(self.sidebar)
        layout.addWidget(self.stack)

    def add_header(self, text):
        lbl = QLabel(text)
        lbl.setStyleSheet("color: #666; font-weight: bold; font-size: 12px; margin-top: 15px; margin-bottom: 5px; margin-left: 10px;")
        self.nav_layout.addWidget(lbl) # Changed to nav_layout

    def add_tool(self, name, widget):
        btn = QPushButton(name)
        btn.setProperty("class", "nav-btn")
        btn.setCheckable(True)
        btn.setCursor(Qt.CursorShape.PointingHandCursor)
        
        idx = self.stack.addWidget(widget)
        btn.clicked.connect(lambda: self.switch_view(idx, btn))
        self.nav_layout.addWidget(btn) # Changed to nav_layout
        self.btns.append(btn)
        
        if idx == 0: btn.click()

    def switch_view(self, idx, active_btn):
        self.stack.setCurrentIndex(idx)
        for b in self.btns: b.setChecked(False)
        active_btn.setChecked(True)
        organizer = self.stack.widget(2)
        if isinstance(organizer, OrganizePage) and idx != 2:
             organizer.cleanup_temp()

if __name__ == "__main__":
    if hasattr(Qt.ApplicationAttribute, 'AA_EnableHighDpiScaling'):
        QApplication.setAttribute(Qt.ApplicationAttribute.AA_EnableHighDpiScaling, True)
    if hasattr(Qt.ApplicationAttribute, 'AA_UseHighDpiPixmaps'):
        QApplication.setAttribute(Qt.ApplicationAttribute.AA_UseHighDpiPixmaps, True)

    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
