import sys
import os
import threading
import tempfile
import shutil
import webbrowser
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QLabel, QListWidget, 
                             QStackedWidget, QFileDialog, QMessageBox, QFrame,
                             QListWidgetItem, QAbstractItemView, QInputDialog, 
                             QLineEdit, QScrollArea, QComboBox, QRadioButton,
                             QButtonGroup, QMenu)
from PyQt6.QtCore import Qt, pyqtSignal, QObject, QSize
from PyQt6.QtGui import QDragEnterEvent, QDropEvent, QIcon, QFont, QPixmap, QKeyEvent, QAction
from pdf2image import convert_from_path

# Import backend engine
from pdf_engine import PDFEngine

# --- MODERN STYLESHEET (Preserved) ---
STYLESHEET = """
QMainWindow { background-color: #181825; }
QFrame#Sidebar { background-color: #1e1e2e; border-right: 1px solid #313244; }
QLabel { color: #cdd6f4; font-family: 'Segoe UI', sans-serif; }

QListWidget {
    background-color: #1e1e2e; 
    border: 2px dashed #45475a; 
    border-radius: 8px;
    color: #cdd6f4; 
    padding: 10px; 
    font-size: 13px;
    outline: none;
}
QListWidget::item { 
    padding: 8px; 
    margin-bottom: 4px; 
    border-radius: 6px; 
    background-color: #313244;
}
QListWidget::item:selected {
    background-color: #f38ba8;
    color: #181825;
}
QListWidget::item:hover { background-color: #45475a; }

QPushButton.nav-btn {
    background-color: transparent; 
    color: #a6adc8; 
    text-align: left;
    padding: 12px 20px; 
    font-size: 14px; 
    border: none; 
    border-radius: 6px;
    margin: 2px 10px;
}
QPushButton.nav-btn:hover { background-color: #313244; color: white; }
QPushButton.nav-btn:checked { 
    background-color: #313244; 
    color: #89b4fa;
    border-left: 3px solid #89b4fa;
    font-weight: bold; 
}

QPushButton.action-btn {
    background-color: #89b4fa; 
    color: #181825; 
    border-radius: 6px;
    padding: 12px 24px; 
    font-size: 15px; 
    font-weight: bold; 
    border: none;
}
QPushButton.action-btn:hover { background-color: #b4befe; }
QPushButton.action-btn:disabled { background-color: #45475a; color: #6c7086; }

QPushButton.upload-btn {
    background-color: #313244; 
    color: #cdd6f4; 
    border: 1px solid #45475a;
    border-radius: 6px;
    padding: 8px 16px; 
    font-size: 13px; 
}
QPushButton.upload-btn:hover { background-color: #45475a; border-color: #585b70; }

QLineEdit, QComboBox {
    background-color: #313244; 
    border: 1px solid #45475a; 
    color: white; 
    padding: 8px; 
    border-radius: 4px;
}
"""

# --- WORKER THREAD ---
class WorkerSignals(QObject):
    finished = pyqtSignal(str)
    error = pyqtSignal(str)
    result_data = pyqtSignal(object)

class TaskWorker(threading.Thread):
    def __init__(self, func, *args, **kwargs):
        super().__init__()
        self.func = func
        self.args = args
        self.kwargs = kwargs
        self.signals = WorkerSignals()

    def run(self):
        try:
            result = self.func(*self.args, **self.kwargs)
            self.signals.finished.emit("Done")
            if result:
                 self.signals.result_data.emit(result)
        except Exception as e:
            self.signals.error.emit(str(e))

# --- WIDGETS ---
class FileDropList(QListWidget):
    def __init__(self, allowed_exts=('.pdf',)):
        super().__init__()
        self.allowed_exts = allowed_exts
        self.setAcceptDrops(True)
        self.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self.setIconSize(QSize(40, 50))

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
            item = QListWidgetItem(os.path.basename(file_path))
            item.setToolTip(file_path)
            item.setData(Qt.ItemDataRole.UserRole, file_path)
            item.setIcon(QIcon.fromTheme("application-pdf")) 
            self.addItem(item)

    def keyPressEvent(self, event: QKeyEvent):
        if event.key() == Qt.Key.Key_Delete:
            for item in self.selectedItems():
                self.takeItem(self.row(item))
        else:
            super().keyPressEvent(event)

class OrganizerGrid(QListWidget):
    def __init__(self):
        super().__init__()
        self.setObjectName("PageGrid")
        self.setViewMode(QListWidget.ViewMode.IconMode)
        self.setResizeMode(QListWidget.ResizeMode.Adjust)
        self.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self.setIconSize(QSize(140, 190))
        self.setSpacing(15)
        self.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.customContextMenuRequested.connect(self.show_context_menu)

    def show_context_menu(self, pos):
        item = self.itemAt(pos)
        if not item: return
        menu = QMenu(self)
        menu.setStyleSheet("QMenu { background-color: #313244; color: white; } QMenu::item:selected { background: #89b4fa; color: black; }")
        rotate_cw = QAction("Rotate Right (90°)", self)
        rotate_ccw = QAction("Rotate Left (-90°)", self)
        delete_page = QAction("Delete Page", self)
        
        rotate_cw.triggered.connect(lambda: self.rotate_item(item, 90))
        rotate_ccw.triggered.connect(lambda: self.rotate_item(item, -90))
        delete_page.triggered.connect(lambda: self.takeItem(self.row(item)))
        
        menu.addAction(rotate_cw)
        menu.addAction(rotate_ccw)
        menu.addSeparator()
        menu.addAction(delete_page)
        menu.exec(self.mapToGlobal(pos))

    def rotate_item(self, item, angle):
        current_rot = item.data(Qt.ItemDataRole.UserRole + 1) or 0
        new_rot = (current_rot + angle) % 360
        item.setData(Qt.ItemDataRole.UserRole + 1, new_rot)
        # Visual rotation logic (simplified)
        icon = item.icon()
        pixmap = icon.pixmap(200, 200)
        from PyQt6.QtGui import QTransform
        transform = pixmap.transformed(QTransform().rotate(angle))
        item.setIcon(QIcon(transform))

# --- BASE PAGE ---
class BaseToolPage(QWidget):
    def __init__(self, title, desc, btn_text="Process", allowed_exts=('.pdf',), use_grid=False):
        super().__init__()
        self.allowed_exts = allowed_exts
        self.worker = None
        layout = QVBoxLayout()
        layout.setContentsMargins(40, 40, 40, 40)
        layout.setSpacing(15)
        
        head_lbl = QLabel(title)
        head_lbl.setFont(QFont("Segoe UI", 24, QFont.Weight.Bold))
        head_lbl.setStyleSheet("color: white;")
        layout.addWidget(head_lbl)
        
        desc_lbl = QLabel(desc)
        desc_lbl.setStyleSheet("color: #a6adc8; font-size: 15px;")
        desc_lbl.setWordWrap(True)
        layout.addWidget(desc_lbl)
        
        self.ctl_layout = QHBoxLayout()
        self.btn_upload = QPushButton("📂 Select Files")
        self.btn_upload.setProperty("class", "upload-btn")
        self.btn_upload.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_upload.clicked.connect(self.open_file_dialog)
        
        self.btn_clear = QPushButton("Clear List")
        self.btn_clear.setProperty("class", "upload-btn")
        self.btn_clear.clicked.connect(self.clear_list)
        
        self.ctl_layout.addWidget(self.btn_upload)
        self.ctl_layout.addWidget(self.btn_clear)
        self.ctl_layout.addStretch()
        layout.addLayout(self.ctl_layout)

        if use_grid: self.file_list = OrganizerGrid()
        else: self.file_list = FileDropList(allowed_exts)
        layout.addWidget(self.file_list)
        
        bot_layout = QHBoxLayout()
        self.btn_process = QPushButton(btn_text)
        self.btn_process.setProperty("class", "action-btn")
        self.btn_process.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_process.setFixedWidth(200)
        
        self.lbl_status = QLabel("")
        self.lbl_status.setStyleSheet("color: #89b4fa; font-weight: bold; margin-left: 15px;")
        
        bot_layout.addWidget(self.btn_process)
        bot_layout.addWidget(self.lbl_status)
        bot_layout.addStretch()
        layout.addLayout(bot_layout)
        
        self.setLayout(layout)

    def clear_list(self): self.file_list.clear()

    def open_file_dialog(self):
        ext_filter = " ".join([f"*{e}" for e in self.allowed_exts])
        files, _ = QFileDialog.getOpenFileNames(self, "Select Files", "", f"Files ({ext_filter})")
        if files: self.file_list.addItems(files)

    def get_files(self):
        return [self.file_list.item(i).data(Qt.ItemDataRole.UserRole) 
                for i in range(self.file_list.count())]

    def run_worker(self, func, *args, **kwargs):
        success_callback = kwargs.pop('success_callback', None)
        self.btn_process.setEnabled(False)
        self.btn_process.setText("Processing...")
        self.lbl_status.setText("Working...")
        
        self.worker = TaskWorker(func, *args, **kwargs)
        self.worker.signals.finished.connect(lambda _: self.on_worker_finished())
        self.worker.signals.error.connect(self.on_worker_error)
        if success_callback: 
            self.worker.signals.result_data.connect(success_callback)
        self.worker.start()

    def on_worker_finished(self):
        self.lbl_status.setText("Success!")
        self.btn_process.setText("Completed")
        self.btn_process.setEnabled(True)
        from PyQt6.QtCore import QTimer
        QTimer.singleShot(3000, lambda: self.btn_process.setText("Process"))

    def on_worker_error(self, err):
        self.lbl_status.setText("Error.")
        self.btn_process.setText("Retry")
        self.btn_process.setEnabled(True)
        QMessageBox.critical(self, "Error", f"Details: {err}")

# --- ENHANCED PAGES ---
class OrganizePage(BaseToolPage):
    def __init__(self):
        super().__init__("Visual Organizer", "Reorder, rotate, or delete pages.", "Save PDF", use_grid=True)
        self.btn_upload.setText("📂 Load PDF")
        self.btn_upload.disconnect()
        self.btn_upload.clicked.connect(self.load_pdf_dialog)
        self.btn_process.clicked.connect(self.save_pdf)
        self.current_pdf_path = None
        self.temp_dir = None

    def load_pdf_dialog(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select PDF", "", "PDF (*.pdf)")
        if path:
            self.current_pdf_path = path
            self.file_list.clear()
            self.cleanup_temp()
            self.temp_dir = tempfile.mkdtemp()
            self.lbl_status.setText("Loading pages...")
            self.run_worker(PDFEngine.pdf_to_images, path, self.temp_dir, 40, "jpeg", success_callback=self.populate_grid)

    def populate_grid(self, image_paths):
        for i, img_path in enumerate(image_paths):
            item = QListWidgetItem(f"Page {i+1}")
            item.setData(Qt.ItemDataRole.UserRole, i) 
            item.setData(Qt.ItemDataRole.UserRole + 1, 0)
            item.setIcon(QIcon(img_path))
            self.file_list.addItem(item)
        self.lbl_status.setText(f"Loaded {len(image_paths)} pages.")

    def save_pdf(self):
        if not self.current_pdf_path or self.file_list.count() == 0: return
        page_order_data = []
        for i in range(self.file_list.count()):
            item = self.file_list.item(i)
            page_order_data.append({
                'original_index': item.data(Qt.ItemDataRole.UserRole),
                'rotation': item.data(Qt.ItemDataRole.UserRole + 1)
            })
        save_path, _ = QFileDialog.getSaveFileName(self, "Save PDF", "organized.pdf", "PDF (*.pdf)")
        if save_path: 
            self.run_worker(PDFEngine.reorder_save_pdf, self.current_pdf_path, save_path, page_order_data)

    def cleanup_temp(self):
        if self.temp_dir and os.path.exists(self.temp_dir):
             shutil.rmtree(self.temp_dir)
             self.temp_dir = None

class SplitPage(BaseToolPage):
    def __init__(self):
        super().__init__("Split PDF", "Split all pages or extract ranges (e.g., '1-5, 8').")
        self.file_list.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.opts_layout = QHBoxLayout()
        self.rb_all = QRadioButton("Split All")
        self.rb_range = QRadioButton("Split Range")
        self.rb_extract = QRadioButton("Extract Range")
        self.rb_all.setChecked(True)
        self.bg = QButtonGroup()
        self.bg.addButton(self.rb_all)
        self.bg.addButton(self.rb_range)
        self.bg.addButton(self.rb_extract)
        self.opts_layout.addWidget(self.rb_all)
        self.opts_layout.addWidget(self.rb_range)
        self.opts_layout.addWidget(self.rb_extract)
        self.range_input = QLineEdit()
        self.range_input.setPlaceholderText("e.g. 1-5, 8")
        self.range_input.setVisible(False)
        self.opts_layout.addWidget(self.range_input)
        self.bg.buttonClicked.connect(lambda: self.range_input.setVisible(not self.rb_all.isChecked()))
        self.ctl_layout.addLayout(self.opts_layout)
        self.btn_process.clicked.connect(self.action)

    def action(self):
        files = self.get_files()
        if not files: return
        mode = "all"
        if self.rb_extract.isChecked(): mode = "extract"
        page_range = self.range_input.text() if not self.rb_all.isChecked() else None
        folder = QFileDialog.getExistingDirectory(self, "Output Folder")
        if folder: self.run_worker(PDFEngine.split_pdf, files[0], folder, mode, page_range)

class CompressPage(BaseToolPage):
    def __init__(self):
        super().__init__("Compress PDF", "Reduce file size.", "Compress")
        lbl = QLabel("Level:")
        lbl.setStyleSheet("color: white;")
        self.combo = QComboBox()
        self.combo.addItems(["Low (Lossless)", "Medium (Optimized)", "Extreme (Rasterize)"])
        self.ctl_layout.addWidget(lbl)
        self.ctl_layout.addWidget(self.combo)
        self.btn_process.clicked.connect(self.action)
    def action(self):
        files = self.get_files()
        if not files: return
        level_map = {0: "low", 1: "medium", 2: "extreme"}
        save_path, _ = QFileDialog.getSaveFileName(self, "Save", "compressed.pdf", "PDF (*.pdf)")
        if save_path: self.run_worker(PDFEngine.compress_pdf, files[0], save_path, level_map[self.combo.currentIndex()])

class ProtectPage(BaseToolPage):
    def __init__(self):
        super().__init__("Protect PDF", "Encrypt PDF with password and algorithm.", "Encrypt")
        
        lbl_algo = QLabel("Encryption:")
        lbl_algo.setStyleSheet("color: white;")
        self.combo_algo = QComboBox()
        self.combo_algo.addItems(["AES-256 (Strongest)", "AES-128 (Standard)", "RC4-128 (Legacy)"])
        
        self.ctl_layout.addWidget(lbl_algo)
        self.ctl_layout.addWidget(self.combo_algo)
        
        self.btn_process.clicked.connect(self.action)

    def action(self):
        files = self.get_files()
        if not files: return
        
        pwd, ok = QInputDialog.getText(self, "Password", "Enter Password:", QLineEdit.EchoMode.Password)
        if ok and pwd:
            algo_map = {0: "AES-256", 1: "AES-128", 2: "RC4-128"}
            selected_algo = algo_map[self.combo_algo.currentIndex()]
            
            save_path, _ = QFileDialog.getSaveFileName(self, "Save", "protected.pdf", "PDF (*.pdf)")
            if save_path: 
                self.run_worker(PDFEngine.protect_pdf, files[0], save_path, pwd, selected_algo)

class OpenProtectedPage(BaseToolPage):
    def __init__(self):
        super().__init__("Open Protected PDF", "Enter password to view secure PDF in default viewer.", "Unlock & Open")
        self.file_list.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.btn_process.clicked.connect(self.action)
        self.temp_file = None

    def action(self):
        files = self.get_files()
        if not files: return
        
        pwd, ok = QInputDialog.getText(self, "Password", "Enter PDF Password:", QLineEdit.EchoMode.Password)
        if not ok or not pwd: return

        # Create temp path
        fd, self.temp_file = tempfile.mkstemp(suffix=".pdf")
        os.close(fd)
        
        self.run_worker(PDFEngine.unlock_pdf, files[0], self.temp_file, pwd, success_callback=self.open_in_viewer)

    def open_in_viewer(self, path):
        try:
            webbrowser.open(path)
            self.lbl_status.setText("Opened in default viewer.")
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Could not open file: {e}")

# --- SIMPLE WRAPPERS ---
class MergePage(BaseToolPage):
    def __init__(self):
        super().__init__("Merge PDF", "Combine PDFs.")
        self.btn_process.clicked.connect(lambda: self.run_worker(PDFEngine.merge_pdfs, self.get_files(), QFileDialog.getSaveFileName(self, "Save", "merged.pdf", "PDF (*.pdf)")[0]) if len(self.get_files()) > 1 else None)

class ImgToPdfPage(BaseToolPage):
    def __init__(self):
        super().__init__("Images to PDF", "JPG/PNG to PDF.", "Convert", ('.jpg', '.png'))
        self.btn_process.clicked.connect(lambda: self.run_worker(PDFEngine.images_to_pdf, self.get_files(), QFileDialog.getSaveFileName(self, "Save", "img.pdf", "PDF (*.pdf)")[0]) if self.get_files() else None)

class PdfToImgPage(BaseToolPage):
    def __init__(self):
        super().__init__("PDF to JPG", "Extract pages as JPG.", "Extract")
        self.btn_process.clicked.connect(lambda: self.run_worker(PDFEngine.pdf_to_images, self.get_files()[0], QFileDialog.getExistingDirectory(self, "Output Folder")) if self.get_files() else None)

class PdfToWordPage(BaseToolPage):
    def __init__(self):
        super().__init__("PDF to Word", "Convert to Docx.")
        self.btn_process.clicked.connect(lambda: self.run_worker(PDFEngine.pdf_to_word, self.get_files()[0], QFileDialog.getSaveFileName(self, "Save", "c.docx", "Word (*.docx)")[0]) if self.get_files() else None)

class WordToPdfPage(BaseToolPage):
    def __init__(self):
        super().__init__("Word to PDF", "Convert Docx to PDF.", "Convert", ('.docx',))
        self.btn_process.clicked.connect(lambda: self.run_worker(PDFEngine.word_to_pdf, self.get_files()[0], QFileDialog.getSaveFileName(self, "Save", "c.pdf", "PDF (*.pdf)")[0]) if self.get_files() else None)

class PptxToPdfPage(BaseToolPage):
    def __init__(self):
        super().__init__("PPT to PDF", "Convert PPTX to PDF.", "Convert", ('.pptx',))
        self.btn_process.clicked.connect(lambda: self.run_worker(PDFEngine.pptx_to_pdf, self.get_files()[0], QFileDialog.getSaveFileName(self, "Save", "c.pdf", "PDF (*.pdf)")[0]) if self.get_files() else None)

class PdfToPptxPage(BaseToolPage):
    def __init__(self):
        super().__init__("PDF to PPT", "Convert PDF to Slides.")
        self.btn_process.clicked.connect(lambda: self.run_worker(PDFEngine.pdf_to_pptx, self.get_files()[0], QFileDialog.getSaveFileName(self, "Save", "c.pptx", "PPTX (*.pptx)")[0]) if self.get_files() else None)

# --- MAIN WINDOW ---
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Local PDF Pro")
        self.resize(1200, 800)
        self.setStyleSheet(STYLESHEET)
        
        central = QWidget()
        self.setCentralWidget(central)
        layout = QHBoxLayout(central)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        self.sidebar = QFrame()
        self.sidebar.setObjectName("Sidebar")
        self.sidebar.setFixedWidth(280)
        
        sidebar_layout = QVBoxLayout(self.sidebar)
        sidebar_layout.setContentsMargins(10, 30, 10, 20)
        sidebar_layout.setSpacing(5)

        title_lbl = QLabel("  PDF TOOLKIT")
        title_lbl.setStyleSheet("font-size: 18px; font-weight: bold; color: #89b4fa; margin-bottom: 20px;")
        sidebar_layout.addWidget(title_lbl)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("background: transparent; border: none;")
        nav_content = QWidget()
        nav_content.setStyleSheet("background: transparent;")
        self.nav_layout = QVBoxLayout(nav_content)
        self.nav_layout.setContentsMargins(0, 0, 0, 0)
        self.nav_layout.setSpacing(5)
        scroll.setWidget(nav_content)
        sidebar_layout.addWidget(scroll)

        self.stack = QStackedWidget()
        self.btns = []
        
        # --- NEW MENU STRUCTURE ---
        self.add_section("MOST USED")
        self.add_nav("Merge PDF", MergePage())
        self.add_nav("Visual Organize", OrganizePage())
        self.add_nav("Split PDF", SplitPage())
        self.add_nav("Compress PDF", CompressPage())
        
        self.add_section("CONVERT TO PDF")
        self.add_nav("Images to PDF", ImgToPdfPage())
        self.add_nav("Word to PDF", WordToPdfPage())
        self.add_nav("PowerPoint to PDF", PptxToPdfPage())
        
        self.add_section("CONVERT FROM PDF")
        self.add_nav("PDF to JPG", PdfToImgPage())
        self.add_nav("PDF to Word", PdfToWordPage())
        self.add_nav("PDF to PowerPoint", PdfToPptxPage())
        
        self.add_section("SECURITY")
        self.add_nav("Protect PDF", ProtectPage())
        self.add_nav("Open Protected PDF", OpenProtectedPage())

        self.nav_layout.addStretch()
        
        layout.addWidget(self.sidebar)
        layout.addWidget(self.stack)

    def add_section(self, text):
        lbl = QLabel(text)
        lbl.setStyleSheet("color: #6c7086; font-weight: bold; font-size: 11px; margin-top: 15px; margin-left: 10px;")
        self.nav_layout.addWidget(lbl)

    def add_nav(self, name, widget):
        btn = QPushButton(name)
        btn.setProperty("class", "nav-btn")
        btn.setCheckable(True)
        btn.setCursor(Qt.CursorShape.PointingHandCursor)
        
        idx = self.stack.addWidget(widget)
        btn.clicked.connect(lambda: self.switch_view(idx, btn))
        self.nav_layout.addWidget(btn)
        self.btns.append(btn)
        if len(self.btns) == 1: btn.click()

    def switch_view(self, idx, active_btn):
        self.stack.setCurrentIndex(idx)
        for b in self.btns: b.setChecked(False)
        active_btn.setChecked(True)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
