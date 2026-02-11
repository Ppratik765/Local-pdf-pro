import sys
import os
import threading
import tempfile
import shutil
import webbrowser
import json
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QLabel, QListWidget, 
                             QStackedWidget, QFileDialog, QMessageBox, QFrame,
                             QListWidgetItem, QAbstractItemView, QInputDialog, 
                             QLineEdit, QScrollArea, QComboBox, QRadioButton,
                             QButtonGroup, QMenu, QDialog, QGridLayout, QCheckBox, QSizePolicy,
                             QInputDialog, QLineEdit, QScrollArea, QTextEdit)
from PyQt6.QtCore import Qt, pyqtSignal, QObject, QSize, QSettings
from PyQt6.QtGui import QDragEnterEvent, QDropEvent, QIcon, QFont, QPixmap, QKeyEvent, QAction, QColor
from pdf2image import convert_from_path

# Import backend engine
from pdf_engine import PDFEngine

# --- THEMES (Unchanged) ---
DARK_THEME = """
QMainWindow { background-color: #181825; }
QFrame#Sidebar { background-color: #1e1e2e; border-right: 1px solid #313244; }
QLabel { color: #cdd6f4; font-family: 'Segoe UI', sans-serif; }
QListWidget { background-color: #1e1e2e; border: 2px dashed #45475a; border-radius: 8px; color: #cdd6f4; padding: 10px; font-size: 13px; }
QListWidget::item { padding: 8px; margin-bottom: 4px; border-radius: 6px; background-color: #313244; }
QListWidget::item:selected { background-color: #f38ba8; color: #181825; }
QPushButton.nav-btn { background-color: transparent; color: #a6adc8; text-align: left; padding: 12px 20px; font-size: 14px; border: none; border-radius: 6px; margin: 2px 10px; }
QPushButton.nav-btn:hover { background-color: #313244; color: white; }
QPushButton.nav-btn:checked { background-color: #313244; color: #89b4fa; border-left: 3px solid #89b4fa; font-weight: bold; }
QPushButton.action-btn { background-color: #89b4fa; color: #181825; border-radius: 6px; padding: 12px 24px; font-size: 15px; font-weight: bold; border: none; }
QPushButton.action-btn:hover { background-color: #b4befe; }
QPushButton.action-btn:disabled { background-color: #45475a; color: #6c7086; }
QPushButton.upload-btn { background-color: #313244; color: #cdd6f4; border: 1px solid #45475a; border-radius: 6px; padding: 8px 16px; font-size: 13px; }
QPushButton.upload-btn:hover { background-color: #45475a; border-color: #585b70; }
QLineEdit, QComboBox { background-color: #313244; border: 1px solid #45475a; color: white; padding: 8px; border-radius: 4px; }
/* Dashboard Grid Buttons */
QPushButton.dash-btn {
    background-color: #313244; color: #cdd6f4; border: 1px solid #45475a; border-radius: 8px;
    padding: 20px; font-size: 14px; font-weight: bold; text-align: center;
}
QPushButton.dash-btn:hover { background-color: #45475a; border-color: #89b4fa; color: white; }
"""

LIGHT_THEME = """
QMainWindow { background-color: #eff1f5; }
QFrame#Sidebar { background-color: #e6e9ef; border-right: 1px solid #bcc0cc; }
QLabel { color: #4c4f69; font-family: 'Segoe UI', sans-serif; }
QListWidget { background-color: white; border: 2px dashed #bcc0cc; border-radius: 8px; color: #4c4f69; padding: 10px; font-size: 13px; }
QListWidget::item { padding: 8px; margin-bottom: 4px; border-radius: 6px; background-color: #e6e9ef; }
QListWidget::item:selected { background-color: #ea76cb; color: white; }
QPushButton.nav-btn { background-color: transparent; color: #5c5f77; text-align: left; padding: 12px 20px; font-size: 14px; border: none; border-radius: 6px; margin: 2px 10px; }
QPushButton.nav-btn:hover { background-color: #dce0e8; color: #4c4f69; }
QPushButton.nav-btn:checked { background-color: #dce0e8; color: #1e66f5; border-left: 3px solid #1e66f5; font-weight: bold; }
QPushButton.action-btn { background-color: #1e66f5; color: white; border-radius: 6px; padding: 12px 24px; font-size: 15px; font-weight: bold; border: none; }
QPushButton.action-btn:hover { background-color: #7287fd; }
QPushButton.action-btn:disabled { background-color: #ccd0da; color: #9ca0b0; }
QPushButton.upload-btn { background-color: white; color: #4c4f69; border: 1px solid #bcc0cc; border-radius: 6px; padding: 8px 16px; font-size: 13px; }
QPushButton.upload-btn:hover { background-color: #eff1f5; border-color: #9ca0b0; }
QLineEdit, QComboBox { background-color: white; border: 1px solid #bcc0cc; color: #4c4f69; padding: 8px; border-radius: 4px; }
QPushButton.dash-btn {
    background-color: white; color: #4c4f69; border: 1px solid #bcc0cc; border-radius: 8px;
    padding: 20px; font-size: 14px; font-weight: bold; text-align: center;
}
QPushButton.dash-btn:hover { background-color: #e6e9ef; border-color: #1e66f5; color: #1e66f5; }
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

# --- CUSTOM WIDGETS ---
class SidebarBtn(QPushButton):
    fileDropped = pyqtSignal(str)
    
    def __init__(self, text):
        super().__init__(text)
        self.setAcceptDrops(True)
        self.setProperty("class", "nav-btn")
        self.setCheckable(True)
        self.setCursor(Qt.CursorShape.PointingHandCursor)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls(): event.accept()
        else: event.ignore()

    def dropEvent(self, event: QDropEvent):
        if event.mimeData().hasUrls():
            files = [u.toLocalFile() for u in event.mimeData().urls()]
            if files: self.fileDropped.emit(files[0])

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
            # Removed Recent Files update call

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
        icon = item.icon()
        pixmap = icon.pixmap(200, 200)
        from PyQt6.QtGui import QTransform
        transform = pixmap.transformed(QTransform().rotate(angle))
        item.setIcon(QIcon(transform))

from PyQt6.QtWidgets import QGraphicsView, QGraphicsScene, QGraphicsPixmapItem, QGraphicsEllipseItem
from PyQt6.QtGui import QPen, QBrush, QPainter

class DraggableScanDialog(QDialog):
    def __init__(self, image_path):
        super().__init__()
        self.setWindowTitle("Adjust Document Corners")
        self.resize(1000, 650) # Smaller window as requested
        self.image_path = image_path
        self.scale_factor = 1.0
        
        # Layout
        main_layout = QVBoxLayout(self)
        
        # Instruction Label
        lbl = QLabel("Drag the RED circles to the 4 corners of the document.")
        lbl.setStyleSheet("color: #a6adc8; font-weight: bold; font-size: 14px;")
        lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(lbl)

        # Graphics View for Image + Handles
        self.scene = QGraphicsScene()
        self.view = QGraphicsView(self.scene)
        self.view.setRenderHint(QPainter.RenderHint.Antialiasing)
        self.view.setStyleSheet("background: #181825; border: 2px solid #45475a; border-radius: 8px;")
        
        # Load Image
        self.pixmap = QPixmap(image_path)
        
        # Fit image to view logic (calculated roughly based on dialog size)
        view_w, view_h = 950, 500
        img_w, img_h = self.pixmap.width(), self.pixmap.height()
        self.scale_factor = min(view_w / img_w, view_h / img_h)
        
        # Scale pixmap for display
        scaled_pix = self.pixmap.scaled(
            int(img_w * self.scale_factor), 
            int(img_h * self.scale_factor), 
            Qt.AspectRatioMode.KeepAspectRatio, 
            Qt.TransformationMode.SmoothTransformation
        )
        
        self.scene_img = QGraphicsPixmapItem(scaled_pix)
        self.scene.addItem(self.scene_img)
        self.view.setSceneRect(0, 0, scaled_pix.width(), scaled_pix.height())
        
        main_layout.addWidget(self.view)

        # Add 4 Handles (Default to corners with slight padding)
        w, h = scaled_pix.width(), scaled_pix.height()
        pad = 50
        self.handles = []
        # Top-Left, Top-Right, Bottom-Right, Bottom-Left
        points = [(pad, pad), (w-pad, pad), (w-pad, h-pad), (pad, h-pad)]
        
        for x, y in points:
            handle = self.create_handle(x, y)
            self.handles.append(handle)
            self.scene.addItem(handle)

        # Buttons
        btn_layout = QHBoxLayout()
        btn_cancel = QPushButton("Cancel")
        btn_cancel.clicked.connect(self.reject)
        
        btn_apply = QPushButton("Apply Crop")
        btn_apply.setStyleSheet("background-color: #89b4fa; color: #1e1e2e; font-weight: bold; padding: 10px 20px; border-radius: 6px;")
        btn_apply.clicked.connect(self.on_apply)
        
        btn_layout.addStretch()
        btn_layout.addWidget(btn_cancel)
        btn_layout.addWidget(btn_apply)
        main_layout.addLayout(btn_layout)

    def create_handle(self, x, y):
        # Create a red draggable circle
        size = 20
        ellipse = QGraphicsEllipseItem(0, 0, size, size)
        ellipse.setPos(x - size/2, y - size/2)
        ellipse.setBrush(QBrush(QColor("#ff4757"))) # Red fill
        ellipse.setPen(QPen(Qt.GlobalColor.white, 2)) # White border
        ellipse.setFlag(QGraphicsEllipseItem.GraphicsItemFlag.ItemIsMovable)
        ellipse.setFlag(QGraphicsEllipseItem.GraphicsItemFlag.ItemSendsGeometryChanges)
        return ellipse

    def on_apply(self):
        # 1. Get positions relative to the SCALED image
        # 2. Convert back to ORIGINAL image coordinates
        final_corners = []
        for h in self.handles:
            # Add radius offset to get center
            center_x = h.pos().x() + 10 
            center_y = h.pos().y() + 10
            
            orig_x = int(center_x / self.scale_factor)
            orig_y = int(center_y / self.scale_factor)
            final_corners.append((orig_x, orig_y))
        
        self.final_corners = final_corners
        self.accept()

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
        layout.addWidget(head_lbl)
        
        desc_lbl = QLabel(desc)
        desc_lbl.setWordWrap(True)
        layout.addWidget(desc_lbl)
        
        self.ctl_layout = QHBoxLayout()
        self.btn_upload = QPushButton("📂 Select Files")
        self.btn_upload.setProperty("class", "upload-btn")
        self.btn_upload.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_upload.clicked.connect(self.open_file_dialog)
        
        self.btn_clear = QPushButton("Clear")
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

# --- DASHBOARD (REDESIGNED) ---
class DashboardPage(QWidget):
    def __init__(self, nav_callback):
        super().__init__()
        self.nav_callback = nav_callback
        
        # Scroll Area for the whole dashboard
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setStyleSheet("background: transparent;")
        
        content_widget = QWidget()
        layout = QVBoxLayout(content_widget)
        layout.setContentsMargins(50, 50, 50, 50)
        layout.setSpacing(20)
        
        # Intro Header
        title = QLabel("Local PDF Pro")
        title.setFont(QFont("Segoe UI", 40, QFont.Weight.Bold))
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("color: #89b4fa;")
        layout.addWidget(title)
        
        desc = QLabel("Your secure, offline, all-in-one PDF toolkit.\nSelect a tool to get started.")
        desc.setAlignment(Qt.AlignmentFlag.AlignCenter)
        desc.setStyleSheet("color: #a6adc8; font-size: 16px;")
        layout.addWidget(desc)
        
        layout.addSpacing(30)

        # Tools Grid
        grid = QGridLayout()
        grid.setSpacing(20)
        
        # Define tools: (Name, Stack Index)
        # Note: Stack indices must match creation order in MainWindow
        tools = [
            ("Merge PDF", 1), ("Visual Organize", 2), ("Split PDF", 3), ("Compress PDF", 4),
            ("Images to PDF", 5), ("Word to PDF", 6), ("PPT to PDF", 7),
            ("PDF to JPG", 8), ("PDF to Word", 9), ("PDF to PPT", 10),
            ("Protect PDF", 11), ("Unlock PDF", 12),
            ("OCR Searchable", 13), ("Watermark", 14), ("Page Numbers", 15), ("Edit Metadata", 16),
            ("HTML to PDF", 17)
        ]

        row, col = 0, 0
        for name, idx in tools:
            btn = QPushButton(name)
            btn.setProperty("class", "dash-btn")
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            # Use closure to capture loop variable
            btn.clicked.connect(lambda _, x=idx: self.nav_callback(x))
            
            grid.addWidget(btn, row, col)
            col += 1
            if col > 3: # 4 columns
                col = 0
                row += 1
        
        layout.addLayout(grid)
        layout.addStretch()
        
        scroll.setWidget(content_widget)
        
        # Main layout for this page widget
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0,0,0,0)
        main_layout.addWidget(scroll)

# --- PAGES (Same as before) ---
class ImgToPdfPage(BaseToolPage):
    def __init__(self):
        super().__init__("Images to PDF", "Smart Scan or convert standard images.", "Convert", ('.jpg', '.png', '.jpeg'))
        
        self.btn_smart = QPushButton("📷 Smart Scan")
        self.btn_smart.setProperty("class", "upload-btn")
        self.btn_smart.setStyleSheet("background-color: #f38ba8; color: #181825; border: none; font-weight: bold;")
        self.btn_smart.clicked.connect(self.smart_scan_dialog)
        self.ctl_layout.insertWidget(1, self.btn_smart)
        
        self.btn_process.clicked.connect(self.action)

    def smart_scan_dialog(self):
            file, _ = QFileDialog.getOpenFileName(self, "Select Image", "", "Images (*.jpg *.jpeg *.png)")
            if not file: return
            
            # Open Manual Editor directly
            dlg = DraggableScanDialog(file)
            if dlg.exec():
                try:
                    # User clicked Apply, we have coordinates
                    corners = dlg.final_corners
                    
                    # Perform the warp using the engine
                    processed_img = PDFEngine.manual_scan_warp(file, corners)
                    
                    # Save result
                    fd, path = tempfile.mkstemp(suffix=".jpg")
                    os.close(fd)
                    processed_img.save(path)
                    self.file_list.addItems([path])
                    
                except Exception as e:
                    QMessageBox.warning(self, "Processing Error", str(e))

    def action(self):
        files = self.get_files()
        if not files: return
        save_path, _ = QFileDialog.getSaveFileName(self, "Save PDF", "images.pdf", "PDF (*.pdf)")
        if save_path: self.run_worker(PDFEngine.images_to_pdf, files, save_path)

# ... (OCRPage, WatermarkPage, PageNumPage, MetadataPage, OrganizePage, SplitPage, CompressPage, ProtectPage, OpenProtectedPage, MergePage, PdfToImgPage, PdfToWordPage, WordToPdfPage, PptxToPdfPage, PdfToPptxPage - KEEP THESE CLASSES AS IS from previous main.py) ...
# For brevity, I am assuming the previous classes are present here. 
# Make sure to copy them over! Below are the simplified versions for context.

class OCRPage(BaseToolPage):
    def __init__(self):
        super().__init__("OCR (Searchable PDF)", "Make scanned documents searchable.", "Run OCR")
        self.btn_process.clicked.connect(self.action)
    def action(self):
        files = self.get_files()
        if not files: return
        save_path, _ = QFileDialog.getSaveFileName(self, "Save", "ocr.pdf", "PDF (*.pdf)")
        if save_path: self.run_worker(PDFEngine.ocr_pdf, files[0], save_path)
class HtmlToPdfPage(BaseToolPage):
    def __init__(self):
        super().__init__("HTML to PDF (Pro)", "Render modern HTML/CSS with emoji support.", "Convert to PDF")
        
        self.file_list.setParent(None)
        self.btn_upload.setParent(None)
        self.btn_clear.setParent(None)
        
        # Modern Code Editor Style
        self.text_area = QTextEdit()
        self.text_area.setPlaceholderText("Paste full HTML or snippets here.\nSupports: Tailwind, Flexbox, Emojis (🚀), Custom Fonts...")
        self.text_area.setAcceptRichText(False) # Force plain text handling
        self.text_area.setStyleSheet("""
            QTextEdit {
                background-color: #1e1e1e; 
                color: #d4d4d4; 
                border: 1px solid #3c3c3c; 
                border-radius: 4px; 
                padding: 10px; 
                font-family: 'Consolas', 'Monaco', monospace;
                font-size: 14px;
                line-height: 1.5;
            }
        """)
        
        self.layout().insertWidget(3, self.text_area)
        self.btn_process.clicked.connect(self.action)

    def action(self):
        html_content = self.text_area.toPlainText()
        if not html_content.strip():
            return QMessageBox.warning(self, "Input Required", "Please paste some HTML code.")
            
        save_path, _ = QFileDialog.getSaveFileName(self, "Save PDF", "modern_render.pdf", "PDF (*.pdf)")
        if save_path:
            # The BaseToolPage worker will handle the threading, keeping the GUI responsive
            self.run_worker(PDFEngine.html_to_pdf, html_content, save_path)
            
class WatermarkPage(BaseToolPage):
    def __init__(self):
        super().__init__("Watermark", "Add text overlay.", "Apply")
        self.txt = QLineEdit()
        self.txt.setPlaceholderText("Text")
        self.ctl_layout.addWidget(self.txt)
        self.btn_process.clicked.connect(lambda: self.run_worker(PDFEngine.add_watermark, self.get_files()[0], QFileDialog.getSaveFileName(self, "S", "w.pdf", "PDF")[0], self.txt.text()) if self.get_files() else None)

class PageNumPage(BaseToolPage):
    def __init__(self):
        super().__init__("Page Numbers", "Add page X of Y.", "Apply")
        self.combo = QComboBox()
        self.combo.addItems(["bottom-center", "bottom-right", "top-right"])
        self.ctl_layout.addWidget(self.combo)
        self.btn_process.clicked.connect(lambda: self.run_worker(PDFEngine.add_page_numbers, self.get_files()[0], QFileDialog.getSaveFileName(self, "S", "n.pdf", "PDF")[0], self.combo.currentText()) if self.get_files() else None)

class MetadataPage(BaseToolPage):
    def __init__(self):
        super().__init__("Edit Metadata", "Select a file below to view and edit its properties.", "Save Metadata")
        
        # 1. Enable selection
        self.file_list.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.file_list.itemClicked.connect(self.load_meta)
        
        # 2. Create the Form Layout (Hidden by default until initialized)
        self.form_container = QWidget()
        # Use Grid Layout for cleaner alignment (Label | Input)
        form_layout = QGridLayout() 
        form_layout.setContentsMargins(10, 10, 10, 10)
        form_layout.setSpacing(10)
        
        self.inputs = {}
        # PDF Metadata keys usually start with /
        fields = [
            ("Title", "/Title"),
            ("Author", "/Author"),
            ("Subject", "/Subject"),
            ("Producer", "/Producer"),
            ("Creator", "/Creator")
        ]
        
        for row, (label_text, key) in enumerate(fields):
            lbl = QLabel(label_text + ":")
            lbl.setStyleSheet("color: #a6adc8; font-weight: bold;")
            
            inp = QLineEdit()
            inp.setPlaceholderText(f"Enter {label_text}...")
            
            form_layout.addWidget(lbl, row, 0)
            form_layout.addWidget(inp, row, 1)
            
            self.inputs[key] = inp
        
        self.form_container.setLayout(form_layout)
        
        # 3. Insert the form into the Main Layout
        # The BaseToolPage layout order is: 
        # 0:Header, 1:Desc, 2:Controls, 3:FileList, 4:ActionButtons
        # We want to insert the form AFTER the file list (index 4)
        self.layout().insertWidget(4, self.form_container)
        
        self.btn_process.clicked.connect(self.action)

    def load_meta(self, item):
        # 1. Get File Path
        path = item.data(Qt.ItemDataRole.UserRole)
        
        # 2. Clear previous inputs
        for inp in self.inputs.values():
            inp.clear()
            
        # 3. Load new data
        try:
            meta = PDFEngine.get_metadata(path)
            if meta:
                for key, inp in self.inputs.items():
                    # PDF metadata can be None or generic objects, ensure string
                    val = meta.get(key, "")
                    if val:
                        inp.setText(str(val))
        except Exception as e:
            print(f"Error reading metadata: {e}")

    def action(self):
        files = self.get_files()
        if not files: return
        
        # Collect data from inputs
        new_meta = {key: inp.text() for key, inp in self.inputs.items() if inp.text()}
        
        save_path, _ = QFileDialog.getSaveFileName(self, "Save PDF", "meta_edited.pdf", "PDF (*.pdf)")
        if save_path: 
            self.run_worker(PDFEngine.update_metadata, files[0], save_path, new_meta)

class OrganizePage(BaseToolPage):
    def __init__(self):
        super().__init__("Visual Organizer", "Reorder pages.", "Save", use_grid=True)
        self.btn_upload.setText("📂 Load PDF")
        self.btn_upload.disconnect()
        self.btn_upload.clicked.connect(self.load)
        self.btn_process.clicked.connect(self.save)
        self.curr = None
        self.temp = None
    def load(self):
        path, _ = QFileDialog.getOpenFileName(self, "PDF", "", "*.pdf")
        if path:
            self.curr = path
            self.file_list.clear()
            if self.temp: shutil.rmtree(self.temp)
            self.temp = tempfile.mkdtemp()
            self.run_worker(PDFEngine.pdf_to_images, path, self.temp, 40, "jpeg", success_callback=self.fill)
    def fill(self, imgs):
        for i, p in enumerate(imgs):
            item = QListWidgetItem(f"{i+1}")
            item.setData(Qt.ItemDataRole.UserRole, i)
            item.setData(Qt.ItemDataRole.UserRole + 1, 0)
            item.setIcon(QIcon(p))
            self.file_list.addItem(item)
    def save(self):
        if not self.curr: return
        data = [{'original_index': self.file_list.item(i).data(Qt.ItemDataRole.UserRole), 'rotation': self.file_list.item(i).data(Qt.ItemDataRole.UserRole+1)} for i in range(self.file_list.count())]
        save, _ = QFileDialog.getSaveFileName(self, "Save", "org.pdf", "PDF")
        if save: self.run_worker(PDFEngine.reorder_save_pdf, self.curr, save, data)

class SplitPage(BaseToolPage):
    def __init__(self):
        super().__init__("Split PDF", "Split or Extract.")
        self.file_list.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.btn_process.clicked.connect(lambda: self.run_worker(PDFEngine.split_pdf, self.get_files()[0], QFileDialog.getExistingDirectory(self, "Folder"), "all", None) if self.get_files() else None)

class CompressPage(BaseToolPage):
    def __init__(self):
        super().__init__("Compress PDF", "Reduce size.")
        self.combo = QComboBox()
        self.combo.addItems(["Low", "Medium", "Extreme"])
        self.ctl_layout.addWidget(self.combo)
        self.btn_process.clicked.connect(lambda: self.run_worker(PDFEngine.compress_pdf, self.get_files()[0], QFileDialog.getSaveFileName(self, "S", "c.pdf", "PDF")[0], ["low","medium","extreme"][self.combo.currentIndex()]) if self.get_files() else None)

class ProtectPage(BaseToolPage):
    def __init__(self):
        super().__init__("Protect PDF", "Encrypt.")
        self.btn_process.clicked.connect(self.act)
    def act(self):
        files = self.get_files()
        if not files: return
        pwd, ok = QInputDialog.getText(self, "Pwd", "Password:", QLineEdit.EchoMode.Password)
        if ok and pwd: self.run_worker(PDFEngine.protect_pdf, files[0], QFileDialog.getSaveFileName(self, "S", "p.pdf", "PDF")[0], pwd)

class OpenProtectedPage(BaseToolPage):
    def __init__(self):
        super().__init__("Unlock PDF", "View Secure PDF.")
        self.file_list.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.btn_process.clicked.connect(self.act)
    def act(self):
        files = self.get_files()
        if not files: return
        pwd, ok = QInputDialog.getText(self, "Pwd", "Password:", QLineEdit.EchoMode.Password)
        if ok and pwd: 
            fd, tmp = tempfile.mkstemp(suffix=".pdf")
            os.close(fd)
            self.run_worker(PDFEngine.unlock_pdf, files[0], tmp, pwd, success_callback=lambda p: webbrowser.open(p))

class MergePage(BaseToolPage):
    def __init__(self):
        super().__init__("Merge PDF", "Combine multiple PDFs. Reorder them using the buttons or drag and drop.")
        
        # 1. Enable Single Selection for easier reordering
        self.file_list.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)

        # 2. Add Reorder Buttons
        self.btn_up = QPushButton("⬆ Move Up")
        self.btn_up.setProperty("class", "upload-btn")
        self.btn_up.clicked.connect(self.move_up)
        
        self.btn_down = QPushButton("⬇ Move Down")
        self.btn_down.setProperty("class", "upload-btn")
        self.btn_down.clicked.connect(self.move_down)
        
        # Insert buttons into the control layout (after Add/Clear buttons)
        self.ctl_layout.insertWidget(2, self.btn_up)
        self.ctl_layout.insertWidget(3, self.btn_down)
        
        self.btn_process.clicked.connect(self.action)

    def move_up(self):
        """Moves the selected file up one position."""
        row = self.file_list.currentRow()
        if row > 0:
            item = self.file_list.takeItem(row)
            self.file_list.insertItem(row - 1, item)
            self.file_list.setCurrentRow(row - 1)

    def move_down(self):
        """Moves the selected file down one position."""
        row = self.file_list.currentRow()
        if row >= 0 and row < self.file_list.count() - 1:
            item = self.file_list.takeItem(row)
            self.file_list.insertItem(row + 1, item)
            self.file_list.setCurrentRow(row + 1)

    def action(self):
        # This gets the files in the EXACT order shown in the list
        files = self.get_files()
        
        if len(files) < 2: 
            return QMessageBox.warning(self, "Info", "Please select at least 2 files to merge.")
            
        save_path, _ = QFileDialog.getSaveFileName(self, "Save Merged PDF", "merged.pdf", "PDF (*.pdf)")
        if save_path: 
            self.run_worker(PDFEngine.merge_pdfs, files, save_path)
class PdfToImgPage(BaseToolPage):
    def __init__(self): super().__init__("PDF to JPG", "Extract.", "Extract"); self.btn_process.clicked.connect(lambda: self.run_worker(PDFEngine.pdf_to_images, self.get_files()[0], QFileDialog.getExistingDirectory(self, "Folder")) if self.get_files() else None)
class PdfToWordPage(BaseToolPage):
    def __init__(self): super().__init__("PDF to Word", "Convert.", "Convert"); self.btn_process.clicked.connect(lambda: self.run_worker(PDFEngine.pdf_to_word, self.get_files()[0], QFileDialog.getSaveFileName(self, "S","c.docx","Word")[0]) if self.get_files() else None)
class WordToPdfPage(BaseToolPage):
    def __init__(self): super().__init__("Word to PDF", "Convert.", "Convert", ('.docx',)); self.btn_process.clicked.connect(lambda: self.run_worker(PDFEngine.word_to_pdf, self.get_files()[0], QFileDialog.getSaveFileName(self, "S","c.pdf","PDF")[0]) if self.get_files() else None)
class PptxToPdfPage(BaseToolPage):
    def __init__(self): super().__init__("PPT to PDF", "Convert.", "Convert", ('.pptx',)); self.btn_process.clicked.connect(lambda: self.run_worker(PDFEngine.pptx_to_pdf, self.get_files()[0], QFileDialog.getSaveFileName(self, "S","c.pdf","PDF")[0]) if self.get_files() else None)
class PdfToPptxPage(BaseToolPage):
    def __init__(self): super().__init__("PDF to PPT", "Convert.", "Convert"); self.btn_process.clicked.connect(lambda: self.run_worker(PDFEngine.pdf_to_pptx, self.get_files()[0], QFileDialog.getSaveFileName(self, "S","c.pptx","PPT")[0]) if self.get_files() else None)

# --- MAIN WINDOW ---
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Local PDF Pro")
        self.setWindowIcon(QIcon("pdf.ico"))
        self.resize(1200, 850)
        self.is_dark = True
        self.setStyleSheet(DARK_THEME)
        
        central = QWidget()
        self.setCentralWidget(central)
        layout = QHBoxLayout(central)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # Sidebar
        self.sidebar = QFrame()
        self.sidebar.setObjectName("Sidebar")
        self.sidebar.setFixedWidth(280)
        
        sidebar_layout = QVBoxLayout(self.sidebar)
        sidebar_layout.setContentsMargins(10, 20, 10, 20)
        sidebar_layout.setSpacing(5)

        title_box = QHBoxLayout()
        title_lbl = QLabel("  PDF TOOLKIT")
        title_lbl.setStyleSheet("font-size: 18px; font-weight: bold; color: #89b4fa;")
        
        self.toggle_btn = QPushButton("🌙")
        self.toggle_btn.setFixedSize(30, 30)
        self.toggle_btn.clicked.connect(self.toggle_theme)
        
        title_box.addWidget(title_lbl)
        title_box.addStretch()
        title_box.addWidget(self.toggle_btn)
        sidebar_layout.addLayout(title_box)
        sidebar_layout.addSpacing(20)

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
        
        # Add Home Button to Sidebar
        btn_home = SidebarBtn("🏠 Home / Dashboard")
        btn_home.clicked.connect(lambda: self.stack.setCurrentIndex(0))
        self.nav_layout.addWidget(btn_home)
        self.nav_layout.addSpacing(10)

        self.stack = QStackedWidget()
        self.btns = []
        
        # 0. Dashboard
        self.stack.addWidget(DashboardPage(self.go_to_tool))

        # 1-4. Most Used
        self.add_section("MOST USED")
        self.add_nav("Merge PDF", MergePage())     # 1
        self.add_nav("Visual Organize", OrganizePage()) # 2
        self.add_nav("Split PDF", SplitPage())     # 3
        self.add_nav("Compress PDF", CompressPage()) # 4
        
        # 5-7. To PDF
        self.add_section("CONVERT TO PDF")
        self.add_nav("Images to PDF", ImgToPdfPage()) # 5
        self.add_nav("Word to PDF", WordToPdfPage()) # 6
        self.add_nav("PPT to PDF", PptxToPdfPage()) # 7
        
        # 8-10. From PDF
        self.add_section("CONVERT FROM PDF")
        self.add_nav("PDF to JPG", PdfToImgPage()) # 8
        self.add_nav("PDF to Word", PdfToWordPage()) # 9
        self.add_nav("PDF to PPT", PdfToPptxPage()) # 10
        
        # 11-12. Security
        self.add_section("SECURITY")
        self.add_nav("Protect PDF", ProtectPage()) # 11
        self.add_nav("Unlock PDF", OpenProtectedPage()) # 12

        # 13-16. Pro (MOVED TO BOTTOM)
        self.add_section("PRO FEATURES")
        self.add_nav("OCR Searchable", OCRPage())  # 13
        self.add_nav("Watermark", WatermarkPage()) # 14
        self.add_nav("Page Numbers", PageNumPage()) # 15
        self.add_nav("Edit Metadata", MetadataPage()) # 16
        self.add_nav("HTML to PDF", HtmlToPdfPage()) # 17

        self.nav_layout.addStretch()        
        layout.addWidget(self.sidebar)
        layout.addWidget(self.stack)

    def toggle_theme(self):
        self.is_dark = not self.is_dark
        self.setStyleSheet(DARK_THEME if self.is_dark else LIGHT_THEME)
        self.toggle_btn.setText("🌙" if self.is_dark else "☀️")

    def add_section(self, text):
        lbl = QLabel(text)
        lbl.setStyleSheet("font-weight: bold; font-size: 11px; margin-top: 15px; margin-left: 10px; opacity: 0.7;")
        self.nav_layout.addWidget(lbl)

    def add_nav(self, name, widget):
        btn = SidebarBtn(name)
        idx = self.stack.addWidget(widget)
        btn.clicked.connect(lambda: self.switch_view(idx, btn))
        btn.fileDropped.connect(lambda f: self.open_tool_with_file(idx, btn, f))
        self.nav_layout.addWidget(btn)
        self.btns.append(btn)

    def switch_view(self, idx, active_btn):
        self.stack.setCurrentIndex(idx)
        # Reset all nav buttons check state
        for b in self.btns: b.setChecked(False)
        active_btn.setChecked(True)

    def open_tool_with_file(self, idx, btn, file_path):
        self.switch_view(idx, btn)
        widget = self.stack.widget(idx)
        if hasattr(widget, 'file_list'):
            widget.file_list.addItems([file_path])
    
    def go_to_tool(self, idx):
        # idx is the stack index. The button index in self.btns is (idx - 1) because stack 0 is dashboard
        if 0 < idx <= len(self.btns):
            # Programmatically click the sidebar button to ensure visual sync
            self.btns[idx-1].click()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
