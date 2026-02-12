import sys
import os
import threading
import tempfile
import shutil
import webbrowser
import json
import qtawesome as qta  # Requires: pip install qtawesome
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QLabel, QListWidget, 
                             QStackedWidget, QFileDialog, QMessageBox, QFrame,
                             QListWidgetItem, QAbstractItemView, QInputDialog, 
                             QLineEdit, QScrollArea, QComboBox, QRadioButton,
                             QButtonGroup, QMenu, QDialog, QGridLayout, QCheckBox, 
                             QSizePolicy, QTextEdit, QToolButton)
from PyQt6.QtCore import Qt, pyqtSignal, QObject, QSize, QSettings
from PyQt6.QtGui import QDragEnterEvent, QDropEvent, QIcon, QFont, QPixmap, QKeyEvent, QAction, QColor
from pdf2image import convert_from_path

# Import backend engine
from pdf_engine import PDFEngine

# --- UPDATED THEMES (Fixed Border Cutoff) ---
DARK_THEME = """
QMainWindow, QWidget#CentralWidget { background-color: #1e1e2e; }
QFrame#Sidebar { background-color: #181825; border-right: 1px solid #313244; }
QScrollArea { background: transparent; border: none; }
QWidget#DashboardContent { background: transparent; }
QWidget#SidebarContent { background: transparent; }

QLabel { color: #cdd6f4; font-family: 'Segoe UI', sans-serif; }
QLabel#DashTitle { color: #89b4fa; font-weight: bold; }
QLabel#DashDesc { color: #a6adc8; }

/* Dashboard Cards (Dark) */
QFrame.dash-card { 
    background-color: #313244; border: 1px solid #45475a; border-radius: 16px; 
}
QFrame.dash-card:hover { 
    background-color: #45475a; border-color: #89b4fa; 
    /* Removed negative margin to fix top border clipping */
}
QLabel#CardTitle { color: #cdd6f4; font-weight: bold; font-size: 15px; }
QLabel#CardDesc { color: #a6adc8; font-size: 12px; }

/* List Widget & Sidebar */
QListWidget { background-color: #181825; border: 2px dashed #45475a; border-radius: 12px; color: #cdd6f4; padding: 10px; }
QListWidget::item { padding: 10px; border-radius: 8px; background-color: #313244; }
QListWidget::item:selected { background-color: #cba6f7; color: #1e1e2e; }

/* Buttons */
QPushButton.nav-btn { background-color: transparent; color: #a6adc8; border: none; text-align: left; padding: 12px 20px; border-radius: 8px; }
QPushButton.nav-btn:checked { background-color: #313244; color: #89b4fa; border-left: 4px solid #89b4fa; }
QPushButton.nav-btn:hover { background-color: #313244; color: white; }
QPushButton.action-btn { background-color: #89b4fa; color: #1e1e2e; border-radius: 8px; padding: 12px; font-weight: bold; border: none; }
QPushButton.upload-btn { background-color: #313244; color: #cdd6f4; border: 1px solid #45475a; border-radius: 8px; padding: 8px 16px; }
QLineEdit, QComboBox, QTextEdit { background-color: #313244; border: 1px solid #45475a; color: white; padding: 10px; border-radius: 8px; }
"""

LIGHT_THEME = """
QMainWindow, QWidget#CentralWidget { background-color: #eff1f5; }
QFrame#Sidebar { background-color: #e6e9ef; border-right: 1px solid #bcc0cc; }
QScrollArea { background: transparent; border: none; }
QWidget#DashboardContent { background: transparent; }
QWidget#SidebarContent { background: transparent; }

QLabel { color: #4c4f69; font-family: 'Segoe UI', sans-serif; }
QLabel#DashTitle { color: #1e66f5; font-weight: bold; }
QLabel#DashDesc { color: #6c6f85; }

/* Dashboard Cards (Light) */
QFrame.dash-card { 
    background-color: white; border: 1px solid #bcc0cc; border-radius: 16px; 
}
QFrame.dash-card:hover { 
    background-color: #e6e9ef; border-color: #1e66f5; 
    /* Removed negative margin to fix top border clipping */
}
QLabel#CardTitle { color: #4c4f69; font-weight: bold; font-size: 15px; }
QLabel#CardDesc { color: #6c6f85; font-size: 12px; }

/* List Widget & Sidebar */
QListWidget { background-color: white; border: 2px dashed #bcc0cc; border-radius: 12px; color: #4c4f69; padding: 10px; }
QListWidget::item { padding: 10px; border-radius: 8px; background-color: #e6e9ef; }
QListWidget::item:selected { background-color: #ea76cb; color: white; }

/* Buttons */
QPushButton.nav-btn { background-color: transparent; color: #5c5f77; border: none; text-align: left; padding: 12px 20px; border-radius: 8px; }
QPushButton.nav-btn:checked { background-color: #dce0e8; color: #1e66f5; border-left: 4px solid #1e66f5; }
QPushButton.nav-btn:hover { background-color: #dce0e8; color: #4c4f69; }
QPushButton.action-btn { background-color: #1e66f5; color: white; border-radius: 8px; padding: 12px; font-weight: bold; border: none; }
QPushButton.upload-btn { background-color: white; color: #4c4f69; border: 1px solid #bcc0cc; border-radius: 8px; padding: 8px 16px; }
QLineEdit, QComboBox, QTextEdit { background-color: white; border: 1px solid #bcc0cc; color: #4c4f69; padding: 10px; border-radius: 8px; }
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

class DashboardCard(QFrame):
    clicked = pyqtSignal()

    def __init__(self, title, desc, icon_name, color="#89b4fa"):
        super().__init__()
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setProperty("class", "dash-card") 
        self.setFrameShape(QFrame.Shape.StyledPanel)
        self.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        
        # Layout
        layout = QVBoxLayout(self)
        layout.setContentsMargins(15, 20, 15, 20)
        layout.setSpacing(5)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        # Icon
        self.icon_lbl = QLabel()
        self.icon_lbl.setPixmap(qta.icon(icon_name, color=color).pixmap(QSize(40, 40)))
        self.icon_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.icon_lbl.setStyleSheet("background: transparent; border: none;")
        
        # Title
        self.title_lbl = QLabel(title)
        self.title_lbl.setObjectName("CardTitle")
        self.title_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        # Description
        self.desc_lbl = QLabel(desc)
        self.desc_lbl.setObjectName("CardDesc")
        self.desc_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.desc_lbl.setWordWrap(True)
        self.desc_lbl.setFixedHeight(45)

        layout.addWidget(self.icon_lbl)
        layout.addWidget(self.title_lbl)
        layout.addWidget(self.desc_lbl)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.clicked.emit()

# --- CUSTOM WIDGETS ---
class SidebarBtn(QPushButton):
    fileDropped = pyqtSignal(str)
    
    def __init__(self, text, icon_name=None):
        super().__init__(text)
        if icon_name:
            self.setIcon(qta.icon(icon_name, color="#a6adc8"))
            self.setIconSize(QSize(20, 20))
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
            ext = os.path.splitext(file_path)[1].lower()
            if ext in ['.pdf']: icon_key = 'fa5s.file-pdf'
            elif ext in ['.docx', '.doc']: icon_key = 'fa5s.file-word'
            elif ext in ['.png', '.jpg']: icon_key = 'fa5s.file-image'
            else: icon_key = 'fa5s.file'
            
            item.setIcon(qta.icon(icon_key, color="#89b4fa"))
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
        
        rotate_cw = QAction(qta.icon('fa5s.redo', color='white'), "Rotate Right (90°)", self)
        rotate_ccw = QAction(qta.icon('fa5s.undo', color='white'), "Rotate Left (-90°)", self)
        delete_page = QAction(qta.icon('fa5s.trash', color='#ff6b6b'), "Delete Page", self)
        
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
        self.resize(1000, 650)
        self.image_path = image_path
        self.scale_factor = 1.0
        
        main_layout = QVBoxLayout(self)
        
        lbl = QLabel("Drag the RED circles to the 4 corners of the document.")
        lbl.setStyleSheet("color: #a6adc8; font-weight: bold; font-size: 14px;")
        lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(lbl)

        self.scene = QGraphicsScene()
        self.view = QGraphicsView(self.scene)
        self.view.setRenderHint(QPainter.RenderHint.Antialiasing)
        self.view.setStyleSheet("background: #181825; border: 2px solid #45475a; border-radius: 8px;")
        
        self.pixmap = QPixmap(image_path)
        view_w, view_h = 950, 500
        img_w, img_h = self.pixmap.width(), self.pixmap.height()
        self.scale_factor = min(view_w / img_w, view_h / img_h)
        
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

        w, h = scaled_pix.width(), scaled_pix.height()
        pad = 50
        self.handles = []
        points = [(pad, pad), (w-pad, pad), (w-pad, h-pad), (pad, h-pad)]
        
        for x, y in points:
            handle = self.create_handle(x, y)
            self.handles.append(handle)
            self.scene.addItem(handle)

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
        size = 20
        ellipse = QGraphicsEllipseItem(0, 0, size, size)
        ellipse.setPos(x - size/2, y - size/2)
        ellipse.setBrush(QBrush(QColor("#ff4757")))
        ellipse.setPen(QPen(Qt.GlobalColor.white, 2))
        ellipse.setFlag(QGraphicsEllipseItem.GraphicsItemFlag.ItemIsMovable)
        ellipse.setFlag(QGraphicsEllipseItem.GraphicsItemFlag.ItemSendsGeometryChanges)
        return ellipse

    def on_apply(self):
        final_corners = []
        for h in self.handles:
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
        head_lbl.setFont(QFont("Segoe UI", 26, QFont.Weight.Bold))
        layout.addWidget(head_lbl)
        
        desc_lbl = QLabel(desc)
        desc_lbl.setStyleSheet("color: #a6adc8; font-size: 14px;")
        desc_lbl.setWordWrap(True)
        layout.addWidget(desc_lbl)
        
        self.ctl_layout = QHBoxLayout()
        self.btn_upload = QPushButton(" Select Files")
        self.btn_upload.setIcon(qta.icon('fa5s.folder-open', color="#cdd6f4"))
        self.btn_upload.setProperty("class", "upload-btn")
        self.btn_upload.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_upload.clicked.connect(self.open_file_dialog)
        
        self.btn_clear = QPushButton(" Clear")
        self.btn_clear.setIcon(qta.icon('fa5s.trash-alt', color="#cdd6f4"))
        self.btn_clear.setProperty("class", "upload-btn")
        self.btn_clear.clicked.connect(self.clear_list)
        
        self.ctl_layout.addWidget(self.btn_upload)
        self.ctl_layout.addWidget(self.btn_clear)
        self.ctl_layout.addStretch()
        layout.addLayout(self.ctl_layout)

        if use_grid: self.file_list = OrganizerGrid()
        else: self.file_list = FileDropList(allowed_exts)
        layout.addWidget(self.file_list)
        
        # We store bot_layout as self.bot_layout so subclasses can access it
        self.bot_layout = QHBoxLayout()
        self.btn_process = QPushButton(btn_text)
        self.btn_process.setIcon(qta.icon('fa5s.play-circle', color="#1e1e2e"))
        self.btn_process.setProperty("class", "action-btn")
        self.btn_process.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_process.setFixedWidth(200)
        
        self.lbl_status = QLabel("")
        self.lbl_status.setStyleSheet("color: #89b4fa; font-weight: bold; margin-left: 15px;")
        
        self.bot_layout.addWidget(self.btn_process)
        self.bot_layout.addWidget(self.lbl_status)
        self.bot_layout.addStretch()
        layout.addLayout(self.bot_layout)
        
        self.setLayout(layout)

    def clear_list(self): self.file_list.clear()

    def open_file_dialog(self):
        ext_filter = " ".join([f"*{e}" for e in self.allowed_exts])
        files, _ = QFileDialog.getOpenFileNames(self, "Select Files", "", f"Files ({ext_filter})")
        if files: self.file_list.addItems(files)

    def get_files(self):
        """Returns the FULL PATHS of files in the list."""
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

class DashboardPage(QWidget):
    def __init__(self, nav_callback):
        super().__init__()
        self.nav_callback = nav_callback
        self.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        
        # Scroll Area Setup
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        
        # Content Widget
        content_widget = QWidget()
        content_widget.setObjectName("DashboardContent")
        content_widget.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        
        layout = QVBoxLayout(content_widget)
        layout.setContentsMargins(50, 50, 50, 50)
        layout.setSpacing(30)
        
        # Header Section
        header_box = QVBoxLayout()
        header_box.setSpacing(10)
        
        title = QLabel("Local PDF Pro")
        title.setObjectName("DashTitle")
        title.setFont(QFont("Segoe UI", 48, QFont.Weight.Bold))
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        desc = QLabel("Your secure, offline, all-in-one PDF toolkit.\nSelect a tool below to get started.")
        desc.setObjectName("DashDesc")
        desc.setAlignment(Qt.AlignmentFlag.AlignCenter)
        desc.setFont(QFont("Segoe UI", 16))
        
        header_box.addWidget(title)
        header_box.addWidget(desc)
        layout.addLayout(header_box)
        layout.addSpacing(20)

        # Tools Grid
        grid = QGridLayout()
        grid.setSpacing(20)
        
        # UPDATED ORDER TO MATCH MAIN WINDOW sidebar logic exactly
        # 1-based index maps to the stack index
        tools = [
            ("Merge PDF", 1, "fa5s.layer-group", "Combine multiple PDFs into one document."),
            ("Visual Organize", 2, "fa5s.th", "Reorder, rotate, or remove pages visually."),
            ("Split PDF", 3, "fa5s.cut", "Split a PDF into separate files or pages."),
            ("Compress PDF", 4, "fa5s.compress-arrows-alt", "Reduce file size while keeping quality."),
            ("Page Numbers", 5, "fa5s.list-ol", "Add page numbering to your document."),
            
            ("Images to PDF", 6, "fa5s.images", "Convert JPG, PNG images to PDF."),
            ("Word to PDF", 7, "fa5s.file-word", "Convert DOCX documents to PDF."),
            ("PPT to PDF", 8, "fa5s.file-powerpoint", "Convert PowerPoint presentations to PDF."),
            ("HTML to PDF", 9, "fa5s.code", "Convert HTML code or files to PDF."),

            ("PDF to JPG", 10, "fa5s.file-image", "Extract pages as high-quality images."),
            ("PDF to Word", 11, "fa5s.file-alt", "Convert PDF to editable Word docs."),
            ("PDF to PPT", 12, "fa5s.file-powerpoint", "Convert PDF to PowerPoint slides."),
            
            ("Protect PDF", 13, "fa5s.lock", "Encrypt PDF with a password."),
            ("Unlock PDF", 14, "fa5s.unlock", "Remove passwords from PDFs."),
            
            ("OCR Searchable", 15, "fa5s.search", "Make scanned documents text-searchable."),
            ("Watermark", 16, "fa5s.stamp", "Add text or image stamps to pages."),
            ("Edit Metadata", 17, "fa5s.info-circle", "Modify title, author, and creator info."),
            ("Extract Images", 18, "fa5s.images", "Rip all embedded images from a file."),
            ("Flatten PDF", 19, "fa5s.clone", "Make fillable forms un-editable."),
            ("Grayscale PDF", 20, "fa5s.tint-slash", "Convert colors to black & white.")
        ]

        row, col = 0, 0
        for name, idx, icon, desc_text in tools:
            card = DashboardCard(name, desc_text, icon)
            # Use closure to capture loop variable 'idx'
            card.clicked.connect(lambda i=idx: self.nav_callback(i))
            
            grid.addWidget(card, row, col)
            
            col += 1
            if col > 3: # 4 columns wide
                col = 0
                row += 1
        
        layout.addLayout(grid)
        layout.addStretch()
        
        scroll.setWidget(content_widget)
        
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0,0,0,0)
        main_layout.addWidget(scroll)
        
# --- PAGE CLASSES ---

class HtmlToPdfPage(BaseToolPage):
    def __init__(self):
        super().__init__("HTML to PDF (Pro)", "Render modern HTML/CSS with emoji support.", "Convert to PDF")
        
        # Hide file upload controls
        self.file_list.setParent(None)
        self.btn_upload.setParent(None)
        self.btn_clear.setParent(None)
        
        # Text Editor
        self.text_area = QTextEdit()
        self.text_area.setPlaceholderText("Paste HTML here...\nSupports: Tailwind, Flexbox, Emojis (🚀)...")
        self.text_area.setAcceptRichText(False)
        self.text_area.setStyleSheet("""
            QTextEdit {
                background-color: #181825; color: #d4d4d4; 
                border: 2px solid #313244; border-radius: 8px; 
                padding: 15px; font-family: 'Consolas', monospace; font-size: 14px;
            }
        """)
        # Insert editor into main layout (index 3 is where file_list was)
        self.layout().insertWidget(3, self.text_area)

        # FIX: Add Preview Button correctly to self.bot_layout
        self.btn_preview = QPushButton(" Preview")
        self.btn_preview.setIcon(qta.icon('fa5s.eye', color="#cdd6f4"))
        self.btn_preview.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_preview.setFixedSize(140, 45) # Match height of action button
        # Style it like a secondary action button
        self.btn_preview.setStyleSheet("""
            QPushButton { background-color: #313244; color: white; border: 1px solid #45475a; border-radius: 8px; font-weight: bold; }
            QPushButton:hover { background-color: #45475a; }
        """)
        
        # Insert before the Process button
        self.bot_layout.insertWidget(0, self.btn_preview)

        self.btn_process.clicked.connect(self.action_save)
        self.btn_preview.clicked.connect(self.action_preview)

    def action_preview(self):
        html_content = self.text_area.toPlainText()
        if not html_content.strip(): return
        temp_pdf = os.path.join(tempfile.gettempdir(), "preview_render.pdf")
        try:
            PDFEngine.html_to_pdf(html_content, temp_pdf)
            webbrowser.open(temp_pdf)
        except Exception as e:
            QMessageBox.critical(self, "Preview Error", str(e))

    def action_save(self):
        html_content = self.text_area.toPlainText()
        if not html_content.strip(): return QMessageBox.warning(self, "Input Required", "Paste HTML first.")
        save_path, _ = QFileDialog.getSaveFileName(self, "Save PDF", "render.pdf", "PDF (*.pdf)")
        if save_path:
            self.run_worker(PDFEngine.html_to_pdf, html_content, save_path)

class ExtractImagesPage(BaseToolPage):
    def __init__(self):
        super().__init__("Extract Images", "Save all images inside the PDF as separate files.", "Extract Images")
        self.btn_process.clicked.connect(self.action)
    def action(self):
        files = self.get_files() # FIX: Use get_files() to get full paths
        if not files: return
        dest = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if dest: self.run_worker(PDFEngine.extract_images, files[0], dest)

class FlattenPdfPage(BaseToolPage):
    def __init__(self):
        super().__init__("Flatten PDF", "Lock forms and annotations permanently.", "Flatten PDF")
        self.btn_process.clicked.connect(self.action)
    def action(self):
        files = self.get_files() # FIX
        if not files: return
        dest, _ = QFileDialog.getSaveFileName(self, "Save File", "flattened.pdf", "PDF (*.pdf)")
        if dest: self.run_worker(PDFEngine.flatten_pdf, files[0], dest)

class GrayscalePdfPage(BaseToolPage):
    def __init__(self):
        super().__init__("Grayscale PDF", "Convert colorful PDFs to Black & White.", "Convert to B&W")
        self.btn_process.clicked.connect(self.action)
    def action(self):
        files = self.get_files() # FIX
        if not files: return
        dest, _ = QFileDialog.getSaveFileName(self, "Save File", "grayscale.pdf", "PDF (*.pdf)")
        if dest: self.run_worker(PDFEngine.convert_grayscale, files[0], dest)

class ImgToPdfPage(BaseToolPage):
    def __init__(self):
        super().__init__("Images to PDF", "Smart Scan or convert standard images.", "Convert", ('.jpg', '.png', '.jpeg'))
        self.btn_smart = QPushButton(" Smart Scan")
        self.btn_smart.setIcon(qta.icon('fa5s.camera', color="#181825"))
        self.btn_smart.setProperty("class", "upload-btn")
        self.btn_smart.setStyleSheet("background-color: #f38ba8; color: #181825; border: none; font-weight: bold;")
        self.btn_smart.clicked.connect(self.smart_scan_dialog)
        self.ctl_layout.insertWidget(1, self.btn_smart)
        self.btn_process.clicked.connect(self.action)

    def smart_scan_dialog(self):
        file, _ = QFileDialog.getOpenFileName(self, "Select Image", "", "Images (*.jpg *.jpeg *.png)")
        if not file: return
        dlg = DraggableScanDialog(file)
        if dlg.exec():
            try:
                processed_img = PDFEngine.manual_scan_warp(file, dlg.final_corners)
                fd, path = tempfile.mkstemp(suffix=".jpg")
                os.close(fd)
                processed_img.save(path)
                self.file_list.addItems([path])
            except Exception as e:
                QMessageBox.warning(self, "Error", str(e))

    def action(self):
        files = self.get_files()
        if not files: return
        save_path, _ = QFileDialog.getSaveFileName(self, "Save PDF", "images.pdf", "PDF (*.pdf)")
        if save_path: self.run_worker(PDFEngine.images_to_pdf, files, save_path)

class OCRPage(BaseToolPage):
    def __init__(self):
        super().__init__("OCR (Searchable PDF)", "Make scanned documents searchable.", "Run OCR")
        self.btn_process.clicked.connect(self.action)
    
    def action(self):
        files = self.get_files()
        if not files: return
        save_path, _ = QFileDialog.getSaveFileName(self, "Save", "ocr.pdf", "PDF (*.pdf)")
        if save_path: self.run_worker(PDFEngine.ocr_pdf, files[0], save_path)

class WatermarkPage(BaseToolPage):
    def __init__(self):
        super().__init__("Watermark", "Add text overlay.", "Apply")
        self.txt = QLineEdit()
        self.txt.setPlaceholderText("Text")
        self.ctl_layout.addWidget(self.txt)
        self.btn_process.clicked.connect(self.action)

    def action(self):
        files = self.get_files()
        if not files: return
        save_path, _ = QFileDialog.getSaveFileName(self, "Save", "w.pdf", "PDF (*.pdf)")
        if save_path: self.run_worker(PDFEngine.add_watermark, files[0], save_path, self.txt.text())

class PageNumPage(BaseToolPage):
    def __init__(self):
        super().__init__("Page Numbers", "Add page X of Y.", "Apply")
        self.combo = QComboBox()
        self.combo.addItems(["bottom-center", "bottom-right", "top-right"])
        self.ctl_layout.addWidget(self.combo)
        self.btn_process.clicked.connect(self.action)

    def action(self):
        files = self.get_files()
        if not files: return
        save_path, _ = QFileDialog.getSaveFileName(self, "Save", "n.pdf", "PDF (*.pdf)")
        if save_path: self.run_worker(PDFEngine.add_page_numbers, files[0], save_path, self.combo.currentText())

class MetadataPage(BaseToolPage):
    def __init__(self):
        super().__init__("Edit Metadata", "Select a file below to view and edit its properties.", "Save Metadata")
        self.file_list.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.file_list.itemClicked.connect(self.load_meta)
        
        self.form_container = QWidget()
        form_layout = QGridLayout() 
        form_layout.setContentsMargins(10, 10, 10, 10)
        self.inputs = {}
        fields = [("Title", "/Title"),("Author", "/Author"),("Subject", "/Subject"),("Producer", "/Producer"),("Creator", "/Creator")]
        
        for row, (label_text, key) in enumerate(fields):
            lbl = QLabel(label_text + ":")
            lbl.setStyleSheet("color: #a6adc8; font-weight: bold;")
            inp = QLineEdit()
            form_layout.addWidget(lbl, row, 0)
            form_layout.addWidget(inp, row, 1)
            self.inputs[key] = inp
        
        self.form_container.setLayout(form_layout)
        self.layout().insertWidget(4, self.form_container)
        self.btn_process.clicked.connect(self.action)

    def load_meta(self, item):
        path = item.data(Qt.ItemDataRole.UserRole)
        for inp in self.inputs.values(): inp.clear()
        try:
            meta = PDFEngine.get_metadata(path)
            if meta:
                for key, inp in self.inputs.items():
                    val = meta.get(key, "")
                    if val: inp.setText(str(val))
        except: pass

    def action(self):
        files = self.get_files()
        if not files: return
        new_meta = {key: inp.text() for key, inp in self.inputs.items() if inp.text()}
        save_path, _ = QFileDialog.getSaveFileName(self, "Save", "meta.pdf", "PDF")
        if save_path: self.run_worker(PDFEngine.update_metadata, files[0], save_path, new_meta)

class OrganizePage(BaseToolPage):
    def __init__(self):
        super().__init__("Visual Organizer", "Reorder pages.", "Save", use_grid=True)
        self.btn_upload.setText(" Load PDF")
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
        self.btn_process.clicked.connect(self.action)
    
    def action(self):
        files = self.get_files()
        if not files: return
        dest = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if dest: self.run_worker(PDFEngine.split_pdf, files[0], dest, "all", None)

class CompressPage(BaseToolPage):
    def __init__(self):
        super().__init__("Compress PDF", "Reduce size.")
        self.combo = QComboBox()
        self.combo.addItems(["Low", "Medium", "Extreme"])
        self.ctl_layout.addWidget(self.combo)
        self.btn_process.clicked.connect(self.action)

    def action(self):
        files = self.get_files()
        if not files: return
        level = ["low","medium","extreme"][self.combo.currentIndex()]
        save_path, _ = QFileDialog.getSaveFileName(self, "Save", "c.pdf", "PDF (*.pdf)")
        if save_path: self.run_worker(PDFEngine.compress_pdf, files[0], save_path, level)

class ProtectPage(BaseToolPage):
    def __init__(self):
        super().__init__("Protect PDF", "Encrypt.")
        self.btn_process.clicked.connect(self.act)

    def act(self):
        files = self.get_files()
        if not files: return
        pwd, ok = QInputDialog.getText(self, "Pwd", "Password:", QLineEdit.EchoMode.Password)
        if ok and pwd:
             save_path, _ = QFileDialog.getSaveFileName(self, "Save", "p.pdf", "PDF (*.pdf)")
             if save_path: self.run_worker(PDFEngine.protect_pdf, files[0], save_path, pwd)

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
            tmp = tempfile.mktemp(".pdf")
            self.run_worker(PDFEngine.unlock_pdf, files[0], tmp, pwd, success_callback=lambda p: webbrowser.open(p))

class MergePage(BaseToolPage):
    def __init__(self):
        super().__init__("Merge PDF", "Combine multiple PDFs.")
        self.file_list.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        
        self.btn_up = QPushButton(" Up")
        self.btn_up.setIcon(qta.icon('fa5s.arrow-up', color="#cdd6f4"))
        self.btn_up.setProperty("class", "upload-btn")
        self.btn_up.clicked.connect(self.move_up)
        
        self.btn_down = QPushButton(" Down")
        self.btn_down.setIcon(qta.icon('fa5s.arrow-down', color="#cdd6f4"))
        self.btn_down.setProperty("class", "upload-btn")
        self.btn_down.clicked.connect(self.move_down)
        
        self.ctl_layout.insertWidget(2, self.btn_up)
        self.ctl_layout.insertWidget(3, self.btn_down)
        self.btn_process.clicked.connect(self.action)

    def move_up(self):
        row = self.file_list.currentRow()
        if row > 0: 
            self.file_list.insertItem(row - 1, self.file_list.takeItem(row))
            self.file_list.setCurrentRow(row - 1)

    def move_down(self):
        row = self.file_list.currentRow()
        if row >= 0 and row < self.file_list.count() - 1: 
            self.file_list.insertItem(row + 1, self.file_list.takeItem(row))
            self.file_list.setCurrentRow(row + 1)

    def action(self):
        files = self.get_files()
        if len(files) < 2: return QMessageBox.warning(self, "Info", "Select 2+ files.")
        save_path, _ = QFileDialog.getSaveFileName(self, "Save", "merged.pdf", "PDF (*.pdf)")
        if save_path: self.run_worker(PDFEngine.merge_pdfs, files, save_path)

class PdfToImgPage(BaseToolPage):
    def __init__(self): 
        super().__init__("PDF to JPG", "Extract.", "Extract")
        self.btn_process.clicked.connect(self.action)
    
    def action(self):
        files = self.get_files()
        if not files: return
        dest = QFileDialog.getExistingDirectory(self, "Select Folder")
        if dest: self.run_worker(PDFEngine.pdf_to_images, files[0], dest)

class PdfToWordPage(BaseToolPage):
    def __init__(self): 
        super().__init__("PDF to Word", "Convert.", "Convert")
        self.btn_process.clicked.connect(self.action)
        
    def action(self):
        files = self.get_files()
        if not files: return
        save_path, _ = QFileDialog.getSaveFileName(self, "Save", "c.docx", "Word (*.docx)")
        if save_path: self.run_worker(PDFEngine.pdf_to_word, files[0], save_path)

class WordToPdfPage(BaseToolPage):
    def __init__(self): 
        super().__init__("Word to PDF", "Convert.", "Convert", ('.docx',))
        self.btn_process.clicked.connect(self.action)
        
    def action(self):
        files = self.get_files()
        if not files: return
        save_path, _ = QFileDialog.getSaveFileName(self, "Save", "c.pdf", "PDF (*.pdf)")
        if save_path: self.run_worker(PDFEngine.word_to_pdf, files[0], save_path)

class PptxToPdfPage(BaseToolPage):
    def __init__(self): 
        super().__init__("PPT to PDF", "Convert.", "Convert", ('.pptx',))
        self.btn_process.clicked.connect(self.action)
        
    def action(self):
        files = self.get_files()
        if not files: return
        save_path, _ = QFileDialog.getSaveFileName(self, "Save", "c.pdf", "PDF (*.pdf)")
        if save_path: self.run_worker(PDFEngine.pptx_to_pdf, files[0], save_path)

class PdfToPptxPage(BaseToolPage):
    def __init__(self): 
        super().__init__("PDF to PPT", "Convert.", "Convert")
        self.btn_process.clicked.connect(self.action)
        
    def action(self):
        files = self.get_files()
        if not files: return
        save_path, _ = QFileDialog.getSaveFileName(self, "Save", "c.pptx", "PPT (*.pptx)")
        if save_path: self.run_worker(PDFEngine.pdf_to_pptx, files[0], save_path)

# --- MAIN WINDOW ---
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Local PDF Pro")
        self.setWindowIcon(qta.icon('fa5s.file-pdf', color="#89b4fa"))
        self.resize(1280, 850)
        self.is_dark = True
        self.setStyleSheet(DARK_THEME)
        
        central = QWidget()
        central.setObjectName("CentralWidget")
        central.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
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
        
        self.toggle_btn = QPushButton()
        self.toggle_btn.setIcon(qta.icon('fa5s.moon', color="white"))
        self.toggle_btn.setFixedSize(30, 30)
        self.toggle_btn.setStyleSheet("border:none;")
        self.toggle_btn.clicked.connect(self.toggle_theme)
        
        title_box.addWidget(title_lbl)
        title_box.addStretch()
        title_box.addWidget(self.toggle_btn)
        sidebar_layout.addLayout(title_box)
        sidebar_layout.addSpacing(20)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        nav_content = QWidget()
        nav_content.setObjectName("SidebarContent") # <--- ID for styling
        nav_content.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True) # <--- Enable styling
        nav_content.setStyleSheet("background-color: transparent;") # <--- Force transparent
        self.nav_layout = QVBoxLayout(nav_content)
        self.nav_layout.setContentsMargins(0, 0, 0, 0)
        self.nav_layout.setSpacing(5)
        scroll.setWidget(nav_content)
        sidebar_layout.addWidget(scroll)
        
        btn_home = SidebarBtn(" Dashboard", "fa5s.home")
        btn_home.clicked.connect(lambda: self.stack.setCurrentIndex(0))
        self.nav_layout.addWidget(btn_home)
        self.nav_layout.addSpacing(10)

        self.stack = QStackedWidget()
        self.btns = []
        
        self.stack.addWidget(DashboardPage(self.go_to_tool))

        self.add_section("MOST USED")
        self.add_nav("Merge PDF", MergePage(), "fa5s.layer-group")
        self.add_nav("Visual Organize", OrganizePage(), "fa5s.th")
        self.add_nav("Split PDF", SplitPage(), "fa5s.cut")
        self.add_nav("Compress PDF", CompressPage(), "fa5s.compress-arrows-alt")
        self.add_nav("Page Numbers", PageNumPage(), "fa5s.list-ol")

        
        self.add_section("CONVERT TO PDF")
        self.add_nav("Images to PDF", ImgToPdfPage(), "fa5s.images")
        self.add_nav("Word to PDF", WordToPdfPage(), "fa5s.file-word")
        self.add_nav("PPT to PDF", PptxToPdfPage(), "fa5s.file-powerpoint")
        self.add_nav("HTML to PDF", HtmlToPdfPage(), "fa5s.code")

        
        self.add_section("CONVERT FROM PDF")
        self.add_nav("PDF to JPG", PdfToImgPage(), "fa5s.file-image")
        self.add_nav("PDF to Word", PdfToWordPage(), "fa5s.file-alt")
        self.add_nav("PDF to PPT", PdfToPptxPage(), "fa5s.file-powerpoint")
        
        self.add_section("SECURITY")
        self.add_nav("Protect PDF", ProtectPage(), "fa5s.lock")
        self.add_nav("Unlock PDF", OpenProtectedPage(), "fa5s.unlock")

        self.add_section("PRO FEATURES")
        self.add_nav("OCR Searchable", OCRPage(), "fa5s.search")
        self.add_nav("Watermark", WatermarkPage(), "fa5s.stamp")
        self.add_nav("Edit Metadata", MetadataPage(), "fa5s.info-circle")
        self.add_nav("Extract Images", ExtractImagesPage(), "fa5s.images")
        self.add_nav("Flatten PDF", FlattenPdfPage(), "fa5s.clone")
        self.add_nav("Grayscale PDF", GrayscalePdfPage(), "fa5s.tint-slash")

        self.nav_layout.addStretch()        
        layout.addWidget(self.sidebar)
        layout.addWidget(self.stack)

    def toggle_theme(self):
        self.is_dark = not self.is_dark
        self.setStyleSheet(DARK_THEME if self.is_dark else LIGHT_THEME)
        self.toggle_btn.setIcon(qta.icon('fa5s.moon' if self.is_dark else 'fa5s.sun', color="#4c4f69" if not self.is_dark else "white"))

    def add_section(self, text):
        lbl = QLabel(text)
        lbl.setStyleSheet("font-weight: bold; font-size: 11px; margin-top: 15px; margin-left: 10px; opacity: 0.7;")
        self.nav_layout.addWidget(lbl)

    def add_nav(self, name, widget, icon=None):
        btn = SidebarBtn(f" {name}", icon)
        idx = self.stack.addWidget(widget)
        btn.clicked.connect(lambda: self.switch_view(idx, btn))
        btn.fileDropped.connect(lambda f: self.open_tool_with_file(idx, btn, f))
        self.nav_layout.addWidget(btn)
        self.btns.append(btn)

    def switch_view(self, idx, active_btn):
        self.stack.setCurrentIndex(idx)
        for b in self.btns: b.setChecked(False)
        active_btn.setChecked(True)

    def open_tool_with_file(self, idx, btn, file_path):
        self.switch_view(idx, btn)
        widget = self.stack.widget(idx)
        if hasattr(widget, 'file_list'):
            widget.file_list.addItems([file_path])
    
    def go_to_tool(self, idx):
        if 0 < idx <= len(self.btns):
            self.btns[idx-1].click()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())