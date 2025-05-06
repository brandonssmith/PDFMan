import sys
import os
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QTabWidget, QWidget,
    QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QFileDialog, QMessageBox, QScrollArea, QSpinBox,
    QStatusBar, QSizePolicy, QListWidget, QGridLayout,
    QSplitter, QToolBar, QMenu, QDialog, QInputDialog,
    QFontComboBox, QColorDialog, QLineEdit, QCheckBox
)
from PyQt6.QtCore import Qt, QSize, QMimeData, QPoint, QThread, pyqtSignal
from PyQt6.QtGui import QDragEnterEvent, QDropEvent, QAction, QImage, QPixmap, QDrag, QIcon, QPainter, QPen, QFont, QColor, QMovie
from pdf_operations import PDFOperations
import logging
import traceback
from PIL import Image
from PyPDF2 import PdfWriter
import fitz  # PyMuPDF
import easyocr
import numpy as np
import json
from pdf2image import convert_from_path
from commands import (
    RotatePagesCommand,
    DuplicatePagesCommand,
    RemovePagesCommand,
    ReorderPagesCommand
)

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class PDFPreviewLabel(QLabel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        self.setMinimumSize(400, 600)
        self.setStyleSheet("QLabel { background-color: #f0f0f0; border: 1px solid #ccc; }")
        self.setText("No PDF loaded")
        self.setScaledContents(False)
        self.original_pixmap = None
        self.zoom_factor = 1.0
        self.min_zoom = 0.25
        self.max_zoom = 4.0
        
        self.dragging = False
        self.last_pos = None
        self.setMouseTracking(True)
        self.setCursor(Qt.CursorShape.OpenHandCursor)
        
        self.edit_mode = False
        self.text_overlays = {}
        
        # Cursor marker
        self.cursor_marker = QLabel(self)
        self.cursor_marker.setFixedSize(20, 20)
        self.cursor_marker.setStyleSheet("""
            QLabel {
                background-color: rgba(255, 0, 0, 0.5);
                border: 2px solid red;
                border-radius: 10px;
            }
        """)
        self.cursor_marker.hide()
        
        # Default text properties
        self.default_font = QFont("Arial", 12)
        self.default_color = QColor(Qt.GlobalColor.red)

    def setPixmap(self, pixmap):
        self.original_pixmap = pixmap
        self.updatePixmap()

    def updatePixmap(self):
        if self.original_pixmap:
            label_size = self.size()
            zoomed_size = QSize(
                int(self.original_pixmap.width() * self.zoom_factor),
                int(self.original_pixmap.height() * self.zoom_factor)
            )
            scaled_pixmap = self.original_pixmap.scaled(
                zoomed_size,
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation
            )
            pixmap = scaled_pixmap.copy()
            
            # --- Highlight search results ---
            main_window = self.window()
            if isinstance(main_window, PDFMan):
                highlights = getattr(main_window, 'current_highlights', [])
                if highlights:
                    painter = QPainter(pixmap)
                    painter.setRenderHint(QPainter.RenderHint.Antialiasing)
                    color = QColor(255, 255, 0, 120)  # semi-transparent yellow
                    painter.setPen(Qt.PenStyle.NoPen)
                    painter.setBrush(color)
                    for rect in highlights:
                        # Scale rectangle to match the pixmap size
                        x = rect[0] * pixmap.width()
                        y = rect[1] * pixmap.height()
                        w = rect[2] * pixmap.width()
                        h = rect[3] * pixmap.height()
                        painter.drawRect(int(x), int(y), int(w), int(h))
                    painter.end()
            # --- End highlight ---
            
            # Add text overlays if any exist for the current page
            main_window = self.window()
            if isinstance(main_window, PDFMan) and main_window.pdf_ops.current_page in self.text_overlays:
                painter = QPainter(pixmap)
                painter.setRenderHint(QPainter.RenderHint.TextAntialiasing)
                painter.setRenderHint(QPainter.RenderHint.Antialiasing)
                
                for overlay in self.text_overlays[main_window.pdf_ops.current_page]:
                    # Set up the font and color
                    font = overlay.get('font', self.default_font)
                    color = overlay.get('color', self.default_color)
                    painter.setPen(QPen(color, 2))
                    painter.setFont(font)
                    
                    # Get text metrics for proper alignment
                    metrics = painter.fontMetrics()
                    text = overlay['text']
                    
                    # Calculate text position with proper baseline alignment
                    x = overlay['x']
                    y = overlay['y'] + metrics.ascent()  # Add ascent to align with baseline
                    
                    # Draw the text
                    painter.drawText(x, y, text)
                painter.end()
            
            # Set the pixmap
            super().setPixmap(pixmap)
            
            # Set minimum size to the zoomed pixmap size
            self.setMinimumSize(pixmap.width(), pixmap.height())
            
            # Force the label to update its size and layout
            self.updateGeometry()
            self.update()
            
            # Update the scroll area to show scroll bars when needed
            if self.parent():
                scroll_area = self.parent()
                if isinstance(scroll_area, QScrollArea):
                    # Force scroll area to update its geometry
                    scroll_area.updateGeometry()
                    scroll_area.update()
                    
                    # Update scroll bar ranges
                    h_scroll = scroll_area.horizontalScrollBar()
                    v_scroll = scroll_area.verticalScrollBar()
                    
                    if h_scroll and v_scroll:
                        # Calculate the content size
                        content_width = self.width()
                        content_height = self.height()
                        
                        # Calculate the viewport size
                        viewport_width = scroll_area.viewport().width()
                        viewport_height = scroll_area.viewport().height()
                        
                        # Set scroll bar ranges
                        h_scroll.setRange(0, max(0, content_width - viewport_width))
                        v_scroll.setRange(0, max(0, content_height - viewport_height))
                        
                        # Force scroll bars to update
                        h_scroll.update()
                        v_scroll.update()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.updatePixmap()

    def setZoom(self, factor):
        """Set the zoom factor for the preview"""
        # Clamp zoom factor between min and max values
        self.zoom_factor = max(self.min_zoom, min(self.max_zoom, factor))
        self.updatePixmap()
        # Do not force scroll to top after zoom

    def wheelEvent(self, event):
        """Handle mouse wheel events for zooming"""
        if event.modifiers() == Qt.KeyboardModifier.ControlModifier:
            # Zoom in/out with Ctrl + mouse wheel
            delta = event.angleDelta().y()
            if delta > 0:
                self.setZoom(self.zoom_factor * 1.1)  # Zoom in
            else:
                self.setZoom(self.zoom_factor / 1.1)  # Zoom out
            event.accept()
        else:
            super().wheelEvent(event)  # Normal scrolling

    def mousePressEvent(self, event):
        """Handle mouse press events for dragging and text overlay"""
        if self.edit_mode and event.button() == Qt.MouseButton.LeftButton:
            # Get the click position
            pos = event.position().toPoint()
            
            # Find the main window
            main_window = self.window()
            if isinstance(main_window, PDFMan):
                # Show text properties dialog
                dialog = TextPropertiesDialog(main_window)
                if dialog.exec():
                    # Get text properties
                    props = dialog.get_text_properties()
                    
                    # Create a text input dialog
                    text, ok = QInputDialog.getText(main_window, 'Add Text', 'Enter text to add:')
                    if ok and text:
                        # Store the text overlay with its position and properties
                        page_num = main_window.pdf_ops.current_page
                        if page_num not in self.text_overlays:
                            self.text_overlays[page_num] = []
                        self.text_overlays[page_num].append({
                            'text': text,
                            'x': pos.x(),
                            'y': pos.y(),
                            'font': props['font'],
                            'color': props['color']
                        })
                        
                        # Update the preview to show the new text
                        self.updatePixmap()
                        
                        # Mark changes as unsaved
                        main_window.pdf_ops.unsaved_changes = True
                        main_window.status_bar.showMessage("Changes pending - Click 'Save' to save")
        elif event.button() == Qt.MouseButton.LeftButton:
            self.dragging = True
            self.last_pos = event.position().toPoint()
            self.setCursor(Qt.CursorShape.ClosedHandCursor)
        event.accept()

    def mouseReleaseEvent(self, event):
        """Handle mouse release events for dragging"""
        if event.button() == Qt.MouseButton.LeftButton:
            self.dragging = False
            self.setCursor(Qt.CursorShape.OpenHandCursor)
            event.accept()

    def mouseMoveEvent(self, event):
        """Handle mouse move events for dragging and cursor marker"""
        if self.edit_mode:
            # Update cursor marker position
            self.cursor_marker.move(event.position().toPoint() - QPoint(10, 10))
            self.cursor_marker.show()
        else:
            self.cursor_marker.hide()
            
        if self.dragging and self.last_pos is not None:
            delta = event.position().toPoint() - self.last_pos
            parent = self.parent()
            if isinstance(parent, QScrollArea):
                scroll_area = parent
                h_scroll = scroll_area.horizontalScrollBar()
                v_scroll = scroll_area.verticalScrollBar()
                
                if h_scroll and v_scroll:
                    new_h_value = h_scroll.value() - delta.x() * 3
                    new_v_value = v_scroll.value() - delta.y() * 3
                    
                    new_h_value = max(0, min(new_h_value, h_scroll.maximum()))
                    new_v_value = max(0, min(new_v_value, v_scroll.maximum()))
                    
                    h_scroll.setValue(new_h_value)
                    v_scroll.setValue(new_v_value)
            self.last_pos = event.position().toPoint()
        event.accept()

    def setEditMode(self, enabled):
        """Enable or disable edit mode"""
        self.edit_mode = enabled
        if enabled:
            self.setCursor(Qt.CursorShape.IBeamCursor)
            self.cursor_marker.show()
        else:
            self.setCursor(Qt.CursorShape.OpenHandCursor)
            self.cursor_marker.hide()

class DraggablePagePreview(QWidget):
    def __init__(self, page_num, parent=None):
        super().__init__(parent)
        self.page_num = page_num
        self.setAcceptDrops(True)
        self.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.customContextMenuRequested.connect(self.show_context_menu)
        self.setStyleSheet("""
            QWidget {
                background-color: white;
                border: 1px solid #ccc;
                border-radius: 4px;
                padding: 5px;
            }
            QWidget:hover {
                border: 2px solid #2196F3;
            }
            QWidget.selected {
                background-color: #E3F2FD;
                border: 2px solid #2196F3;
            }
        """)
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(5, 5, 5, 5)
        
        # Create label for the preview
        self.preview_label = QLabel()
        self.preview_label.setMinimumSize(150, 200)
        self.preview_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        # Add page number label
        self.page_num_label = QLabel(f"Page {page_num + 1}")
        self.page_num_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.page_num_label.setStyleSheet("font-weight: bold;")
        
        layout.addWidget(self.preview_label)
        layout.addWidget(self.page_num_label)
        
        # Selection state
        self.is_selected = False
        self.rotation = 0  # Track rotation in degrees
    
    def show_context_menu(self, position):
        """Show the context menu on right-click"""
        menu = QMenu(self)
        
        # Create rotation submenu
        rotate_menu = menu.addMenu("Rotate")
        rotate_90 = rotate_menu.addAction("Rotate 90° Clockwise")
        rotate_180 = rotate_menu.addAction("Rotate 180°")
        rotate_270 = rotate_menu.addAction("Rotate 90° Counter-clockwise")
        
        # Create duplicate action
        duplicate_action = menu.addAction("Duplicate Selected Pages")
        
        # Add extract action
        extract_action = menu.addAction("Extract Selected Pages...")
        
        menu.addSeparator()
        
        # Create remove action
        remove_action = menu.addAction("Remove Selected Pages")
        
        # Create export images action
        export_images_action = menu.addAction("Export Selected Pages as Images")
        
        # Show menu and handle action
        action = menu.exec(self.mapToGlobal(position))
        
        # Get the main window
        main_window = self.window()
        if isinstance(main_window, PDFMan):
            if action == rotate_90:
                main_window.rotate_selected_pages(90)
            elif action == rotate_180:
                main_window.rotate_selected_pages(180)
            elif action == rotate_270:
                main_window.rotate_selected_pages(270)
            elif action == duplicate_action:
                main_window.duplicate_selected_pages()
            elif action == remove_action:
                main_window.remove_selected_pages()
            elif action == extract_action:
                main_window.extract_selected_pages()
            elif action == export_images_action:
                main_window.export_selected_pages_as_images()

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            # Get the main window
            main_window = self.window()
            if isinstance(main_window, PDFMan):
                # Handle selection
                if event.modifiers() == Qt.KeyboardModifier.ShiftModifier:
                    # Multi-select mode
                    main_window.select_pages_range(self.page_num)
                else:
                    # Single select mode
                    main_window.select_page(self.page_num)
            
            # Start drag operation
            drag = QDrag(self)
            mime_data = QMimeData()
            mime_data.setText(str(self.page_num))
            drag.setMimeData(mime_data)
            drag.exec()
    
    def setSelected(self, selected):
        """Set the selection state of this preview"""
        self.is_selected = selected
        if selected:
            self.setStyleSheet("""
                QWidget {
                    background-color: #E3F2FD;
                    border: 2px solid #2196F3;
                    border-radius: 4px;
                    padding: 5px;
                }
            """)
        else:
            self.setStyleSheet("""
                QWidget {
                    background-color: white;
                    border: 1px solid #ccc;
                    border-radius: 4px;
                    padding: 5px;
                }
                QWidget:hover {
                    border: 2px solid #2196F3;
                }
            """)
    
    def dragEnterEvent(self, event):
        if event.mimeData().hasText():
            event.acceptProposedAction()
    
    def dropEvent(self, event):
        source_page = int(event.mimeData().text())
        if source_page != self.page_num:
            # Find the main window by traversing up the widget hierarchy
            main_window = self.window()
            if isinstance(main_window, PDFMan):
                # Swap pages
                main_window.swap_pages(source_page, self.page_num)

class TextPropertiesDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Text Properties")
        self.setModal(True)
        
        layout = QVBoxLayout(self)
        
        # Font family selection
        font_layout = QHBoxLayout()
        font_layout.addWidget(QLabel("Font:"))
        self.font_combo = QFontComboBox()
        self.font_combo.setCurrentFont(QFont("Arial"))
        font_layout.addWidget(self.font_combo)
        layout.addLayout(font_layout)
        
        # Font size selection
        size_layout = QHBoxLayout()
        size_layout.addWidget(QLabel("Size:"))
        self.size_spin = QSpinBox()
        self.size_spin.setRange(8, 72)
        self.size_spin.setValue(12)
        size_layout.addWidget(self.size_spin)
        layout.addLayout(size_layout)
        
        # Color selection
        color_layout = QHBoxLayout()
        color_layout.addWidget(QLabel("Color:"))
        self.color_button = QPushButton()
        self.color_button.setFixedSize(30, 30)
        self.color_button.setStyleSheet("background-color: red;")
        self.color_button.clicked.connect(self.choose_color)
        color_layout.addWidget(self.color_button)
        layout.addLayout(color_layout)
        
        # Preview text
        preview_layout = QVBoxLayout()
        preview_layout.addWidget(QLabel("Preview:"))
        self.preview_label = QLabel("Sample Text")
        self.preview_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.preview_label.setMinimumHeight(50)
        self.preview_label.setStyleSheet("border: 1px solid #ccc;")
        preview_layout.addWidget(self.preview_label)
        layout.addLayout(preview_layout)
        
        # Connect signals for live preview
        self.font_combo.currentFontChanged.connect(self.update_preview)
        self.size_spin.valueChanged.connect(self.update_preview)
        
        # Buttons
        button_layout = QHBoxLayout()
        ok_button = QPushButton("OK")
        ok_button.clicked.connect(self.accept)
        cancel_button = QPushButton("Cancel")
        cancel_button.clicked.connect(self.reject)
        button_layout.addWidget(ok_button)
        button_layout.addWidget(cancel_button)
        layout.addLayout(button_layout)
        
        # Initialize color
        self.selected_color = QColor(Qt.GlobalColor.red)
        self.update_preview()
    
    def choose_color(self):
        color = QColorDialog.getColor(self.selected_color, self)
        if color.isValid():
            self.selected_color = color
            self.color_button.setStyleSheet(f"background-color: {color.name()};")
            self.update_preview()
    
    def update_preview(self):
        font = self.font_combo.currentFont()
        font.setPointSize(self.size_spin.value())
        self.preview_label.setFont(font)
        self.preview_label.setStyleSheet(f"border: 1px solid #ccc; color: {self.selected_color.name()};")
    
    def get_text_properties(self):
        font = self.font_combo.currentFont()
        font.setPointSize(self.size_spin.value())
        return {
            'font': font,
            'color': self.selected_color
        }

class OCRSearchThread(QThread):
    finished = pyqtSignal(list)
    def __init__(self, reader, previews, term):
        super().__init__()
        self.reader = reader
        self.previews = previews
        self.term = term
    def run(self):
        import numpy as np
        results = []
        for i, preview in enumerate(self.previews):
            if preview is None:
                continue
            img = np.array(preview)
            ocr_results = self.reader.readtext(img, detail=1, paragraph=False)
            for bbox, text, conf in ocr_results:
                words = text.split()
                for word in words:
                    if self.term.lower() == word.lower():
                        results.append(i)
                        break
                else:
                    continue
                break
        self.finished.emit(results)

class SpinnerDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Processing...")
        self.setModal(True)
        self.setWindowFlags(self.windowFlags() | Qt.WindowType.FramelessWindowHint)
        layout = QVBoxLayout(self)
        self.label = QLabel("Processing OCR... Please wait.")
        layout.addWidget(self.label)
        # Try to use a spinner GIF if available
        try:
            self.spinner = QLabel()
            self.movie = QMovie("spinner.gif")
            if self.movie.isValid():
                self.spinner.setMovie(self.movie)
                self.movie.start()
                layout.addWidget(self.spinner)
        except Exception:
            pass
        self.setFixedSize(220, 100)

class PDFMan(QMainWindow):
    RECENT_FILES_PATH = "recent_files.json"
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDFMan")
        self.setMinimumSize(800, 600)
        
        # Initialize PDF operations
        self.pdf_ops = PDFOperations()
        
        # Initialize selection tracking
        self.selected_pages = set()
        self.last_selected_page = None
        
        # Initialize undo/redo stacks
        self.undo_stack = []
        self.redo_stack = []
        
        # Initialize recent files
        self.recent_files = []
        self.max_recent_files = 10
        self.load_recent_files()  # Load from disk
        
        # Load DPI setting
        self.load_preview_dpi()
        
        # Create menu bar
        self.create_menu_bar()
        
        # Create toolbar
        self.create_toolbar()
        
        # Set up the UI
        self.setup_ui()
    
    def load_recent_files(self):
        try:
            if os.path.exists(self.RECENT_FILES_PATH):
                with open(self.RECENT_FILES_PATH, 'r', encoding='utf-8') as f:
                    self.recent_files = json.load(f)
        except Exception:
            self.recent_files = []

    def save_recent_files(self):
        try:
            with open(self.RECENT_FILES_PATH, 'w', encoding='utf-8') as f:
                json.dump(self.recent_files, f)
        except Exception:
            pass

    def create_menu_bar(self):
        menubar = self.menuBar()
        
        # File menu
        file_menu = menubar.addMenu("File")
        
        # Open action
        open_action = QAction("Open", self)
        open_action.setShortcut("Ctrl+O")
        open_action.triggered.connect(self.browse_pdf)
        file_menu.addAction(open_action)
        
        # Recent files submenu
        self.recent_menu = QMenu("Recent Files", self)
        file_menu.addMenu(self.recent_menu)
        self.update_recent_files_menu()
        
        # Save action
        save_action = QAction("Save", self)
        save_action.setShortcut("Ctrl+S")
        save_action.triggered.connect(self.save_pdf)
        file_menu.addAction(save_action)
        
        # Save As action
        save_as_action = QAction("Save As...", self)
        save_as_action.setShortcut("Ctrl+Shift+S")
        save_as_action.triggered.connect(self.save_as_pdf)
        file_menu.addAction(save_as_action)
        
        file_menu.addSeparator()
        
        # Export submenu
        export_menu = file_menu.addMenu("Export")
        
        # Export as TXT
        export_txt_action = QAction("Export as Text...", self)
        export_txt_action.triggered.connect(self.export_as_txt)
        export_menu.addAction(export_txt_action)
        
        # Export as DOC
        export_doc_action = QAction("Export as DOC...", self)
        export_doc_action.triggered.connect(self.export_as_doc)
        export_menu.addAction(export_doc_action)
        
        file_menu.addSeparator()
        
        # Close action
        close_action = QAction("Close", self)
        close_action.setShortcut("Ctrl+W")
        close_action.triggered.connect(self.close_document)
        file_menu.addAction(close_action)
        
        file_menu.addSeparator()
        
        # Exit action
        exit_action = QAction("Exit", self)
        exit_action.setShortcut("Ctrl+Q")
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)
        
        # Edit menu
        edit_menu = menubar.addMenu("Edit")
        
        # Undo action
        self.undo_action = QAction("Undo", self)
        self.undo_action.setShortcut("Ctrl+Z")
        self.undo_action.triggered.connect(self.undo)
        self.undo_action.setEnabled(False)
        edit_menu.addAction(self.undo_action)
        
        # Redo action
        self.redo_action = QAction("Redo", self)
        self.redo_action.setShortcut("Ctrl+Y")
        self.redo_action.triggered.connect(self.redo)
        self.redo_action.setEnabled(False)
        edit_menu.addAction(self.redo_action)
        
        edit_menu.addSeparator()
        
        # Rotate actions
        rotate_menu = edit_menu.addMenu("Rotate")
        rotate_90 = rotate_menu.addAction("Rotate 90° Clockwise")
        rotate_180 = rotate_menu.addAction("Rotate 180°")
        rotate_270 = rotate_menu.addAction("Rotate 90° Counter-clockwise")
        
        rotate_90.triggered.connect(lambda: self.rotate_selected_pages(90))
        rotate_180.triggered.connect(lambda: self.rotate_selected_pages(180))
        rotate_270.triggered.connect(lambda: self.rotate_selected_pages(270))
        
        # Duplicate action
        duplicate_action = QAction("Duplicate Selected Pages", self)
        duplicate_action.setShortcut("Ctrl+D")
        duplicate_action.triggered.connect(self.duplicate_selected_pages)
        edit_menu.addAction(duplicate_action)
        
        # Settings menu
        settings_menu = menubar.addMenu("Settings")
        
        # Set Poppler Path action
        poppler_action = QAction("Set Poppler Path...", self)
        poppler_action.triggered.connect(self.set_poppler_path)
        settings_menu.addAction(poppler_action)
        
        # Add Set Preview DPI
        set_dpi_action = QAction("Set Preview DPI", self)
        set_dpi_action.triggered.connect(self.set_preview_dpi)
        settings_menu.addAction(set_dpi_action)
        
        # Export options
        export_current_action = QAction("Export Current Page as Image", self)
        export_current_action.triggered.connect(self.export_current_page_as_image)
        edit_menu.addAction(export_current_action)
        export_all_action = QAction("Export All Pages as Images", self)
        export_all_action.triggered.connect(self.export_all_pages_as_images)
        edit_menu.addAction(export_all_action)
    
    def create_toolbar(self):
        toolbar = QToolBar()
        toolbar.setMovable(False)
        toolbar.setIconSize(QSize(24, 24))
        self.addToolBar(toolbar)
        # Open action
        open_action = QAction(QIcon("icons/open.png"), "Open", self)
        open_action.setShortcut("Ctrl+O")
        open_action.triggered.connect(self.browse_pdf)
        toolbar.addAction(open_action)
        # Save action
        save_action = QAction(QIcon("icons/save.png"), "Save", self)
        save_action.setShortcut("Ctrl+S")
        save_action.triggered.connect(self.save_pdf)
        toolbar.addAction(save_action)
        # Save As action
        save_as_action = QAction(QIcon("icons/save_as.png"), "Save As", self)
        save_as_action.setShortcut("Ctrl+Shift+S")
        save_as_action.triggered.connect(self.save_as_pdf)
        toolbar.addAction(save_as_action)
        # Edit action
        edit_action = QAction(QIcon("icons/edit.png"), "Edit", self)
        edit_action.setShortcut("Ctrl+E")
        edit_action.triggered.connect(self.toggle_edit_mode)
        toolbar.addAction(edit_action)
        # Close action
        close_action = QAction(QIcon("icons/close.png"), "Close", self)
        close_action.setShortcut("Ctrl+W")
        close_action.triggered.connect(self.close_document)
        toolbar.addAction(close_action)
        # Separator
        toolbar.addSeparator()
        # --- Search Bar ---
        self.search_bar = QLineEdit()
        self.search_bar.setPlaceholderText("Search PDF text...")
        self.search_bar.setFixedWidth(200)
        self.search_bar.returnPressed.connect(self.perform_search)
        toolbar.addWidget(self.search_bar)
        self.ocr_checkbox = QCheckBox("OCR Search")
        toolbar.addWidget(self.ocr_checkbox)
        search_btn = QPushButton("Search")
        search_btn.clicked.connect(self.perform_search)
        toolbar.addWidget(search_btn)
        clear_search_btn = QPushButton("Clear Search")
        clear_search_btn.clicked.connect(self.clear_search)
        toolbar.addWidget(clear_search_btn)
        prev_btn = QPushButton("<")
        prev_btn.clicked.connect(self.goto_prev_match)
        toolbar.addWidget(prev_btn)
        next_btn = QPushButton(">")
        next_btn.clicked.connect(self.goto_next_match)
        toolbar.addWidget(next_btn)
        self.match_label = QLabel("")
        toolbar.addWidget(self.match_label)
        self.search_results = []
        self.current_match_index = -1
        self.current_highlights = []
        self.ocr_reader = None  # Will be initialized on first use

    def clear_search(self):
        self.search_bar.setText("")
        self.match_label.setText("")
        self.search_results = []
        self.current_match_index = -1
        self.current_highlights = []
        self.update_preview()

    def perform_search(self):
        term = self.search_bar.text().strip()
        self.search_results = []
        self.current_match_index = -1
        self.current_highlights = []
        if not term or not self.pdf_ops.current_pdf:
            self.match_label.setText("")
            self.current_highlights = []
            self.update_preview()
            return
        if self.ocr_checkbox.isChecked():
            self.perform_ocr_search(term)
        else:
            self.perform_standard_search(term)

    def perform_standard_search(self, term):
        # Search all pages using extractable text
        for i, page in enumerate(self.pdf_ops.current_pdf.pages):
            try:
                text = page.extract_text() or ""
                if term.lower() in text.lower():
                    self.search_results.append(i)
            except Exception:
                continue
        if self.search_results:
            self.current_match_index = 0
            self.match_label.setText(f"1/{len(self.search_results)}")
            self.show_highlights_for_page(self.search_results[0], term)
            self.go_to_page(self.search_results[0] + 1)
        else:
            self.match_label.setText("0/0")
            self.current_highlights = []
            self.update_preview()

    def perform_ocr_search(self, term):
        # Use EasyOCR to search all pages in a background thread
        if self.ocr_reader is None:
            self.ocr_reader = easyocr.Reader(['en'], gpu=False)
        previews = [self.pdf_ops.get_preview(i) for i in range(self.pdf_ops.get_total_pages())]
        self.spinner_dialog = SpinnerDialog(self)
        self.ocr_thread = OCRSearchThread(self.ocr_reader, previews, term)
        self.ocr_thread.finished.connect(self.ocr_search_finished)
        self.ocr_thread.start()
        self.spinner_dialog.show()

    def ocr_search_finished(self, results):
        self.spinner_dialog.close()
        self.search_results = results
        if self.search_results:
            self.current_match_index = 0
            self.match_label.setText(f"1/{len(self.search_results)}")
            self.show_ocr_highlights_for_page(self.search_results[0], self.search_bar.text().strip())
            self.go_to_page(self.search_results[0] + 1)
        else:
            self.match_label.setText("0/0")
            self.current_highlights = []
            self.update_preview()

    def show_ocr_highlights_for_page(self, page_num, term):
        self.current_highlights = []
        if self.ocr_reader is None:
            self.ocr_reader = easyocr.Reader(['en'], gpu=False)
        preview = self.pdf_ops.get_preview(page_num)
        if preview is None:
            self.update_preview()
            return
        img = np.array(preview)
        results = self.ocr_reader.readtext(img, detail=1, paragraph=False)
        img_width, img_height = preview.size
        for bbox, text, conf in results:
            # Split detected text into words and check each word
            words = text.split()
            for idx, word in enumerate(words):
                if term.lower() == word.lower():
                    # Estimate word bbox within the detected bbox
                    # Assume words are evenly spaced in the bbox
                    x_coords = [pt[0] for pt in bbox]
                    y_coords = [pt[1] for pt in bbox]
                    x0 = min(x_coords)
                    y0 = min(y_coords)
                    x1 = max(x_coords)
                    y1 = max(y_coords)
                    total_words = len(words)
                    word_width = (x1 - x0) / total_words
                    wx0 = x0 + idx * word_width
                    wx1 = wx0 + word_width
                    # Normalize
                    x = wx0 / img_width
                    y = y0 / img_height
                    w = (wx1 - wx0) / img_width
                    h = (y1 - y0) / img_height
                    self.current_highlights.append((x, y, w, h))
        self.update_preview()

    def goto_next_match(self):
        if not self.search_results:
            return
        self.current_match_index = (self.current_match_index + 1) % len(self.search_results)
        self.match_label.setText(f"{self.current_match_index+1}/{len(self.search_results)}")
        page = self.search_results[self.current_match_index]
        if self.ocr_checkbox.isChecked():
            self.show_ocr_highlights_for_page(page, self.search_bar.text().strip())
        else:
            self.show_highlights_for_page(page, self.search_bar.text().strip())
        self.go_to_page(page + 1)

    def goto_prev_match(self):
        if not self.search_results:
            return
        self.current_match_index = (self.current_match_index - 1) % len(self.search_results)
        self.match_label.setText(f"{self.current_match_index+1}/{len(self.search_results)}")
        page = self.search_results[self.current_match_index]
        if self.ocr_checkbox.isChecked():
            self.show_ocr_highlights_for_page(page, self.search_bar.text().strip())
        else:
            self.show_highlights_for_page(page, self.search_bar.text().strip())
        self.go_to_page(page + 1)

    def go_to_page(self, page_number):
        if self.pdf_ops.go_to_page(page_number - 1):
            if hasattr(self, 'search_bar') and self.search_bar.text().strip():
                if hasattr(self, 'search_results') and self.pdf_ops.current_page in self.search_results:
                    if self.ocr_checkbox.isChecked():
                        self.show_ocr_highlights_for_page(self.pdf_ops.current_page, self.search_bar.text().strip())
                    else:
                        self.show_highlights_for_page(self.pdf_ops.current_page, self.search_bar.text().strip())
                else:
                    self.current_highlights = []
            self.update_preview()
            self.update_page_controls()
            # Scroll to top after page change
            if hasattr(self, 'scroll_area'):
                self.scroll_area.verticalScrollBar().setValue(0)

    def setup_ui(self):
        """Set up the main UI"""
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        main_layout.setStretch(0, 1)
        self.splitter = QSplitter(Qt.Orientation.Horizontal)
        self.splitter.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        # Left panel (PDF viewer)
        left_panel = QWidget()
        left_panel.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(0)
        left_layout.setStretch(0, 1)
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(False)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll_area.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        self.preview_label = PDFPreviewLabel()
        scroll_area.setWidget(self.preview_label)
        self.scroll_area = scroll_area
        left_layout.addWidget(scroll_area, stretch=1)
        controls_container = QWidget()
        controls_container.setStyleSheet("QWidget { background-color: #f0f0f0; border-top: 1px solid #ccc; }")
        controls_layout = QVBoxLayout(controls_container)
        controls_layout.setContentsMargins(5, 5, 5, 5)
        info_layout = QHBoxLayout()
        self.file_info_label = QLabel("No file loaded")
        info_layout.addWidget(self.file_info_label)
        controls_layout.addLayout(info_layout)
        controls_row = QHBoxLayout()
        nav_layout = QHBoxLayout()
        prev_btn = QPushButton("Previous")
        prev_btn.clicked.connect(self.previous_page)
        nav_layout.addWidget(prev_btn)
        self.page_spin = QSpinBox()
        self.page_spin.setMinimum(1)
        self.page_spin.setMaximum(1)
        self.page_spin.valueChanged.connect(self.go_to_page)
        nav_layout.addWidget(self.page_spin)
        self.page_count_label = QLabel("/ 1")
        nav_layout.addWidget(self.page_count_label)
        next_btn = QPushButton("Next")
        next_btn.clicked.connect(self.next_page)
        nav_layout.addWidget(next_btn)
        controls_row.addLayout(nav_layout)
        controls_row.addStretch()
        zoom_layout = QHBoxLayout()
        zoom_out_btn = QPushButton("-")
        zoom_out_btn.clicked.connect(lambda: self.preview_label.setZoom(self.preview_label.zoom_factor / 1.2))
        zoom_layout.addWidget(zoom_out_btn)
        zoom_reset_btn = QPushButton("100%")
        zoom_reset_btn.clicked.connect(lambda: self.preview_label.setZoom(1.0))
        zoom_layout.addWidget(zoom_reset_btn)
        zoom_in_btn = QPushButton("+")
        zoom_in_btn.clicked.connect(lambda: self.preview_label.setZoom(self.preview_label.zoom_factor * 1.2))
        zoom_layout.addWidget(zoom_in_btn)
        controls_row.addLayout(zoom_layout)
        controls_layout.addLayout(controls_row)
        button_layout = QHBoxLayout()
        browse_btn = QPushButton("Browse PDF")
        browse_btn.clicked.connect(self.browse_pdf)
        button_layout.addWidget(browse_btn)
        save_btn = QPushButton("Save")
        save_btn.clicked.connect(self.save_pdf)
        button_layout.addWidget(save_btn)
        save_as_btn = QPushButton("Save As...")
        save_as_btn.clicked.connect(self.save_as_pdf)
        button_layout.addWidget(save_as_btn)
        controls_layout.addLayout(button_layout)
        left_layout.addWidget(controls_container)
        self.splitter.addWidget(left_panel)
        self.left_panel = left_panel
        self.right_panel = QWidget()
        self.right_panel.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        right_layout = QVBoxLayout(self.right_panel)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(0)
        right_layout.setStretch(0, 1)
        self.tabs = QTabWidget()
        self.tabs.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        right_layout.addWidget(self.tabs, stretch=1)
        self.arrange_tab = QWidget()
        self.compare_tab = QWidget()
        self.combine_tab = QWidget()
        self.tabs.addTab(self.arrange_tab, "Edit Pages")
        self.tabs.addTab(self.compare_tab, "Compare PDFs")
        self.tabs.addTab(self.combine_tab, "Combine PDFs")
        self.setup_arrange_tab()
        self.setup_compare_tab()
        self.setup_combine_tab()
        self.splitter.setSizes([self.width(), 0])
        main_layout.addWidget(self.splitter, stretch=1)
        self.toggle_right_panel_btn = QPushButton("▶")
        self.toggle_right_panel_btn.setFixedSize(20, 60)
        self.toggle_right_panel_btn.setStyleSheet("""
            QPushButton {
                background-color: #f0f0f0;
                border: 1px solid #ccc;
                border-radius: 2px;
                padding: 2px;
                font-size: 12px;
            }
            QPushButton:hover {
                background-color: #e0e0e0;
            }
        """)
        self.toggle_right_panel_btn.clicked.connect(self.toggle_right_panel)
        main_layout.addWidget(self.toggle_right_panel_btn)
        self.setAcceptDrops(True)
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("No file loaded")
        if not self.pdf_ops.poppler_path:
            self.show_poppler_warning()
        self.edit_mode = False
        self.text_overlays = {}
        self.right_panel_in_splitter = False

    def toggle_right_panel(self):
        """Toggle the visibility of the right panel by adding/removing it from the splitter."""
        if self.right_panel_in_splitter:
            # Remove the right panel from the splitter
            idx = self.splitter.indexOf(self.right_panel)
            if idx != -1:
                self.splitter.widget(idx).setParent(None)
            self.toggle_right_panel_btn.setText("▶")
            self.splitter.setSizes([self.width(), 0])
            self.right_panel_in_splitter = False
        else:
            # Add the right panel back to the splitter
            self.splitter.addWidget(self.right_panel)
            self.toggle_right_panel_btn.setText("◀")
            self.splitter.setSizes([int(self.width() * 0.8), int(self.width() * 0.2)])
            self.right_panel_in_splitter = True

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if not self.right_panel_in_splitter:
            self.splitter.setSizes([self.width(), 0])

    def update_recent_files_menu(self):
        """Update the recent files menu"""
        self.recent_menu.clear()
        
        for file_path in self.recent_files:
            action = QAction(os.path.basename(file_path), self)
            action.setData(file_path)
            action.triggered.connect(lambda checked, path=file_path: self.handle_pdf_file(path))
            self.recent_menu.addAction(action)
        self.save_recent_files()  # Save whenever the menu is updated
    
    def add_recent_file(self, file_path):
        """Add a file to the recent files list"""
        if file_path in self.recent_files:
            self.recent_files.remove(file_path)
        self.recent_files.insert(0, file_path)
        self.recent_files = self.recent_files[:self.max_recent_files]
        self.update_recent_files_menu()
    
    def execute_command(self, command):
        """Execute a command and add it to the undo stack"""
        if command.execute():
            self.undo_stack.append(command)
            self.redo_stack.clear()  # Clear redo stack when new command is executed
            self.update_undo_redo_actions()
            self.update_preview()
            self.update_arrange_tab()
            return True
        return False

    def undo(self):
        """Undo the last command"""
        if self.undo_stack:
            command = self.undo_stack.pop()
            if command.undo():
                self.redo_stack.append(command)
                self.update_undo_redo_actions()
                self.update_preview()
                self.update_arrange_tab()
                return True
        return False

    def redo(self):
        """Redo the last undone command"""
        if self.redo_stack:
            command = self.redo_stack.pop()
            if command.execute():
                self.undo_stack.append(command)
                self.update_undo_redo_actions()
                self.update_preview()
                self.update_arrange_tab()
                return True
        return False

    def update_undo_redo_actions(self):
        """Update the enabled state of undo/redo actions"""
        self.undo_action.setEnabled(len(self.undo_stack) > 0)
        self.redo_action.setEnabled(len(self.redo_stack) > 0)

    def rotate_selected_pages(self, degrees):
        """Rotate selected pages by the specified degrees"""
        if not self.pdf_ops.current_pdf or not self.selected_pages:
            return
        
        command = RotatePagesCommand(self.pdf_ops, self.selected_pages, degrees)
        if self.execute_command(command):
            self.status_bar.showMessage(f"Rotated {len(self.selected_pages)} page(s) by {degrees}°")

    def duplicate_selected_pages(self):
        """Duplicate selected pages"""
        if not self.pdf_ops.current_pdf or not self.selected_pages:
            return
        
        command = DuplicatePagesCommand(self.pdf_ops, self.selected_pages)
        if self.execute_command(command):
            self.status_bar.showMessage(f"Duplicated {len(self.selected_pages)} page(s)")

    def remove_selected_pages(self):
        """Remove all selected pages"""
        if not self.pdf_ops.current_pdf or not self.selected_pages:
            return
        
        # Sort pages in reverse order to avoid index shifting
        pages_to_remove = sorted(self.selected_pages, reverse=True)
        
        # Confirm with user
        reply = QMessageBox.question(
            self,
            "Remove Pages",
            f"Are you sure you want to remove {len(pages_to_remove)} page(s)?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            command = RemovePagesCommand(self.pdf_ops, self.selected_pages)
            if self.execute_command(command):
                # Clear selection
                self.selected_pages.clear()
                self.last_selected_page = None
                self.status_bar.showMessage(f"Removed {len(pages_to_remove)} page(s)")

    def handle_pdf_file(self, file_path):
        try:
            logger.debug(f"Handling PDF file: {file_path}")
            if self.pdf_ops.load_pdf(file_path):
                self.update_file_info(file_path)
                self.update_preview()
                self.update_page_controls()
                self.update_arrange_tab()  # Update the arrange tab with new previews
                self.status_bar.showMessage(f"Loaded: {file_path}")
                
                # Add to recent files
                self.add_recent_file(file_path)
                
                # Check if previews were generated successfully
                if not self.pdf_ops.preview_images:
                    logger.warning("No preview images generated")
                    QMessageBox.warning(
                        self,
                        "Preview Generation Warning",
                        "Could not generate PDF previews. This might be due to missing Poppler installation.\n"
                        "Please ensure Poppler is installed on your system:\n"
                        "- Windows: Download from http://blog.alivate.com.au/poppler-windows/\n"
                        "- Linux: sudo apt-get install poppler-utils\n"
                        "- macOS: brew install poppler"
                    )
            else:
                logger.error(f"Failed to load PDF file: {file_path}")
                QMessageBox.critical(self, "Error", 
                                   f"Failed to load PDF file: {file_path}")
        except Exception as e:
            logger.error(f"Error handling PDF file: {str(e)}")
            logger.error(traceback.format_exc())
            QMessageBox.critical(self, "Error", 
                               f"An error occurred while handling the PDF file: {str(e)}")
        # Clear search state
        self.clear_search()
    
    def save_pdf(self):
        if not self.pdf_ops.current_pdf:
            QMessageBox.warning(self, "No File", "No PDF file is currently loaded.")
            return False
        
        # If the file hasn't been saved before, use Save As
        if not self.pdf_ops.current_path:
            return self.save_as_pdf()
        
        if self.pdf_ops.save_pdf():
            self.status_bar.showMessage(f"Saved: {self.pdf_ops.current_path}")
            return True
        else:
            QMessageBox.critical(self, "Error", "Failed to save PDF file.")
            return False
    
    def save_as_pdf(self):
        if not self.pdf_ops.current_pdf:
            QMessageBox.warning(self, "No File", "No PDF file is currently loaded.")
            return False
        
        file_name, _ = QFileDialog.getSaveFileName(
            self, "Save PDF As", "", "PDF Files (*.pdf)")
        if file_name:
            if self.pdf_ops.save_as_pdf(file_name):
                self.update_file_info(file_name)
                self.status_bar.showMessage(f"Saved as: {file_name}")
                return True
            else:
                QMessageBox.critical(self, "Error", "Failed to save PDF file.")
                return False
        return False
    
    def update_file_info(self, file_path):
        if self.pdf_ops.current_pdf:
            page_count = self.pdf_ops.get_page_count()
            self.file_info_label.setText(
                f"File: {os.path.basename(file_path)} | Pages: {page_count}")
    
    def close_document(self):
        """Close the current document"""
        if not self.pdf_ops.current_pdf:
            return
        
        if self.pdf_ops.has_unsaved_changes():
            reply = QMessageBox.question(
                self, "Unsaved Changes",
                "You have unsaved changes. Do you want to save them before closing?",
                QMessageBox.StandardButton.Save | 
                QMessageBox.StandardButton.Discard | 
                QMessageBox.StandardButton.Cancel
            )
            
            if reply == QMessageBox.StandardButton.Save:
                if not self.save_pdf():
                    return
            elif reply == QMessageBox.StandardButton.Cancel:
                return
        
        # Clear the current document
        self.pdf_ops.current_pdf = None
        self.pdf_ops.current_path = None
        self.pdf_ops.preview_images = []
        self.pdf_ops.current_page = 0
        self.pdf_ops.modified = False
        self.pdf_ops.unsaved_changes = False
        
        # Update the UI
        self.preview_label.clear()
        self.preview_label.setText("No PDF loaded")
        self.file_info_label.setText("No file loaded")
        self.page_spin.setValue(1)
        self.page_spin.setMaximum(1)
        self.page_count_label.setText("/ 1")
        self.status_bar.showMessage("No file loaded")
        # Clear search state
        self.clear_search()
    
    def show_poppler_warning(self):
        """Show warning about Poppler not being available"""
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Icon.Warning)
        msg.setWindowTitle("Poppler Not Found")
        msg.setText("Poppler is not installed or not found in your system.")
        msg.setInformativeText(
            "To enable PDF preview functionality, please install Poppler:\n\n"
            "Windows:\n"
            "1. Download Poppler from: https://github.com/oschwartz10612/poppler-windows/releases/\n"
            "2. Extract to C:\\Program Files\\poppler\n"
            "3. Add C:\\Program Files\\poppler\\bin to your system PATH\n\n"
            "Linux:\n"
            "sudo apt-get install poppler-utils\n\n"
            "macOS:\n"
            "brew install poppler"
        )
        msg.setStandardButtons(QMessageBox.StandardButton.Ok)
        msg.exec()

    def set_poppler_path(self):
        """Open dialog to set Poppler path"""
        current_path = self.pdf_ops.get_poppler_path()
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Icon.Information)
        msg.setWindowTitle("Set Poppler Path")
        msg.setText("Enter the path to your Poppler installation:")
        msg.setInformativeText(
            "This should be the directory containing pdfinfo.exe.\n"
            "Common locations:\n"
            "- C:\\Program Files\\poppler\\bin\n"
            "- C:\\Program Files (x86)\\poppler\\bin"
        )
        
        # Create a text input dialog
        dialog = QFileDialog()
        dialog.setFileMode(QFileDialog.FileMode.Directory)
        dialog.setDirectory(current_path if current_path else "C:\\Program Files")
        
        if dialog.exec():
            selected_path = dialog.selectedFiles()[0]
            if self.pdf_ops.set_poppler_path(selected_path):
                msg.setText("Poppler path set successfully!")
                msg.setInformativeText(f"Path: {selected_path}")
                msg.exec()
                # Reload current PDF if one is loaded
                if self.pdf_ops.current_pdf:
                    self.update_preview()
            else:
                msg.setIcon(QMessageBox.Icon.Warning)
                msg.setText("Invalid Poppler path!")
                msg.setInformativeText(
                    "The selected path does not contain pdfinfo.exe.\n"
                    "Please select the directory containing the Poppler binaries."
                )
                msg.exec()

    def swap_pages(self, source_page, target_page):
        """Swap two pages in the preview list"""
        # Find indices of the pages
        source_idx = self.page_previews.index(source_page)
        target_idx = self.page_previews.index(target_page)
        
        # Swap the page numbers
        self.page_previews[source_idx], self.page_previews[target_idx] = \
            self.page_previews[target_idx], self.page_previews[source_idx]
        
        # Remove all widgets from the grid
        while self.pages_grid.count():
            item = self.pages_grid.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        
        # Clear the page labels list
        self.page_labels.clear()
        
        # Re-add widgets in the new order
        for i, page_num in enumerate(self.page_previews):
            # Create a new container for this page
            container = DraggablePagePreview(page_num)
            
            # Get the preview image
            preview = self.pdf_ops.get_preview(page_num)
            if preview:
                try:
                    # Resize the preview to fit the label
                    preview = preview.resize((150, 200), Image.Resampling.LANCZOS)
                    
                    # Convert to RGB if needed
                    if preview.mode != 'RGB':
                        preview = preview.convert('RGB')
                    
                    # Convert to QPixmap
                    img_data = preview.tobytes("raw", "RGB")
                    qimg = QImage(img_data, preview.size[0], preview.size[1], QImage.Format.Format_RGB888)
                    pixmap = QPixmap.fromImage(qimg)
                    container.preview_label.setPixmap(pixmap)
                except Exception as e:
                    logger.error(f"Error creating preview for page {page_num + 1}: {str(e)}")
                    container.preview_label.setText(f"Page {page_num + 1}")
            
            # Add to grid
            row = i // 3
            col = i % 3
            self.pages_grid.addWidget(container, row, col)
            
            # Store reference
            self.page_labels.append(container)
        
        # Mark changes as unsaved
        self.pdf_ops.unsaved_changes = True
        self.status_bar.showMessage("Changes pending - Click 'Apply Changes' to save")

    def remove_page(self, page_num):
        """Remove a page from the PDF"""
        if not self.pdf_ops.current_pdf:
            return
        
        # Confirm with user
        reply = QMessageBox.question(
            self,
            "Remove Page",
            f"Are you sure you want to remove page {page_num + 1}?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            try:
                # Create a new PDF writer
                writer = PdfWriter()
                
                # Add all pages except the one to be removed
                for i in range(self.pdf_ops.get_total_pages()):
                    if i != page_num:
                        page = self.pdf_ops.get_page(i)
                        if page:
                            writer.add_page(page)
                
                # Save to a temporary file
                temp_file = "temp_removed.pdf"
                with open(temp_file, 'wb') as output_file:
                    writer.write(output_file)
                
                # Reload the PDF with the page removed
                self.handle_pdf_file(temp_file)
                
                # Delete the temporary file
                os.remove(temp_file)
                
                self.status_bar.showMessage(f"Page {page_num + 1} removed")
            except Exception as e:
                logger.error(f"Error removing page: {str(e)}")
                logger.error(traceback.format_exc())
                QMessageBox.critical(
                    self,
                    "Error",
                    f"An error occurred while removing the page: {str(e)}"
                )

    def toggle_edit_mode(self):
        """Toggle edit mode for adding text overlays"""
        self.edit_mode = not self.edit_mode
        if self.edit_mode:
            self.preview_label.setEditMode(True)
            self.status_bar.showMessage("Edit mode: Click anywhere to add text")
        else:
            self.preview_label.setEditMode(False)
            self.status_bar.showMessage("Edit mode disabled")

    def select_page(self, page_num):
        """Select a single page"""
        # Clear existing selection
        self.selected_pages.clear()
        self.selected_pages.add(page_num)
        self.last_selected_page = page_num
        
        # Update UI
        self.update_page_selection()
    
    def select_pages_range(self, page_num):
        """Select a range of pages"""
        if self.last_selected_page is not None:
            # Select all pages between last_selected_page and page_num
            start = min(self.last_selected_page, page_num)
            end = max(self.last_selected_page, page_num)
            self.selected_pages.update(range(start, end + 1))
        else:
            # If no previous selection, just select this page
            self.selected_pages.add(page_num)
        
        self.last_selected_page = page_num
        
        # Update UI
        self.update_page_selection()
    
    def update_page_selection(self):
        """Update the visual selection state of all page previews"""
        for label in self.page_labels:
            label.setSelected(label.page_num in self.selected_pages)
    
    def apply_page_arrangement(self):
        """Apply the new page order"""
        if not self.pdf_ops.current_pdf or not self.page_previews:
            return
        
        command = ReorderPagesCommand(self.pdf_ops, self.page_previews)
        if self.execute_command(command):
            QMessageBox.information(self, "Success", "Page order has been updated successfully!")
            self.status_bar.showMessage("Page order updated")
            # Reload the PDF to show the new order
            self.handle_pdf_file(self.pdf_ops.current_path)

    def setup_arrange_tab(self):
        """Set up the arrange tab for page management"""
        layout = QVBoxLayout(self.arrange_tab)
        layout.setContentsMargins(10, 10, 10, 10)
        
        # Title
        title_label = QLabel("Arrange PDF Pages")
        title_label.setStyleSheet("font-size: 14px; font-weight: bold; margin-bottom: 10px;")
        layout.addWidget(title_label)
        
        # Description
        desc_label = QLabel("Drag and drop pages to reorder them. Click 'Apply Changes' to save the new order.")
        desc_label.setStyleSheet("color: #666; margin-bottom: 10px;")
        layout.addWidget(desc_label)
        
        # Create scroll area for the grid of pages
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        
        # Create container widget for the grid
        self.arrange_container = QWidget()
        self.arrange_layout = QVBoxLayout(self.arrange_container)
        self.arrange_layout.setContentsMargins(0, 0, 0, 0)
        
        # Create grid layout for page previews
        self.pages_grid = QGridLayout()
        self.pages_grid.setSpacing(10)
        self.arrange_layout.addLayout(self.pages_grid)
        
        # Add stretch to push content to the top
        self.arrange_layout.addStretch()
        
        # Set the container as the scroll area's widget
        scroll_area.setWidget(self.arrange_container)
        layout.addWidget(scroll_area)
        
        # Button to apply changes
        self.apply_arrange_btn = QPushButton("Apply Changes")
        self.apply_arrange_btn.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                border: none;
                padding: 8px 20px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
            QPushButton:disabled {
                background-color: #BDBDBD;
            }
        """)
        self.apply_arrange_btn.clicked.connect(self.apply_page_arrangement)
        self.apply_arrange_btn.setEnabled(False)
        layout.addWidget(self.apply_arrange_btn)
        
        # Store references to page previews
        self.page_previews = []
        self.page_labels = []

    def update_arrange_tab(self):
        """Update the arrange tab with current page previews"""
        # Clear existing previews
        for label in self.page_labels:
            self.pages_grid.removeWidget(label)
            label.deleteLater()
        self.page_labels.clear()
        self.page_previews.clear()
        
        if not self.pdf_ops.current_pdf:
            return
        
        # Get all page previews
        total_pages = self.pdf_ops.get_total_pages()
        for page_num in range(total_pages):
            preview = self.pdf_ops.get_preview(page_num)
            if preview:
                # Create draggable preview widget
                container = DraggablePagePreview(page_num)
                
                # Convert PIL Image to QPixmap
                try:
                    # Resize the preview to fit the label while maintaining aspect ratio
                    target_size = (150, 200)
                    preview = preview.resize(target_size, Image.Resampling.LANCZOS)
                    
                    # Convert to RGB if needed
                    if preview.mode != 'RGB':
                        preview = preview.convert('RGB')
                    
                    # Convert to QPixmap directly
                    img_data = preview.tobytes("raw", "RGB")
                    qimg = QImage(img_data, preview.size[0], preview.size[1], preview.size[0] * 3, QImage.Format.Format_RGB888)
                    pixmap = QPixmap.fromImage(qimg)
                    container.preview_label.setPixmap(pixmap)
                except Exception as e:
                    logger.error(f"Error creating preview for page {page_num + 1}: {str(e)}")
                    container.preview_label.setText(f"Page {page_num + 1}")
                
                # Add container to grid
                row = page_num // 3  # 3 previews per row
                col = page_num % 3
                self.pages_grid.addWidget(container, row, col)
                
                # Store references
                self.page_labels.append(container)
                self.page_previews.append(page_num)
        
        # Enable the apply button
        self.apply_arrange_btn.setEnabled(True)
        
        # Update the layout
        self.arrange_container.updateGeometry()
        self.arrange_container.update()

    def setup_compare_tab(self):
        """Set up the Compare PDFs tab"""
        # Create main layout
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(10)  # Add spacing between elements
        
        # Add title and description
        title = QLabel("Compare PDFs")
        title.setStyleSheet("font-size: 24px; font-weight: bold; margin-bottom: 10px;")
        description = QLabel("Select a second PDF to compare with the current document")
        description.setStyleSheet("color: #666; margin-bottom: 20px;")
        
        layout.addWidget(title)
        layout.addWidget(description)
        
        # Add browse button for second PDF
        browse_button = QPushButton("Browse Second PDF")
        browse_button.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
        """)
        browse_button.clicked.connect(self.browse_second_pdf)
        layout.addWidget(browse_button)
        
        # Create scroll area for second PDF viewer
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(False)  # Disable widget resizing to allow scrolling
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll_area.setMinimumSize(400, 600)  # Set minimum size for the scroll area
        
        # Create preview label for second PDF
        self.second_pdf_viewer = PDFPreviewLabel()
        scroll_area.setWidget(self.second_pdf_viewer)
        layout.addWidget(scroll_area)
        
        # Create controls container
        controls_container = QWidget()
        controls_container.setStyleSheet("QWidget { background-color: #f0f0f0; border-top: 1px solid #ccc; }")
        controls_layout = QVBoxLayout(controls_container)
        controls_layout.setContentsMargins(5, 5, 5, 5)
        controls_layout.setSpacing(5)
        
        # Second PDF navigation
        second_nav_layout = QHBoxLayout()
        self.prev_page_second = QPushButton("Previous")
        self.next_page_second = QPushButton("Next")
        self.page_label_second = QLabel("Page: 0/0")
        
        second_nav_layout.addWidget(self.prev_page_second)
        second_nav_layout.addWidget(self.page_label_second)
        second_nav_layout.addWidget(self.next_page_second)
        controls_layout.addLayout(second_nav_layout)
        
        # Add zoom controls
        zoom_layout = QHBoxLayout()
        zoom_layout.addStretch()
        
        # Zoom out button
        zoom_out_btn = QPushButton("-")
        zoom_out_btn.clicked.connect(lambda: self.second_pdf_viewer.setZoom(self.second_pdf_viewer.zoom_factor / 1.2))
        zoom_layout.addWidget(zoom_out_btn)
        
        # Zoom reset button
        zoom_reset_btn = QPushButton("100%")
        zoom_reset_btn.clicked.connect(lambda: self.second_pdf_viewer.setZoom(1.0))
        zoom_layout.addWidget(zoom_reset_btn)
        
        # Zoom in button
        zoom_in_btn = QPushButton("+")
        zoom_in_btn.clicked.connect(lambda: self.second_pdf_viewer.setZoom(self.second_pdf_viewer.zoom_factor * 1.2))
        zoom_layout.addWidget(zoom_in_btn)
        
        controls_layout.addLayout(zoom_layout)
        
        # Add controls container to main layout
        layout.addWidget(controls_container)
        
        # Add compare button
        compare_button = QPushButton("Compare PDFs")
        compare_button.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #388E3C;
            }
        """)
        compare_button.clicked.connect(self.compare_pdfs)
        layout.addWidget(compare_button)
        
        # Create widget and set layout
        self.compare_tab.setLayout(layout)
        
        # Initialize second PDF variables
        self.second_pdf_path = None
        self.second_pdf_ops = None
        self.current_page_second = 0
        
        # Connect navigation signals
        self.prev_page_second.clicked.connect(self.prev_page_second_pdf)
        self.next_page_second.clicked.connect(self.next_page_second_pdf)
        
        # Initially disable navigation buttons
        self.prev_page_second.setEnabled(False)
        self.next_page_second.setEnabled(False)

    def setup_combine_tab(self):
        """Set up the Combine PDFs tab"""
        layout = QVBoxLayout(self.combine_tab)
        layout.setContentsMargins(10, 10, 10, 10)
        
        # Title
        title_label = QLabel("Combine Multiple PDF Files")
        title_label.setStyleSheet("font-size: 14px; font-weight: bold; margin-bottom: 10px;")
        layout.addWidget(title_label)
        
        # Description
        desc_label = QLabel("Add multiple PDF files to combine them into a single document.")
        desc_label.setStyleSheet("color: #666; margin-bottom: 10px;")
        layout.addWidget(desc_label)
        
        # List widget to show PDFs to be combined
        self.combine_list = QListWidget()
        self.combine_list.setStyleSheet("""
            QListWidget {
                border: 1px solid #ccc;
                border-radius: 4px;
                padding: 5px;
            }
            QListWidget::item {
                padding: 5px;
                border-bottom: 1px solid #eee;
            }
            QListWidget::item:selected {
                background-color: #e0e0e0;
            }
        """)
        layout.addWidget(self.combine_list)
        
        # Buttons layout
        button_layout = QHBoxLayout()
        
        # Add PDF button
        add_pdf_btn = QPushButton("Add PDF")
        add_pdf_btn.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                padding: 5px 15px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
        add_pdf_btn.clicked.connect(self.add_pdf_to_combine)
        button_layout.addWidget(add_pdf_btn)
        
        # Remove PDF button
        remove_pdf_btn = QPushButton("Remove Selected")
        remove_pdf_btn.setStyleSheet("""
            QPushButton {
                background-color: #f44336;
                color: white;
                border: none;
                padding: 5px 15px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #da190b;
            }
        """)
        remove_pdf_btn.clicked.connect(self.remove_pdf_from_combine)
        button_layout.addWidget(remove_pdf_btn)
        
        # Clear all button
        clear_btn = QPushButton("Clear All")
        clear_btn.setStyleSheet("""
            QPushButton {
                background-color: #607d8b;
                color: white;
                border: none;
                padding: 5px 15px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #455a64;
            }
        """)
        clear_btn.clicked.connect(self.clear_combine_list)
        button_layout.addWidget(clear_btn)
        
        layout.addLayout(button_layout)
        
        # Combine button
        self.combine_btn = QPushButton("Combine PDFs")  # Store the reference
        self.combine_btn.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                border: none;
                padding: 8px 20px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
            QPushButton:disabled {
                background-color: #BDBDBD;
            }
        """)
        self.combine_btn.clicked.connect(self.combine_pdfs)
        self.combine_btn.setEnabled(False)  # Initially disabled
        layout.addWidget(self.combine_btn)
        
        # Add stretch to push everything to the top
        layout.addStretch()

    def add_pdf_to_combine(self):
        """Add a PDF file to the combine list"""
        file_name, _ = QFileDialog.getOpenFileName(
            self, "Select PDF to Add", "", "PDF Files (*.pdf)")
        if file_name:
            # Add the file to the list
            self.combine_list.addItem(file_name)
            # Enable the combine button if we have at least 2 PDFs
            self.combine_btn.setEnabled(self.combine_list.count() >= 2)

    def remove_pdf_from_combine(self):
        """Remove the selected PDF from the combine list"""
        current_item = self.combine_list.currentItem()
        if current_item:
            self.combine_list.takeItem(self.combine_list.row(current_item))
            # Disable the combine button if we have less than 2 PDFs
            self.combine_btn.setEnabled(self.combine_list.count() >= 2)

    def clear_combine_list(self):
        """Clear all PDFs from the combine list"""
        self.combine_list.clear()
        self.combine_btn.setEnabled(False)

    def combine_pdfs(self):
        """Combine the selected PDFs into a single file"""
        if self.combine_list.count() < 2:
            QMessageBox.warning(self, "Not Enough PDFs", 
                              "Please add at least 2 PDF files to combine.")
            return
        
        # Get the output file name
        output_file, _ = QFileDialog.getSaveFileName(
            self, "Save Combined PDF", "", "PDF Files (*.pdf)")
        if not output_file:
            return
        
        # Get all PDF files from the list
        pdf_files = []
        for i in range(self.combine_list.count()):
            pdf_files.append(self.combine_list.item(i).text())
        
        try:
            # Combine the PDFs
            if self.pdf_ops.combine_pdfs(pdf_files, output_file):
                QMessageBox.information(self, "Success", 
                                      "PDFs have been combined successfully!")
                self.status_bar.showMessage(f"Combined PDFs saved as: {output_file}")
            else:
                QMessageBox.critical(self, "Error", 
                                   "Failed to combine PDFs.")
        except Exception as e:
            logger.error(f"Error combining PDFs: {str(e)}")
            logger.error(traceback.format_exc())
            QMessageBox.critical(self, "Error", 
                               f"An error occurred while combining PDFs: {str(e)}")

    def browse_second_pdf(self):
        """Open file dialog to select second PDF for comparison"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Second PDF",
            "",
            "PDF Files (*.pdf)"
        )
        
        if file_path:
            try:
                # Create new PDF operations instance for second PDF
                self.second_pdf_ops = PDFOperations()
                self.second_pdf_ops.load_pdf(file_path)
                self.second_pdf_path = file_path
                
                # Update UI
                self.current_page_second = 0
                self.update_second_pdf_viewer()
                
                # Enable navigation buttons
                self.prev_page_second.setEnabled(True)
                self.next_page_second.setEnabled(True)
                
                self.status_bar.showMessage(f"Second PDF loaded: {os.path.basename(file_path)}")
            except Exception as e:
                logger.error(f"Error loading second PDF: {str(e)}")
                logger.error(traceback.format_exc())
                QMessageBox.critical(
                    self,
                    "Error",
                    f"An error occurred while loading the PDF: {str(e)}"
                )

    def update_second_pdf_viewer(self):
        """Update the second PDF viewer with current page"""
        if not self.second_pdf_ops or not self.second_pdf_ops.current_pdf:
            return
        
        try:
            # Get the current page
            page = self.second_pdf_ops.get_page(self.current_page_second)
            if page:
                # Convert page to image
                image = self.second_pdf_ops.get_preview(self.current_page_second)
                if image:
                    try:
                        # Resize the image to a more manageable size before conversion
                        max_size = 1200  # Maximum dimension size for better quality
                        if image.size[0] > max_size or image.size[1] > max_size:
                            ratio = min(max_size / image.size[0], max_size / image.size[1])
                            new_size = (int(image.size[0] * ratio), int(image.size[1] * ratio))
                            image = image.resize(new_size, Image.Resampling.LANCZOS)
                        
                        # Convert to RGB if not already
                        if image.mode != 'RGB':
                            image = image.convert('RGB')
                        
                        # Convert PIL Image to QPixmap directly
                        img_data = image.tobytes("raw", "RGB")
                        qimg = QImage(img_data, image.size[0], image.size[1], image.size[0] * 3, QImage.Format.Format_RGB888)
                        pixmap = QPixmap.fromImage(qimg)
                        
                        # Set the pixmap
                        self.second_pdf_viewer.setPixmap(pixmap)
                        
                        # Update page label
                        total_pages = self.second_pdf_ops.get_total_pages()
                        self.page_label_second.setText(f"Page: {self.current_page_second + 1}/{total_pages}")
                        
                        # Update navigation buttons
                        self.prev_page_second.setEnabled(self.current_page_second > 0)
                        self.next_page_second.setEnabled(self.current_page_second < total_pages - 1)
                        
                    except Exception as e:
                        logger.error(f"Error converting preview image: {str(e)}")
                        logger.error(traceback.format_exc())
                        self.second_pdf_viewer.setText("Error displaying PDF preview")
        except Exception as e:
            logger.error(f"Error updating second PDF viewer: {str(e)}")
            logger.error(traceback.format_exc())

    def prev_page_second_pdf(self):
        """Show previous page of second PDF"""
        if self.current_page_second > 0:
            self.current_page_second -= 1
            self.update_second_pdf_viewer()

    def next_page_second_pdf(self):
        """Show next page of second PDF"""
        total_pages = self.second_pdf_ops.get_total_pages()
        if self.current_page_second < total_pages - 1:
            self.current_page_second += 1
            self.update_second_pdf_viewer()

    def compare_pdfs(self):
        """Compare the two PDFs and show differences"""
        if not self.pdf_ops.current_pdf or not self.second_pdf_ops or not self.second_pdf_ops.current_pdf:
            QMessageBox.warning(
                self,
                "Warning",
                "Please load both PDFs before comparing"
            )
            return
            
        try:
            # Reset both documents to page 1 and update their previews
            self.pdf_ops.current_page = 0  # Reset main PDF to page 1
            self.current_page_second = 0  # Reset second PDF to page 1
            
            # Update both viewers
            self.update_preview()  # Update main viewer
            self.update_second_pdf_viewer()  # Update second PDF viewer
            
            # Create difference viewer window
            diff_window = QDialog(self)
            diff_window.setWindowTitle("PDF Differences")
            diff_window.setMinimumSize(800, 600)
            
            layout = QVBoxLayout(diff_window)
            
            # Add page navigation controls
            nav_layout = QHBoxLayout()
            
            # First PDF navigation
            first_pdf_nav = QHBoxLayout()
            prev_page1_btn = QPushButton("←")
            next_page1_btn = QPushButton("→")
            page1_label = QLabel(f"Page {self.pdf_ops.current_page + 1}")
            first_pdf_nav.addWidget(prev_page1_btn)
            first_pdf_nav.addWidget(page1_label)
            first_pdf_nav.addWidget(next_page1_btn)
            
            # Second PDF navigation
            second_pdf_nav = QHBoxLayout()
            prev_page2_btn = QPushButton("←")
            next_page2_btn = QPushButton("→")
            page2_label = QLabel(f"Page {self.current_page_second + 1}")
            second_pdf_nav.addWidget(prev_page2_btn)
            second_pdf_nav.addWidget(page2_label)
            second_pdf_nav.addWidget(next_page2_btn)
            
            # Add navigation controls to layout
            nav_layout.addLayout(first_pdf_nav)
            nav_layout.addStretch()
            nav_layout.addLayout(second_pdf_nav)
            layout.addLayout(nav_layout)
            
            # Create scroll area for the difference image
            scroll_area = QScrollArea()
            scroll_area.setWidgetResizable(False)
            scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
            scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
            
            # Create preview label for the difference image
            diff_label = PDFPreviewLabel()
            
            def update_difference_view():
                """Update the difference view with current pages"""
                try:
                    # Get current pages from both PDFs
                    current_page1 = self.pdf_ops.current_page
                    current_page2 = self.current_page_second
                    
                    # Update page labels
                    page1_label.setText(f"Page {current_page1 + 1}/{self.pdf_ops.get_total_pages()}")
                    page2_label.setText(f"Page {current_page2 + 1}/{self.second_pdf_ops.get_total_pages()}")
                    
                    # Update navigation buttons
                    prev_page1_btn.setEnabled(current_page1 > 0)
                    next_page1_btn.setEnabled(current_page1 < self.pdf_ops.get_total_pages() - 1)
                    prev_page2_btn.setEnabled(current_page2 > 0)
                    next_page2_btn.setEnabled(current_page2 < self.second_pdf_ops.get_total_pages() - 1)
                    
                    # Get pages and convert to images
                    page1 = self.pdf_ops.get_page(current_page1)
                    page2 = self.second_pdf_ops.get_page(current_page2)
                    
                    if not page1 or not page2:
                        return
                    
                    img1 = self.pdf_ops.get_preview(current_page1)
                    img2 = self.second_pdf_ops.get_preview(current_page2)
                    
                    if not img1 or not img2:
                        return
                    
                    # Resize images to same size for comparison
                    size = (800, 1000)  # Standard size for comparison
                    img1 = img1.resize(size, Image.Resampling.LANCZOS)
                    img2 = img2.resize(size, Image.Resampling.LANCZOS)
                    
                    # Convert to RGB if needed
                    if img1.mode != 'RGB':
                        img1 = img1.convert('RGB')
                    if img2.mode != 'RGB':
                        img2 = img2.convert('RGB')
                    
                    # Compare images pixel by pixel
                    diff = Image.new('RGB', size)
                    diff_data = diff.load()
                    img1_data = img1.load()
                    img2_data = img2.load()
                    
                    for x in range(size[0]):
                        for y in range(size[1]):
                            r1, g1, b1 = img1_data[x, y]
                            r2, g2, b2 = img2_data[x, y]
                            
                            # If pixels are different, mark in red
                            if (r1, g1, b1) != (r2, g2, b2):
                                diff_data[x, y] = (255, 0, 0)  # Red for differences
                            else:
                                diff_data[x, y] = (r1, g1, b1)
                    
                    # Convert difference image to QPixmap
                    img_data = diff.tobytes("raw", "RGB")
                    qimg = QImage(img_data, size[0], size[1], QImage.Format.Format_RGB888)
                    pixmap = QPixmap.fromImage(qimg)
                    diff_label.setPixmap(pixmap)
                    
                except Exception as e:
                    logger.error(f"Error updating difference view: {str(e)}")
                    logger.error(traceback.format_exc())
            
            # Connect navigation buttons
            def next_page1():
                if self.pdf_ops.current_page < self.pdf_ops.get_total_pages() - 1:
                    self.pdf_ops.current_page += 1
                    self.update_preview()  # Update main viewer
                    update_difference_view()
            
            def prev_page1():
                if self.pdf_ops.current_page > 0:
                    self.pdf_ops.current_page -= 1
                    self.update_preview()  # Update main viewer
                    update_difference_view()
            
            def next_page2():
                if self.current_page_second < self.second_pdf_ops.get_total_pages() - 1:
                    self.current_page_second += 1
                    self.update_second_pdf_viewer()  # Update second PDF viewer
                    update_difference_view()
            
            def prev_page2():
                if self.current_page_second > 0:
                    self.current_page_second -= 1
                    self.update_second_pdf_viewer()  # Update second PDF viewer
                    update_difference_view()
            
            prev_page1_btn.clicked.connect(prev_page1)
            next_page1_btn.clicked.connect(next_page1)
            prev_page2_btn.clicked.connect(prev_page2)
            next_page2_btn.clicked.connect(next_page2)
            
            # Set the preview label as the scroll area's widget
            scroll_area.setWidget(diff_label)
            layout.addWidget(scroll_area)
            
            # Add zoom controls
            zoom_layout = QHBoxLayout()
            zoom_layout.addStretch()
            
            # Zoom out button
            zoom_out_btn = QPushButton("-")
            zoom_out_btn.clicked.connect(lambda: diff_label.setZoom(diff_label.zoom_factor / 1.2))
            zoom_layout.addWidget(zoom_out_btn)
            
            # Zoom reset button
            zoom_reset_btn = QPushButton("100%")
            zoom_reset_btn.clicked.connect(lambda: diff_label.setZoom(1.0))
            zoom_layout.addWidget(zoom_reset_btn)
            
            # Zoom in button
            zoom_in_btn = QPushButton("+")
            zoom_in_btn.clicked.connect(lambda: diff_label.setZoom(diff_label.zoom_factor * 1.2))
            zoom_layout.addWidget(zoom_in_btn)
            
            layout.addLayout(zoom_layout)
            
            # Add close button
            close_button = QPushButton("Close")
            close_button.clicked.connect(diff_window.close)
            layout.addWidget(close_button)
            
            # Initial update of the difference view
            update_difference_view()
            
            diff_window.exec()
            
        except Exception as e:
            logger.error(f"Error comparing PDFs: {str(e)}")
            logger.error(traceback.format_exc())
            QMessageBox.critical(
                self,
                "Error",
                f"An error occurred while comparing the PDFs: {str(e)}"
            )

    def show_highlights_for_page(self, page_num, term):
        """Extract bounding boxes for the search term on the given page and store as normalized rectangles (using PyMuPDF)."""
        self.current_highlights = []
        if not self.pdf_ops.current_path or not term:
            self.update_preview()
            return
        try:
            doc = fitz.open(self.pdf_ops.current_path)
            page = doc.load_page(page_num)
            text_instances = page.search_for(term, quads=False)
            preview = self.pdf_ops.get_preview(page_num)
            if not preview:
                self.update_preview()
                return
            img_width, img_height = preview.size
            for rect in text_instances:
                x = rect.x0 / img_width
                y = rect.y0 / img_height
                w = (rect.x1 - rect.x0) / img_width
                h = (rect.y1 - rect.y0) / img_height
                self.current_highlights.append((x, y, w, h))
        except Exception:
            self.current_highlights = []
        self.update_preview()

    def extract_selected_pages(self):
        if not self.pdf_ops.current_pdf or not self.selected_pages:
            QMessageBox.warning(self, "No Selection", "Please select one or more pages to extract.")
            return
        file_name, _ = QFileDialog.getSaveFileName(self, "Save Extracted PDF", "", "PDF Files (*.pdf)")
        if not file_name:
            return
        try:
            from PyPDF2 import PdfWriter
            writer = PdfWriter()
            for i in sorted(self.selected_pages):
                page = self.pdf_ops.get_page(i)
                if page:
                    writer.add_page(page)
            with open(file_name, 'wb') as output_file:
                writer.write(output_file)
            QMessageBox.information(self, "Success", f"Extracted {len(self.selected_pages)} page(s) to {file_name}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to extract pages: {str(e)}")

    def export_current_page_as_image(self):
        preview = self.pdf_ops.get_current_preview()
        if not preview:
            QMessageBox.warning(self, "No Preview", "No page preview available.")
            return
        file_name, _ = QFileDialog.getSaveFileName(self, "Save Image", "", "PNG Files (*.png);;JPEG Files (*.jpg)")
        if not file_name:
            return
        try:
            ext = os.path.splitext(file_name)[1].lower()
            fmt = 'PNG' if ext == '.png' else 'JPEG'
            preview.save(file_name, fmt)
            QMessageBox.information(self, "Success", f"Page saved as {file_name}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save image: {str(e)}")

    def export_all_pages_as_images(self):
        if not self.pdf_ops.preview_images:
            QMessageBox.warning(self, "No Previews", "No page previews available.")
            return
        folder = QFileDialog.getExistingDirectory(self, "Select Folder to Save Images")
        if not folder:
            return
        try:
            for i, preview in enumerate(self.pdf_ops.preview_images):
                file_path = os.path.join(folder, f"page_{i+1}.png")
                preview.save(file_path, 'PNG')
            QMessageBox.information(self, "Success", f"All pages saved as images in {folder}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save images: {str(e)}")

    def export_selected_pages_as_images(self):
        if not self.pdf_ops.current_pdf or not self.selected_pages:
            QMessageBox.warning(self, "No Selection", "Please select one or more pages to export.")
            return
        folder = QFileDialog.getExistingDirectory(self, "Select Folder to Save Images")
        if not folder:
            return
        try:
            for i in sorted(self.selected_pages):
                preview = self.pdf_ops.get_preview(i)
                if preview:
                    file_path = os.path.join(folder, f"page_{i+1}.png")
                    preview.save(file_path, 'PNG')
            QMessageBox.information(self, "Success", f"Selected pages saved as images in {folder}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save images: {str(e)}")

    def set_preview_dpi(self):
        dpi, ok = QInputDialog.getInt(self, "Set Preview DPI", "Enter DPI (e.g., 100-600):", value=getattr(self, 'preview_dpi', 150), min=50, max=600)
        if ok:
            self.preview_dpi = dpi
            self.save_preview_dpi()
            # Regenerate previews for current PDF
            if self.pdf_ops.current_pdf:
                self.pdf_ops.generate_previews(dpi_override=dpi)
                self.update_preview()
                self.update_arrange_tab()

    def save_preview_dpi(self):
        try:
            with open('preview_dpi.txt', 'w') as f:
                f.write(str(getattr(self, 'preview_dpi', 150)))
        except Exception:
            pass

    def load_preview_dpi(self):
        try:
            if os.path.exists('preview_dpi.txt'):
                with open('preview_dpi.txt', 'r') as f:
                    self.preview_dpi = int(f.read().strip())
            else:
                self.preview_dpi = 150
        except Exception:
            self.preview_dpi = 150

    def export_as_txt(self):
        """Export the current PDF as a text file"""
        if not self.pdf_ops.current_pdf:
            QMessageBox.warning(self, "No File", "No PDF file is currently loaded.")
            return
        
        file_name, _ = QFileDialog.getSaveFileName(
            self, "Export as Text", "", "Text Files (*.txt)")
        if file_name:
            if self.pdf_ops.export_to_txt(file_name):
                self.status_bar.showMessage(f"Exported as text: {file_name}")
                QMessageBox.information(self, "Success", "PDF exported as text successfully!")
            else:
                QMessageBox.critical(self, "Error", "Failed to export PDF as text.")

    def export_as_doc(self):
        """Export the current PDF as a DOC file"""
        if not self.pdf_ops.current_pdf:
            QMessageBox.warning(self, "No File", "No PDF file is currently loaded.")
            return
        
        file_name, _ = QFileDialog.getSaveFileName(
            self, "Export as DOC", "", "Word Documents (*.docx)")
        if file_name:
            if self.pdf_ops.export_to_doc(file_name):
                self.status_bar.showMessage(f"Exported as DOC: {file_name}")
                QMessageBox.information(self, "Success", "PDF exported as DOC successfully!")
            else:
                QMessageBox.critical(self, "Error", "Failed to export PDF as DOC.")

    def browse_pdf(self):
        """Open file dialog to select a PDF file"""
        if self.pdf_ops.has_unsaved_changes():
            reply = QMessageBox.question(
                self, "Unsaved Changes",
                "You have unsaved changes. Do you want to save them before opening a new file?",
                QMessageBox.StandardButton.Save | 
                QMessageBox.StandardButton.Discard | 
                QMessageBox.StandardButton.Cancel
            )
            
            if reply == QMessageBox.StandardButton.Save:
                if not self.save_pdf():
                    return
            elif reply == QMessageBox.StandardButton.Cancel:
                return
        
        file_name, _ = QFileDialog.getOpenFileName(
            self, "Open PDF File", "", "PDF Files (*.pdf)")
        if file_name:
            self.handle_pdf_file(file_name)

    def previous_page(self):
        """Move to previous page"""
        if self.pdf_ops.previous_page():
            self.update_preview()
            self.update_page_controls()

    def next_page(self):
        """Move to next page"""
        if self.pdf_ops.next_page():
            self.update_preview()
            self.update_page_controls()

    def update_page_controls(self):
        """Update page navigation controls"""
        total_pages = self.pdf_ops.get_total_pages()
        current_page = self.pdf_ops.get_current_page_number()
        
        self.page_spin.setMaximum(total_pages)
        self.page_spin.setValue(current_page)
        self.page_count_label.setText(f"/ {total_pages}")

    def update_preview(self):
        """Update the PDF preview display"""
        try:
            if not self.pdf_ops.current_pdf:
                self.preview_label.clear()
                self.preview_label.setText("No PDF loaded")
                return
            
            preview = self.pdf_ops.get_current_preview()
            if preview:
                try:
                    # Convert PIL Image to QPixmap
                    logger.debug(f"Converting preview image: size={preview.size}, mode={preview.mode}")
                    
                    # Resize the image to a more manageable size before conversion
                    max_size = 1200  # Maximum dimension size for better quality
                    if preview.size[0] > max_size or preview.size[1] > max_size:
                        ratio = min(max_size / preview.size[0], max_size / preview.size[1])
                        new_size = (int(preview.size[0] * ratio), int(preview.size[1] * ratio))
                        preview = preview.resize(new_size, Image.Resampling.LANCZOS)
                        logger.debug(f"Resized image to: {new_size}")
                    
                    # Convert to RGB if not already
                    if preview.mode != 'RGB':
                        preview = preview.convert('RGB')
                    
                    # Convert PIL Image to QPixmap directly
                    img_data = preview.tobytes("raw", "RGB")
                    qimg = QImage(img_data, preview.size[0], preview.size[1], preview.size[0] * 3, QImage.Format.Format_RGB888)
                    pixmap = QPixmap.fromImage(qimg)
                    
                    if pixmap.isNull():
                        raise Exception("Failed to create QPixmap from QImage")
                    
                    # Set the pixmap
                    self.preview_label.setPixmap(pixmap)
                    self.preview_label.setText("")  # Clear any error text
                    
                    logger.debug("Successfully converted and displayed preview")
                    
                except Exception as e:
                    logger.error(f"Error converting preview image: {str(e)}")
                    logger.error(traceback.format_exc())
                    self.preview_label.setText("Error displaying PDF preview")
            else:
                logger.warning("No preview available")
                self.preview_label.setText("Error: Could not generate preview")
        except Exception as e:
            logger.error(f"Error in update_preview: {str(e)}")
            logger.error(traceback.format_exc())
            self.preview_label.setText("Error updating preview")

def main():
    app = QApplication(sys.argv)
    window = PDFMan()
    window.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main() 