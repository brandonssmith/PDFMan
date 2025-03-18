import sys
import os
from PyQt6.QtWidgets import (QApplication, QMainWindow, QTabWidget, QWidget,
                           QVBoxLayout, QPushButton, QFileDialog, QLabel,
                           QMessageBox, QHBoxLayout, QMenuBar, QMenu,
                           QScrollArea, QSpinBox, QStatusBar, QSizePolicy,
                           QListWidget, QGridLayout)
from PyQt6.QtCore import Qt, QSize, QMimeData
from PyQt6.QtGui import QDragEnterEvent, QDropEvent, QAction, QImage, QPixmap, QDrag
from pdf_operations import PDFOperations
import logging
import traceback
from PIL import Image
from PyPDF2 import PdfWriter

# Set up logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

class PDFPreviewLabel(QLabel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setMinimumSize(400, 600)
        self.setStyleSheet("QLabel { background-color: #f0f0f0; border: 1px solid #ccc; }")
        self.setText("No PDF loaded")
        self.setScaledContents(False)  # Changed to False to maintain aspect ratio
        self.original_pixmap = None
        self.zoom_factor = 1.0  # Initial zoom factor
        self.min_zoom = 0.25   # Minimum zoom level
        self.max_zoom = 4.0    # Maximum zoom level
        
        # Set size policy to allow the label to grow beyond its minimum size
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        
        # Mouse drag functionality
        self.dragging = False
        self.last_pos = None
        self.setMouseTracking(True)
        self.setCursor(Qt.CursorShape.OpenHandCursor)

    def setPixmap(self, pixmap):
        self.original_pixmap = pixmap
        self.updatePixmap()

    def updatePixmap(self):
        if self.original_pixmap:
            # Calculate the scaled size while maintaining aspect ratio
            label_size = self.size()
            
            # Apply zoom factor to the label size
            zoomed_size = QSize(
                int(label_size.width() * self.zoom_factor),
                int(label_size.height() * self.zoom_factor)
            )
            
            # Scale the pixmap
            scaled_pixmap = self.original_pixmap.scaled(
                zoomed_size,
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation
            )
            
            # Set the pixmap
            super().setPixmap(scaled_pixmap)
            
            # Update the label's minimum size to accommodate the zoomed image
            self.setMinimumSize(
                int(400 * self.zoom_factor),
                int(600 * self.zoom_factor)
            )
            
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
        """Handle mouse press events for dragging"""
        if event.button() == Qt.MouseButton.LeftButton:
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
        """Handle mouse move events for dragging"""
        if self.dragging and self.last_pos is not None:
            delta = event.position().toPoint() - self.last_pos
            parent = self.parent()
            if isinstance(parent, QScrollArea):
                scroll_area = parent
                h_scroll = scroll_area.horizontalScrollBar()
                v_scroll = scroll_area.verticalScrollBar()
                
                if h_scroll and v_scroll:
                    # Calculate new scroll positions with increased sensitivity
                    new_h_value = h_scroll.value() - delta.x() * 3
                    new_v_value = v_scroll.value() - delta.y() * 3
                    
                    # Clamp values to valid ranges
                    new_h_value = max(0, min(new_h_value, h_scroll.maximum()))
                    new_v_value = max(0, min(new_v_value, v_scroll.maximum()))
                    
                    # Update scroll positions
                    h_scroll.setValue(new_h_value)
                    v_scroll.setValue(new_v_value)
            self.last_pos = event.position().toPoint()
        event.accept()

class DraggablePagePreview(QWidget):
    def __init__(self, page_num, parent=None):
        super().__init__(parent)
        self.page_num = page_num
        self.setAcceptDrops(True)
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
    
    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            drag = QDrag(self)
            mime_data = QMimeData()
            mime_data.setText(str(self.page_num))
            drag.setMimeData(mime_data)
            drag.exec()
    
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

class PDFMan(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDFMan")
        self.setMinimumSize(800, 600)
        
        # Initialize PDF operations
        self.pdf_ops = PDFOperations()
        
        # Create menu bar
        self.create_menu_bar()
        
        # Create central widget and layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)  # Changed to horizontal layout
        
        # Create left panel for PDF viewer
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(0)
        
        # Create scroll area for PDF preview
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(False)  # Disable widget resizing to allow scrolling
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll_area.setMinimumSize(400, 600)  # Set minimum size for the scroll area
        
        # Create preview label
        self.preview_label = PDFPreviewLabel()
        scroll_area.setWidget(self.preview_label)
        
        # Store scroll area reference for later use
        self.scroll_area = scroll_area
        
        # Add scroll area to left panel layout
        left_layout.addWidget(scroll_area)
        
        # Create controls container for PDF viewer
        controls_container = QWidget()
        controls_container.setStyleSheet("QWidget { background-color: #f0f0f0; border-top: 1px solid #ccc; }")
        controls_layout = QVBoxLayout(controls_container)
        controls_layout.setContentsMargins(5, 5, 5, 5)
        
        # File info section
        info_layout = QHBoxLayout()
        self.file_info_label = QLabel("No file loaded")
        info_layout.addWidget(self.file_info_label)
        controls_layout.addLayout(info_layout)
        
        # Navigation and zoom controls in a single row
        controls_row = QHBoxLayout()
        
        # Navigation controls
        nav_layout = QHBoxLayout()
        
        # Previous page button
        prev_btn = QPushButton("Previous")
        prev_btn.clicked.connect(self.previous_page)
        nav_layout.addWidget(prev_btn)
        
        # Page number spin box
        self.page_spin = QSpinBox()
        self.page_spin.setMinimum(1)
        self.page_spin.setMaximum(1)
        self.page_spin.valueChanged.connect(self.go_to_page)
        nav_layout.addWidget(self.page_spin)
        
        # Page count label
        self.page_count_label = QLabel("/ 1")
        nav_layout.addWidget(self.page_count_label)
        
        # Next page button
        next_btn = QPushButton("Next")
        next_btn.clicked.connect(self.next_page)
        nav_layout.addWidget(next_btn)
        
        controls_row.addLayout(nav_layout)
        
        # Add stretch to push zoom controls to the right
        controls_row.addStretch()
        
        # Zoom controls
        zoom_layout = QHBoxLayout()
        
        # Zoom out button
        zoom_out_btn = QPushButton("-")
        zoom_out_btn.clicked.connect(lambda: self.preview_label.setZoom(self.preview_label.zoom_factor / 1.2))
        zoom_layout.addWidget(zoom_out_btn)
        
        # Zoom reset button
        zoom_reset_btn = QPushButton("100%")
        zoom_reset_btn.clicked.connect(lambda: self.preview_label.setZoom(1.0))
        zoom_layout.addWidget(zoom_reset_btn)
        
        # Zoom in button
        zoom_in_btn = QPushButton("+")
        zoom_in_btn.clicked.connect(lambda: self.preview_label.setZoom(self.preview_label.zoom_factor * 1.2))
        zoom_layout.addWidget(zoom_in_btn)
        
        controls_row.addLayout(zoom_layout)
        controls_layout.addLayout(controls_row)
        
        # File operations buttons in a single row
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
        
        # Add the controls container to the left panel layout
        left_layout.addWidget(controls_container)
        
        # Add left panel to main layout
        main_layout.addWidget(left_panel, stretch=2)  # Give it more stretch than the right panel
        
        # Create right panel for tabs
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(0, 0, 0, 0)
        
        # Create tab widget
        self.tabs = QTabWidget()
        right_layout.addWidget(self.tabs)
        
        # Create tabs for different functionalities
        self.arrange_tab = QWidget()
        self.edit_tab = QWidget()
        self.compare_tab = QWidget()
        self.combine_tab = QWidget()
        
        # Add tabs to widget
        self.tabs.addTab(self.arrange_tab, "Arrange Pages")
        self.tabs.addTab(self.edit_tab, "Edit PDF")
        self.tabs.addTab(self.compare_tab, "Compare PDFs")
        self.tabs.addTab(self.combine_tab, "Combine PDFs")
        
        # Setup each tab
        self.setup_arrange_tab()
        self.setup_edit_tab()
        self.setup_compare_tab()
        self.setup_combine_tab()
        
        # Add right panel to main layout
        main_layout.addWidget(right_panel, stretch=1)
        
        # Enable drag and drop
        self.setAcceptDrops(True)
        
        # Status bar for showing current file and unsaved changes
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("No file loaded")
        
        # Check for Poppler availability
        if not self.pdf_ops.poppler_path:
            self.show_poppler_warning()
    
    def create_menu_bar(self):
        menubar = self.menuBar()
        
        # File menu
        file_menu = menubar.addMenu("File")
        
        # Open action
        open_action = QAction("Open", self)
        open_action.setShortcut("Ctrl+O")
        open_action.triggered.connect(self.browse_pdf)
        file_menu.addAction(open_action)
        
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
        
        # Settings menu
        settings_menu = menubar.addMenu("Settings")
        
        # Set Poppler Path action
        poppler_action = QAction("Set Poppler Path...", self)
        poppler_action.triggered.connect(self.set_poppler_path)
        settings_menu.addAction(poppler_action)
    
    def setup_arrange_tab(self):
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

    def apply_page_arrangement(self):
        """Apply the new page order"""
        if not self.pdf_ops.current_pdf or not self.page_previews:
            return
        
        try:
            # Get the output file name
            file_name, _ = QFileDialog.getSaveFileName(
                self, "Save Reordered PDF", "", "PDF Files (*.pdf)")
            if not file_name:
                return
            
            # Create a new PDF writer
            writer = PdfWriter()
            
            # Add pages in the new order
            for page_num in self.page_previews:
                page = self.pdf_ops.get_page(page_num)
                if page:
                    writer.add_page(page)
            
            # Write to file
            with open(file_name, 'wb') as output_file:
                writer.write(output_file)
            
            QMessageBox.information(self, "Success", "Page order has been updated successfully!")
            self.status_bar.showMessage("Page order updated")
            # Reload the PDF to show the new order
            self.handle_pdf_file(file_name)
        except Exception as e:
            logger.error(f"Error applying page arrangement: {str(e)}")
            logger.error(traceback.format_exc())
            QMessageBox.critical(self, "Error", f"An error occurred while updating the page order: {str(e)}")

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
                    max_size = 1200  # Increased maximum dimension size for better quality
                    if preview.size[0] > max_size or preview.size[1] > max_size:
                        ratio = min(max_size / preview.size[0], max_size / preview.size[1])
                        new_size = (int(preview.size[0] * ratio), int(preview.size[1] * ratio))
                        preview = preview.resize(new_size, Image.Resampling.LANCZOS)
                        logger.debug(f"Resized image to: {new_size}")
                    
                    # Convert to RGB if not already
                    if preview.mode != 'RGB':
                        preview = preview.convert('RGB')
                    
                    # Use a temporary file for conversion to avoid memory issues
                    temp_file = None
                    try:
                        import tempfile
                        temp_file = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
                        temp_file.close()
                        
                        # Save with high quality settings
                        preview.save(temp_file.name, 'PNG', optimize=True, quality=95)
                        logger.debug(f"Saved temporary PNG file: {temp_file.name}")
                        
                        # Load the PNG file into QImage
                        qimg = QImage(temp_file.name)
                        if qimg.isNull():
                            raise Exception("Failed to create QImage from temporary file")
                        
                        # Convert to QPixmap
                        pixmap = QPixmap.fromImage(qimg)
                        if pixmap.isNull():
                            raise Exception("Failed to create QPixmap from QImage")
                        
                        # Set the pixmap
                        self.preview_label.setPixmap(pixmap)
                        self.preview_label.setText("")  # Clear any error text
                        
                        logger.debug("Successfully converted and displayed preview")
                    except Exception as e:
                        logger.error(f"Error in PNG conversion method: {str(e)}")
                        logger.error(traceback.format_exc())
                        raise
                    finally:
                        # Clean up the temporary file
                        if temp_file and os.path.exists(temp_file.name):
                            try:
                                # Force garbage collection of QImage and QPixmap
                                del qimg
                                del pixmap
                                import gc
                                gc.collect()
                                # Now try to delete the file
                                os.unlink(temp_file.name)
                            except Exception as e:
                                logger.warning(f"Could not delete temporary file: {str(e)}")
                    
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
    
    def update_page_controls(self):
        """Update page navigation controls"""
        total_pages = self.pdf_ops.get_total_pages()
        current_page = self.pdf_ops.get_current_page_number()
        
        self.page_spin.setMaximum(total_pages)
        self.page_spin.setValue(current_page)
        self.page_count_label.setText(f"/ {total_pages}")
    
    def next_page(self):
        """Move to next page"""
        if self.pdf_ops.next_page():
            self.update_preview()
            self.update_page_controls()
    
    def previous_page(self):
        """Move to previous page"""
        if self.pdf_ops.previous_page():
            self.update_preview()
            self.update_page_controls()
    
    def go_to_page(self, page_number):
        """Go to specific page"""
        if self.pdf_ops.go_to_page(page_number - 1):
            self.update_preview()
    
    def setup_edit_tab(self):
        layout = QVBoxLayout(self.edit_tab)
        label = QLabel("Edit PDF content")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(label)
    
    def setup_compare_tab(self):
        layout = QVBoxLayout(self.compare_tab)
        label = QLabel("Compare two PDF files")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(label)
    
    def setup_combine_tab(self):
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
    
    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
    
    def dropEvent(self, event: QDropEvent):
        files = [url.toLocalFile() for url in event.mimeData().urls()]
        for file in files:
            if file.lower().endswith('.pdf'):
                self.handle_pdf_file(file)
            else:
                QMessageBox.warning(self, "Invalid File", 
                                  f"{file} is not a PDF file.")
    
    def browse_pdf(self):
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
    
    def handle_pdf_file(self, file_path):
        try:
            logger.debug(f"Handling PDF file: {file_path}")
            if self.pdf_ops.load_pdf(file_path):
                self.update_file_info(file_path)
                self.update_preview()
                self.update_page_controls()
                self.update_arrange_tab()  # Update the arrange tab with new previews
                self.status_bar.showMessage(f"Loaded: {file_path}")
                
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

def main():
    app = QApplication(sys.argv)
    window = PDFMan()
    window.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main() 