from PyPDF2 import PdfReader, PdfWriter
from pathlib import Path
import os
from pdf2image import convert_from_path
from PIL import Image
import io
import logging
import platform
import subprocess
import winreg
import traceback
from docx import Document

# Set up logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

class PDFOperations:
    def __init__(self):
        self.current_pdf = None
        self.current_path = None
        self.modified = False
        self.unsaved_changes = False
        self.preview_images = []
        self.current_page = 0
        self.poppler_path = None
        self._init_poppler()
    
    def _get_windows_path(self):
        """Get the system PATH from Windows registry"""
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SYSTEM\CurrentControlSet\Control\Session Manager\Environment", 0, winreg.KEY_READ) as key:
                path, _ = winreg.QueryValueEx(key, "Path")
                return path.split(os.pathsep)
        except Exception as e:
            logger.error(f"Error reading PATH from registry: {str(e)}")
            logger.error(traceback.format_exc())
            return []
    
    def _init_poppler(self):
        """Initialize Poppler path"""
        try:
            # First try to find poppler in common locations
            system = platform.system().lower()
            if system == 'windows':
                common_paths = [
                    r'C:\Program Files\poppler\bin',
                    r'C:\Program Files (x86)\poppler\bin',
                    os.path.expanduser('~\\AppData\\Local\\Programs\\poppler\\bin'),
                ]
                
                # Add paths from system PATH
                common_paths.extend(self._get_windows_path())
                
                logger.debug(f"Searching for Poppler in paths: {common_paths}")
                
                for path in common_paths:
                    if os.path.exists(path):
                        # Check if pdfinfo.exe exists in this path
                        pdfinfo_path = os.path.join(path, 'pdfinfo.exe')
                        if os.path.exists(pdfinfo_path):
                            self.poppler_path = path
                            logger.info(f"Found Poppler at: {path}")
                            return
                        # Also check if the path contains poppler
                        if 'poppler' in path.lower():
                            self.poppler_path = path
                            logger.info(f"Found Poppler at: {path}")
                            return
            elif system == 'linux':
                if os.path.exists('/usr/bin/pdfinfo'):
                    self.poppler_path = '/usr/bin'
                    logger.info("Found Poppler in /usr/bin")
                    return
            elif system == 'darwin':
                if os.path.exists('/usr/local/bin/pdfinfo'):
                    self.poppler_path = '/usr/local/bin'
                    logger.info("Found Poppler in /usr/local/bin")
                    return
            
            # If not found in common locations, try PATH
            try:
                subprocess.run(['pdfinfo', '--version'], capture_output=True, check=True)
                logger.info("Found Poppler in system PATH")
            except Exception as e:
                logger.warning(f"Poppler not found in PATH: {str(e)}")
                logger.warning("Poppler not found. PDF previews will not be available.")
        except Exception as e:
            logger.error(f"Error initializing Poppler: {str(e)}")
            logger.error(traceback.format_exc())
    
    def set_poppler_path(self, path):
        """Manually set the Poppler path"""
        try:
            if os.path.exists(path):
                pdfinfo_path = os.path.join(path, 'pdfinfo.exe')
                if os.path.exists(pdfinfo_path):
                    self.poppler_path = path
                    logger.info(f"Manually set Poppler path to: {path}")
                    return True
                else:
                    logger.warning(f"Path {path} does not contain pdfinfo.exe")
            else:
                logger.warning(f"Path {path} does not exist")
            return False
        except Exception as e:
            logger.error(f"Error setting Poppler path: {str(e)}")
            logger.error(traceback.format_exc())
            return False
    
    def get_poppler_path(self):
        """Get the current Poppler path"""
        return self.poppler_path
    
    def load_pdf(self, file_path):
        """Load a PDF file and return True if successful"""
        try:
            logger.debug(f"Loading PDF file: {file_path}")
            self.current_pdf = PdfReader(file_path)
            self.current_path = file_path
            self.modified = False
            self.unsaved_changes = False
            self.current_page = 0
            self.generate_previews()
            return True
        except Exception as e:
            logger.error(f"Error loading PDF: {str(e)}")
            logger.error(traceback.format_exc())
            return False
    
    def generate_previews(self, dpi_override=None):
        """Generate preview images for all pages"""
        if not self.current_pdf:
            logger.warning("No PDF loaded, cannot generate previews")
            return
        
        try:
            logger.debug("Generating PDF previews")
            if not self.poppler_path:
                logger.warning("Poppler not available, previews will not be generated")
                return
            
            logger.debug(f"Using Poppler path: {self.poppler_path}")
            logger.debug(f"Converting PDF: {self.current_path}")
            
            dpi = dpi_override if dpi_override else getattr(self, 'preview_dpi', 150)
            self.preview_images = convert_from_path(
                self.current_path,
                dpi=dpi,
                fmt='jpeg',
                poppler_path=self.poppler_path
            )
            logger.debug(f"Successfully generated {len(self.preview_images)} preview images")
        except Exception as e:
            logger.error(f"Error generating previews: {str(e)}")
            logger.error(traceback.format_exc())
            self.preview_images = []
    
    def get_preview(self, page_number):
        """Get preview image for a specific page"""
        if not self.preview_images:
            return None
        if 0 <= page_number < len(self.preview_images):
            return self.preview_images[page_number]
        logger.warning(f"Invalid page number: {page_number}")
        return None
    
    def get_current_preview(self):
        """Get preview image for current page"""
        return self.get_preview(self.current_page)
    
    def next_page(self):
        """Move to next page"""
        if self.current_page < len(self.preview_images) - 1:
            self.current_page += 1
            return True
        return False
    
    def previous_page(self):
        """Move to previous page"""
        if self.current_page > 0:
            self.current_page -= 1
            return True
        return False
    
    def go_to_page(self, page_number):
        """Go to specific page"""
        if 0 <= page_number < len(self.preview_images):
            self.current_page = page_number
            return True
        return False
    
    def get_current_page_number(self):
        """Get current page number (1-based)"""
        return self.current_page + 1
    
    def get_total_pages(self):
        """Get total number of pages"""
        return len(self.preview_images)
    
    def save_pdf(self, file_path=None):
        """Save the current PDF to the specified path or current path"""
        if not self.current_pdf:
            logger.error("No PDF loaded to save")
            return False
        
        try:
            save_path = file_path or self.current_path
            if not save_path:
                logger.error("No save path specified")
                return False
            
            # Create a new PDF writer
            writer = PdfWriter()
            
            # Add all pages from the current PDF
            for page in self.current_pdf.pages:
                writer.add_page(page)
            
            # Write to file
            with open(save_path, 'wb') as output_file:
                writer.write(output_file)
            
            # Update current path if this was a save as operation
            if file_path:
                self.current_path = file_path
            
            self.modified = False
            self.unsaved_changes = False
            logger.info(f"Successfully saved PDF to: {save_path}")
            return True
        except Exception as e:
            logger.error(f"Error saving PDF: {str(e)}")
            logger.error(traceback.format_exc())
            return False
    
    def save_as_pdf(self, file_path):
        """Save the current PDF to a new location"""
        if not file_path:
            logger.error("No file path specified for save as operation")
            return False
        
        return self.save_pdf(file_path)
    
    def has_unsaved_changes(self):
        """Check if there are unsaved changes"""
        # If no file is loaded, there are no unsaved changes
        if not self.current_pdf:
            return False
        # If there's no current path, it means the file hasn't been saved yet
        if not self.current_path:
            return True
        # Otherwise, check if there are actual modifications
        return self.unsaved_changes
    
    def mark_modified(self):
        """Mark the PDF as modified"""
        self.modified = True
        self.unsaved_changes = True
    
    def get_page_count(self):
        """Get the number of pages in the current PDF"""
        if self.current_pdf:
            return len(self.current_pdf.pages)
        return 0
    
    def get_page(self, page_number):
        """Get a specific page from the current PDF"""
        if self.current_pdf and 0 <= page_number < len(self.current_pdf.pages):
            return self.current_pdf.pages[page_number]
        return None

    def export_to_txt(self, output_file):
        """Export the current PDF to a text file"""
        try:
            if not self.current_pdf:
                logger.error("No PDF loaded to export")
                return False

            with open(output_file, 'w', encoding='utf-8') as f:
                for page in self.current_pdf.pages:
                    text = page.extract_text()
                    if text:
                        f.write(text + '\n\n')
            
            logger.info(f"Successfully exported PDF to text file: {output_file}")
            return True
        except Exception as e:
            logger.error(f"Error exporting to text: {str(e)}")
            logger.error(traceback.format_exc())
            return False

    def export_to_doc(self, output_file):
        """Export the current PDF to a DOC file"""
        try:
            if not self.current_pdf:
                logger.error("No PDF loaded to export")
                return False

            doc = Document()

            for page in self.current_pdf.pages:
                text = page.extract_text()
                if text:
                    doc.add_paragraph(text)
                doc.add_page_break()
            
            doc.save(output_file)
            logger.info(f"Successfully exported PDF to DOC file: {output_file}")
            return True
        except Exception as e:
            logger.error(f"Error exporting to DOC: {str(e)}")
            logger.error(traceback.format_exc())
            return False

    def combine_pdfs(self, pdf_files, output_file):
        """Combine multiple PDF files into a single PDF"""
        try:
            logger.debug(f"Combining PDFs: {pdf_files} into {output_file}")
            
            # Create a new PDF writer
            writer = PdfWriter()
            
            # Process each PDF file
            for pdf_file in pdf_files:
                try:
                    # Read the PDF file
                    reader = PdfReader(pdf_file)
                    
                    # Add all pages from this PDF
                    for page in reader.pages:
                        writer.add_page(page)
                    
                    logger.debug(f"Added pages from: {pdf_file}")
                except Exception as e:
                    logger.error(f"Error processing PDF file {pdf_file}: {str(e)}")
                    logger.error(traceback.format_exc())
                    raise
            
            # Write the combined PDF to the output file
            with open(output_file, 'wb') as output:
                writer.write(output)
            
            logger.info(f"Successfully combined PDFs into: {output_file}")
            return True
        except Exception as e:
            logger.error(f"Error combining PDFs: {str(e)}")
            logger.error(traceback.format_exc())
            return False 