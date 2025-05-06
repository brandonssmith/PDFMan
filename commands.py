from abc import ABC, abstractmethod
from PyPDF2 import PdfReader, PdfWriter
import os
import logging

logger = logging.getLogger(__name__)

class Command(ABC):
    """Base command class for the command pattern"""
    @abstractmethod
    def execute(self):
        pass
    
    @abstractmethod
    def undo(self):
        pass

class RotatePagesCommand(Command):
    """Command for rotating pages"""
    def __init__(self, pdf_ops, page_numbers, degrees):
        self.pdf_ops = pdf_ops
        self.page_numbers = page_numbers
        self.degrees = degrees
        self.original_pdf_path = None
        self.temp_file = "temp_rotated.pdf"
    
    def execute(self):
        try:
            # Save current state
            self.original_pdf_path = self.pdf_ops.current_path
            writer = PdfWriter()
            
            # Process all pages
            for i in range(self.pdf_ops.get_total_pages()):
                page = self.pdf_ops.get_page(i)
                if page:
                    if i in self.page_numbers:
                        page.rotate(self.degrees)
                    writer.add_page(page)
            
            # Save to temporary file
            with open(self.temp_file, 'wb') as output_file:
                writer.write(output_file)
            
            # Reload the PDF
            self.pdf_ops.load_pdf(self.temp_file)
            return True
        except Exception as e:
            logger.error(f"Error executing rotate command: {str(e)}")
            return False
    
    def undo(self):
        try:
            if self.original_pdf_path:
                self.pdf_ops.load_pdf(self.original_pdf_path)
                if os.path.exists(self.temp_file):
                    os.remove(self.temp_file)
                return True
        except Exception as e:
            logger.error(f"Error undoing rotate command: {str(e)}")
        return False

class DuplicatePagesCommand(Command):
    """Command for duplicating pages"""
    def __init__(self, pdf_ops, page_numbers):
        self.pdf_ops = pdf_ops
        self.page_numbers = page_numbers
        self.original_pdf_path = None
        self.temp_file = "temp_duplicated.pdf"
    
    def execute(self):
        try:
            # Save current state
            self.original_pdf_path = self.pdf_ops.current_path
            writer = PdfWriter()
            
            # Process all pages
            for i in range(self.pdf_ops.get_total_pages()):
                page = self.pdf_ops.get_page(i)
                if page:
                    writer.add_page(page)
                    if i in self.page_numbers:
                        writer.add_page(page)
            
            # Save to temporary file
            with open(self.temp_file, 'wb') as output_file:
                writer.write(output_file)
            
            # Reload the PDF
            self.pdf_ops.load_pdf(self.temp_file)
            return True
        except Exception as e:
            logger.error(f"Error executing duplicate command: {str(e)}")
            return False
    
    def undo(self):
        try:
            if self.original_pdf_path:
                self.pdf_ops.load_pdf(self.original_pdf_path)
                if os.path.exists(self.temp_file):
                    os.remove(self.temp_file)
                return True
        except Exception as e:
            logger.error(f"Error undoing duplicate command: {str(e)}")
        return False

class RemovePagesCommand(Command):
    """Command for removing pages"""
    def __init__(self, pdf_ops, page_numbers):
        self.pdf_ops = pdf_ops
        self.page_numbers = page_numbers
        self.original_pdf_path = None
        self.temp_file = "temp_removed.pdf"
    
    def execute(self):
        try:
            # Save current state
            self.original_pdf_path = self.pdf_ops.current_path
            writer = PdfWriter()
            
            # Add all pages except the ones to be removed
            for i in range(self.pdf_ops.get_total_pages()):
                if i not in self.page_numbers:
                    page = self.pdf_ops.get_page(i)
                    if page:
                        writer.add_page(page)
            
            # Save to temporary file
            with open(self.temp_file, 'wb') as output_file:
                writer.write(output_file)
            
            # Reload the PDF
            self.pdf_ops.load_pdf(self.temp_file)
            return True
        except Exception as e:
            logger.error(f"Error executing remove command: {str(e)}")
            return False
    
    def undo(self):
        try:
            if self.original_pdf_path:
                self.pdf_ops.load_pdf(self.original_pdf_path)
                if os.path.exists(self.temp_file):
                    os.remove(self.temp_file)
                return True
        except Exception as e:
            logger.error(f"Error undoing remove command: {str(e)}")
        return False

class ReorderPagesCommand(Command):
    """Command for reordering pages"""
    def __init__(self, pdf_ops, new_order):
        self.pdf_ops = pdf_ops
        self.new_order = new_order
        self.original_pdf_path = None
        self.temp_file = "temp_reordered.pdf"
    
    def execute(self):
        try:
            # Save current state
            self.original_pdf_path = self.pdf_ops.current_path
            writer = PdfWriter()
            
            # Add pages in the new order
            for page_num in self.new_order:
                page = self.pdf_ops.get_page(page_num)
                if page:
                    writer.add_page(page)
            
            # Save to temporary file
            with open(self.temp_file, 'wb') as output_file:
                writer.write(output_file)
            
            # Reload the PDF
            self.pdf_ops.load_pdf(self.temp_file)
            return True
        except Exception as e:
            logger.error(f"Error executing reorder command: {str(e)}")
            return False
    
    def undo(self):
        try:
            if self.original_pdf_path:
                self.pdf_ops.load_pdf(self.original_pdf_path)
                if os.path.exists(self.temp_file):
                    os.remove(self.temp_file)
                return True
        except Exception as e:
            logger.error(f"Error undoing reorder command: {str(e)}")
        return False 