# PDFMan

PDFMan is a powerful, user-friendly PDF management tool built with Python and PyQt6. It allows you to view, search, organize, edit, and export PDF documents with ease.

## Features

- **PDF Viewing**: High-quality, zoomable PDF preview with smooth scrolling.
- **Search**: Fast text search and OCR-based search (EasyOCR) for scanned/image PDFs, with highlighting and navigation.
- **Arrange Pages**: Drag-and-drop page reordering, multi-page selection, and context menu actions (rotate, duplicate, remove, extract, export as images).
- **Combine PDFs**: Merge multiple PDFs into a single document.
- **Compare PDFs**: Side-by-side comparison with visual difference highlighting.
- **Export Options**:
  - Export current page as image (PNG/JPEG)
  - Export all pages as images
  - Export selected pages as images
- **Extract Pages**: Save selected pages as a new PDF.
- **Recent Files**: Persistent recent files list.
- **Customizable Preview DPI**: Choose the resolution for PDF previews for best balance of clarity and performance.
- **Settings**: Set Poppler path (for Windows), set preview DPI.
- **Cross-platform**: Works on Windows, macOS, and Linux.

## Installation

1. **Clone the repository**
   ```sh
   git clone <your-repo-url>
   cd PDFMan
   ```

2. **Set up a virtual environment (recommended)**
   ```sh
   python -m venv .venv
   source .venv/bin/activate  # On Windows: .venv\Scripts\activate
   ```

3. **Install dependencies**
   ```sh
   pip install -r requirements.txt
   ```

4. **Install Poppler** (required for PDF preview/export)
   - **Windows**: Download from [Poppler for Windows](https://github.com/oschwartz10612/poppler-windows/releases/), extract, and set the path in Settings > Set Poppler Path.
   - **macOS**: `brew install poppler`
   - **Linux**: `sudo apt-get install poppler-utils`

## Usage

1. **Run the application**
   ```sh
   python main.py
   ```

2. **Open a PDF** using the File menu, toolbar, or drag-and-drop.
3. **Navigate, search, zoom, and organize** your PDF using the intuitive interface.
4. **Right-click** on pages in the Arrange tab for advanced options (extract, export, rotate, etc.).
5. **Use the Edit and Settings menus** for export and configuration options.

## Dependencies
- PyQt6
- PyMuPDF (fitz)
- pdf2image
- Pillow
- PyPDF2
- easyocr
- torch, torchvision, torchaudio
- numpy
- poppler-utils (system dependency)

## Notes
- For OCR search, EasyOCR and PyTorch are required.
- For PDF preview/export, Poppler must be installed and the path set (on Windows).
- All export and extract features use high-quality images generated from the PDF.

## License
MIT License

---

For questions or contributions, please open an issue or pull request on GitHub.