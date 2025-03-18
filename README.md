# PDFMan

PDFMan is a comprehensive PDF management application that allows you to:
- Load and view PDF files
- Arrange pages within PDF files
- Remove pages from PDF files
- Add pages to PDF files
- Add text to PDF files
- Analyze differences between two PDF files
- Combine multiple PDF files

## Features
- Modern, tabbed interface for easy navigation
- Drag-and-drop support for PDF files
- Preview of PDF pages
- Batch operations support
- Cross-platform compatibility

## Requirements
- Python 3.8 or higher
- Poppler (for PDF preview functionality)
  - Windows: Download and install from [poppler releases](http://blog.alivate.com.au/poppler-windows/)
  - Linux: `sudo apt-get install poppler-utils`
  - macOS: `brew install poppler`

## Installation
1. Clone this repository
2. Create a virtual environment (recommended):
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```
3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage
Run the application:
```bash
python main.py
```

## License
MIT License 