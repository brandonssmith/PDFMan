# SVG to PNG Converter

A simple GUI application to convert SVG files to PNG format with customizable dimensions.

## Features

- Preview SVG files before conversion
- Customize output PNG dimensions
- Simple and intuitive interface
- Real-time preview of the conversion

## Requirements

- Python 3.8 or higher
- PyQt6
- cairosvg

## Installation

1. Clone this repository or download the source code
2. Install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. Run the application:
   ```bash
   python main.py
   ```
2. Click "Browse SVG" to select an SVG file
3. Adjust the width and height as needed
4. Click "Convert to PNG" to save the converted file

## License

This project is licensed under the MIT License - see the LICENSE file for details. 

Versions
0.2
0.2 Fixes and enhancements
Page Operations:
-Rotate pages (90°, 180°, 270°)
-Duplicate selected pages
-Basic search functionality with text highlighting for found text items using PyMuPDF - If standard searching does not work OCR conversion and searching can be used.
- Added the ability to set the DPI for document previews in Settings
- Extract selected pages into a new PDF by right clicking selected pages
File Management:
-Recent files list
Fixed:
- scrolling was not working properly when a document is zoomed in.
- removed "set Popplar path" icon from the menu bar.
Export Option:
- export selected pages to images using the right button context menu
- “Export current page as image” and “Export all pages as images” from the Edit menu
Personal note: Now it is usable as an everyday application. I have created a shortcut on my desktop to be readily usable.

0.1 Original Commit