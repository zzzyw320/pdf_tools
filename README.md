# PDF Tools

A simple desktop PDF tool built with Python and PySide6.

This software currently provides five main functions:

- Merge PDF
- Split PDF
- Process PDF pages
- PDF to Word
- PDF to JPG

Everyone is welcome to use it.  

You can download the Simplified Chinese version of the soft ware through the provided below:

https://pan.baidu.com/s/1FL3U8HRLb-1Pvqtki1wo9A?pwd=8qv2

If you have any suggestions for improvement, please feel free to contact me at:

**353959580@qq.com**

---

## Features

### 1. Merge PDF
Merge multiple PDF files into one PDF file.

### 2. Split PDF
Split a PDF file in two ways:

- **Automatic split**: split every N pages
- **Manual split**: choose split positions page by page

### 3. Process PDF Pages
Reorganize PDF pages with page thumbnails.

You can:

- move the selected page to the first position
- move the selected page up by one
- delete the selected page
- move the selected page down by one
- move the selected page to the last position

### 4. PDF to Word
Convert a PDF file into a Word document (`.docx`).

### 5. PDF to JPG
Convert each page of a PDF file into JPG images.

---

## Interface Overview

The main window includes:

- **Add PDF File**
- **Merge**
- **Split**
- **Process**
- **PDF to Word**
- **PDF to JPG**

You can also drag PDF files directly into the file list area.

---

## Requirements

This project is based on Python and mainly uses the following libraries:

- PySide6
- PyMuPDF
- pypdf
- pdf2docx
- python-docx
- Pillow
- PyInstaller

Install dependencies with:

```bash
pip install -r requirements.txt