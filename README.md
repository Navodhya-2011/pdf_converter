# File Converter

This Python application allows you to convert multiple `.pptx` (PowerPoint) and `.docx` (Word) files to PDF format. The application supports Windows, macOS, and Linux.

## Features

- Convert `.pptx` and `.docx` files to PDF.
- Select multiple files for batch conversion.
- Save converted PDFs to a user-selected directory.
- Easy-to-use graphical user interface (GUI).

## Prerequisites

### General Requirements

1. **Python 3.x**: Ensure Python 3.x is installed on your system. You can download it from the [official Python website](https://www.python.org/).

2. **pip**: Ensure `pip` is installed and updated:
   ```bash
   python -m ensurepip --upgrade
   ```

## Installation Guide

Follow these instructions to set up the application on your system.

### Windows

1. **Install LibreOffice**:
   Download and install LibreOffice from the [official website](https://www.libreoffice.org/download/download-libreoffice/).
   
3. **Install unoconv**:
	•	Download unoconv from [unoconv GitHub repository](https://github.com/unoconv/unoconv).
	•	Place the unoconv executable in a directory that’s included in your system’s PATH.
4. **Install Python dependencies**:
   
   Open Command Prompt and run:
   ```bash
   pip install tk
   ```

### macOS

1. **Install Homebrew (if not already installed)**:
   
   Open Terminal and run:
   ```bash
   /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
   ```
   
2. **Install LibreOffice and unoconv:**:

   Run the following commands in Terminal:
   ```bash
   brew install --cask libreoffice
   brew install unoconv
   ```
   
3. **Install Python dependencies**:
   
   Open Command Prompt and run:
   ```bash
   pip install tk
   ```

### Linux

1. **Install LibreOffice and unoconv**:
   
   Use your package manager (e.g., apt for Debian/Ubuntu, dnf for Fedora):
   
   ```bash
   # For Debian/Ubuntu
   sudo apt update
   sudo apt install libreoffice unoconv
   ```

   ```bash
   # For Fedora
   sudo dnf install libreoffice unoconv
   ```
   
2. **Install Python dependencies**:
   
   Open Command Prompt and run:
      ```bash
   pip install tk
   ```

## How to Run the Application

1. **Clone the Repository**:

   Open your terminal or command prompt and run:
   ```bash
   git clone https://github.com/IT22085580BGDNBeligolla/pdf_converter.git
   cd pdf_converter
   ```

2. **Run the Application**:
   
   Run the Python script:
   
   ```bash
   # for single file converter
   python converter.py
   ```

     ```bash
   # for multiple file converter
   python converterm.py
   ```

3. **Select Files and Convert**:

   •	Use the GUI to select .pptx or .docx files.
   •	Choose the directory where you want to save the converted PDFs.
   •	Click “Convert to PDF” to start the conversion process.

## Acknowledgments

   •	Thanks to the developers of [LibreOffice](https://www.libreoffice.org/) and [unoconv](https://github.com/unoconv/unoconv) for their great tools.
