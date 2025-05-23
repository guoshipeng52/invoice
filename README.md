# Invoice Processor with OCR and Ollama

## Overview

This application processes invoice PDF files to extract key information. It uses PaddleOCR for Optical Character Recognition (if the PDF is scanned) and Ollama (with a model like qwen2.5:14b) for natural language processing to understand and structure the extracted text. The extracted data is then used to populate a user-provided Excel template.

The application features a graphical user interface (GUI) built with Tkinter and has been modularized for better maintainability.

## Prerequisites

1.  **Python:** Python 3.9 or newer is recommended.
2.  **Ollama:**
    *   Ollama must be installed and running on your system. You can download it from [https://ollama.com/](https://ollama.com/).
    *   The application is configured by default to use the `qwen2.5:14b` model. You need to pull this model using the command:
        ```bash
        ollama pull qwen2.5:14b
        ```
    *   Ensure Ollama is serving on its default address (`http://localhost:11434`). This can be configured in `config.json` if needed.
3.  **Python Dependencies:**
    *   Clone this repository or download the source code.
    *   Navigate to the project's root directory in your terminal.
    *   Install the required Python packages using the `requirements.txt` file:
        ```bash
        pip install -r requirements.txt
        ```
        (Note: If `requirements.txt` is not yet present, it should be created based on project imports: `tkinter` (usually part of Python), `paddleocr`, `fitz` (PyMuPDF), `requests`, `pandas`, `openpyxl`, `Pillow`, `numpy`, `GPUtil`).

## Running the Application

The main entry point for the application is `main_app.py`.

*   **From the Command Line:**
    1.  Navigate to the root directory of the project in your terminal.
    2.  Run the command:
        ```bash
        python main_app.py
        ```
*   **Using an IDE (e.g., VS Code, PyCharm):**
    1.  Open the project folder in your IDE.
    2.  Locate and open the `main_app.py` file.
    3.  Run `main_app.py` using your IDE's Python interpreter.

## Basic Usage

1.  **Load Ollama Model (First Run or if Ollama was Restarted):**
    *   When you first start the application, a small window will appear prompting you to "Load Ollama qwen2.5". Click this button.
    *   Wait for the confirmation message "模型加载成功！" (Model loaded successfully!). The main application window will then appear.
2.  **Select PDF File:**
    *   In the main window, click the "选择PDF文件" (Select PDF file) button.
    *   Browse and select the invoice PDF file you want to process.
3.  **Run Processing:**
    *   Once a file is selected, the "运行处理" (Run processing) button will be enabled. Click it.
    *   The application will determine if the PDF is scanned (uses OCR) or text-based (extracts text directly).
    *   It will then send the text to Ollama for information extraction. Status messages will update you on the progress.
4.  **View Extracted Data:**
    *   The JSON data returned by Ollama will be displayed in the text area at the bottom of the window.
5.  **Select Excel Template:**
    *   After Ollama processing, you will be prompted to "选择Excel模板文件" (Select Excel template file). Choose your `.xlsx` template.
6.  **Save Output Excel File:**
    *   Next, you will be prompted to "选择保存Excel文件的位置" (Select save location for Excel file). Choose where you want to save the populated Excel file and provide a name.
7.  **Open Saved File (Optional):**
    *   A dialog will ask if you want to open the newly saved Excel file.

## Configuration

The application uses a `config.json` file (created in the same directory as `main_app.py` when settings are first saved or the app runs) to store its settings. You can manually edit this file (if you know what you're doing) or wait for the Settings UI (planned feature) to manage these.

Key configurable items include:
*   Ollama API URL and model name.
*   Default paths for PDF input, Excel templates, and output files.
*   Hardware VRAM threshold for warnings.
*   OCR language.
*   Options like auto-opening the generated Excel file.

## Modules Overview (For Developers)

The codebase has been structured into the following Python modules:
*   `main_app.py`: Main application script, acts as the controller.
*   `app_ui.py`: Handles all Tkinter GUI elements and user interactions.
*   `pdf_handler.py`: Manages PDF processing, including OCR (PaddleOCR) and direct text extraction (PyMuPDF).
*   `ollama_client.py`: Handles all communication with the Ollama API.
*   `excel_writer.py`: Responsible for writing extracted data to Excel templates.
*   `hardware_checker.py`: Contains logic for checking GPU VRAM.
*   `config_manager.py`: Manages loading and saving application settings from/to `config.json`.
*   `hook-paddleocr.py`: PyInstaller hook file to aid in EXE packaging.
*   `ocr-ollama_OLD.py`: The original monolithic script, preserved for reference.
