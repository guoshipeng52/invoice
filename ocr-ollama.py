import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from paddleocr import PaddleOCR
import fitz  # PyMuPDF
import os
import GPUtil
import subprocess
import re
import requests
import json
import pandas as pd
from openpyxl import Workbook, load_workbook
from PIL import Image
import numpy as np
from openpyxl import styles


# Main application class for processing invoices.
# Handles UI, OCR, interaction with Ollama API, and Excel report generation.
class InvoiceProcessor:
    def __init__(self):
        """
        Initializes the InvoiceProcessor.
        Sets up and displays a loading window to ensure the Ollama model is loaded
        before the main application window appears.
        """
        # Create and configure the initial model loading window
        self.load_window = tk.Tk()
        self.load_window.title("加载模型")
        self.load_window.geometry("300x150")
        
        # 居中显示
        self.load_window.eval('tk::PlaceWindow . center')
        
        # 添加标签
        label = tk.Label(self.load_window, text="请先加载模型", font=('Arial', 12))
        label.pack(pady=20)
        
        # 添加加载按钮
        self.load_button = ttk.Button(
            self.load_window, 
            text="加载 Ollama qwen2.5", 
            command=self.load_model
        )
        self.load_button.pack(pady=10)
        
        # 添加进度标签
        self.load_status = tk.Label(self.load_window, text="")
        self.load_status.pack(pady=10)
        
        # 等待加载窗口关闭后再创建主窗口
        self.load_window.mainloop()

    def load_model(self):
        """
        Handles the Ollama model loading process.
        Sends a test request to the Ollama API to ensure the model is available.
        Updates the loading window status and transitions to the main application window on success.
        Shows an error message if loading fails.
        """
        self.load_button.config(state=tk.DISABLED) # Disable button during loading
        self.load_status.config(text="正在加载模型...")
        
        try:
            # Send a test request to Ollama API to check model availability
            response = requests.post(
                "http://localhost:11434/api/generate", # Ollama API endpoint
                json={
                    "model": "qwen2.5:14b",
                    "prompt": "test",
                    "stream": False
                }
            )
            
            if response.status_code == 200: # HTTP OK
                self.load_status.config(text="模型加载成功！")
                # Wait 1 second then close loading window and open main app window
                self.load_window.after(1000, self.start_main_window)
            else:
                raise Exception("模型加载失败") # Trigger error for non-200 status
                
        except Exception as e:
            # Display error message if model loading fails
            messagebox.showerror("错误", f"模型加载失败: {str(e)}\n请确保Ollama服务正在运行且已安装qwen2.5模型")
            self.load_button.config(state=tk.NORMAL) # Re-enable button
            self.load_status.config(text="加载失败，请重试")

    def start_main_window(self):
        """
        Closes the loading window, initializes OCR, and sets up the main application UI.
        """
        self.load_window.destroy() # Close the loading window
        
        # Initialize PaddleOCR for Chinese language, using angle classification
        self.ocr = PaddleOCR(use_angle_cls=True, lang='ch')
        
        # Setup and display the main application window and its UI elements
        self.setup_ui()
        self.root.mainloop() # Start the Tkinter event loop for the main window

    def check_gpu_vram(self, required_vram_gb=4):
        """
        Checks available GPU VRAM using GPUtil and nvidia-smi.
        Shows a warning if VRAM is below required_vram_gb.

        Args:
            required_vram_gb (int, optional): The minimum recommended VRAM in GB. Defaults to 4.
        """
        gpu_vram_gb = 0
        checked_method = "N/A" # Method used to get VRAM info
        try:
            # Attempt to get GPU info using GPUtil
            gpus = GPUtil.getGPUs()
            if gpus:
                gpu_vram_gb = gpus[0].memoryTotal / 1024  # Convert MB to GB
                checked_method = "GPUtil"
            else: # If GPUtil finds no GPUs, try nvidia-smi
                raise Exception("No GPUs found by GPUtil") 
        except Exception as e_gputil:
            # GPUtil failed, try nvidia-smi (often available on systems with NVIDIA GPUs)
            # print(f"GPUtil failed: {e_gputil}, trying nvidia-smi...") # For debugging
            try:
                # For Windows, subprocess.CREATE_NO_WINDOW prevents a console window from appearing
                cflags = subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
                # Run nvidia-smi command to query GPU memory
                result = subprocess.run(
                    ['nvidia-smi', '--query-gpu=memory.total', '--format=csv,noheader,nounits'],
                    capture_output=True, text=True, check=True, creationflags=cflags
                )
                # Output is like "8192" (MiB). Takes the first line if multiple GPUs.
                first_line = result.stdout.strip().split('\n')[0]
                if first_line:
                    gpu_vram_gb = int(first_line) / 1024 # Convert MiB to GB
                    checked_method = "nvidia-smi"
            except Exception as e_nvidia_smi:
                # Both GPUtil and nvidia-smi failed
                # print(f"nvidia-smi failed: {e_nvidia_smi}") # For debugging
                pass # VRAM remains 0, checked_method "N/A"

        # Display warning or info based on VRAM check
        if gpu_vram_gb > 0 and gpu_vram_gb < required_vram_gb:
            # Warn if detected VRAM is positive but less than required
            messagebox.showwarning(
                "Hardware Warning",
                f"Detected {gpu_vram_gb:.1f}GB VRAM using {checked_method}. "
                f"This is below the recommended {required_vram_gb}GB. "
                "Performance might be affected or OCR/AI models may fail."
            )
        elif checked_method == "N/A": # If no method could determine VRAM
             messagebox.showinfo(
                "Hardware Info",
                "Could not automatically determine GPU VRAM. "
                "Please ensure your hardware meets model requirements if you encounter issues."
            )
        elif gpu_vram_gb >= required_vram_gb :
            # VRAM meets requirements, print to console (for debugging/logging)
            print(f"VRAM check: {gpu_vram_gb:.1f}GB detected via {checked_method}. Meets {required_vram_gb}GB requirement.")
        # This case handles where gpu_vram_gb is 0 but a method ran (e.g., integrated graphics).
        # It's covered by the first conditional if required_vram_gb > 0.

    def setup_ui(self):
        """Sets up the main graphical user interface (GUI) for the application."""
        self.root = tk.Tk() # Main application window
        self.root.title("发票信息提取") # Window title
        self.root.geometry("400x200") # Initial window size

        # Button to select a PDF file
        self.select_button = tk.Button(self.root, text="选择PDF文件", command=self.select_pdf)
        self.select_button.pack(pady=20)

        # Button to start processing the selected PDF, initially disabled
        self.run_button = tk.Button(self.root, text="运行处理", command=self.process_pdf, state=tk.DISABLED)
        self.run_button.pack(pady=10)

        # Label to display current status messages
        self.status_label = tk.Label(self.root, text="等待选择文件...")
        self.status_label.pack(pady=10)

        # Label to display the path of the selected file
        self.file_label = tk.Label(self.root, text="未选择文件", wraplength=350) # wraplength for long paths
        self.file_label.pack(pady=10)

        # Label for the Ollama API response display area
        self.result_label = tk.Label(self.root, text="Ollama 返回结果：")
        self.result_label.pack(pady=(10,0))

        # Text area to display results from Ollama, with a scrollbar
        self.result_text = tk.Text(self.root, height=15, width=60)
        self.result_text.pack(pady=5, padx=10)
        
        # Scrollbar for the result_text Text widget
        scrollbar = tk.Scrollbar(self.root, command=self.result_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.result_text.config(yscrollcommand=scrollbar.set)

    def update_status(self, message):
        """
        Updates the status label in the UI with the given message.

        Args:
            message (str): The message to display in the status label.
        """
        self.status_label.config(text=message)
        self.root.update() # Force UI update

    def pdf_to_images(self, pdf_path):
        """
        Converts each page of a PDF file into a list of images (as numpy arrays).
        Used for OCR processing of scanned PDFs.

        Args:
            pdf_path (str): The file path to the PDF.

        Returns:
            list: A list of images, where each image is a NumPy array.
        """
        doc = fitz.open(pdf_path) # Open PDF using PyMuPDF
        images = []
        for page in doc:
            pix = page.get_pixmap() # Render page to a pixmap
            # Convert pixmap to PIL Image, then to NumPy array
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            images.append(np.array(img))
        doc.close()
        return images

    def select_pdf(self):
        """
        Opens a file dialog for the user to select a PDF file.
        Updates UI elements to reflect the selected file and enables the run button.
        """
        # Open file dialog to select a PDF
        self.pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if self.pdf_path:
            # If a file is selected, update UI labels and enable run button
            self.file_label.config(text=f"已选择: {self.pdf_path}")
            self.run_button.config(state=tk.NORMAL) # Enable run button
            self.update_status("文件已选择，请点击运行处理")
        else:
            # If no file is selected, reset UI elements
            self.file_label.config(text="未选择文件")
            self.run_button.config(state=tk.DISABLED) # Disable run button
            self.update_status("等待选择文件...")

    def is_scanned_pdf(self, pdf_path):
        """
        Checks if a PDF is likely a scanned document by trying to extract text.
        If the amount of extractable text is very small, it's considered scanned.

        Args:
            pdf_path (str): The file path to the PDF.

        Returns:
            bool: True if the PDF is likely scanned, False otherwise.
        """
        doc = fitz.open(pdf_path)
        text_content = ""
        for page in doc:
            text_content += page.get_text()
        doc.close()
        # If total text extracted is less than 100 characters, assume it's a scanned PDF.
        return len(text_content.strip()) < 100

    def process_pdf(self):
        """
        Main method to process the selected PDF.
        It handles OCR/text extraction, calls the Ollama API for information extraction,
        parses the JSON response, and populates an Excel template with the data.
        """
        # Ensure a PDF file has been selected
        if not hasattr(self, 'pdf_path') or not self.pdf_path:
            self.update_status("请先选择PDF文件")
            return

        self.update_status("正在处理PDF...")
        self.run_button.config(state=tk.DISABLED) # Disable run button during processing
        
        # Create a 'working' directory if it doesn't exist, for temporary files
        work_dir = "working"
        if not os.path.exists(work_dir):
            os.makedirs(work_dir)

        text_content = [] # Stores extracted text lines from the PDF
        
        # Determine if PDF is scanned or text-based, then extract text accordingly.
        if self.is_scanned_pdf(self.pdf_path):
            self.update_status("检测到扫描PDF，正在进行OCR识别...")
            images = self.pdf_to_images(self.pdf_path) # Convert PDF to images
            for img in images:
                # Perform OCR on each image
                result = self.ocr.ocr(img)
                # Append recognized text lines to text_content
                for line_info in result[0]: # result[0] contains list of [bbox, (text, confidence)]
                    text_content.append(line_info[1][0]) # Append only the text
        else:
            # If PDF is text-based, extract text directly
            self.update_status("检测到可复制PDF，正在直接提取文本...")
            doc = fitz.open(self.pdf_path)
            extracted_pages_text = []
            for page_num in range(len(doc)):
                page = doc.load_page(page_num)
                extracted_pages_text.append(page.get_text("text"))
            doc.close()
            # Split combined text from all pages into lines for consistency with OCR output
            full_text = "\n".join(extracted_pages_text)
            text_content.extend(full_text.splitlines())

        # Save the raw extracted text to a .txt file (for debugging/logging)
        txt_path = os.path.join(work_dir, "ocr_result.txt")
        with open(txt_path, "w", encoding="utf-8") as f:
            f.write("\n".join(text_content))

        self.update_status("正在调用Ollama API...")

        # This prompt instructs the Ollama model (qwen2.5:14b) to act as a JSON extractor.
        # It specifies the desired JSON structure with keys corresponding to invoice fields.
        # The model should return *only* the JSON object.
        # Key names in this JSON structure are used in `excel_keys_order` for mapping.
        prompt = """Please analyze the following OCR text from an invoice. Extract the specified information and return it as a single, valid JSON object. Do not include any text or explanations before or after the JSON object. The JSON object should conform to the following structure:
{
  "buyer_name": "购买方名称",
  "buyer_tin": "购买方的纳税人识别号",
  "buyer_bank": "购买方开户行",
  "buyer_bank_account": "购买方开户行账号",
  "invoice_code": "发票代码",
  "item_name": "货物或应税劳务、服务名称",
  "item_quantity": "数量 (If not present, default to '1')",
  "item_unit_price": "单价 (If not present, use the value from '金额')",
  "seller_name": "销售方名称",
  "seller_tin": "销售方的纳税人识别号",
  "seller_bank": "销售方开户行",
  "seller_bank_account": "销售方开户行账号",
  "invoice_date": "开票日期 (Return in 'YYYYMMDD' format)",
  "invoice_number": "发票号码 (If 'No.xxx' format exists, return the number part; otherwise, return a blank string ' ')",
  "total_amount_with_tax": "价税合计 (Return as a number/numeric string)",
  "tax_rate": "税率 (e.g., '0%', '免税', '3%')",
  "tax_amount": "税额 (If not present, default to '0')",
  "total_amount_without_tax": "金额 (Return as a number/numeric string)",
  "invoice_type": "发票类型 (Return only '普通发票' or '专用发票')"
}

Text content:
""" + "\\n".join(text_content) # Append the extracted text to the prompt

        try:
            # Call the Ollama API
            response = requests.post(
                "http://localhost:11434/api/generate", # Ollama API endpoint
                json={
                    "model": "qwen2.5:14b", # Specify the model to use
                    "prompt": prompt,       # The constructed prompt
                    "stream": False         # Expect a single response, not a stream
                }
            )
            
            if response.status_code == 200: # HTTP OK
                result = response.json() # Get JSON response from Ollama
                try:
                    # Attempt to parse the 'response' field from Ollama, which should be a JSON string.
                    extracted_data_str = result['response']
                    
                    # Clean the JSON string: some models might wrap JSON in markdown ```json ... ```
                    if extracted_data_str.startswith("```json"):
                        extracted_data_str = extracted_data_str[7:] # Remove ```json
                        if extracted_data_str.endswith("```"):
                            extracted_data_str = extracted_data_str[:-3] # Remove ```
                    extracted_data_str = extracted_data_str.strip() # Remove leading/trailing whitespace
                    
                    # Parse the cleaned string into a Python dictionary
                    extracted_info = json.loads(extracted_data_str)
                    
                    # Display the extracted key-value pairs and raw JSON in the UI text area
                    self.result_text.delete('1.0', tk.END) # Clear previous results
                    self.result_text.insert('1.0', "提取结果 (JSON):\n" + "="*50 + "\n")
                    for key, value in extracted_info.items():
                        self.result_text.insert(tk.END, f"{key}: {value}\n")
                    # Also display the raw JSON for debugging/verification
                    self.result_text.insert(tk.END, "="*50 + "\nRaw JSON:\n" + json.dumps(extracted_info, indent=2, ensure_ascii=False))
                    self.result_text.see("1.0") # Scroll to the top of the text area

                    # `excel_keys_order` defines the exact order of JSON keys as they should appear
                    # as columns (A1, B1, C1, ...) in the temporary Excel file.
                    # This ensures consistent mapping when these values are later read for the final template.
                    excel_keys_order = [
                        "buyer_name", "buyer_tin", "buyer_bank", "buyer_bank_account",
                        "invoice_code", "item_name", "item_quantity", "item_unit_price",
                        "seller_name", "seller_tin", "seller_bank", "seller_bank_account",
                        "invoice_date", "invoice_number", "total_amount_with_tax",
                        "tax_rate", "tax_amount", "total_amount_without_tax", "invoice_type"
                    ]

                    # Create a temporary Excel workbook to store all 19 extracted fields.
                    # This file can be used for debugging or as an intermediary step if needed,
                    # though the final user template is now populated directly from `extracted_info`.
                    temp_wb = Workbook()
                    temp_ws = temp_wb.active # Get the active sheet
                    
                    # Write data from `extracted_info` to the first row of `temp_ws`
                    # according to the order defined in `excel_keys_order`.
                    for i, key in enumerate(excel_keys_order):
                        col_letter = chr(65 + i)  # Convert index to Excel column letter (A, B, C...)
                        cell_pos = f"{col_letter}1" # e.g., A1, B1
                        # Use .get(key, "") to default to an empty string if a key is missing in the JSON
                        temp_ws[cell_pos] = extracted_info.get(key, "") 
                    
                except json.JSONDecodeError as e:
                    # Handle error if Ollama response is not valid JSON
                    self.update_status(f"JSON解析错误: {str(e)}")
                    self.result_text.delete('1.0', tk.END)
                    self.result_text.insert('1.0', f"无法解析返回的JSON: {result['response']}\nError: {str(e)}")
                    self.run_button.config(state=tk.NORMAL) # Re-enable run button
                    return # Stop further processing
                except KeyError as e:
                    # Handle error if the 'response' key is missing from Ollama's output
                    self.update_status(f"API响应格式错误: Missing key {str(e)}")
                    self.result_text.delete('1.0', tk.END)
                    self.result_text.insert('1.0', f"API响应中缺少预期的 'response' 键: {result}\nError: {str(e)}")
                    self.run_button.config(state=tk.NORMAL) # Re-enable run button
                    return # Stop further processing

                # Save the temporary Excel file with all 19 extracted fields
                temp_file = "temp_data.xlsx"
                temp_wb.save(temp_file)
                
                # Ask user to select an Excel template file to populate
                template_path = filedialog.askopenfilename(
                    filetypes=[("Excel files", "*.xlsx")],
                    title="选择Excel模板文件"
                )
                
                if template_path: # If a template file is selected
                    wb = load_workbook(template_path) # Load the user's template
                    ws = wb.active # Get the active sheet of the template
                    
                    # Clear existing data from the template (rows 3 onwards)
                    # This loop iterates from max_row down to 3 to correctly delete rows.
                    # If ws.max_row < 3, this loop won't run.
                    for row_idx in range(ws.max_row, 2, -1): # Start from ws.max_row, end at 3, step -1
                        ws.delete_rows(row_idx)

                    # Ensure row 3 exists (it might have been deleted if template had less than 3 rows)
                    # If row 3 doesn't exist after deletion, this implies an empty or very small template.
                    # We'll proceed assuming row 3 is the target, which is consistent with current logic.
                    # If ws.max_row became < 3, cell access like ws['E3'] would create the row.

                    # Unlock cells in the third row of the template (if they exist)
                    if ws.max_row >= 3: # Check if row 3 actually exists or will be created
                        for cell in ws[3]: # ws[3] provides all cells in row 3
                            cell.protection = styles.Protection(locked=False)
                    
                    # Populate the user's Excel template (ws) directly from `extracted_info` dictionary.
                    # This mapping follows the original logic which used specific cells from temp_ws.
                    # The JSON keys correspond to the fields defined in `excel_keys_order`.
                    # For example, temp_ws['A1'] (buyer_name) is now extracted_info.get('buyer_name', '').
                    
                    ws['E3'] = extracted_info.get('buyer_name', '')
                    ws['F3'] = extracted_info.get('buyer_tin', '')
                    # buyer_bank and buyer_bank_account are present in `extracted_info` and `temp_data.xlsx`
                    # but were not part of the original mapping to the user template (ws).
                    
                    ws['B3'] = extracted_info.get('invoice_code', '')
                    ws['I3'] = f"{extracted_info.get('item_name', '')}*00010-" # Append suffix
                    ws['L3'] = extracted_info.get('item_quantity', '')
                    ws['M3'] = extracted_info.get('item_unit_price', '')
                    
                    ws['G3'] = extracted_info.get('seller_name', '')
                    ws['H3'] = extracted_info.get('seller_tin', '')
                    # seller_bank and seller_bank_account are similar to buyer_bank above.
                    
                    ws['D3'] = extracted_info.get('invoice_date', '')
                    ws['C3'] = extracted_info.get('invoice_number', '')
                    ws['Q3'] = extracted_info.get('total_amount_with_tax', '')
                    ws['O3'] = extracted_info.get('tax_rate', '')
                    ws['P3'] = extracted_info.get('tax_amount', '')
                    ws['N3'] = extracted_info.get('total_amount_without_tax', '')
                    ws['A3'] = extracted_info.get('invoice_type', '')
                    
                    # Ask user where to save the populated Excel file
                    save_path = filedialog.asksaveasfilename(
                        defaultextension=".xlsx",
                        filetypes=[("Excel files", "*.xlsx")],
                        title="选择保存Excel文件的位置"
                    )
                    
                    if save_path: # If a save location is chosen
                        wb.save(save_path) # Save the populated template
                        self.update_status(f"处理完成！结果已保存到: {save_path}")
                        
                        # Create a confirmation dialog asking to open the saved file
                        open_dialog = tk.Toplevel(self.root)
                        open_dialog.title("文件已保存")
                        open_dialog.geometry("300x150")
                        
                        open_dialog.transient(self.root) # Set dialog to be on top of the main window
                        open_dialog.grab_set() # Make dialog modal
                        
                        tk.Label(
                            open_dialog, 
                            text="文件保存成功！\n是否立即打开文件？",
                            font=('Arial', 10)
                        ).pack(pady=20)
                        
                        btn_frame = tk.Frame(open_dialog) # Frame for buttons
                        btn_frame.pack(pady=10)
                        
                        def open_file_action(): # Renamed to avoid conflict with os.open
                            os.startfile(save_path) # Open the saved file
                            open_dialog.destroy()
                            
                        def close_dialog_action(): # Renamed for clarity
                            open_dialog.destroy()
                        
                        # "Open" button
                        tk.Button(
                            btn_frame, 
                            text="打开", 
                            command=open_file_action, # Use renamed function
                            width=10,
                            relief="raised", # Standard button appearance
                            bg="#e1e1e1"     # Light grey background
                        ).pack(side=tk.LEFT, padx=10)
                        
                        # "Close" button
                        tk.Button(
                            btn_frame, 
                            text="关闭", 
                            command=close_dialog_action, # Use renamed function
                            width=10,
                            relief="raised",
                            bg="#e1e1e1"
                        ).pack(side=tk.LEFT, padx=10)
                        
                        # Center the confirmation dialog
                        window_width = 300
                        window_height = 150
                        screen_width = self.root.winfo_screenwidth()
                        screen_height = self.root.winfo_screenheight()
                        x = (screen_width - window_width) // 2
                        y = (screen_height - window_height) // 2
                        open_dialog.geometry(f"{window_width}x{window_height}+{x}+{y}")
                        
                        self.root.wait_window(open_dialog) # Wait for the dialog to close
                        
                    else: # If user cancels saving
                        self.update_status("已取消保存Excel文件")
                    
                    # Delete the temporary Excel file
                    os.remove(temp_file)
            else: # If Ollama API call failed (not status 200)
                self.update_status("API调用失败")

        except Exception as e: # Catch any other exceptions during the process
            self.update_status(f"发生错误: {str(e)}")

        self.run_button.config(state=tk.NORMAL) # Re-enable run button after processing

    def run(self):
        """Starts the Tkinter main event loop for the application."""
        self.root.mainloop()

# Standard Python entry point: create an instance of InvoiceProcessor and run it.
if __name__ == "__main__":
    processor = InvoiceProcessor()
    # The run method is currently empty in the provided code, 
    # but typically it would start the Tkinter main loop (self.root.mainloop()).
    # For this class structure, __init__ already starts a mainloop for the load_window,
    # and start_main_window starts the mainloop for self.root.
    # processor.run() # This call is not strictly necessary if mainloop is handled within methods.
    # However, if run() is intended to be the primary entry point to start the UI,
    # then self.root.mainloop() should be called within it.
    # Current structure: load_window.mainloop() in __init__, then root.mainloop() in start_main_window.
    # The run() method as defined (self.root.mainloop()) is redundant if called after setup_ui,
    # as setup_ui is called by start_main_window which then calls self.root.mainloop().
    # For clarity, if processor.run() is the intended public method to start the app,
    # it should encompass the logic that eventually calls self.root.mainloop().
    # Given the current flow, processor instance creation itself kicks off the UI.
    # No explicit processor.run() is needed if __init__ directly or indirectly starts the main loop.
    # The current code calls self.root.mainloop() in start_main_window().
    # The final self.root.mainloop() in the `run` method is therefore okay.
    processor.run() # This will call self.root.mainloop()
