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



# Main class for the Invoice Processing application.
# Handles UI, OCR, Ollama interaction, and PDF/Excel processing.
class InvoiceProcessor:
    def __init__(self):
        # Initialize the loading window.
        # This window is shown first to allow the user to load the Ollama model
        # before the main application UI appears.
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
        
        # Start the Tkinter event loop for the loading window.
        # The main application window will be created after this loop ends (on model load).
        self.load_window.mainloop()

    def load_model(self):
        # This method is called when the "Load Ollama qwen2.5" button is clicked.
        # It disables the button, updates status, and attempts to connect to Ollama.
        self.load_button.config(state=tk.DISABLED) # Disable button during loading
        self.load_status.config(text="正在加载模型...") # Update status label
        
        try:
            # Test if Ollama is running and the specified model is available.
            # Sends a simple test prompt to the Ollama API.
            response = requests.post(
                "http://localhost:11434/api/generate", # Ollama API endpoint
                json={
                    "model": "qwen2.5:14b",
                    "prompt": "test",
                    "stream": False
                }
            )
            
            if response.status_code == 200:
                # If the API call is successful, update status and schedule main window creation.
                self.load_status.config(text="模型加载成功！")
                # Call start_main_window after 1 second (1000 ms) to allow user to see the success message.
                self.load_window.after(1000, self.start_main_window)
            else:
                # If API call fails, raise an exception to be caught by the except block.
                raise Exception("模型加载失败")
                
        except Exception as e:
            # Handle exceptions during model loading (e.g., Ollama not running, model not found).
            messagebox.showerror("错误", f"模型加载失败: {str(e)}\n请确保Ollama服务正在运行且已安装qwen2.5模型")
            self.load_button.config(state=tk.NORMAL) # Re-enable the load button
            self.load_status.config(text="加载失败，请重试") # Update status label

    def start_main_window(self):
        # This method is called after the Ollama model is successfully checked/loaded.
        # It closes the loading window, initializes critical components, and sets up the main UI.
        
        # Close the loading window.
        self.load_window.destroy()
        
        # Initialize PaddleOCR with Chinese language support and angle classification.
        self.ocr = PaddleOCR(use_angle_cls=True, lang='ch')
        
        # Perform GPU VRAM check before setting up the main UI.
        self.check_gpu_vram()

        # Create and display the main application UI.
        self.setup_ui()
        # Start the Tkinter event loop for the main application window.
        self.root.mainloop()

    def setup_ui(self):
        # Sets up the main User Interface for the application.
        # This includes creating the main window, buttons, labels, and text areas.
        self.root = tk.Tk() # Create the main Tkinter window.
        self.root.title("发票信息提取") # Set the window title.
        self.root.geometry("400x200") # Set the initial window size.

        # Create and pack the "Select PDF File" button.
        # When clicked, it calls the self.select_pdf method.
        self.select_button = tk.Button(self.root, text="选择PDF文件", command=self.select_pdf)
        self.select_button.pack(pady=20) # Add padding in the y-direction.

        # Create and pack the "Run Processing" button.
        # It's initially disabled and enabled after a PDF is selected.
        # When clicked, it calls the self.process_pdf method.
        self.run_button = tk.Button(self.root, text="运行处理", command=self.process_pdf, state=tk.DISABLED)
        self.run_button.pack(pady=10)

        # Create and pack the status label.
        # This label displays messages about the application's current state.
        self.status_label = tk.Label(self.root, text="等待选择文件...")
        self.status_label.pack(pady=10)

        # Create and pack the file label.
        # This label displays the path of the selected PDF file.
        # wraplength ensures text wraps if it's too long.
        self.file_label = tk.Label(self.root, text="未选择文件", wraplength=350)
        self.file_label.pack(pady=10)

        # Create and pack the label for Ollama results.
        self.result_label = tk.Label(self.root, text="Ollama 返回结果：")
        self.result_label.pack(pady=(10,0)) # Padding: 10px top, 0px bottom.

        # Create and pack the text area for displaying Ollama's response.
        # Height is 15 lines, width is 60 characters.
        self.result_text = tk.Text(self.root, height=15, width=60)
        self.result_text.pack(pady=5, padx=10) # Add padding.
        
        # Create and add a vertical scrollbar to the result_text text area.
        scrollbar = tk.Scrollbar(self.root, command=self.result_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y) # Pack scrollbar to the right, fill vertically.
        self.result_text.config(yscrollcommand=scrollbar.set) # Link scrollbar to text area.

    # Updates the status label with the given message.
    # Updates the status label with the given message.
    def update_status(self, message):
        # Parameters:
        #   message (str): The message to display in the status label.
        self.status_label.config(text=message) # Set the text of the status label.
        self.root.update() # Force Tkinter to update the UI immediately.

    # Opens a file dialog for the user to select a PDF file.
    # Updates UI elements based on whether a file was selected.
    def select_pdf(self):
        # Ask the user to select a PDF file using a standard file dialog.
        # Only files with the ".pdf" extension are shown by default.
        self.pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        
        if self.pdf_path: # If a file was selected (path is not empty).
            self.file_label.config(text=f"已选择: {self.pdf_path}") # Update file label with selected path.
            self.run_button.config(state=tk.NORMAL) # Enable the "Run Processing" button.
            self.update_status("文件已选择，请点击运行处理") # Update status message.
        else: # If no file was selected (dialog was cancelled).
            self.file_label.config(text="未选择文件") # Reset file label.
            self.run_button.config(state=tk.DISABLED) # Keep "Run Processing" button disabled.
            self.update_status("等待选择文件...") # Reset status message.

    # Processes the selected PDF: performs OCR if scanned, or direct text extraction if not.
    # Then, sends the text to Ollama for information extraction and processes the response
    # to populate an Excel template.
    def process_pdf(self):
        # Check if a PDF path has been selected.
        if not hasattr(self, 'pdf_path') or not self.pdf_path:
            self.update_status("请先选择PDF文件") # Prompt user to select a file.
            return

        self.update_status("正在处理PDF...") # Update status message.
        self.run_button.config(state=tk.DISABLED) # Disable run button during processing.
        
        # Create a working directory if it doesn't exist.
        # This directory is used to store intermediate files like ocr_result.txt.
        work_dir = "working"
        if not os.path.exists(work_dir):
            os.makedirs(work_dir)

        text_content = [] # Initialize a list to store lines of text from the PDF.
        
        # Determine if the PDF is scanned and process accordingly.
        if self.is_scanned_pdf(self.pdf_path):
            # If PDF is scanned, perform OCR.
            self.update_status("检测到扫描PDF，正在进行OCR识别...")
            # Convert PDF pages to images.
            images = self.pdf_to_images(self.pdf_path)
            for img in images: # For each image (page).
                # Perform OCR using PaddleOCR.
                result = self.ocr.ocr(img)
                # PaddleOCR returns a list of results per image.
                # Each result contains bounding boxes, text, and confidence.
                # We are interested in the text part (line[1][0]).
                for line in result[0]: 
                    text_content.append(line[1][0]) # Append extracted text line.
        else:
            # If PDF is not scanned (i.e., contains selectable text).
            self.update_status("检测到可复制PDF，正在直接提取文本...")
            doc = fitz.open(self.pdf_path) # Open PDF with PyMuPDF.
            extracted_pages_text = [] # Store text from each page.
            for page_num in range(len(doc)): # Iterate through pages.
                page = doc.load_page(page_num) # Load page.
                extracted_pages_text.append(page.get_text("text")) # Extract raw text.
            doc.close() # Close PDF.
            
            # The OCR branch results in a list of text lines.
            # To maintain consistency for the subsequent prompt construction,
            # we combine all extracted text and then split it into lines.
            full_text = "\n".join(extracted_pages_text) # Join text from all pages.
            text_content.extend(full_text.splitlines()) # Split into lines and add to text_content list.

        # Save the extracted text (either from OCR or direct extraction) to a .txt file.
        # This can be useful for debugging or intermediate inspection.
        txt_path = os.path.join(work_dir, "ocr_result.txt")
        with open(txt_path, "w", encoding="utf-8") as f: # Open file in write mode with UTF-8 encoding.
            f.write("\n".join(text_content)) # Write each text line, joined by newlines.

        self.update_status("正在调用Ollama API...") # Update status message.

        # Construct the prompt for the Ollama API to request JSON output.
        prompt = f'''Please extract the following information from the provided text and return it as a single, minified JSON object. Do not include any explanatory text before or after the JSON object. The JSON keys should be exactly as specified below:
"buyer_name", "buyer_tax_id", "buyer_bank_name", "buyer_bank_account", "invoice_code", "item_name", "item_quantity", "item_unit_price", "seller_name", "seller_tax_id", "seller_bank_name", "seller_bank_account", "invoice_date_YYYYMMDD", "invoice_number", "total_amount_with_tax", "tax_rate", "tax_amount", "total_amount_without_tax", "invoice_type"

If a value is not found or not applicable, use an empty string "" or null for that key. For numerical fields like quantity, unit_price, amounts, tax_rate, if not found, use 0 or null. Ensure numerical values are returned as numbers, not strings, if possible. 'item_quantity' should default to 1 if not found. 'invoice_date_YYYYMMDD' should be in YYYYMMDD format. 'invoice_type' should be "普通发票" or "专用发票".

Text content:
{ "\n".join(text_content) }'''

        try:
            # Call the Ollama API using a POST request.
            response = requests.post(
                "http://localhost:11434/api/generate", # Ollama API endpoint.
                json={ # Request payload.
                    "model": "qwen2.5:14b", # Specify the Ollama model to use.
                    "prompt": prompt,       # The constructed prompt.
                    "stream": False,        # Do not stream the response.
                    "format": "json"        # Request JSON format from Ollama
                }
            )
            
            if response.status_code == 200: # If the API call was successful.
                result = response.json() # Parse the JSON response from Ollama's outer layer.
                
                try:
                    ollama_response_json_str = result['response'].strip()
                    # Sometimes Ollama might wrap its JSON in backticks or add stray text, try to clean it.
                    if ollama_response_json_str.startswith("```json"):
                        ollama_response_json_str = ollama_response_json_str[7:]
                    if ollama_response_json_str.endswith("```"):
                        ollama_response_json_str = ollama_response_json_str[:-3]
                    ollama_response_json_str = ollama_response_json_str.strip()
                    
                    extracted_data = json.loads(ollama_response_json_str) # Parse the actual JSON content

                    # Update UI display
                    self.result_text.delete('1.0', tk.END)
                    self.result_text.insert('1.0', "提取结果 (JSON):\n" + "="*50 + "\n")
                    self.result_text.insert(tk.END, json.dumps(extracted_data, indent=2, ensure_ascii=False)) # Pretty print JSON
                    self.result_text.insert(tk.END, "\n" + "="*50)
                    self.result_text.see("1.0")

                    # Update temporary Excel file creation
                    temp_wb = Workbook()
                    temp_ws = temp_wb.active
                    
                    # Define the order of keys as expected by the original Excel mapping logic
                    keys_in_order = [
                        "buyer_name", "buyer_tax_id", "buyer_bank_name", "buyer_bank_account", 
                        "invoice_code", "item_name", "item_quantity", "item_unit_price", 
                        "seller_name", "seller_tax_id", "seller_bank_name", "seller_bank_account", 
                        "invoice_date_YYYYMMDD", "invoice_number", "total_amount_with_tax", 
                        "tax_rate", "tax_amount", "total_amount_without_tax", "invoice_type"
                    ]

                    for i, key in enumerate(keys_in_order):
                        cell_char = chr(65 + i) # A, B, C...
                        temp_ws[f"{cell_char}1"] = extracted_data.get(key, "") # Use .get for safety, write to A1, B1, etc.

                    temp_file = "temp_data.xlsx"
                    temp_wb.save(temp_file)

                except json.JSONDecodeError as e:
                    self.update_status(f"Ollama返回JSON解析失败: {str(e)}")
                    self.result_text.delete('1.0', tk.END)
                    self.result_text.insert('1.0', f"Ollama did not return valid JSON.\nResponse:\n{result.get('response', 'No response content found in result.')}")
                    self.run_button.config(state=tk.NORMAL)
                    return # Stop further processing if JSON is invalid
                except KeyError as e: # This would catch issues if 'response' key itself is missing from 'result'
                    self.update_status(f"Ollama返回数据结构错误: {str(e)}")
                    self.result_text.delete('1.0', tk.END)
                    self.result_text.insert('1.0', f"Ollama's response was missing an expected key: {str(e)}.\nFull Response:\n{result}")
                    self.run_button.config(state=tk.NORMAL)
                    return
                
                # Ask the user to select an Excel template file.
                template_path = filedialog.askopenfilename(
                    filetypes=[("Excel files", "*.xlsx")], # Filter for .xlsx files.
                    title="选择Excel模板文件" # Dialog title.
                )
                
                if template_path: # If a template file was selected.
                    wb = load_workbook(template_path) # Load the selected Excel template.
                    ws = wb.active # Get the active worksheet.
                    
                    # Clear existing data from the template, keeping only the first two rows (presumably headers).
                    # It deletes rows starting from the 3rd row to the max row.
                    for row_idx in range(ws.max_row, 2, -1): # Iterate downwards to avoid issues with row deletion
                        ws.delete_rows(row_idx)

                    # Ensure the third row exists for writing (it might have been deleted if template had <3 rows)
                    if ws.max_row < 2: # If template had 0 or 1 row
                        # This case might need more robust handling if headers are expected to be preserved
                        # For now, we just ensure a row 3 exists if it was deleted.
                         pass # Or create empty rows up to row 3 if necessary
                    elif ws.max_row == 2: # If template had exactly 2 rows
                        pass # Row 3 will be the next available row.
                    
                    # Unlock cells in the third row of the template (if they exist).
                    # This assumes data will be written to the 3rd row.
                    # If ws.max_row was < 3, this might error or do nothing.
                    # A more robust way would be to ensure row 3 exists or append a new row.
                    # For this commenting task, I'll assume row 3 is the target.
                    if ws.max_row >= 3: # Check if row 3 actually exists
                        for cell in ws[3]: # Iterate through cells in the third row.
                            cell.protection = styles.Protection(locked=False) # Unlock cell.
                    
                    # Map and write data from the temporary Excel (Ollama's output) to specific cells in the template.
                    # The comments A1->E3 etc. indicate source cell in temp_ws and target cell in ws.
                    ws['E3'] = temp_ws['A1'].value  # Buyer Name
                    ws['F3'] = temp_ws['B1'].value  # Buyer Tax ID
                    ws['B3'] = temp_ws['E1'].value  # Invoice Code
                    ws['I3'] = f"{temp_ws['F1'].value}*00010-"  # Goods/Service Name + "*00010-"
                    ws['L3'] = temp_ws['G1'].value  # Quantity
                    ws['M3'] = temp_ws['H1'].value  # Unit Price
                    ws['G3'] = temp_ws['I1'].value  # Seller Name
                    ws['H3'] = temp_ws['J1'].value  # Seller Tax ID
                    ws['D3'] = temp_ws['M1'].value  # Invoice Date (YYYYMMDD)
                    ws['C3'] = temp_ws['N1'].value  # Invoice Number
                    ws['Q3'] = temp_ws['O1'].value  # Total Amount (incl. tax)
                    ws['O3'] = temp_ws['P1'].value  # Tax Rate
                    ws['P3'] = temp_ws['Q1'].value  # Tax Amount
                    ws['N3'] = temp_ws['R1'].value  # Amount (excl. tax)
                    ws['A3'] = temp_ws['S1'].value  # Invoice Type ("普通发票" or "专用发票")

                    
                    # Ask the user where to save the newly populated Excel file.
                    save_path = filedialog.asksaveasfilename(
                        defaultextension=".xlsx", # Default file extension.
                        filetypes=[("Excel files", "*.xlsx")], # Filter for .xlsx files.
                        title="选择保存Excel文件的位置" # Dialog title.
                    )
                    
                    if save_path: # If a save path was provided.
                        wb.save(save_path) # Save the workbook to the chosen path.
                        self.update_status(f"处理完成！结果已保存到: {save_path}") # Update status.
                        
                        # Create a confirmation dialog asking if the user wants to open the saved file.
                        open_dialog = tk.Toplevel(self.root) # Create a new top-level window.
                        open_dialog.title("文件已保存") # Dialog title.
                        open_dialog.geometry("300x150") # Dialog size.
                        
                        # Make the dialog modal (blocks interaction with the main window).
                        open_dialog.transient(self.root)
                        open_dialog.grab_set()
                        
                        # Add a label to the dialog.
                        tk.Label(
                            open_dialog, 
                            text="文件保存成功！\n是否立即打开文件？", # Confirmation message.
                            font=('Arial', 10)
                        ).pack(pady=20)
                        
                        # Create a frame to hold the "Open" and "Close" buttons.
                        btn_frame = tk.Frame(open_dialog)
                        btn_frame.pack(pady=10)
                        
                        # Define function to open the saved file.
                        def open_file():
                            os.startfile(save_path) # Use os.startfile for cross-platform opening.
                            open_dialog.destroy()   # Close the confirmation dialog.
                            
                        # Define function to simply close the dialog.
                        def close_dialog():
                            open_dialog.destroy()
                        
                        # Create the "Open" button.
                        tk.Button(
                            btn_frame, 
                            text="打开",      # Button text.
                            command=open_file, # Command to execute on click.
                            width=10,        # Button width.
                            relief="raised", # Button style.
                            bg="#e1e1e1"     # Button background color.
                        ).pack(side=tk.LEFT, padx=10) # Pack button to the left with padding.
                        
                        # Create the "Close" button.
                        tk.Button(
                            btn_frame, 
                            text="关闭",          # Button text.
                            command=close_dialog, # Command to execute on click.
                            width=10,            # Button width.
                            relief="raised",     # Button style.
                            bg="#e1e1e1"         # Button background color.
                        ).pack(side=tk.LEFT, padx=10) # Pack button to the left with padding.
                        
                        # Center the confirmation dialog on the screen.
                        window_width = 300
                        window_height = 150
                        screen_width = self.root.winfo_screenwidth()    # Get screen width.
                        screen_height = self.root.winfo_screenheight()  # Get screen height.
                        # Calculate x and y coordinates for centering.
                        x = (screen_width - window_width) // 2
                        y = (screen_height - window_height) // 2
                        open_dialog.geometry(f"{window_width}x{window_height}+{x}+{y}") # Set dialog position.
                        
                        # Wait for the confirmation dialog to be closed before continuing.
                        self.root.wait_window(open_dialog)
                        
                    else: # If the user cancelled the "Save As" dialog for the Excel file.
                        self.update_status("已取消保存Excel文件") # Update status.
                    
                    # Delete the temporary Excel file.
                    os.remove(temp_file)
            else: # If Ollama API call failed (status code not 200).
                self.update_status(f"API调用失败: {response.status_code} {response.text}") # Update status.
                self.result_text.delete('1.0', tk.END)
                self.result_text.insert('1.0', f"Ollama API Error: {response.status_code}\n{response.text}")


        except requests.exceptions.RequestException as e: # Catch network errors
            self.update_status(f"Ollama API连接错误: {str(e)}")
            self.result_text.delete('1.0', tk.END)
            self.result_text.insert('1.0', f"Could not connect to Ollama API at http://localhost:11434.\nPlease ensure Ollama is running.\nError: {str(e)}")
        except Exception as e: # Catch any other unhandled exceptions during processing.
            self.update_status(f"处理过程中发生未知错误: {str(e)}") # Display error message in status.
            self.result_text.delete('1.0', tk.END)
            self.result_text.insert('1.0', f"An unexpected error occurred: {str(e)}")

        # Re-enable the "Run Processing" button after processing is complete or an error occurs.
        self.run_button.config(state=tk.NORMAL)

    # Starts the Tkinter main event loop for the primary application window.
    def run(self):
        self.root.mainloop() # This call blocks until the main window is closed.

# This block executes if the script is run directly (not imported as a module).
if __name__ == "__main__":
    # Create an instance of the InvoiceProcessor class.
    # This initializes the loading window and starts its event loop.
    processor = InvoiceProcessor()
    # The processor.run() method is called implicitly if the main window
    # is set up correctly after model loading.
    # If start_main_window is reached, it calls self.root.mainloop().
    # If the loading window is closed before model loading, then processor.run() might be needed,
    # but current logic has start_main_window handle the main root's mainloop.
    # The run() method defined in the class is effectively the main UI loop starter.
    # The call processor.run() here is technically not needed if the main window's
    # mainloop is started from within start_main_window, as it is.
    # However, it doesn't hurt as self.root.mainloop() would have already been called.
    # For clarity, it could be removed if start_main_window always leads to root.mainloop().
    # processor.run() # This line is redundant due to self.root.mainloop() in start_main_window
