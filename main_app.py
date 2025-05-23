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
from ollama_client import OllamaClient # Import the new OllamaClient
from excel_writer import ExcelWriter, KEYS_IN_ORDER # Import ExcelWriter and KEYS_IN_ORDER


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
        
        # Center the window
        self.load_window.eval('tk::PlaceWindow . center')
        
        # Label for the loading window
        label = tk.Label(self.load_window, text="请先加载模型", font=('Arial', 12))
        label.pack(pady=20)
        
        # Instantiate OllamaClient
        self.ollama_client = OllamaClient()
        # Instantiate ExcelWriter
        self.excel_writer = ExcelWriter()

        # Load button in the loading window
        self.load_button = ttk.Button(
            self.load_window, 
            text="加载 Ollama qwen2.5", 
            command=self.handle_load_model_button # Command now calls the handler
        )
        self.load_button.pack(pady=10)
        
        # Status label for loading progress
        self.load_status = tk.Label(self.load_window, text="")
        self.load_status.pack(pady=10)
        
        # Start the Tkinter event loop for the loading window.
        self.load_window.mainloop()

    def handle_load_model_button(self):
        # This method is called when the "Load Ollama qwen2.5" button is clicked.
        self.load_button.config(state=tk.DISABLED) # Disable button during loading
        self.load_status.config(text="正在加载模型...") # Update status label
        
        # Use OllamaClient to check connection and model
        if self.ollama_client.check_connection_and_model():
            self.load_status.config(text="模型加载成功！")
            # Call start_main_window after 1 second (1000 ms)
            self.load_window.after(1000, self.start_main_window)
        else:
            # Error messages are handled by check_connection_and_model in OllamaClient
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

    # setup_ui has been moved to AppUI class.

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

        self.update_status("正在调用Ollama API...") # Update status message.

        try:
            # Call OllamaClient to extract data
            extracted_data = self.ollama_client.extract_invoice_data_from_text(text_content)

            if extracted_data is not None:
                # Update UI display
                self.result_text.delete('1.0', tk.END)
                self.result_text.insert('1.0', "提取结果 (JSON):\n" + "="*50 + "\n")
                self.result_text.insert(tk.END, json.dumps(extracted_data, indent=2, ensure_ascii=False)) # Pretty print JSON
                self.result_text.insert(tk.END, "\n" + "="*50)
                self.result_text.see("1.0")

                # Ask the user to select an Excel template file.
                template_path = filedialog.askopenfilename(
                    filetypes=[("Excel files", "*.xlsx")], # Filter for .xlsx files.
                    title="选择Excel模板文件" # Dialog title.
                )
                
                if template_path: # If a template file was selected.
                    # Ask the user where to save the newly populated Excel file.
                    save_path = filedialog.asksaveasfilename(
                        defaultextension=".xlsx", # Default file extension.
                        filetypes=[("Excel files", "*.xlsx")], # Filter for .xlsx files.
                        title="选择保存Excel文件的位置" # Dialog title.
                    )
                    
                    if save_path: # If a save path was provided.
                        # Use ExcelWriter to create the final Excel file
                        success = self.excel_writer.create_invoice_excel(save_path, template_path, extracted_data)
                        if success:
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
                        else:
                            self.update_status("Excel文件保存失败。")
                            messagebox.showerror("保存错误", "无法保存Excel文件。请检查文件路径或权限。")
                    else: # If the user cancelled the "Save As" dialog for the Excel file.
                        self.update_status("已取消保存Excel文件") # Update status.
                else: # If the user cancelled the template selection dialog.
                    self.update_status("未选择Excel模板，操作取消。")
            else:
                # Error messages are handled by extract_invoice_data_from_text or displayed generally below
                self.update_status("Ollama API调用失败或数据提取失败.")
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
            else:
                # Error messages are handled by extract_invoice_data_from_text or displayed generally below
                self.update_status("Ollama API调用失败或数据提取失败.")
                self.result_text.delete('1.0', tk.END)
                self.result_text.insert('1.0', "Ollama API call failed or data extraction returned None.")
                self.run_button.config(state=tk.NORMAL)
                return

        except Exception as e: # Catch any other unhandled exceptions during processing.
            self.update_status(f"处理过程中发生未知错误: {str(e)}") # Display error message in status.
            self.result_text.delete('1.0', tk.END)
            self.result_text.insert('1.0', f"An unexpected error occurred: {str(e)}")
            # Ensure run button is re-enabled in case of unexpected error
            self.run_button.config(state=tk.NORMAL)
            return # Stop processing

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
