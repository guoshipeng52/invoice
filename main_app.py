import tkinter as tk
import os
import json # For json.dumps in process_pdf_action

# Custom module imports
from app_ui import AppUI
from ollama_client import OllamaClient
from excel_writer import ExcelWriter # Comment cleaned, KEYS_IN_ORDER is internal to ExcelWriter
from hardware_checker import HardwareChecker
from pdf_handler import PDFHandler
# config_manager will be used later if needed
# Unused imports (PaddleOCR, fitz, PIL, numpy, GPUtil, subprocess, re, requests) 
# are confirmed to be removed from this file as their functionality is encapsulated
# in their respective handler/client classes.

class InvoiceProcessor: # Acts as the Controller
    def __init__(self):
        # Initialize UI first to display any critical init errors
        self.root = tk.Tk()
        self.root.withdraw() # Hide main window initially
        self.app_ui = AppUI(self.root, self) # Pass root and self (as controller)

        # Initialize backend components
        self.hardware_checker = HardwareChecker(app_ui=self.app_ui) # Pass AppUI instance
        try:
            # Pass a method from app_ui to PDFHandler for status updates during its init
            self.pdf_handler = PDFHandler(update_status_callback=self.app_ui.update_loading_status) 
        except RuntimeError as e:
            self.app_ui.show_error_message("严重初始化错误", f"无法初始化PDF处理组件 (OCR引擎): {e}\n应用程序将无法正常工作。请检查PaddleOCR安装和环境。")
            self.root.destroy() # Close the hidden root window if init fails
            raise # Re-raise to stop execution
            
        self.ollama_client = OllamaClient(app_ui=self.app_ui) # Pass AppUI instance
        self.excel_writer = ExcelWriter()
        self.pdf_path = None # Stores the path to the currently selected PDF
        
        self.app_ui.show_loading_window(self.handle_load_model_button_action)
        self.root.mainloop() # Start the main Tkinter event loop

    def handle_load_model_button_action(self):
        self.app_ui.enable_load_button(False)
        self.app_ui.update_loading_status("正在加载模型...")
        
        if self.ollama_client.check_connection_and_model(): 
            self.app_ui.update_loading_status("模型加载成功！")
            self.root.after(1000, self.start_main_application_flow) 
        else:
            # OllamaClient's check_connection_and_model now uses its AppUI reference for errors
            self.app_ui.update_loading_status("加载失败，请重试")
            self.app_ui.enable_load_button(True)

    def start_main_application_flow(self):
        self.app_ui.close_loading_window()
        # OCR is initialized in PDFHandler.__init__
        # HardwareChecker's check_gpu_vram uses its AppUI reference
        self.hardware_checker.check_gpu_vram() 
        self.app_ui.setup_main_ui() 
        self.app_ui.update_status("等待选择文件...")

    def select_pdf_action(self):
        filepath = self.app_ui.ask_open_pdf_file()
        if filepath:
            self.pdf_path = filepath
            self.app_ui.update_file_label(self.pdf_path)
            self.app_ui.enable_run_button(True)
            self.app_ui.update_status("文件已选择，请点击运行处理")
        else:
            self.pdf_path = None 
            self.app_ui.update_file_label("未选择文件")
            self.app_ui.enable_run_button(False)
            self.app_ui.update_status("等待选择文件...")

    def process_pdf_action(self):
        if not self.pdf_path:
            self.app_ui.show_error_message("错误", "请先选择PDF文件")
            return

        self.app_ui.update_status("正在处理PDF...")
        self.app_ui.enable_run_button(False)
        
        work_dir = "working" 
        if not os.path.exists(work_dir):
            os.makedirs(work_dir)

        text_content_lines = [] # Initialize

        # This callback allows PDFHandler to update the UI's status label during its processing
        def status_update_callback(message):
            self.app_ui.update_status(message)
        
        try:
            # PDFHandler now encapsulates all PDF text extraction logic
            text_content_lines = self.pdf_handler.extract_text_from_pdf(self.pdf_path, status_update_callback)
        except Exception as e:
            self.app_ui.update_status(f"处理PDF时出错: {str(e)}")
            self.app_ui.show_error_in_results(f"Error during PDF processing: {str(e)}") # Show error in results area
            self.app_ui.enable_run_button(True) # Re-enable run button
            return

        if not text_content_lines: # If extraction failed or PDF was empty
            # PDFHandler's extract_text_from_pdf should have called update_status to inform the user.
            self.app_ui.enable_run_button(True) # Re-enable run button
            return

        # Save the extracted text to a .txt file (kept for now as per instruction)
        txt_path = os.path.join(work_dir, "ocr_result.txt")
        try:
            with open(txt_path, "w", encoding="utf-8") as f:
                f.write("\n".join(text_content_lines))
        except Exception as e:
            self.app_ui.update_status(f"保存文本文件失败: {e}")
            # Optionally show this minor error to user, or just log it
            # For now, continue processing even if this save fails.

        self.app_ui.update_status("正在调用Ollama API...")
        # The rest of the method (Ollama call, Excel processing) continues here,
        # using text_content_lines for the Ollama prompt.
        # Note: The prompt construction is inside ollama_client.extract_invoice_data_from_text
        try:
            extracted_data = self.ollama_client.extract_invoice_data_from_text(text_content_lines) 

            if extracted_data:
                self.app_ui.show_results(json.dumps(extracted_data, indent=2, ensure_ascii=False))
                
                template_path = self.app_ui.ask_excel_template_file()
                if template_path:
                    save_path = self.app_ui.ask_save_excel_file()
                    if save_path:
                        success = self.excel_writer.create_invoice_excel(save_path, template_path, extracted_data)
                        if success:
                            self.app_ui.update_status(f"处理完成！结果已保存到: {save_path}")
                            self.app_ui.prompt_and_open_file(save_path)
                        else:
                            self.app_ui.update_status("Excel文件保存失败。")
                            self.app_ui.show_error_message("保存错误", "无法保存Excel文件。请检查文件路径或权限。")
                    else:
                        self.app_ui.update_status("已取消保存Excel文件")
                else:
                    self.app_ui.update_status("未选择Excel模板，操作取消。")
            else:
                # OllamaClient's methods now directly call AppUI's show_error_message
                self.app_ui.update_status("Ollama API调用失败或数据提取失败.")
        
        except RuntimeError as e: 
            self.app_ui.show_error_message("严重错误", str(e))
            self.app_ui.update_status(f"关键组件初始化失败: {e}")
        except Exception as e: 
            self.app_ui.show_error_message("未知错误", f"处理过程中发生未知错误: {str(e)}")
            self.app_ui.update_status(f"处理过程中发生未知错误: {str(e)}")
        
        finally: 
            self.app_ui.enable_run_button(True)

# This block executes if the script is run directly
if __name__ == "__main__":
    try:
        app_controller = InvoiceProcessor()
    except RuntimeError as e_init:
        print(f"Application failed to initialize: {e_init}")
    except tk.TclError as e_tk: # Catch Tkinter specific errors if display is not available
        print(f"Tkinter error: {e_tk}. Ensure a display environment is available for GUI applications.")
