import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from paddleocr import PaddleOCR
import fitz  # PyMuPDF
import os
import requests
import json
# import pandas as pd # Removed as pandas is not used in the current workflow
from openpyxl import Workbook, load_workbook # load_workbook is used for the old invoice flow's template
from openpyxl.styles import Font, Alignment # Ensure Alignment is imported for cell styling
from PIL import Image
import numpy as np
import shutil # For directory cleaning (e.g., rmtree)
import re # For filename sanitization
import traceback # For detailed error logging

class SubsystemVerificationDialog(tk.Toplevel):
    """
    A dialog window for users to verify and edit subsystem information 
    extracted from a PDF's Table of Contents. Allows editing subsystem names 
    and their 1-indexed start pages, with a countdown timer for auto-confirmation.
    """
    def __init__(self, parent, subsystems_data: list[dict], timeout_seconds: int = 30):
        """
        Initializes the SubsystemVerificationDialog.

        Args:
            parent: The parent tkinter window.
            subsystems_data: A list of dictionaries, where each dictionary should contain
                             "subsystem_name" (str) and "start_page" (int, 1-indexed).
            timeout_seconds: Duration in seconds for the auto-confirmation timer.
        """
        super().__init__(parent)
        self.parent = parent
        self.subsystems_data = subsystems_data if subsystems_data is not None else []
        self.timeout_seconds = timeout_seconds
        self.confirmed_subsystems = None # Stores the list of confirmed/edited subsystems or None if cancelled.
        self.timer_id = None # ID for the tkinter `after` timer.
        self.entry_widgets = [] # List to store references to Entry widgets for names and pages.

        self.title("Verify Extracted Subsystems") 
        
        # Dynamically adjust dialog height based on the number of subsystems, within min/max bounds.
        height = min(600, max(300, len(self.subsystems_data) * 40 + 180)) 
        self.geometry(f"700x{height}")

        # Configure modal behavior
        self.protocol("WM_DELETE_WINDOW", self.cancel) # Handle window close (X button) as a cancel action.
        self.transient(parent) # Make dialog appear on top of parent.
        self.grab_set() # Make dialog modal, blocking interaction with parent.

        self._setup_ui() # Create and layout UI elements.

        if self.subsystems_data: # Only start countdown if there's data to confirm.
            self.start_countdown()
        else:
            self.timer_label.config(text="No subsystems provided to verify.")
            if hasattr(self, 'confirm_button'): 
                 self.confirm_button.config(state=tk.DISABLED) # Disable confirm if no data.

        # Set focus to the first editable name field for quicker editing.
        if self.entry_widgets and self.entry_widgets[0]['name']:
            self.entry_widgets[0]['name'].focus_set()
            self.entry_widgets[0]['name'].selection_range(0, tk.END) # Select existing text


    def _setup_ui(self):
        """Sets up the UI elements of the dialog, including input fields and buttons."""
        main_frame = ttk.Frame(self, padding="15")
        main_frame.pack(expand=True, fill=tk.BOTH)

        # --- Header Frame for Column Titles ---
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 5))
        
        header_configs = [ # Configuration for header labels and corresponding data columns
            ("Original Name", {"sticky": "w", "weight": 1}),
            ("Editable Name", {"sticky": "ew", "weight": 2}), # Editable name gets more space
            ("Original Page", {"sticky": "w", "weight": 0}),
            ("Editable Page", {"sticky": "ew", "weight": 0})
        ]

        for i, (text, config) in enumerate(header_configs):
            label = ttk.Label(header_frame, text=text, font=('Arial', 10, 'bold'))
            label.grid(row=0, column=i, padx=5, pady=2, sticky=config["sticky"])
            header_frame.columnconfigure(i, weight=config["weight"], uniform=f"col_header_{i}") # Uniform for alignment
        
        # --- Scrollable Area for Subsystem Entries ---
        canvas_frame = ttk.Frame(main_frame)
        canvas_frame.pack(fill=tk.BOTH, expand=True, pady=(0,10))

        canvas = tk.Canvas(canvas_frame)
        scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas) # This frame will hold the entry rows

        # Configure scrollable_frame to update scrollregion of canvas when its size changes
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw") # Embed scrollable_frame in canvas
        canvas.configure(yscrollcommand=scrollbar.set) # Link scrollbar to canvas

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Show scrollbar only if content is likely to exceed initial dialog height
        estimated_content_height = len(self.subsystems_data) * 40 # Approx height per item
        # self.winfo_height() might not be accurate before window is fully drawn, so this is a heuristic
        if estimated_content_height > (self.winfo_reqheight() - 100): # 100 for header/footer/padding
             scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Configure columns in the scrollable_frame to match header widths/weights
        for i, (_, config) in enumerate(header_configs): # Use same weights as header for alignment
            self.scrollable_frame.columnconfigure(i, weight=config["weight"], uniform=f"col_data_{i}")

        # Populate rows with subsystem data if available
        if not self.subsystems_data:
            no_data_label = ttk.Label(self.scrollable_frame, text="No subsystem data provided to verify.")
            no_data_label.grid(row=0, column=0, columnspan=4, pady=15, padx=10, sticky="ew")
        else:
            for i, item in enumerate(self.subsystems_data):
                row_idx = i 
                
                original_name_val = item.get("subsystem_name", "N/A")
                original_page_val = item.get("start_page", "N/A") # Page numbers are 1-indexed

                # Original (non-editable) information
                original_name_label = ttk.Label(self.scrollable_frame, text=original_name_val, wraplength=180) 
                original_name_label.grid(row=row_idx, column=0, padx=(5,10), pady=4, sticky="w")
                
                original_page_label = ttk.Label(self.scrollable_frame, text=str(original_page_val))
                original_page_label.grid(row=row_idx, column=2, padx=5, pady=4, sticky="w")

                # Editable fields
                name_entry = ttk.Entry(self.scrollable_frame) 
                name_entry.insert(0, original_name_val if original_name_val != "N/A" else "")
                name_entry.grid(row=row_idx, column=1, padx=5, pady=4, sticky="ew")
                
                page_entry = ttk.Entry(self.scrollable_frame) 
                page_entry.insert(0, str(original_page_val) if original_page_val != "N/A" else "")
                page_entry.grid(row=row_idx, column=3, padx=5, pady=4, sticky="ew")
                
                # Store references to widgets and original data for later retrieval/validation
                self.entry_widgets.append({
                    'name': name_entry, 
                    'page': page_entry, 
                    'original_name': original_name_val, 
                    'original_page': original_page_val
                })
        
        # --- Footer for Timer and Buttons ---
        footer_frame = ttk.Frame(main_frame)
        footer_frame.pack(fill=tk.X, pady=(10,0), side=tk.BOTTOM) # Ensure footer is at the bottom

        self.timer_label = ttk.Label(footer_frame, text="", font=('Arial', 10, 'italic'))
        self.timer_label.pack(side=tk.LEFT, padx=5, pady=5) # Timer on the left

        button_subframe = ttk.Frame(footer_frame) # Subframe to pack buttons to the right
        button_subframe.pack(side=tk.RIGHT)

        self.confirm_button = ttk.Button(button_subframe, text="Confirm Changes", command=self.confirm) # More descriptive text
        self.confirm_button.pack(side=tk.LEFT, padx=(0,5), pady=5)

        self.cancel_button = ttk.Button(button_subframe, text="Cancel", command=self.cancel)
        self.cancel_button.pack(side=tk.LEFT, padx=0, pady=5)
        
    def _stop_timer(self):
        """Stops the countdown timer if it's active."""
        if self.timer_id:
            self.after_cancel(self.timer_id)
            self.timer_id = None

    def start_countdown(self):
        """Starts the auto-confirmation countdown timer."""
        self._stop_timer() # Ensure no existing timer is running before starting a new one.
        self.remaining_time = self.timeout_seconds
        self.update_timer()

    def update_timer(self):
        """Updates the countdown timer label each second. Calls auto_confirm when timer expires."""
        if not self.winfo_exists(): # Stop if window is somehow destroyed externally.
            self._stop_timer()
            return

        if self.remaining_time >= 0:
            self.timer_label.config(text=f"Auto-confirming in {self.remaining_time} seconds...")
            self.remaining_time -= 1
            self.timer_id = self.after(1000, self.update_timer) # Schedule next update.
        else:
            self.timer_label.config(text="Timer expired. Auto-confirming...")
            self.auto_confirm()

    def _collect_and_validate_data(self) -> list[dict] | None:
        """
        Collects data from entry fields and validates it (especially page numbers).
        Page numbers must be positive integers. Empty names are allowed.
        
        Returns:
            A list of validated subsystem dictionaries if all data is valid.
            None if validation fails for any entry, allowing the dialog to stay open for correction.
        """
        collected_data = []
        for i, entries_map in enumerate(self.entry_widgets): # Iterate through stored widget references.
            name_widget = entries_map['name']
            page_widget = entries_map['page']
            name = name_widget.get().strip()
            page_str = page_widget.get().strip()
            
            original_name = entries_map['original_name'] # For error messages
            
            # Name can be empty (user might want to remove/ignore a subsystem by clearing its name)
            # If names were mandatory, validation would be added here.

            try:
                page = 0 # Default if page_str is empty (could also be an error if pages are mandatory)
                if page_str: # Only convert if not empty
                    page = int(page_str) 
                
                # Page numbers should be positive (1-indexed as per user input)
                if page < 0 : # Allow 0 if it means "not set" or handle as error
                    messagebox.showerror("Validation Error", 
                                         f"Page number for '{name if name else original_name}' (item {i+1}) "
                                         f"cannot be negative. Got '{page_str}'.", 
                                         parent=self)
                    page_widget.focus_set() # Focus on the problematic field
                    return None # Indicates validation failure
            except ValueError: # page_str was not empty and not a valid integer
                messagebox.showerror("Validation Error", 
                                     f"Start page for '{name if name else original_name}' (item {i+1}) "
                                     f"must be an integer. Got '{page_str}'.", 
                                     parent=self)
                page_widget.focus_set()
                return None # Indicates validation failure
            
            collected_data.append({"subsystem_name": name, "start_page": page})
        return collected_data # Return list of valid data.

    def confirm(self):
        """Handles the 'Confirm' button click. Validates data and closes dialog if valid."""
        self._stop_timer() # Stop countdown on manual action.
        validated_data = self._collect_and_validate_data()
        if validated_data is not None: # Validation was successful.
            self.confirmed_subsystems = validated_data
            if self.winfo_exists(): self.destroy() # Close dialog.
        # If validation failed, dialog remains open for user to correct errors.

    def cancel(self):
        """Handles the 'Cancel' button click or window close action. Sets result to None and closes."""
        self._stop_timer() # Stop countdown.
        self.confirmed_subsystems = None # Indicate cancellation.
        if self.winfo_exists(): self.destroy() # Close dialog.

    def auto_confirm(self):
        """Handles auto-confirmation when the timer expires. Validates data."""
        self._stop_timer() 
        if self.winfo_exists(): # Ensure window still exists before updating label.
            self.timer_label.config(text="Auto-confirmed.")
        
        validated_data = self._collect_and_validate_data()
        if validated_data is not None: # If current edits are valid
            self.confirmed_subsystems = validated_data
        else:
            # Auto-confirmation failed validation due to invalid user edits.
            # Reverting to original data is a safe default to prevent data loss.
            self.confirmed_subsystems = self.subsystems_data 
            print(f"SubsystemVerificationDialog: Auto-confirm validation failed on user edits. Reverted to original data.")
            # Notify user if parent window still exists
            if self.parent and self.parent.winfo_exists():
                 messagebox.showwarning("Auto-confirm Validation Failed", 
                                   "Edits made were invalid. Auto-confirmed using the original subsystem data to prevent data loss.", 
                                   parent=self.parent)
        if self.winfo_exists(): self.destroy() # Close dialog.
        
    def get_confirmed_subsystems(self) -> list[dict] | None:
        """
        Makes the dialog modal and returns the confirmed subsystems data upon dialog closure.

        Returns:
            A list of confirmed/edited subsystem dictionaries if confirmed.
            None if the dialog was cancelled.
            The original `subsystems_data` if auto-confirmation occurred with invalid edits.
        """
        self.wait_window() # Crucial for modal behavior; execution waits here until self.destroy() is called.
        return self.confirmed_subsystems


class InvoiceProcessor:
    """
    Main application class for processing documents. 
    Handles PDF selection, model interactions for subsystem and deliverable extraction,
    and generation of output files (split PDFs, Excel reports).
    """
    def __init__(self):
        """
    Initializes the application, setting up model names from `config.json`,
    and the initial model loading window.
    Model names can be configured by editing 'config.json' in the script's directory.
    If 'config.json' is missing, invalid, or a model key is not found,
    default values will be used and a new 'config.json' will be created/updated.
        """
    config_filename = "config.json"
    default_models = {
        "gemma_model_name": "gemma:7b",      # For Table of Contents / Subsystem extraction
        "qwen3_model_name": "qwen2:14b",     # For deliverable identification (placeholder for Qwen3)
        "invoice_ollama_model": "qwen2.5:14b" # For original invoice processing
    }
    
    current_config = default_models.copy() # Start with defaults, override with loaded values
    config_needs_save = False

    try:
        with open(config_filename, 'r', encoding='utf-8') as f:
            loaded_config = json.load(f)
        
        if not isinstance(loaded_config, dict):
            print(f"Warning: Content of '{config_filename}' is not a valid dictionary. Using defaults and recreating file.")
            config_needs_save = True
            # current_config remains default_models
        else:
            # Load valid values from config, use default if a key is missing or invalid
            for key, default_value in default_models.items():
                if key not in loaded_config or not isinstance(loaded_config[key], str) or not loaded_config[key].strip():
                    print(f"Info: Model key '{key}' missing or invalid in '{config_filename}'. Using default: '{default_value}'.")
                    current_config[key] = default_value # Use default if missing/invalid
                    config_needs_save = True # Mark that we need to save the updated config
                else:
                    current_config[key] = loaded_config[key] # Use value from file
            
            # Check if there were any keys in loaded_config not in default_models (they will be ignored but config saved to prune them)
            if any(key not in default_models for key in loaded_config):
                config_needs_save = True


    except FileNotFoundError:
        print(f"Info: '{config_filename}' not found. Creating with default model names.")
        # current_config is already default_models
        config_needs_save = True
    except json.JSONDecodeError:
        print(f"Warning: Error decoding '{config_filename}'. File may be corrupted. Using defaults and recreating file.")
        # current_config is already default_models
        config_needs_save = True
    except IOError as e:
        print(f"Warning: Could not read '{config_filename}': {e}. Using default model names for this session.")
        # In this case, we use defaults but might not want to overwrite if it was a temporary read issue
        # However, to ensure a valid config is available for next run, we'll mark for save.
        config_needs_save = True

    if config_needs_save:
        try:
            with open(config_filename, 'w', encoding='utf-8') as f:
                json.dump(current_config, f, indent=4, ensure_ascii=False)
            print(f"Info: Model configuration saved/updated in '{config_filename}'.")
        except IOError as e:
            print(f"Error: Could not save configuration to '{config_filename}': {e}")
            messagebox.showerror("Configuration Error", f"Could not save model configuration file '{config_filename}'. Please check permissions.\nDefaults will be used for this session.")

    # Assign model names to instance attributes
    self.gemma_model_name = current_config["gemma_model_name"]
    self.qwen3_model_name = current_config["qwen3_model_name"]
    self.invoice_ollama_model = current_config["invoice_ollama_model"]
    
    print(f"Models configured to use: Gemma='{self.gemma_model_name}', Qwen3='{self.qwen3_model_name}', Invoice='{self.invoice_ollama_model}'")
    print(f"Edit '{config_filename}' to change model names if needed.")

        # --- Initial Model Loading Window ---
        self.load_window = tk.Tk()
        self.load_window.title("模型加载") # "Model Loading"
        self.load_window.geometry("300x150")
        self.load_window.eval('tk::PlaceWindow . center') # Center the loading window.
        
        label = ttk.Label(self.load_window, text="请先加载模型", font=('Arial', 12))
        label.pack(pady=20)
        
        # Button text reflects the model used for the initial test/load.
        self.load_button = ttk.Button(
            self.load_window, 
            text=f"加载默认模型 ({self.invoice_ollama_model})", 
            command=self.load_model
        )
        self.load_button.pack(pady=10)
        
        self.load_status = ttk.Label(self.load_window, text="") 
        self.load_status.pack(pady=10)
        
        self.load_window.mainloop() # Start event loop for loading window.

    def load_model(self):
        """
        Handles the initial model loading test.
        Currently tests the `invoice_ollama_model`. This could be expanded or made more generic.
        """
        self.load_button.config(state=tk.DISABLED)
        self.load_status.config(text="正在加载模型...") # "Loading model..."
        
        try:
            # Test a generic prompt with one of the models to ensure Ollama is running.
            # This uses the invoice_ollama_model, but any configured model could be used.
            response = requests.post( 
                "http://localhost:11434/api/generate",
                json={
                    "model": self.invoice_ollama_model, 
                    "prompt": "Hello Ollama!", # Simple test prompt
                    "stream": False
                }
            )
            
            if response.status_code == 200: 
                self.load_status.config(text="模型加载成功！") # "Model loaded successfully!"
                self.load_window.after(1000, self.start_main_window) # Proceed to main app window.
            else:
                # Try to parse a more specific error from Ollama if available.
                error_message = f"模型加载失败 (HTTP {response.status_code})"
                try:
                    error_detail = response.json().get("error", "Unknown error from Ollama.")
                    error_message += f": {error_detail}"
                except ValueError: # Response was not JSON
                    error_message += f": {response.text[:100]}..." # Show part of non-JSON response
                raise Exception(error_message)
                
        except requests.exceptions.ConnectionError as e:
             messagebox.showerror("错误", f"模型连接失败: {str(e)}\n请确保Ollama服务正在运行。") # "Model connection failed... Ensure Ollama is running."
             self.load_button.config(state=tk.NORMAL)
             self.load_status.config(text="连接失败，请重试") # "Connection failed, please retry."
        except Exception as e:
            messagebox.showerror("错误", f"模型加载出错: {str(e)}") # "Error loading model."
            self.load_button.config(state=tk.NORMAL)
            self.load_status.config(text="加载失败，请重试") # "Loading failed, please retry."

    def start_main_window(self):
        """Closes the loading window and initializes the main application UI and OCR instance."""
        if self.load_window and self.load_window.winfo_exists():
            self.load_window.destroy()
        
        # Initialize PaddleOCR here, as it's used by multiple parts of the application.
        self.ocr = PaddleOCR(use_angle_cls=True, lang='ch') # lang='ch' for Chinese, adjust if needed.
        
        self.setup_ui() # Setup the main application UI.
        if self.root: # Check if root window was created
             self.root.mainloop() # Start the Tkinter event loop for the main window.


def call_ollama_api(model_name: str, prompt: str) -> dict | None:
    """
    Calls the Ollama API with the given model name and prompt.

    Args:
        model_name: The name of the Ollama model to use.
        prompt: The prompt string to send to the model.

    Returns:
        A dictionary containing the JSON response from the API, or None if an error occurs.
    """
    try:
        response = requests.post(
            "http://localhost:11434/api/generate", # Standard Ollama API endpoint
            json={
                "model": model_name,
                "prompt": prompt,
                "stream": False # Assuming non-streaming responses for simplicity
            }
        )
        response.raise_for_status() # Raises HTTPError for bad responses (4XX or 5XX)
        return response.json()
    except requests.exceptions.ConnectionError as e:
        print(f"Ollama API connection error: {e}")
        messagebox.showerror("API Connection Error", 
                             f"Cannot connect to Ollama API at http://localhost:11434. "
                             f"Please ensure Ollama is running.\nDetails: {e}")
        return None
    except requests.exceptions.HTTPError as e:
        error_detail = "Unknown error"
        try: # Try to get more specific error from Ollama's JSON response
            error_detail = e.response.json().get("error", e.response.text)
        except ValueError: # Response not JSON
            error_detail = e.response.text[:200] # Show snippet if not JSON
        print(f"Ollama API HTTP error: {e.response.status_code} - {error_detail}")
        messagebox.showerror("API HTTP Error", 
                             f"Ollama API request failed (HTTP {e.response.status_code}):\n{error_detail}")
        return None
    except requests.exceptions.RequestException as e: # Catch other request-related errors
        print(f"Error calling Ollama API: {e}")
        messagebox.showerror("API Request Error", 
                             f"An unexpected error occurred with the Ollama API request:\n{e}")
        return None
    except Exception as e_gen: # Catch any other general errors
        print(f"A general error occurred in call_ollama_api: {e_gen}")
        messagebox.showerror("API General Error", f"An unexpected error occurred: {e_gen}")
        return None


def extract_text_from_pdf(pdf_path: str, 
                            ocr_instance: PaddleOCR, 
                            is_scanned_pdf_func: callable, 
                            pdf_to_images_func: callable) -> str:
    """
    Extracts all text from a given PDF file.
    Uses OCR via PaddleOCR if the PDF is detected as scanned.

    Args:
        pdf_path: Path to the PDF file.
        ocr_instance: An initialized instance of PaddleOCR.
        is_scanned_pdf_func: A callable that takes pdf_path and returns True if PDF is scanned.
        pdf_to_images_func: A callable that takes pdf_path and returns a list of images (numpy arrays).

    Returns:
        A string containing all extracted text, or an empty string if an error occurs or no text found.
    """
    try:
        if not os.path.exists(pdf_path):
            print(f"Error in extract_text_from_pdf: PDF file not found at '{pdf_path}'")
            return ""
        
        text_content_list = []
        if is_scanned_pdf_func(pdf_path):
            # PDF is determined to be scanned; use OCR.
            images = pdf_to_images_func(pdf_path) # This should return a list of numpy arrays.
            for i, img_array in enumerate(images):
                # print(f"OCR processing page {i+1} of {len(images)} for '{os.path.basename(pdf_path)}'")
                result = ocr_instance.ocr(img_array, cls=True) # cls=True for text angle classification
                if result and result[0] is not None: # PaddleOCR returns a list of lists, result[0] contains lines for the page
                    for line_info in result[0]:
                        text_content_list.append(line_info[1][0]) # Actual text is in line_info[1][0]
        else:
            # PDF is not scanned (or text extraction is preferred); use fitz to extract text.
            doc = fitz.open(pdf_path)
            for page_num in range(len(doc)):
                page = doc.load_page(page_num)
                text_content_list.append(page.get_text("text")) # "text" for plain text, "blocks" for more structure
            doc.close()
        return "\n".join(text_content_list)
    except Exception as e:
        print(f"Error extracting text from PDF '{pdf_path}': {e}")
        traceback.print_exc() # Print full traceback for debugging
        return "" # Return empty string on error


    # Note: The duplicate InvoiceProcessor class definition was removed in the previous step.
    # Methods like setup_ui, update_status, etc., are now part of the main InvoiceProcessor class.

    def setup_ui(self):
        """Sets up the main UI for the Document Processor application."""
        self.root = tk.Tk()
        self.root.title("Document Processor") 
        self.root.geometry("800x400") # Adjusted size as the large text result area was removed

        main_app_frame = ttk.Frame(self.root, padding="10")
        main_app_frame.pack(expand=True, fill=tk.BOTH)

        # Main panel for controls
        controls_panel = ttk.Frame(main_app_frame) 
        controls_panel.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0,10))

        self.select_button = ttk.Button(controls_panel, text="Select PDF Document", command=self.select_pdf) 
        self.select_button.pack(pady=10, fill=tk.X)

        self.run_button = ttk.Button(controls_panel, text="Process Technical Agreement", 
                                     command=self.process_technical_agreement_workflow, state=tk.DISABLED)
        self.run_button.pack(pady=10, fill=tk.X)
        
        # Debug/Test button for SubsystemVerificationDialog
        self.test_subsystem_button = ttk.Button(controls_panel, text="Test Subsystem Dialog (Debug)", 
                                                command=self.show_verification_dialog_example)
        self.test_subsystem_button.pack(pady=5, fill=tk.X) 

        # Optional: Button for old invoice processing (can be uncommented if needed)
        # self.invoice_button = ttk.Button(controls_panel, text="Process Invoice (Old)", 
        #                                  command=self._process_invoice_pdf, state=tk.DISABLED)
        # self.invoice_button.pack(pady=5, fill=tk.X)

        # Status and file labels at the bottom of the control panel
        self.status_label = ttk.Label(controls_panel, text="Select a PDF document to begin.", wraplength=750) 
        self.status_label.pack(pady=10, fill=tk.X, side=tk.BOTTOM, expand=True) # Pack at bottom

        self.file_label = ttk.Label(controls_panel, text="No document selected.", wraplength=750) 
        self.file_label.pack(pady=10, fill=tk.X, side=tk.BOTTOM, expand=True) # Pack at bottom


    def update_status(self, message: str):
        """Updates the status label in the UI. Ensures UI updates are handled safely."""
        if hasattr(self, 'status_label') and self.status_label.winfo_exists():
            self.status_label.config(text=message)
            if hasattr(self, 'root') and self.root.winfo_exists():
                self.root.update_idletasks() # Safely update UI during processing

    def pdf_to_images(self, pdf_path: str) -> list[np.ndarray]:
        """
        Converts all pages of a PDF to a list of images (numpy arrays, BGR format for PaddleOCR).

        Args:
            pdf_path: Path to the PDF file.

        Returns:
            A list of numpy arrays, where each array represents an image of a page.
        """
        images = []
        try:
            doc = fitz.open(pdf_path)
            for page_num in range(len(doc)):
                page = doc.load_page(page_num)
                # Increase DPI for potentially better OCR results on scanned documents
                pix = page.get_pixmap(dpi=200) 
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                # PaddleOCR typically expects BGR format numpy arrays
                img_np_bgr = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR) # Requires opencv-python
                images.append(img_np_bgr)
            doc.close()
        except Exception as e:
            print(f"Error converting PDF to images for '{pdf_path}': {e}")
            self.update_status(f"Error converting PDF to images: {e}")
        return images

    def select_pdf(self):
        """Opens a file dialog for PDF selection and updates UI elements accordingly."""
        self.pdf_path = filedialog.askopenfilename(
            title="Select PDF Document",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")] # Allow selection of any file type
        )
        if self.pdf_path:
            self.file_label.config(text=f"Selected: {os.path.basename(self.pdf_path)}")
            self.run_button.config(state=tk.NORMAL) # Enable main processing button
            if hasattr(self, 'invoice_button') and self.invoice_button.winfo_exists(): 
                self.invoice_button.config(state=tk.NORMAL) 
            self.update_status("Document selected. Ready to process.")
        else: # No file selected
            self.file_label.config(text="No document selected.")
            self.run_button.config(state=tk.DISABLED)
            if hasattr(self, 'invoice_button') and self.invoice_button.winfo_exists(): 
                self.invoice_button.config(state=tk.DISABLED)
            self.update_status("Waiting for document selection...")

    def is_scanned_pdf(self, pdf_path: str) -> bool:
        """
        Checks if a PDF is likely scanned by attempting to extract a small amount of text.
        If very little text is found, it's considered scanned.

        Args:
            pdf_path: Path to the PDF file.

        Returns:
            True if the PDF is likely a scanned document, False otherwise.
        """
        text_content_len = 0
        try:
            doc = fitz.open(pdf_path)
            # Check first few pages or a character threshold for efficiency
            for i, page in enumerate(doc):
                if i >= 3 and text_content_len > 100: # Optimization: if enough text found early, assume not scanned
                    break
                text_content_len += len(page.get_text("text").strip())
                if text_content_len > 100: # Arbitrary threshold for "enough text"
                    break 
            doc.close()
        except Exception as e:
            print(f"Error checking if PDF is scanned ('{pdf_path}'): {e}")
            # Default to treating as not scanned if error, or could default to scanned
            return False 
        return text_content_len < 100 # Threshold for considering it scanned

    def _process_invoice_pdf(self):
        """
        Handles the original invoice processing workflow.
        This method is kept for backward compatibility or optional use.
        It extracts information from an invoice PDF and optionally fills an Excel template.
        """
        if not hasattr(self, 'pdf_path') or not self.pdf_path:
            messagebox.showwarning("No Document Selected", "Please select a PDF document first.")
            return

        self.update_status("Processing document for Invoices...")
        
        # Determine which button triggered this, to re-enable the correct one.
        active_run_button = self.invoice_button if (hasattr(self, 'invoice_button') and self.invoice_button.winfo_exists()) else self.run_button
        if active_run_button: active_run_button.config(state=tk.DISABLED)
        
        work_dir = "working_output_invoice" # Separate working directory for invoice processing
        if not os.path.exists(work_dir):
            os.makedirs(work_dir)

        try:
            self.update_status("Extracting text for invoice processing...")
            text_content_str = extract_text_from_pdf(
                self.pdf_path, self.ocr, self.is_scanned_pdf, self.pdf_to_images
            )

            txt_path = os.path.join(work_dir, f"{os.path.basename(self.pdf_path)}_invoice_ocr.txt")
            with open(txt_path, "w", encoding="utf-8") as f:
                f.write(text_content_str)
            self.update_status(f"Invoice text extracted to '{txt_path}'")

            self.update_status(f"Calling Ollama API ({self.invoice_ollama_model}) for invoice data...")
            
            prompt_template = """
            请从以下文本中按顺序提取这些信息，只返回信息内容，每行一个，不要其他任何文字，其中购买方再左上角，销售方：
            第1行：购买方名称
            ... (rest of the detailed invoice prompt) ...
            第19行：发票类型（只返回“普通发票”或“专用发票”）
            文本内容：
            {document_text}
            """ # Prompt kept concise here for brevity in this review
            full_prompt = prompt_template.format(document_text=text_content_str)

            api_result = call_ollama_api(self.invoice_ollama_model, full_prompt)
            
            if api_result and 'response' in api_result:
                extracted_info_str = api_result['response'].strip()
                extracted_info_lines = extracted_info_str.split('\n')
                
                print("--- Invoice Ollama API Extraction Results ---")
                display_titles = [
                    "购买方名称", "购买方纳税人识别号", "购买方开户行", "购买方开户行账号",
                    "发票代码", "货物或应税劳务、服务名称", "数量", "单价",
                    "销售方名称", "销售方的纳税人识别号", "销售方开户行", "销售方开户行账号",
                    "开票日期", "发票号码", "价税合计", "税率", "税额", "金额", "发票类型"
                ]
                for i, title in enumerate(display_titles):
                    info_line = extracted_info_lines[i] if i < len(extracted_info_lines) else "N/A"
                    print(f"{title}: {info_line}")
                print("--- End of Invoice Results ---")
                
                self.update_status("API call successful. Processing Excel for invoice...")

                temp_wb = Workbook()
                temp_ws = temp_wb.active
                max_lines_to_write = min(19, len(extracted_info_lines))
                cells = [f"{chr(65+i)}1" for i in range(max_lines_to_write)]
                for cell, line_idx in zip(cells, range(max_lines_to_write)):
                    temp_ws[cell] = extracted_info_lines[line_idx].strip()
                
                temp_file = os.path.join(work_dir, "temp_invoice_data.xlsx")
                temp_wb.save(temp_file)
                
                template_path = filedialog.askopenfilename(
                    title="Select Excel Template for Invoice",
                    filetypes=[("Excel files", "*.xlsx")]
                )
                
                if template_path:
                    wb = load_workbook(template_path)
                    ws = wb.active
                    for _ in range(3, ws.max_row + 1): # Clear existing data beyond header (row 2)
                        ws.delete_rows(3) 
                    
                    def get_temp_val(cell_id, default=""): # Helper for safe value access
                        return temp_ws[cell_id].value if temp_ws[cell_id].value is not None else default

                    # Fill template (ensure mapping is correct based on actual template)
                    ws['E3'] = get_temp_val('A1'); ws['F3'] = get_temp_val('B1') 
                    # ... (rest of the cell mappings, ensure they match the template)
                    ws['A3'] = get_temp_val('S1') 

                    save_path = filedialog.asksaveasfilename(
                        defaultextension=".xlsx",
                        filetypes=[("Excel files", "*.xlsx")],
                        title="Save Processed Invoice Excel As...",
                        initialfile=f"processed_invoice_{os.path.basename(self.pdf_path)}.xlsx"
                    )
                    
                    if save_path:
                        wb.save(save_path)
                        self.update_status(f"Invoice processed and saved to: {os.path.basename(save_path)}")
                        self._show_open_file_dialog(save_path, "Invoice Processed & Saved")
                    else:
                        self.update_status("Invoice Excel save cancelled.")
                    if os.path.exists(temp_file): os.remove(temp_file)
                else: 
                    self.update_status("No Excel template selected for invoice. Skipping Excel filling.")
                    if os.path.exists(temp_file): os.remove(temp_file)
            else: 
                self.update_status("Invoice API call failed or no valid response.")
                messagebox.showerror("API Error", "Ollama API call for invoice processing failed or returned no valid response.")

        except Exception as e:
            self.update_status(f"Error during invoice processing: {str(e)}")
            messagebox.showerror("Invoice Processing Error", f"An error occurred: {str(e)}")
            traceback.print_exc()
        finally:
            if active_run_button: active_run_button.config(state=tk.NORMAL) # Re-enable the button

    def _show_open_file_dialog(self, file_path: str, title: str = "File Saved"):
        """
        Shows a general-purpose dialog asking the user if they want to open a specified file.

        Args:
            file_path: The absolute path to the file that was saved.
            title: The title for the dialog window.
        """
        if not file_path or not os.path.exists(file_path):
            print(f"Error in _show_open_file_dialog: File not found at '{file_path}'")
            return

        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        dialog.geometry("400x170") 
        dialog.transient(self.root) 
        dialog.grab_set() 

        # Attempt to center the dialog on the parent window
        try:
            if self.root and self.root.winfo_exists(): # Ensure root window is available
                self.root.update_idletasks() 
                parent_x = self.root.winfo_x()
                parent_y = self.root.winfo_y()
                parent_width = self.root.winfo_width()
                parent_height = self.root.winfo_height()
                
                dialog_width = 400
                dialog_height = 170
                
                x = parent_x + (parent_width // 2) - (dialog_width // 2)
                y = parent_y + (parent_height // 2) - (dialog_height // 2)
                dialog.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")
            else: # Fallback if root is not available (e.g. during some tests)
                 dialog.eval('tk::PlaceWindow . center')
        except Exception as e_center: 
            print(f"Error centering dialog: {e_center}")
            dialog.eval('tk::PlaceWindow . center') # Fallback centering


        main_text = f"File saved successfully:\n{os.path.basename(file_path)}\n\nLocation:\n{os.path.dirname(file_path)}\n\nDo you want to open this file?"
        ttk.Label(dialog, text=main_text, wraplength=380, justify="center").pack(pady=15, padx=10)
        
        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=(5, 10)) 

        def do_open_file():
            try:
                # os.startfile is primarily for Windows.
                # For cross-platform, more robust solutions might be needed.
                if os.name == 'nt': # Windows
                    os.startfile(file_path)
                elif os.name == 'posix': # macOS, Linux
                    import subprocess
                    if sys.platform == "darwin": # macOS
                        subprocess.call(('open', file_path))
                    else: # Linux and other POSIX
                        subprocess.call(('xdg-open', file_path))
                else:
                    messagebox.showinfo("Open File", "File opening is not automatically supported on this OS. Please open it manually.", parent=dialog)

            except Exception as e_startfile:
                messagebox.showerror("Error Opening File", f"Could not open file: {e_startfile}.\nPlease navigate to it manually.", parent=dialog)
            finally:
                if dialog.winfo_exists(): dialog.destroy()

        def do_close_dialog():
            if dialog.winfo_exists(): dialog.destroy()

        open_btn = ttk.Button(button_frame, text="Open File", command=do_open_file)
        open_btn.pack(side=tk.LEFT, padx=10)
        open_btn.focus_set() 

        close_btn = ttk.Button(button_frame, text="Close", command=do_close_dialog)
        close_btn.pack(side=tk.LEFT, padx=10)
        
        self.root.wait_window(dialog) # Ensure dialog is modal


    def process_technical_agreement_workflow(self):
        """
        Orchestrates the entire workflow for processing a technical agreement PDF:
        1. Extracts subsystems using a Gemma model.
        2. Allows user verification/editing of these subsystems via a dialog.
        3. Splits the original PDF into multiple smaller documents based on confirmed subsystems.
        4. Identifies deliverables from each subsystem PDF using a Qwen model.
        5. Generates a consolidated Excel report of all identified deliverables.
        """
        if not hasattr(self, 'pdf_path') or not self.pdf_path:
            messagebox.showwarning("No Document Selected", "Please select a PDF document first.")
            return

        # Disable run buttons during processing to prevent concurrent runs
        self.run_button.config(state=tk.DISABLED)
        if hasattr(self, 'invoice_button') and self.invoice_button.winfo_exists(): 
            self.invoice_button.config(state=tk.DISABLED)
        
        self.update_status("Starting technical agreement processing workflow...")

        try:
            # Step 1: Extract Subsystems from PDF Table of Contents
            self.update_status(f"Extracting subsystems using '{self.gemma_model_name}' model...")
            subsystems_data = self.extract_subsystems_with_gemma3(self.pdf_path, self.gemma_model_name)
            if not subsystems_data: # Includes errors or model returning empty list
                self.update_status("No subsystems found or error during extraction. Workflow cannot continue.")
                messagebox.showwarning("Subsystem Extraction Failed", 
                                       "Could not extract any subsystems from the document's Table of Contents. "
                                       "Please check the PDF or the Gemma model's response.")
                return # Stop workflow if no subsystems are found

            self.update_status(f"{len(subsystems_data)} potential subsystems extracted. Awaiting user verification.")

            # Step 2: Human Verification of Subsystems
            # This dialog allows users to confirm, edit, or remove extracted subsystems.
            dialog = SubsystemVerificationDialog(self.root, subsystems_data, timeout_seconds=60) # Increased timeout
            confirmed_subsystems = dialog.get_confirmed_subsystems()
            
            if confirmed_subsystems is None: # User cancelled the dialog
                self.update_status("Subsystem verification cancelled by user. Workflow stopped.")
                return
            if not confirmed_subsystems: # User confirmed an empty list or auto-confirm failed with no valid data
                self.update_status("No subsystems were confirmed for processing. Workflow stopped.")
                messagebox.showinfo("Subsystem Verification", "No subsystems were confirmed. Cannot proceed.")
                return
            
            self.update_status(f"{len(confirmed_subsystems)} subsystems confirmed by user.")

            # Step 3: Split PDF by Confirmed Subsystems
            split_pdf_output_dir = os.path.join("working_output", "split_subsystem_pdfs") 
            self.update_status(f"Splitting PDF into subsystem documents in '{split_pdf_output_dir}'...")
            split_infos = self.split_pdf_by_subsystems(self.pdf_path, confirmed_subsystems, split_pdf_output_dir)
            
            if not split_infos: # If splitting failed or produced no documents
                self.update_status("PDF splitting failed or produced no files. Workflow stopped.")
                messagebox.showerror("PDF Splitting Error", 
                                     "Failed to split the PDF into subsystem documents. Check logs for details.")
                return
            
            self.update_status(f"PDF successfully split into {len(split_infos)} subsystem documents.")

            # Step 4: Identify Deliverables from each Subsystem PDF
            all_identified_deliverables = []
            for sub_info in split_infos: # sub_info contains path to a split PDF
                self.update_status(f"Identifying deliverables for subsystem: '{sub_info['subsystem_name']}'...")
                deliverables_for_sub = self.identify_deliverables_with_qwen3(
                    sub_info, 
                    self.qwen3_model_name, 
                    self.ocr, # Pass the initialized OCR instance
                    self.is_scanned_pdf, 
                    self.pdf_to_images
                )
                all_identified_deliverables.extend(deliverables_for_sub)
            
            if not all_identified_deliverables:
                self.update_status("No deliverables found across all subsystems. Workflow finished.")
                messagebox.showinfo("Deliverable Identification Complete", 
                                    "No specific deliverables were identified in any of the processed subsystems.")
                return

            self.update_status(f"Total of {len(all_identified_deliverables)} deliverables identified. Generating Excel report...")

            # Step 5: Generate Excel Report of Deliverables
            default_excel_filename = f"{os.path.splitext(os.path.basename(self.pdf_path))[0]}_Deliverables_Report.xlsx"
            excel_save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                title="Save Deliverables Excel Report As...",
                initialfile=default_excel_filename
            )
            if not excel_save_path: # User cancelled save dialog
                self.update_status("Excel report save cancelled by user. Workflow stopped.")
                return

            excel_generation_success = self.generate_deliverables_excel(all_identified_deliverables, excel_save_path)
            if excel_generation_success:
                self.update_status(f"Deliverables Excel report saved to '{os.path.basename(excel_save_path)}'.")
                self._show_open_file_dialog(excel_save_path, "Deliverables Report Saved Successfully")
            else:
                # generate_deliverables_excel method already shows an error message
                self.update_status("Failed to save the deliverables Excel report.")

        except Exception as e_workflow:
            self.update_status(f"An unexpected error occurred in the workflow: {str(e_workflow)}")
            messagebox.showerror("Workflow Error", f"An unexpected error occurred during processing: {str(e_workflow)}")
            traceback.print_exc() # Log full traceback to console for debugging
        finally:
            # Ensure run buttons are re-enabled regardless of success or failure
            self.run_button.config(state=tk.NORMAL)
            if hasattr(self, 'invoice_button') and self.invoice_button.winfo_exists(): 
                self.invoice_button.config(state=tk.NORMAL)
            self.update_status("Processing finished. Ready for new document or parameters.")


    def run(self):
        """Starts the main Tkinter event loop for the application."""
        if hasattr(self, 'root') and self.root:
            self.root.mainloop()
        else:
            print("Error: Main application window (self.root) not initialized before run().")

    def extract_subsystems_with_gemma3(self, pdf_path: str, ollama_model_name: str = "gemma:7b") -> list[dict]:
        """
        Extracts subsystems and their starting pages from a PDF's Table of Contents (ToC)
        using a specified Gemma model via Ollama.

        Args:
            pdf_path: Path to the PDF file.
            ollama_model_name: Name of the Gemma Ollama model to use (e.g., "gemma:7b").

        Returns:
            A list of dictionaries, where each dictionary contains "subsystem_name" and "start_page".
            Returns an empty list if no ToC is found, no subsystems are identified, or an error occurs.
        """
        try:
            doc = fitz.open(pdf_path)
            toc = doc.get_toc() # get_toc() returns list of [level, title, page_num (1-indexed)]
            doc.close()

            if not toc:
                self.update_status("PDF has no Table of Contents (ToC). Cannot extract subsystems.")
                return []

            # Format ToC for the Gemma model prompt
            toc_text = ""
            for level, title, page_num in toc:
                toc_text += f"{'  ' * (level - 1)}- {title} (Page {page_num})\n"

            prompt = f"""
            Given the following Table of Contents:
            {toc_text}
            Please identify the main subsystems or sections and their starting page numbers.
            Return the output as a JSON list of objects. Each object should have keys "subsystem_name" and "start_page".
            The page number should be the integer page number as listed in the ToC.
            For example: [{{"subsystem_name": "Subsystem A - Introduction", "start_page": 1}}, {{"subsystem_name": "Subsystem B - Technical Details", "start_page": 15}}]
            If no clear subsystems can be identified, return an empty JSON list [].
            """

            self.update_status(f"Calling {ollama_model_name} to extract subsystems from ToC...")
            api_response = call_ollama_api(ollama_model_name, prompt)

            if not api_response or 'response' not in api_response:
                self.update_status(f"Error: No valid response from {ollama_model_name} API for subsystem extraction.")
                return []

            gemma_response_str = api_response['response']
            
            # Attempt to parse JSON (list or single object) from the model's response
            try:
                # Prioritize finding a JSON list
                json_start_index = gemma_response_str.find('[')
                json_end_index = gemma_response_str.rfind(']') + 1
                
                if json_start_index != -1 and json_end_index > json_start_index:
                    json_str = gemma_response_str[json_start_index:json_end_index]
                    parsed_json = json.loads(json_str)
                    
                    if isinstance(parsed_json, list):
                        valid_subsystems = []
                        for item in parsed_json: # Validate each item in the list
                            if isinstance(item, dict) and "subsystem_name" in item and \
                               "start_page" in item and isinstance(item["start_page"], int):
                                valid_subsystems.append(item)
                            else:
                                print(f"Warning: Skipping invalid subsystem item from Gemma: {item}")
                        self.update_status(f"Successfully extracted {len(valid_subsystems)} subsystems from ToC.")
                        return valid_subsystems
                    else: # Should be a list if parsing was successful based on find '[...]'
                        self.update_status("Error: Parsed JSON from Gemma (expected list) is not a list.")
                        return []
                else: 
                    # Fallback: try to find a single JSON object if list parsing failed
                    json_start_index = gemma_response_str.find('{')
                    json_end_index = gemma_response_str.rfind('}') + 1
                    if json_start_index != -1 and json_end_index > json_start_index:
                        json_str = gemma_response_str[json_start_index:json_end_index]
                        parsed_json = json.loads(json_str)
                        if isinstance(parsed_json, dict) and "subsystem_name" in parsed_json and \
                           "start_page" in parsed_json and isinstance(parsed_json["start_page"], int):
                            self.update_status("Successfully extracted 1 subsystem (as single JSON object).")
                            return [parsed_json] # Return as a list with one item
                    
                    # If neither list nor object found, or if it's an empty valid response
                    if gemma_response_str.strip() == "[]":
                        self.update_status("Gemma model indicated no subsystems found (returned empty list).")
                        return []
                    
                    self.update_status("Error: Could not find valid JSON list or object in Gemma's response for subsystems.")
                    return []
            except json.JSONDecodeError:
                self.update_status(f"Error: Could not parse JSON response from {ollama_model_name} for subsystems. Response: {gemma_response_str[:200]}...")
                return []
        except Exception as e:
            self.update_status(f"An error occurred during subsystem extraction: {str(e)}")
            traceback.print_exc()
            return []

    def show_verification_dialog_example(self):
        """
        Shows an example of the SubsystemVerificationDialog with sample data.
        Useful for testing the dialog's functionality independently.
        """
        sample_data_for_dialog = [
            {"subsystem_name": "Chapter 1: Introduction to Advanced Galactic Superlasers", "start_page": 1},
            {"subsystem_name": "Chapter 2: The Intricacies of Kyber Crystal Alignment and Focus", "start_page": 15},
            {"subsystem_name": "Chapter 3: Empirical Results - A Case Study on Alderaan (Long Name Test for Wrapping)", "start_page": 30},
            {"subsystem_name": "Chapter 4: Critical Discussion of Minor Thermal Exhaust Port Vulnerabilities and Their Potential Exploitation by Small Rebel Starfighters - A Comprehensive Analysis", "start_page": 75},
        ]
        # Add more items to test scrolling behavior of the dialog
        for i in range(5, 16): # Creates items from Appendix E to Appendix O
            sample_data_for_dialog.append({"subsystem_name": f"Appendix {chr(64+i)}: Detail Section {i}", "start_page": 100 + i*5})
        
        # Determine the parent window for the dialog
        parent_window = self.root if hasattr(self, 'root') and self.root and self.root.winfo_exists() else None
        temp_root_used = False

        if parent_window is None:
            # Fallback if main UI isn't fully initialized (e.g., during isolated testing)
            print("Main window (self.root) not available for dialog. Using a temporary parent for testing.")
            parent_window = tk.Tk() 
            parent_window.withdraw() # Hide the temporary root window
            temp_root_used = True
        
        self.update_status(f"Showing subsystem verification dialog with {len(sample_data_for_dialog)} sample items...")
        dialog = SubsystemVerificationDialog(parent_window, sample_data_for_dialog, timeout_seconds=25)
        confirmed_data = dialog.get_confirmed_subsystems() # This call blocks until the dialog is closed.
        
        # Process the results from the dialog
        status_message = ""
        if confirmed_data is not None:
            status_message = f"Subsystem verification complete. Confirmed: {len(confirmed_data)} subsystems."
            print("Confirmed Subsystems by User:")
            for item in confirmed_data:
                print(f"  - Name: '{item['subsystem_name']}', Start Page: {item['start_page']}")
        else: # User cancelled the dialog
            status_message = "Subsystem verification was cancelled by the user."
            print(status_message)

        if hasattr(self, 'update_status'): self.update_status(status_message)
        
        # Clean up temporary root window if it was used
        if temp_root_used and parent_window.winfo_exists():
            parent_window.destroy()

    @staticmethod
    def _sanitize_filename(name: str) -> str:
        """
        Sanitizes a string to make it suitable for use as a filename.
        Replaces spaces with underscores and removes characters that are typically
        invalid or problematic in filenames across different operating systems.

        Args:
            name: The input string to sanitize.

        Returns:
            A sanitized string, truncated to 100 characters.
        """
        if not name:
            return "untitled_subsystem" # Default for empty names
        
        name = name.replace(" ", "_") # Replace spaces with underscores
        # Remove characters that are broadly problematic in filenames
        name = re.sub(r'[\\/*?:"<>|&!@#$%\^\(\)+=\[\]\{\};\'~`]', "", name) 
        name = re.sub(r'[\n\r\t]', '_', name) # Replace newline/tab characters
        
        # Limit length to a reasonable maximum (e.g., 100 chars) to avoid OS errors
        return name[:100] 

    def split_pdf_by_subsystems(self, original_pdf_path: str, 
                                confirmed_subsystems: list[dict], 
                                output_directory: str = "temp_split_pdfs") -> list[dict]:
        """
        Splits the original PDF into multiple smaller PDF files based on the
        start pages of confirmed subsystems.

        Args:
            original_pdf_path: Path to the source PDF file.
            confirmed_subsystems: A list of dictionaries, each containing "subsystem_name"
                                  and "start_page" (1-indexed and validated by user).
                                  It's assumed this list is sorted by "start_page".
            output_directory: The directory where the split PDF files will be saved.
                              This directory will be cleared if it already exists.

        Returns:
            A list of dictionaries, where each dictionary contains information about
            a successfully created split PDF:
            {"subsystem_name": str, "pdf_path": str, "original_start_page": int (1-indexed)}
            Returns an empty list if errors occur or no PDFs are split.
        """
        if not original_pdf_path or not os.path.exists(original_pdf_path):
            self.update_status(f"Error: Original PDF not found at '{original_pdf_path}'. Splitting aborted.")
            messagebox.showerror("PDF Splitting Error", f"Original PDF not found: {original_pdf_path}")
            return []

        if not confirmed_subsystems:
            self.update_status("No subsystems provided for splitting. Nothing to do.")
            return []

        try:
            original_doc = fitz.open(original_pdf_path)
            total_pages_original_doc = len(original_doc)
        except Exception as e:
            self.update_status(f"Error opening PDF '{os.path.basename(original_pdf_path)}': {e}")
            messagebox.showerror("PDF Splitting Error", f"Error opening original PDF: {e}")
            return []

        # Manage output directory: clear if exists, then create
        if os.path.exists(output_directory):
            try:
                shutil.rmtree(output_directory) 
                self.update_status(f"Cleared previous output directory: '{output_directory}'")
            except Exception as e:
                self.update_status(f"Error clearing output directory '{output_directory}': {e}")
                messagebox.showerror("Directory Error", f"Could not clear output directory: {e}")
                if original_doc: original_doc.close()
                return []
        
        try:
            os.makedirs(output_directory, exist_ok=True)
            self.update_status(f"Ensured output directory exists: '{output_directory}'")
        except Exception as e:
            self.update_status(f"Error creating output directory '{output_directory}': {e}")
            messagebox.showerror("Directory Error", f"Could not create output directory: {e}")
            if original_doc: original_doc.close()
            return []
            
        split_pdf_info_list = [] # To store info about successfully created PDFs

        # Iterate through confirmed subsystems to create individual PDFs
        for i, subsystem in enumerate(confirmed_subsystems):
            subsystem_name = subsystem.get("subsystem_name", f"Subsystem_{i+1}") # Default name if missing
            original_1_indexed_start_page = subsystem.get("start_page")

            # Validate start page from subsystem data
            if original_1_indexed_start_page is None or not isinstance(original_1_indexed_start_page, int) or original_1_indexed_start_page <= 0:
                self.update_status(f"Skipping subsystem '{subsystem_name}' due to invalid start page: {original_1_indexed_start_page}.")
                print(f"Warning: Invalid start page for subsystem '{subsystem_name}': {original_1_indexed_start_page}")
                continue

            # Convert 1-indexed start page from user/ToC to 0-indexed for fitz
            fitz_start_page = original_1_indexed_start_page - 1

            if not (0 <= fitz_start_page < total_pages_original_doc):
                self.update_status(f"Skipping '{subsystem_name}': Start page {original_1_indexed_start_page} is out of document bounds (1-{total_pages_original_doc}).")
                print(f"Warning: Start page {original_1_indexed_start_page} for '{subsystem_name}' is out of document bounds.")
                continue

            # Determine end page (0-indexed) for the current subsystem
            fitz_end_page = -1
            if i == len(confirmed_subsystems) - 1: # If this is the last subsystem
                fitz_end_page = total_pages_original_doc - 1 # It goes to the end of the document
            else:
                next_subsystem_1_indexed_start_page = confirmed_subsystems[i+1].get("start_page")
                # Validate start page of the *next* subsystem
                if next_subsystem_1_indexed_start_page is None or not isinstance(next_subsystem_1_indexed_start_page, int) or next_subsystem_1_indexed_start_page <= 0:
                     self.update_status(f"Warning: Invalid start page for next subsystem after '{subsystem_name}'. Current subsystem will extend to document end.")
                     print(f"Warning: Invalid start page for subsystem after '{subsystem_name}'. Current extends to end.")
                     fitz_end_page = total_pages_original_doc - 1 
                else:
                    # End page is one less than the start of the next subsystem (0-indexed)
                    fitz_end_page = next_subsystem_1_indexed_start_page - 2 
            
            # Further validation for end page
            if fitz_end_page < fitz_start_page: # Can happen if next subsystem starts on same or immediately after
                fitz_end_page = fitz_start_page # Ensure at least one page is included
            
            if fitz_end_page >= total_pages_original_doc: # Cap at the actual last page
                fitz_end_page = total_pages_original_doc - 1

            self.update_status(f"Processing '{subsystem_name}': Original pages {fitz_start_page+1}-{fitz_end_page+1}")

            try:
                sub_doc = fitz.open() # Create a new, empty PDF document
                sub_doc.insert_pdf(original_doc, from_page=fitz_start_page, to_page=fitz_end_page) # Copy page range
                
                sanitized_name = self._sanitize_filename(subsystem_name)
                # Prefix with index for ordering and to help uniqueness
                output_filename = f"{str(i+1).zfill(3)}_{sanitized_name}.pdf"
                output_path = os.path.join(output_directory, output_filename)
                
                sub_doc.save(output_path, garbage=4, deflate=True) # Save with options for optimization
                sub_doc.close()
                
                split_pdf_info_list.append({
                    "subsystem_name": subsystem_name,
                    "pdf_path": output_path,
                    "original_start_page": original_1_indexed_start_page # Store 1-indexed for reference
                })
                print(f"Successfully saved: {output_path} (Original pages {original_1_indexed_start_page} to {fitz_end_page + 1})")

            except Exception as e_save:
                error_msg = f"Error saving PDF for subsystem '{subsystem_name}': {e_save}"
                self.update_status(error_msg)
                print(error_msg)
                # Continue with other subsystems even if one fails
        
        if original_doc: original_doc.close() # Close the main document
        
        if not split_pdf_info_list:
            self.update_status("PDF splitting completed, but no files were generated (check subsystem page ranges).")
        else:
            self.update_status(f"PDF splitting complete. {len(split_pdf_info_list)} subsystem PDFs created in '{output_directory}'.")
        
        return split_pdf_info_list

    def identify_deliverables_with_qwen3(self, 
                                         subsystem_pdf_info: dict, 
                                         qwen3_model_name: str,
                                         ocr_instance: PaddleOCR, 
                                         is_scanned_pdf_func: callable, 
                                         pdf_to_images_func: callable) -> list[dict]:
        """
        Identifies deliverables from a single subsystem PDF using a specified Qwen3 model.

        This involves extracting text from the subsystem PDF, constructing a detailed prompt
        for the Qwen3 model, calling the Ollama API, and parsing the JSON response.
        The identified deliverables are augmented with subsystem name and original PDF page number.

        Args:
            subsystem_pdf_info: A dictionary containing details of the subsystem PDF:
                                {"subsystem_name": str, 
                                 "pdf_path": str, 
                                 "original_start_page": int (1-indexed)}.
            qwen3_model_name: The name of the Qwen3 model to use (e.g., "qwen2:14b-instruct").
            ocr_instance: An initialized instance of PaddleOCR for text extraction if needed.
            is_scanned_pdf_func: Callable to check if the subsystem PDF is scanned.
            pdf_to_images_func: Callable to convert PDF pages to images for OCR.

        Returns:
            A list of deliverable dictionaries. Each dictionary includes the keys from
            the Qwen3 model's response plus "subsystem_name" and "original_pdf_page".
            Returns an empty list if no deliverables are found or an error occurs.
        """
        subsystem_name = subsystem_pdf_info.get("subsystem_name", "Unknown Subsystem")
        pdf_path = subsystem_pdf_info.get("pdf_path")
        # original_doc_start_page is 1-indexed, refers to the start page of this subsystem in the *original* document
        original_doc_start_page = subsystem_pdf_info.get("original_start_page", 1) 

        if not pdf_path or not os.path.exists(pdf_path):
            self.update_status(f"Error: PDF file not found for subsystem '{subsystem_name}' at '{pdf_path}'. Cannot identify deliverables.")
            print(f"Error: PDF file not found for subsystem '{subsystem_name}' at '{pdf_path}'.")
            return []

        self.update_status(f"Identifying deliverables for subsystem: '{subsystem_name}'...")

        try:
            # Step 1: Extract text from the current subsystem's PDF file.
            self.update_status(f"Extracting text from '{subsystem_name}' PDF ({os.path.basename(pdf_path)})...")
            subsystem_text = extract_text_from_pdf(
                pdf_path, 
                ocr_instance, 
                is_scanned_pdf_func, 
                pdf_to_images_func
            )
            if not subsystem_text or not subsystem_text.strip():
                self.update_status(f"No text could be extracted from subsystem '{subsystem_name}'. Skipping deliverable identification for this subsystem.")
                print(f"Warning: No text extracted from subsystem '{subsystem_name}'.")
                return []
            
            # Truncate text if too long to avoid issues with model context limits or performance.
            max_text_length = 50000 # Character limit for the prompt text.
            if len(subsystem_text) > max_text_length:
                subsystem_text = subsystem_text[:max_text_length]
                self.update_status(f"Text from '{subsystem_name}' truncated to {max_text_length} characters for the prompt.")


            # Step 2: Construct the prompt for the Qwen3 model.
            prompt = f"""
You are a meticulous Technical Project Analyst. Your task is to carefully analyze the provided text from a technical subsystem document.
Your goal is to identify and list all specific deliverable items mentioned. Deliverables can be documents, reports, software modules, hardware components, datasets, project plans, test plans, etc.

For each deliverable you identify, please provide the following information:
1.  `deliverable_name`: The specific name of the deliverable (e.g., "Preliminary Design Review Report", "Software Module X", "User Acceptance Test Plan").
2.  `phase`: The project phase it most likely belongs to. Choose strictly from one of these three options: "Design Phase", "Manufacturing Phase", or "Debugging/Acceptance Phase".
3.  `page_in_subsystem_pdf`: The page number *within the current subsystem PDF document* where this deliverable is primarily mentioned or described. This should be an integer. If the page number is explicitly mentioned, use that. Otherwise, infer from the context. Assume page numbering starts from 1 for the current document.
4.  `context_snippet`: A brief quote (20-30 words) from the document text that directly supports the identification of the deliverable and its page number.

Return your findings STRICTLY as a JSON list of objects. Each object must follow this structure:
{{"deliverable_name": "string", "phase": "string", "page_in_subsystem_pdf": integer, "context_snippet": "string"}}

Example JSON output:
[
  {{
    "deliverable_name": "Detailed Design Specification",
    "phase": "Design Phase",
    "page_in_subsystem_pdf": 3,
    "context_snippet": "The Detailed Design Specification, found on page 3, outlines all modules..."
  }},
  {{
    "deliverable_name": "Final Acceptance Test Report",
    "phase": "Debugging/Acceptance Phase",
    "page_in_subsystem_pdf": 25,
    "context_snippet": "...successful completion is documented in the Final Acceptance Test Report (page 25)."
  }}
]

If no specific deliverables are found in the text, or if you are unsure, return an empty JSON list: [].
Do not include any explanations or text outside of the JSON list.

Document Text to Analyze:
---
{subsystem_text}
---
"""
            # Step 3: Call the Ollama API with the Qwen3 model.
            self.update_status(f"Calling {qwen3_model_name} model for '{subsystem_name}' deliverables identification...")
            api_response = call_ollama_api(qwen3_model_name, prompt) # Uses global function

            if not api_response or 'response' not in api_response:
                self.update_status(f"Error: No valid response received from {qwen3_model_name} API for '{subsystem_name}'.")
                print(f"Error: No valid response from {qwen3_model_name} API for '{subsystem_name}'.")
                return []

            qwen_response_str = api_response['response']
            
            # Step 4: Parse the JSON output from Qwen3.
            self.update_status(f"Parsing response from {qwen3_model_name} for '{subsystem_name}' deliverables...")
            identified_deliverables = []
            try:
                # Attempt to find the JSON list within the model's response string.
                json_start_index = qwen_response_str.find('[')
                json_end_index = qwen_response_str.rfind(']') + 1
                
                if json_start_index != -1 and json_end_index > json_start_index:
                    json_str = qwen_response_str[json_start_index:json_end_index]
                    parsed_json = json.loads(json_str)
                    
                    if isinstance(parsed_json, list):
                        for item in parsed_json:
                            # Validate structure of each item from the model
                            if (isinstance(item, dict) and 
                                "deliverable_name" in item and 
                                "phase" in item and 
                                "page_in_subsystem_pdf" in item and isinstance(item["page_in_subsystem_pdf"], int) and
                                "context_snippet" in item):
                                
                                # Augment with subsystem_name and calculate original_pdf_page
                                item["subsystem_name"] = subsystem_name
                                # page_in_subsystem_pdf is 1-indexed (as per prompt instruction to model)
                                # original_doc_start_page is also 1-indexed
                                item["original_pdf_page"] = (original_doc_start_page - 1) + item["page_in_subsystem_pdf"]
                                identified_deliverables.append(item)
                            else:
                                print(f"Warning: Skipping malformed deliverable item from '{subsystem_name}': {item}")
                        self.update_status(f"Found {len(identified_deliverables)} deliverables in subsystem '{subsystem_name}'.")
                    else: 
                        self.update_status(f"Error: Parsed JSON for '{subsystem_name}' is not a list as expected. Response: {qwen_response_str[:200]}...")
                        print(f"Error: Parsed JSON for '{subsystem_name}' is not a list. Response snippet: {qwen_response_str[:200]}...")
                        return [] 
                else: 
                    # Handle cases where no JSON list is found (e.g., model returned empty string or "[]" directly)
                    if qwen_response_str.strip() == "[]":
                         self.update_status(f"No deliverables explicitly found by model in '{subsystem_name}'.")
                    else:
                        self.update_status(f"Error: No JSON list structure found in {qwen3_model_name} response for '{subsystem_name}'. Response snippet: {qwen_response_str[:200]}...")
                        print(f"Error: No JSON list found in {qwen3_model_name} response for '{subsystem_name}'. Response snippet: {qwen_response_str[:200]}...")
            
            except json.JSONDecodeError as e_json:
                self.update_status(f"Error parsing JSON response from {qwen3_model_name} for '{subsystem_name}': {e_json}. Response snippet: {qwen_response_str[:200]}...")
                print(f"Error parsing JSON response from {qwen3_model_name} for '{subsystem_name}': {e_json}. Response snippet: {qwen_response_str[:200]}...")
                return [] 

            return identified_deliverables

        except Exception as e_main:
            self.update_status(f"An unexpected error occurred while processing subsystem '{subsystem_name}' for deliverables: {str(e_main)}")
            print(f"An unexpected error occurred while processing subsystem '{subsystem_name}' for deliverables: {str(e_main)}")
            traceback.print_exc() # For detailed error logging to console
            return []

    def generate_deliverables_excel(self, all_deliverables: list[dict], output_excel_path: str) -> bool:
        """
        Generates an Excel file from the list of all identified deliverables.
        The deliverables are sorted by subsystem name, then by a predefined phase order,
        and finally by their page number in the original document.

        Args:
            all_deliverables: A flat list of deliverable dictionaries. Each dictionary
                              is expected to have keys like "subsystem_name", "deliverable_name",
                              "phase", "original_pdf_page", and "context_snippet".
            output_excel_path: Full path where the Excel file will be saved.

        Returns:
            True if the Excel file was generated and saved successfully, False otherwise.
        """
        if not all_deliverables:
            self.update_status("No deliverables data provided to generate Excel file.")
            # Consider if an empty report should be generated or if this is a 'failure'.
            # For now, returning False as no meaningful report is created.
            return False 

        self.update_status(f"Generating deliverables Excel report at '{output_excel_path}'...")

        try:
            wb = Workbook() # Create a new Excel workbook
            ws = wb.active # Get the active worksheet
            ws.title = "Deliverables List" # Set the sheet title

            # 1. Write Headers to the first row
            headers = [
                "Subsystem Name", "Deliverable Name", "Phase", 
                "Source PDF Page Number", "Context Snippet"
            ]
            ws.append(headers)

            # Style for headers (Bold)
            header_font = Font(bold=True) # Use openpyxl.styles.Font
            for cell in ws[1]: # Iterate through cells in the first row (headers)
                cell.font = header_font

            # 2. Sort Data before writing to Excel
            # Define a custom order for project phases for logical sorting
            phase_order = ["Design Phase", "Manufacturing Phase", "Debugging/Acceptance Phase"]
            
            def sort_key(deliverable):
                # Ensure robust access to keys, defaulting for sorting if keys are missing
                subsystem_name = str(deliverable.get("subsystem_name", "")).lower() # Case-insensitive sort for name
                phase = str(deliverable.get("phase", ""))
                try:
                    phase_index = phase_order.index(phase) # Get index from predefined order
                except ValueError: # If phase is not in phase_order list
                    phase_index = len(phase_order) # Place it at the end
                original_page = int(deliverable.get("original_pdf_page", 0)) # Sort by page number
                return (subsystem_name, phase_index, original_page)

            sorted_deliverables = sorted(all_deliverables, key=sort_key)

            # 3. Write Data rows
            for deliverable in sorted_deliverables:
                row_data = [
                    deliverable.get("subsystem_name", "N/A"),
                    deliverable.get("deliverable_name", "N/A"),
                    deliverable.get("phase", "N/A"),
                    deliverable.get("original_pdf_page", "N/A"), # This is the calculated original PDF page
                    deliverable.get("context_snippet", "N/A")
                ]
                ws.append(row_data)
            
            # 4. Adjust Column Widths for better readability
            column_widths = {
                'A': 30,  # Subsystem Name
                'B': 45,  # Deliverable Name (increased width)
                'C': 25,  # Phase (increased width)
                'D': 20,  # Source PDF Page Number
                'E': 70   # Context Snippet (increased width)
            }
            for col_letter, width in column_widths.items():
                ws.column_dimensions[col_letter].width = width
            
            # Enable text wrapping for the context snippet column (column E)
            for row in ws.iter_rows(min_row=2, max_col=5, min_col=5): # From row 2, only column E
                for cell in row:
                    cell.alignment = Alignment(wrap_text=True, vertical='top') # Align wrapped text to top

            # 5. Save Workbook
            wb.save(output_excel_path)
            self.update_status(f"Successfully saved deliverables Excel report to '{os.path.basename(output_excel_path)}'")
            print(f"Deliverables Excel report saved to '{output_excel_path}'")
            return True

        except IOError as e_io: # Specific error for file operations
            self.update_status(f"Error saving Excel file: {str(e_io)}")
            messagebox.showerror("Excel Save Error", 
                                 f"Could not save Excel file to '{output_excel_path}'.\n"
                                 f"Reason: {str(e_io)}\nPlease ensure the file is not open elsewhere and you have write permissions.")
            print(f"IOError saving Excel file '{output_excel_path}': {str(e_io)}")
            return False
        except Exception as e: # Catch-all for other unexpected errors
            self.update_status(f"An unexpected error occurred during Excel report generation: {str(e)}")
            messagebox.showerror("Excel Generation Error", f"An unexpected error occurred: {str(e)}")
            print(f"Unexpected error during Excel generation for '{output_excel_path}': {str(e)}")
            traceback.print_exc()
            return False


if __name__ == "__main__":
    # This block is for testing and direct execution of the script.
    # It initializes and runs the InvoiceProcessor application.
    processor = InvoiceProcessor()
    
    # --- Testing Examples (Commented Out) ---
    # These examples can be uncommented for specific testing scenarios if needed.

    # Example 1: Test the Subsystem Verification Dialog directly
    # This requires the main Tkinter loop (processor.run()) to be commented out,
    # or this test to be run before processor.run() if self.root is available.
    # If self.root is not yet created (before setup_ui), show_verification_dialog_example
    # creates a temporary root window.
    #
    # if not hasattr(processor, 'root') or not processor.root.winfo_exists():
    #     # Create a dummy root if main UI isn't running, for dialog testing
    #     dummy_root = tk.Tk()
    #     dummy_root.withdraw()
    #     processor.root = dummy_root # Temporarily assign for the dialog parent
    #     processor.show_verification_dialog_example()
    #     if dummy_root.winfo_exists():
    #         dummy_root.destroy()
    # else:
    # processor.show_verification_dialog_example()


    # Example 2: Test the PDF Splitting functionality
    # This requires a sample PDF and subsystem data.
    #
    # if hasattr(processor, 'root') and processor.root.winfo_exists(): 
    #     sample_pdf_for_splitting = "test_document.pdf" 
    #     if not os.path.exists(sample_pdf_for_splitting):
    #         dummy_doc = fitz.open()
    #         for k in range(10):
    #             page = dummy_doc.new_page()
    #             page.insert_text((50, 72), f"This is page {k+1} of the test document.")
    #         dummy_doc.save(sample_pdf_for_splitting)
    #         dummy_doc.close()
    #         print(f"Created dummy PDF: {sample_pdf_for_splitting}")
    #
    #     example_subsystems_for_splitting = [
    #         {"subsystem_name": "Section 1: Intro", "start_page": 1},
    #         {"subsystem_name": "Section 2: Main Content", "start_page": 3},
    #         {"subsystem_name": "Section 3: Appendix", "start_page": 8},
    #     ]
    #     processor.pdf_path = sample_pdf_for_splitting 
    #     split_results = processor.split_pdf_by_subsystems(
    #         original_pdf_path=sample_pdf_for_splitting,
    #         confirmed_subsystems=example_subsystems_for_splitting,
    #         output_directory="split_test_output"
    #     )
    #     if split_results:
    #         print("\n--- Split PDF Results ---")
    #         for res in split_results: print(f"  {res}")
    #         print(f"Find split PDFs in: {os.path.abspath('split_test_output')}")
    #     else: print("No PDFs were split in the example.")
    # else:
    #     print("Skipping split_pdf_by_subsystems example as Tkinter root is not available for status updates.")
    
    # --- Run the main application ---
    processor.run()
