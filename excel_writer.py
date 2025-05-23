import os
from openpyxl import Workbook, load_workbook
from openpyxl import styles # For cell protection

# Define the order of keys as expected by the original Excel mapping logic
# This order corresponds to how data is written to temp_data.xlsx (A1, B1, ...)
# and then mapped to the final user template.
KEYS_IN_ORDER = [
    "buyer_name", "buyer_tax_id", "buyer_bank_name", "buyer_bank_account", 
    "invoice_code", "item_name", "item_quantity", "item_unit_price", 
    "seller_name", "seller_tax_id", "seller_bank_name", "seller_bank_account", 
    "invoice_date_YYYYMMDD", "invoice_number", "total_amount_with_tax", 
    "tax_rate", "tax_amount", "total_amount_without_tax", "invoice_type"
]

class ExcelWriter:
    def __init__(self):
        # The working directory for temporary files.
        self.work_dir = "working"
        if not os.path.exists(self.work_dir):
            os.makedirs(self.work_dir)

    def create_invoice_excel(self, output_path, template_path, invoice_data_dict):
        """
        Creates the final invoice Excel file by populating a template with extracted data.

        Args:
            output_path (str): The path where the final Excel file will be saved.
            template_path (str): The path to the user-selected Excel template file.
            invoice_data_dict (dict): A dictionary containing the extracted invoice data
                                      from OllamaClient.
        Returns:
            bool: True if the Excel file was created successfully, False otherwise.
        """
        temp_file_path = os.path.join(self.work_dir, "temp_data.xlsx")

        try:
            # 1. Create and save the temporary Excel file with data from Ollama
            temp_wb = Workbook()
            temp_ws = temp_wb.active
            
            for i, key in enumerate(KEYS_IN_ORDER):
                cell_char = chr(65 + i)  # A, B, C...
                temp_ws[f"{cell_char}1"] = invoice_data_dict.get(key, "")
            
            temp_wb.save(temp_file_path)

            # 2. Load the user's template Excel file
            if not template_path: # Should be handled by UI, but double check
                # print("Error: Template path is missing.") # For debugging
                return False
                
            wb = load_workbook(template_path)
            ws = wb.active

            # 3. Clear existing data from the template (from 3rd row onwards)
            # Iterate downwards to avoid issues with row deletion
            for row_idx in range(ws.max_row, 2, -1): 
                ws.delete_rows(row_idx)
            
            # 4. Unlock cells in the third row (target row for new data)
            # This assumes the template structure where data is written to the 3rd row.
            # A more robust solution might involve configuration or smarter row detection.
            if ws.max_row >= 2: # Ensure there are at least header rows
                # Check if row 3 exists, if not, it implies template has < 2 rows or they were all deleted.
                # This logic implicitly assumes row 3 is the target for new data.
                # If template has only 1 or 0 rows, ws[3] would error.
                # If template has 2 rows, ws[3] is the next available row.
                # A more robust approach might be to append data or use named ranges/tables.
                # For now, matching original logic:
                try:
                    for cell in ws[3]: # Attempt to access cells in row 3
                        cell.protection = styles.Protection(locked=False)
                except IndexError:
                    # This can happen if the sheet has fewer than 3 rows after deletion.
                    # Or if the sheet was empty.
                    # print(f"Warning: Worksheet '{ws.title}' has fewer than 3 rows. Cannot unlock cells in row 3.")
                    pass # Continue, data writing below will use ws.append or specific cell refs

            # 5. Map data from temp_ws (which holds Ollama output in A1:S1) to the main worksheet (ws)
            # This mapping is based on the original script's logic.
            # Using .get from invoice_data_dict directly now, as temp_file was just for original structure.
            
            # Helper to get value from invoice_data_dict using KEYS_IN_ORDER index
            def get_val_by_original_temp_cell(temp_cell_idx): # 0 for A1, 1 for B1 etc.
                return invoice_data_dict.get(KEYS_IN_ORDER[temp_cell_idx], "")

            ws['E3'] = get_val_by_original_temp_cell(0)   # Buyer Name (Original A1)
            ws['F3'] = get_val_by_original_temp_cell(1)   # Buyer Tax ID (Original B1)
            ws['B3'] = get_val_by_original_temp_cell(4)   # Invoice Code (Original E1)
            ws['I3'] = f"{get_val_by_original_temp_cell(5)}*00010-"  # Goods/Service Name (Original F1) + "*00010-"
            ws['L3'] = get_val_by_original_temp_cell(6)   # Quantity (Original G1)
            ws['M3'] = get_val_by_original_temp_cell(7)   # Unit Price (Original H1)
            ws['G3'] = get_val_by_original_temp_cell(8)   # Seller Name (Original I1)
            ws['H3'] = get_val_by_original_temp_cell(9)   # Seller Tax ID (Original J1)
            ws['D3'] = get_val_by_original_temp_cell(12)  # Invoice Date (YYYYMMDD) (Original M1)
            ws['C3'] = get_val_by_original_temp_cell(13)  # Invoice Number (Original N1)
            ws['Q3'] = get_val_by_original_temp_cell(14)  # Total Amount (incl. tax) (Original O1)
            ws['O3'] = get_val_by_original_temp_cell(15)  # Tax Rate (Original P1)
            ws['P3'] = get_val_by_original_temp_cell(16)  # Tax Amount (Original Q1)
            ws['N3'] = get_val_by_original_temp_cell(17)  # Amount (excl. tax) (Original R1)
            ws['A3'] = get_val_by_original_temp_cell(18)  # Invoice Type (Original S1)

            # 6. Save the final workbook
            wb.save(output_path)
            return True

        except Exception as e:
            # print(f"Error in ExcelWriter: {e}") # For debugging
            # Consider using tkinter.messagebox here or raising exception for main_app to handle
            return False
        finally:
            # 7. Delete the temporary Excel file
            if os.path.exists(temp_file_path):
                try:
                    os.remove(temp_file_path)
                except Exception as e_remove:
                    # print(f"Error deleting temporary file {temp_file_path}: {e_remove}") # For debugging
                    pass # Log or handle as needed, but don't let it hide main error.
