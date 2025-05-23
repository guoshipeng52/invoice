import requests
import json
from tkinter import messagebox # Required for error popups in check_connection_and_model

class OllamaClient:
    def __init__(self, api_url="http://localhost:11434/api/generate", model_name="qwen2.5:14b"):
        self.api_url = api_url
        self.model_name = model_name

    def check_connection_and_model(self):
        """
        Tests if Ollama is running and the specified model is available.
        Sends a simple test prompt to the Ollama API.
        Returns:
            bool: True if connection and model are okay, False otherwise.
        """
        try:
            response = requests.post(
                self.api_url,
                json={
                    "model": self.model_name,
                    "prompt": "test", # A simple prompt to check model availability
                    "stream": False
                }
            )
            if response.status_code == 200:
                return True
            else:
                messagebox.showerror("Ollama Error", f"Failed to connect to Ollama or model '{self.model_name}' not available. Status: {response.status_code}\n{response.text}")
                return False
        except requests.exceptions.RequestException as e:
            messagebox.showerror("Ollama Connection Error", f"Could not connect to Ollama at {self.api_url}.\nPlease ensure Ollama is running.\nError: {str(e)}")
            return False
        except Exception as e:
            messagebox.showerror("Ollama Error", f"An unexpected error occurred while checking Ollama: {str(e)}")
            return False

    def extract_invoice_data_from_text(self, text_content_list):
        """
        Sends text content to Ollama to extract invoice data using a JSON prompt.
        Args:
            text_content_list (list): A list of strings representing lines of text from the invoice.
        Returns:
            dict: Parsed JSON data extracted from the invoice, or None if an error occurs.
        """
        prompt = f'''Please extract the following information from the provided text and return it as a single, minified JSON object. Do not include any explanatory text before or after the JSON object. The JSON keys should be exactly as specified below:
"buyer_name", "buyer_tax_id", "buyer_bank_name", "buyer_bank_account", "invoice_code", "item_name", "item_quantity", "item_unit_price", "seller_name", "seller_tax_id", "seller_bank_name", "seller_bank_account", "invoice_date_YYYYMMDD", "invoice_number", "total_amount_with_tax", "tax_rate", "tax_amount", "total_amount_without_tax", "invoice_type"

If a value is not found or not applicable, use an empty string "" or null for that key. For numerical fields like quantity, unit_price, amounts, tax_rate, if not found, use 0 or null. Ensure numerical values are returned as numbers, not strings, if possible. 'item_quantity' should default to 1 if not found. 'invoice_date_YYYYMMDD' should be in YYYYMMDD format. 'invoice_type' should be "普通发票" or "专用发票".

Text content:
{ "\n".join(text_content_list) }'''

        try:
            response = requests.post(
                self.api_url,
                json={
                    "model": self.model_name,
                    "prompt": prompt,
                    "stream": False,
                    "format": "json"
                }
            )
            
            if response.status_code == 200:
                result = response.json()
                ollama_response_json_str = result.get('response', '{}').strip()
                
                if ollama_response_json_str.startswith("```json"):
                    ollama_response_json_str = ollama_response_json_str[7:]
                if ollama_response_json_str.endswith("```"):
                    ollama_response_json_str = ollama_response_json_str[:-3]
                ollama_response_json_str = ollama_response_json_str.strip()
                
                if not ollama_response_json_str: # Handle empty string case
                    # print("Ollama returned an empty JSON string.") # For debugging
                    return {} # Return empty dict if response is empty after stripping

                extracted_data = json.loads(ollama_response_json_str)
                return extracted_data
            else:
                # print(f"Ollama API Error: {response.status_code}\n{response.text}") # For debugging
                # Consider raising an exception or returning a more specific error object
                return None
        except json.JSONDecodeError as e:
            # print(f"JSON Decode Error: {e}. Response: {ollama_response_json_str}") # For debugging
            return None
        except requests.exceptions.RequestException as e:
            # print(f"Ollama Request Exception: {e}") # For debugging
            return None
        except Exception as e:
            # print(f"Unexpected error in Ollama client: {e}") # For debugging
            return None
