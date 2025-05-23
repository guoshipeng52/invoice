import json
import os

class ConfigManager:
    def __init__(self, config_filepath='config.json'):
        """
        Initializes the ConfigManager.
        Args:
            config_filepath (str): The path to the configuration file.
        """
        self.config_filepath = config_filepath
        self.defaults = {
            # Ollama Settings
            "ollama_api_url": "http://localhost:11434/api/generate",
            "ollama_model_name": "qwen2.5:14b",
            "ollama_keep_alive": "5m", # Duration Ollama keeps models in memory
            "ollama_temperature": 0.7, # Controls randomness, lower is more deterministic
            
            # OCR Settings (PaddleOCR specific)
            "ocr_lang": "ch", # Language for OCR (e.g., "ch" for Chinese, "en" for English)
            "ocr_use_gpu": True, # Whether to use GPU for OCR if available
            "ocr_det_model_dir": "", # Path to custom detection model (empty for default)
            "ocr_rec_model_dir": "", # Path to custom recognition model (empty for default)
            "ocr_cls_model_dir": "", # Path to custom classification model (empty for default)
            
            # Hardware Settings
            "hardware_vram_threshold_gb": 4, # VRAM threshold in GB for warnings
            
            # Path Settings
            "default_pdf_input_dir": "", # Default directory for opening PDF files
            "default_excel_template_path": "", # Default path to an Excel template
            "default_output_dir": "./output", # Default directory for saving output files
            
            # Processing Settings
            "processing_auto_open_excel": True, # Auto-open generated Excel file
            "processing_save_ollama_json": False, # Save raw Ollama JSON response
            "processing_intermediate_text_file": "ocr_result.txt", # Name for intermediate text file
            
            # Database Logging Settings (Future use)
            "database_enabled": False,
            "database_filepath": "invoices.db"
        }
        self.settings = self.defaults.copy()
        self.load_config()

    def load_config(self):
        """
        Loads configuration from the JSON file if it exists, updating defaults.
        If the file doesn't exist or is invalid, defaults (or current settings) are used.
        """
        if os.path.exists(self.config_filepath):
            try:
                with open(self.config_filepath, 'r', encoding='utf-8') as f:
                    loaded_settings = json.load(f)
                
                # Update settings with loaded values, only for keys present in defaults
                for key in self.defaults.keys():
                    if key in loaded_settings:
                        self.settings[key] = loaded_settings[key]
                # print(f"Config loaded from {self.config_filepath}") # For debugging
            except json.JSONDecodeError:
                print(f"Warning: Configuration file {self.config_filepath} is corrupted. Using default settings.")
            except Exception as e:
                print(f"Warning: Error loading configuration from {self.config_filepath}: {e}. Using default settings.")
        else:
            # print(f"Info: Configuration file {self.config_filepath} not found. Using default settings and creating it on save.")
            # Optionally, save defaults immediately if file doesn't exist:
            # self.save_config() 
            pass

    def save_config(self):
        """
        Saves the current settings to the configuration JSON file.
        """
        try:
            # Ensure the directory for the config file exists
            config_dir = os.path.dirname(self.config_filepath)
            if config_dir and not os.path.exists(config_dir):
                os.makedirs(config_dir, exist_ok=True)
                
            with open(self.config_filepath, 'w', encoding='utf-8') as f:
                json.dump(self.settings, f, indent=4, ensure_ascii=False)
            # print(f"Config saved to {self.config_filepath}") # For debugging
        except Exception as e:
            print(f"Warning: Error saving configuration to {self.config_filepath}: {e}")

    def get_setting(self, key):
        """
        Retrieves a setting value.
        Args:
            key (str): The key of the setting to retrieve.
        Returns:
            The value of the setting, or the default value if the key is not found.
        """
        return self.settings.get(key, self.defaults.get(key))

    def update_setting(self, key, value):
        """
        Updates a setting value if the key is valid (present in defaults).
        Args:
            key (str): The key of the setting to update.
            value: The new value for the setting.
        Returns:
            bool: True if update was successful, False if key was invalid.
        """
        if key in self.defaults:
            self.settings[key] = value
            # Could add self.save_config() here if settings should be persisted immediately
            return True
        # print(f"Warning: Attempted to update invalid setting key: {key}") # For debugging
        return False

    def reset_to_defaults(self):
        """
        Resets all settings to their default values and saves the configuration.
        """
        self.settings = self.defaults.copy()
        self.save_config()
        # print("Configuration reset to defaults.") # For debugging

# Example Usage (for testing the ConfigManager directly):
if __name__ == '__main__':
    # Create a config manager instance (will load from or create 'config.json')
    config = ConfigManager(config_filepath='test_config.json')
    
    print("Initial settings (after load or defaults):")
    for k, v in config.settings.items():
        print(f"  {k}: {v}")

    # Get a specific setting
    print(f"\nOllama API URL: {config.get_setting('ollama_api_url')}")
    print(f"A non-existent key: {config.get_setting('non_existent_key')}") # Should be None

    # Update a setting
    config.update_setting('ollama_temperature', 0.9)
    print(f"Updated Ollama Temperature: {config.get_setting('ollama_temperature')}")
    
    # Attempt to update an invalid setting
    config.update_setting('invalid_setting_key', 'some_value')

    # Save the current configuration (includes the updated temperature)
    config.save_config()
    print("\nSettings saved to test_config.json")

    # Create a new instance to see if it loads the saved settings
    config2 = ConfigManager(config_filepath='test_config.json')
    print(f"\nOllama Temperature from new instance: {config2.get_setting('ollama_temperature')}") # Should be 0.9

    # Reset to defaults
    config.reset_to_defaults()
    print("\nSettings after reset_to_defaults:")
    for k, v in config.settings.items():
        print(f"  {k}: {v}")
    
    # Clean up test file
    if os.path.exists('test_config.json'):
        os.remove('test_config.json')
        print("\nCleaned up test_config.json")
