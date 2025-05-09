import json
import os

class ConfigManager:
    def __init__(self):
        self.config_dir = os.path.join(os.path.expanduser('~'), '.barcode_scanner')
        self.config_file = os.path.join(self.config_dir, 'config.json')
        self.default_config = {
            'input_port1': '',
            'input_port2': '',
            'output_port': '',
            'printer_port': '',  # Add printer port
            'input_baud1': '9600',
            'input_baud2': '9600',
            'output_baud': '9600',
            'printer_baud': '9600',  # Add printer baud rate
            'switch_number': '1'
        }
    
    def load_config(self):
        try:
            if not os.path.exists(self.config_dir):
                os.makedirs(self.config_dir)
                
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r') as f:
                    config = json.load(f)
                    # Ensure all required keys exist by updating with defaults
                    return {**self.default_config, **config}
            return self.default_config.copy()
        except Exception:
            return self.default_config.copy()
    
    def save_config(self, config):
        try:
            if not os.path.exists(self.config_dir):
                os.makedirs(self.config_dir)
                
            with open(self.config_file, 'w') as f:
                json.dump(config, f, indent=4)
            return True
        except Exception:
            return False