import json
import os
from datetime import datetime

class CustomLogger:
    _instance = None

    def __new__(cls, *args, **kwargs):
        if not cls._instance:
            cls._instance = super(CustomLogger, cls).__new__(cls, *args, **kwargs)
            cls._instance._initialized = False
        return cls._instance

    def __init__(self):
        if getattr(self, "_initialized", False):
            return
            
        self.log_dir = "logs"
        self.json_log_dir = os.path.join(self.log_dir, "json")
        os.makedirs(self.json_log_dir, exist_ok=True)
        
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        self.log_file = os.path.join(self.json_log_dir, f"execution_log_{timestamp}.json")
        self.log_list = [] # In-memory list to dump (or we can append line by line)
        
        # Initialize file as an empty list or just ready for appending objects?
        # User requested "saved to json", usually means a valid JSON file.
        # We will append line-by-line JSON objects (JSONL style) for safety against crashes,
        # but wraps them in a list if the user wants strictly valid JSON array (harder to do streaming).
        # Let's do JSON Lines (NDJSON) as it's cleaner for logs, but to satisfy "save to json", 
        # I will just append to a list and write the file on every log (inefficient but safe for small batches) 
        # OR simply append line-by-line. 
        # Let's stick to appending line by line for now, easy to parse.
        
        print(f"[LOGGER] Logs serão salvos em: {self.log_file}")
        self._initialized = True

    def log(self, level, message, details=None):
        entry = {
            "timestamp": datetime.now().isoformat(),
            "level": level.upper(),
            "message": message,
            "details": details or {}
        }

        # 1. Print to Terminal
        print(f"[{level.upper()}] {message}")
        if details:
            # Only print simple details to terminal to avoid clutter, or skipped
            pass 

        # 2. Save to JSON File (Append mode)
        # We will write it as a list of objects. To do this efficiently without rewriting the valid JSON array every time:
        # A common trick is to write Newline Delimited JSON, but if the user wants strict JSON:
        # We can read, append, write. (Slow for many logs)
        # Let's do Newline Delimited JSON (.jsonl) but name it .json for simplicity if user opens in text editor.
        try:
            with open(self.log_file, "a", encoding="utf-8") as f:
                f.write(json.dumps(entry, ensure_ascii=False) + "\n")
        except Exception as e:
            print(f"[LOGGER ERROR] Failed to write log: {e}")

    def info(self, message, details=None):
        self.log("INFO", message, details)

    def warn(self, message, details=None):
        self.log("WARN", message, details)

    def error(self, message, details=None):
        self.log("ERROR", message, details)

# Singleton instance
logger = CustomLogger()
