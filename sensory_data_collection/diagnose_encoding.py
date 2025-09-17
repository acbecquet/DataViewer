# Create this as 'diagnose_encoding.py' in your sensory_data_collection folder
import os

def diagnose_file_encoding(filepath):
    """Diagnose encoding issues in a specific file."""
    print(f"Diagnosing: {filepath}")
    
    # Read raw bytes
    try:
        with open(filepath, 'rb') as f:
            raw_bytes = f.read(100)  # First 100 bytes
        
        print(f"First 20 raw bytes: {raw_bytes[:20]}")
        print(f"Hex representation: {raw_bytes[:20].hex()}")
        
        # Check for BOM
        if raw_bytes.startswith(b'\xef\xbb\xbf'):
            print("ERROR: UTF-8 BOM detected!")
            return "BOM"
        elif raw_bytes.startswith(b'\xff\xfe'):
            print("ERROR: UTF-16 LE BOM detected!")
            return "BOM"
        elif raw_bytes.startswith(b'\xfe\xff'):
            print("ERROR: UTF-16 BE BOM detected!")
            return "BOM"
        
        # Try different encodings
        encodings = ['utf-8', 'latin1', 'cp1252', 'ascii']
        for encoding in encodings:
            try:
                with open(filepath, 'r', encoding=encoding) as f:
                    content = f.read(50)
                print(f"SUCCESS with {encoding}: {repr(content[:50])}")
                return encoding
            except UnicodeDecodeError as e:
                print(f"FAILED with {encoding}: {e}")
        
        return "UNKNOWN"
        
    except Exception as e:
        print(f"ERROR reading file: {e}")
        return "ERROR"

# Run the diagnostic
filepath = "file_io.py"  # or full path if needed
result = diagnose_file_encoding(filepath)
print(f"\nResult: {result}")