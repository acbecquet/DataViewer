# Create this as 'fix_encoding_issue.py' in your sensory_data_collection folder
import os

def fix_specific_encoding_issue(filepath):
    """Fix the specific encoding issue at position 7894."""
    print(f"Fixing encoding issue in: {filepath}")
    
    # Create backup first
    backup_path = filepath + ".backup"
    with open(filepath, 'rb') as src:
        with open(backup_path, 'wb') as dst:
            dst.write(src.read())
    print(f"Backup created: {backup_path}")
    
    try:
        # Read as latin1 (which works according to diagnostic)
        with open(filepath, 'r', encoding='latin1') as f:
            content = f.read()
        
        print(f"File length: {len(content)} characters")
        
        # Check around position 7894 to see what's there
        start_pos = max(0, 7894 - 50)
        end_pos = min(len(content), 7894 + 50)
        context = content[start_pos:end_pos]
        
        print(f"Context around position 7894:")
        print(repr(context))
        
        # Common problematic Windows-1252 characters and their replacements
        replacements = {
            '\x91': "'",    # Left single quotation mark
            '\x92': "'",    # Right single quotation mark  
            '\x93': '"',    # Left double quotation mark
            '\x94': '"',    # Right double quotation mark
            '\x95': '•',    # Bullet point
            '\x96': '–',    # En dash
            '\x97': '—',    # Em dash
            '\x85': '…',    # Horizontal ellipsis
        }
        
        # Apply replacements
        original_content = content
        for bad_char, replacement in replacements.items():
            if bad_char in content:
                count = content.count(bad_char)
                content = content.replace(bad_char, replacement)
                print(f"Replaced {count} instances of {repr(bad_char)} with {repr(replacement)}")
        
        # Write the fixed content as UTF-8
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(content)
        
        print("File fixed and saved as UTF-8!")
        
        # Verify the fix worked
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                test_content = f.read()
            print("✓ Verification successful - file can now be read as UTF-8!")
            return True
        except UnicodeDecodeError as e:
            print(f"✗ Verification failed: {e}")
            # Restore backup
            with open(backup_path, 'rb') as src:
                with open(filepath, 'wb') as dst:
                    dst.write(src.read())
            return False
            
    except Exception as e:
        print(f"Error during fix: {e}")
        # Restore backup
        try:
            with open(backup_path, 'rb') as src:
                with open(filepath, 'wb') as dst:
                    dst.write(src.read())
        except:
            pass
        return False

# Run the fix
if __name__ == "__main__":
    result = fix_specific_encoding_issue("file_io.py")
    print(f"\nFix successful: {result}")
    
    if result:
        print("\nYou can now try importing the module again!")
        print("The backup file has been saved as file_io.py.backup")
    else:
        print("\nFix failed - original file has been restored from backup")