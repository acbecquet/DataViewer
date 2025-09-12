#!/usr/bin/env python3
"""Remove trailing whitespace from Python file"""
import sys

def remove_trailing_whitespace(filepath):
    """Remove trailing whitespace and fix final newline"""
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            lines = f.readlines()
    except FileNotFoundError:
        print(f"Error: File '{filepath}' not found.")
        return False
    except Exception as e:
        print(f"Error reading file '{filepath}': {e}")
        return False
    
    # Remove trailing whitespace from each line
    cleaned_lines = [line.rstrip() + '\n' for line in lines]
    
    # Ensure file ends with newline
    if cleaned_lines and not cleaned_lines[-1].endswith('\n'):
        cleaned_lines[-1] += '\n'
    
    try:
        with open(filepath, 'w', encoding='utf-8') as f:
            f.writelines(cleaned_lines)
    except Exception as e:
        print(f"Error writing to file '{filepath}': {e}")
        return False
    
    print(f"Cleaned trailing whitespace in {filepath}")
    return True

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python clean_whitespace.py <filepath>")
        print("Example: python clean_whitespace.py main_gui.py")
        sys.exit(1)
    
    filepath = sys.argv[1]
    success = remove_trailing_whitespace(filepath)
    sys.exit(0 if success else 1)