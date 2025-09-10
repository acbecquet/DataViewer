#!/usr/bin/env python3
"""Remove trailing whitespace from Python file"""

def remove_trailing_whitespace(filepath):
    """Remove trailing whitespace and fix final newline"""
    with open(filepath, 'r', encoding='utf-8') as f:
        lines = f.readlines()
    
    # Remove trailing whitespace from each line
    cleaned_lines = [line.rstrip() + '\n' for line in lines]
    
    # Ensure file ends with newline
    if cleaned_lines and not cleaned_lines[-1].endswith('\n'):
        cleaned_lines[-1] += '\n'
    
    with open(filepath, 'w', encoding='utf-8') as f:
        f.writelines(cleaned_lines)
    
    print(f"Cleaned trailing whitespace in {filepath}")

if __name__ == "__main__":
    remove_trailing_whitespace('main_gui.py')