#!/usr/bin/env python3
"""Find and replace special characters in all project files"""

import os
import sys
import re

def scan_file_for_special_chars(filepath):
    """Scan a file for non-ASCII characters and return details"""
    issues = []
    
    try:
        # Try UTF-8 first
        with open(filepath, 'r', encoding='utf-8') as f:
            content = f.read()
    except UnicodeDecodeError:
        try:
            # Fallback to latin-1 which can read any byte sequence
            with open(filepath, 'r', encoding='latin-1') as f:
                content = f.read()
        except Exception as e:
            return [{"line": 0, "char": "FILE READ ERROR", "error": str(e)}]
    except Exception as e:
        return [{"line": 0, "char": "FILE ACCESS ERROR", "error": str(e)}]
    
    lines = content.split('\n')
    
    for line_num, line in enumerate(lines, 1):
        for char_pos, char in enumerate(line):
            if ord(char) > 127:  # Non-ASCII character
                issues.append({
                    "line": line_num,
                    "pos": char_pos,
                    "char": char,
                    "code": ord(char),
                    "context": line.strip()[:50]
                })
    
    return issues

def fix_file_special_chars(filepath, replacement="?"):
    """Replace special characters in a file"""
    try:
        # Read file
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                content = f.read()
        except UnicodeDecodeError:
            with open(filepath, 'r', encoding='latin-1') as f:
                content = f.read()
        
        # Replace non-ASCII characters
        fixed_content = ""
        for char in content:
            if ord(char) > 127:
                fixed_content += replacement
            else:
                fixed_content += char
        
        # Write back
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(fixed_content)
        
        return True
        
    except Exception as e:
        print("ERROR fixing file " + filepath + ": " + str(e))
        return False

def scan_project_directory(directory="."):
    """Scan all Python files in project for special characters"""
    
    print("SCANNING PROJECT FOR SPECIAL CHARACTERS")
    print("=" * 50)
    print("Directory: " + os.path.abspath(directory))
    
    # File extensions to check
    extensions = ['.py', '.txt', '.md', '.rst', '.cfg', '.ini']
    
    all_issues = {}
    total_files = 0
    files_with_issues = 0
    
    for root, dirs, files in os.walk(directory):
        # Skip certain directories
        skip_dirs = ['__pycache__', '.git', 'build', 'dist', 'clean_env', '.env', 'venv']
        dirs[:] = [d for d in dirs if d not in skip_dirs]
        
        for file in files:
            # Check if file has relevant extension
            if any(file.endswith(ext) for ext in extensions):
                filepath = os.path.join(root, file)
                total_files += 1
                
                print("Checking: " + filepath)
                
                issues = scan_file_for_special_chars(filepath)
                
                if issues:
                    all_issues[filepath] = issues
                    files_with_issues += 1
                    
                    print("  FOUND " + str(len(issues)) + " special characters")
                    for issue in issues[:3]:  # Show first 3
                        if "error" in issue:
                            print("    ERROR: " + issue["error"])
                        else:
                            print("    Line " + str(issue["line"]) + ": '" + issue["char"] + "' (code " + str(issue["code"]) + ")")
                            print("      Context: " + issue["context"])
                    
                    if len(issues) > 3:
                        print("    ... and " + str(len(issues) - 3) + " more")
                else:
                    print("  OK")
    
    print("\n" + "=" * 50)
    print("SCAN RESULTS")
    print("=" * 50)
    print("Files scanned: " + str(total_files))
    print("Files with special characters: " + str(files_with_issues))
    
    return all_issues

def main():
    print("SPECIAL CHARACTER FINDER AND FIXER")
    print("This will scan your project for non-ASCII characters")
    print()
    
    # Scan first
    issues = scan_project_directory()
    
    if not issues:
        print("\nGOOD NEWS: No special characters found!")
        return
    
    print("\nFILES WITH SPECIAL CHARACTERS:")
    for filepath, file_issues in issues.items():
        print("  " + filepath + " (" + str(len(file_issues)) + " issues)")
    
    # Ask what to do
    print("\nOPTIONS:")
    print("1. Replace all special characters with '?' and save")
    print("2. Show detailed report only")
    print("3. Exit without changes")
    
    choice = input("\nChoice (1-3): ").strip()
    
    if choice == "1":
        print("\nREPLACING SPECIAL CHARACTERS...")
        backup_dir = "backup_before_char_fix"
        
        # Create backup directory
        if not os.path.exists(backup_dir):
            os.makedirs(backup_dir)
            print("Created backup directory: " + backup_dir)
        
        fixed_files = 0
        for filepath in issues.keys():
            # Create backup
            filename = os.path.basename(filepath)
            backup_path = os.path.join(backup_dir, filename)
            
            try:
                import shutil
                shutil.copy2(filepath, backup_path)
                print("Backed up: " + filepath + " -> " + backup_path)
            except Exception as e:
                print("WARNING: Could not backup " + filepath + ": " + str(e))
            
            # Fix the file
            if fix_file_special_chars(filepath, "?"):
                print("Fixed: " + filepath)
                fixed_files += 1
            else:
                print("FAILED to fix: " + filepath)
        
        print("\nCOMPLETED!")
        print("Fixed files: " + str(fixed_files))
        print("Backups saved in: " + backup_dir)
        
    elif choice == "2":
        print("\nDETAILED REPORT:")
        for filepath, file_issues in issues.items():
            print("\nFile: " + filepath)
            for issue in file_issues:
                if "error" in issue:
                    print("  ERROR: " + issue["error"])
                else:
                    print("  Line " + str(issue["line"]) + ", Pos " + str(issue["pos"]) + ": '" + issue["char"] + "' (Unicode " + str(issue["code"]) + ")")
                    print("    Context: " + issue["context"])
    
    print("\nDone.")

if __name__ == "__main__":
    main()