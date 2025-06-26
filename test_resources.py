#!/usr/bin/env python3
"""Test script to verify resource loading works correctly."""

import os
from utils import get_resource_path

def test_resources():
    """Test that resources can be found."""
    print("Testing resource loading...")
    
    # Test template file
    template_path = get_resource_path("resources/Standardized Test Template - LATEST VERSION - 2025 Jan.xlsx")
    print(f"Template path: {template_path}")
    print(f"Template exists: {os.path.exists(template_path)}")
    
    # List all files in resources directory
    resources_dir = get_resource_path("resources")
    if os.path.exists(resources_dir):
        print(f"\nFiles in resources directory:")
        for file in os.listdir(resources_dir):
            print(f"  - {file}")
    else:
        print(f"Resources directory not found: {resources_dir}")

if __name__ == "__main__":
    test_resources()