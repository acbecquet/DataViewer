#!/usr/bin/env python3
"""
license_validator.py
Advanced license key validation system
Include debugging output for troubleshooting
"""
import hashlib
import hmac
import base64
import json
import time
from datetime import datetime, timedelta
from typing import Optional, Dict, Any

class LicenseValidator:
    """Advanced license validation with debugging"""

    def __init__(self, secret_key: str = "your-secret-key-change-this"):
        self.secret_key = secret_key.encode('utf-8')
        self.debug_mode = True

    def debug_print(self, message: str):
        """Print debug information"""
        if self.debug_mode:
            print(f"LICENSE_DEBUG: {message}")

    def generate_license_key(self,
                           user_id: str,
                           expiry_days: int = 365,
                           license_type: str = "FULL") -> str:
        """Generate a license key for a user"""
        self.debug_print(f"Generating license for user: {user_id}")

        # Create license data
        expiry_date = datetime.now() + timedelta(days=expiry_days)
        license_data = {
            'user_id': user_id,
            'license_type': license_type,
            'expiry': expiry_date.isoformat(),
            'issued': datetime.now().isoformat(),
            'version': '3.0.0'
        }

        self.debug_print(f"License data: {license_data}")

        # Encode and sign
        encoded_data = base64.b64encode(
            json.dumps(license_data).encode('utf-8')
        ).decode('utf-8')

        signature = hmac.new(
            self.secret_key,
            encoded_data.encode('utf-8'),
            hashlib.sha256
        ).hexdigest()[:8]  # First 8 chars of signature

        # Format as XXXX-XXXX-XXXX-XXXX
        key_parts = [
            license_type[:4].upper(),
            signature[:4].upper(),
            signature[4:8].upper(),
            encoded_data[:4].upper()
        ]

        license_key = '-'.join(key_parts)
        self.debug_print(f"Generated license key: {license_key}")

        return license_key

    def validate_license_key(self, license_key: str) -> Dict[str, Any]:
        """Validate a license key and return details"""
        self.debug_print(f"Validating license key: {license_key}")

        try:
            # Parse key format
            if not license_key or len(license_key.replace('-', '')) != 16:
                self.debug_print("Invalid key format - wrong length")
                return {'valid': False, 'reason': 'Invalid format'}

            parts = license_key.split('-')
            if len(parts) != 4:
                self.debug_print("Invalid key format - wrong parts")
                return {'valid': False, 'reason': 'Invalid format'}

            license_type = parts[0]
            signature_part = parts[1] + parts[2]

            self.debug_print(f"Parsed - Type: {license_type}, Signature: {signature_part}")

            # Demo keys (for testing)
            demo_keys = {
                'DEMO-1234-5678-ABCD': {
                    'user_id': 'demo_user',
                    'license_type': 'DEMO',
                    'expiry': (datetime.now() + timedelta(days=30)).isoformat(),
                    'features': ['basic_analysis', 'limited_reports']
                },
                'FULL-9876-5432-WXYZ': {
                    'user_id': 'full_user',
                    'license_type': 'FULL',
                    'expiry': (datetime.now() + timedelta(days=365)).isoformat(),
                    'features': ['all_features', 'unlimited_reports', 'advanced_analysis']
                }
            }

            if license_key in demo_keys:
                self.debug_print(f"Found demo key: {license_key}")
                license_data = demo_keys[license_key]

                # Check expiry
                expiry = datetime.fromisoformat(license_data['expiry'])
                if datetime.now() > expiry:
                    self.debug_print("Demo key expired")
                    return {'valid': False, 'reason': 'License expired'}

                self.debug_print("Demo key validated successfully")
                return {
                    'valid': True,
                    'license_type': license_data['license_type'],
                    'user_id': license_data['user_id'],
                    'expiry': license_data['expiry'],
                    'features': license_data['features']
                }

            # For production, implement proper signature validation here
            self.debug_print("Key not found in demo keys, would validate signature in production")

            return {'valid': False, 'reason': 'Invalid license key'}

        except Exception as e:
            self.debug_print(f"Validation error: {e}")
            return {'valid': False, 'reason': f'Validation error: {str(e)}'}

    def get_demo_keys(self) -> Dict[str, str]:
        """Get demo keys for testing"""
        return {
            'DEMO-1234-5678-ABCD': 'Demo License (30 days)',
            'FULL-9876-5432-WXYZ': 'Full License (1 year)',
            'DEV-AAAA-BBBB-CCCC': 'Developer License (unlimited)'
        }

# Test the validator
if __name__ == "__main__":
    validator = LicenseValidator()

    print("Testing License Validator")
    print("=" * 40)

    # Test demo keys
    demo_keys = validator.get_demo_keys()
    for key, description in demo_keys.items():
        print(f"\nTesting: {key} ({description})")
        result = validator.validate_license_key(key)
        print(f"Result: {result}")

    # Test invalid key
    print(f"\nTesting invalid key:")
    result = validator.validate_license_key("INVALID-KEY-1234")
    print(f"Result: {result}")
