#!/usr/bin/env python3
"""
Configuration file for Excel to CSV Converter Agent
Edit this file to set default values
"""

import os
from pathlib import Path

# ==========================================
# DEFAULT SETTINGS
# ==========================================

# ==========================================
# PATHS CONFIGURATION
# ==========================================

# Excel Source Path (Network or Local)
EXCEL_SOURCE_PATH = r"\\10.210.32.5\IGS\SINYAR\HSES\PowerBIReports"

# Default directory for CSV output
DEFAULT_OUTPUT_PATH = r"C:\Users\VemanasaiPratapBillu\OneDrive - Unify Group\Desktop\HSE CSV Files\CSV"

# Service Account Configuration
# Path to your service account JSON key file
# Leave empty if not using service account authentication
SERVICE_ACCOUNT_KEY_PATH = None  # Example: "C:\\keys\\service_account.json"

# Auto-save settings
AUTO_SAVE_ENABLED = True
CREATE_LOG_FILE = True

# ==========================================
# ADVANCED SETTINGS
# ==========================================

# Enable debugging mode
DEBUG_MODE = False

# Maximum file size to process (in MB)
# Set to None for unlimited
MAX_FILE_SIZE_MB = None

# Supported file formats
SUPPORTED_FORMATS = ['.xlsx', '.xls']

# CSV encoding (usually 'utf-8')
CSV_ENCODING = 'utf-8'

# Include index in CSV output
INCLUDE_INDEX = False

# ==========================================
# SERVICE ACCOUNT SETTINGS
# ==========================================

# Service account scopes for cloud storage (if using Google Drive, Azure, etc.)
SERVICE_ACCOUNT_SCOPES = [
    'https://www.googleapis.com/auth/drive',
]

# ==========================================
# HELPER FUNCTIONS
# ==========================================

def get_config():
    """Get current configuration as dictionary"""
    return {
        'excel_source_path': EXCEL_SOURCE_PATH,
        'default_output_path': DEFAULT_OUTPUT_PATH,
        'service_account_key': SERVICE_ACCOUNT_KEY_PATH,
        'auto_save': AUTO_SAVE_ENABLED,
        'create_log': CREATE_LOG_FILE,
        'debug': DEBUG_MODE,
        'max_file_size_mb': MAX_FILE_SIZE_MB,
        'supported_formats': SUPPORTED_FORMATS,
        'encoding': CSV_ENCODING,
        'include_index': INCLUDE_INDEX,
    }


def validate_config():
    """Validate configuration settings"""
    errors = []
    
    # Check if default path exists
    if not Path(DEFAULT_OUTPUT_PATH).exists():
        errors.append(f"Default output path does not exist: {DEFAULT_OUTPUT_PATH}")
    
    # Check if service account key exists
    if SERVICE_ACCOUNT_KEY_PATH and not Path(SERVICE_ACCOUNT_KEY_PATH).exists():
        errors.append(f"Service account key file not found: {SERVICE_ACCOUNT_KEY_PATH}")
    
    return errors


if __name__ == "__main__":
    # Display current configuration
    print("=" * 50)
    print("Excel to CSV Converter - Configuration")
    print("=" * 50)
    
    config = get_config()
    for key, value in config.items():
        print(f"{key:.<40} {value}")
    
    print("\n" + "=" * 50)
    
    # Validate configuration
    errors = validate_config()
    if errors:
        print("⚠️  Configuration Issues:")
        for error in errors:
            print(f"  - {error}")
    else:
        print("✅ Configuration is valid!")
