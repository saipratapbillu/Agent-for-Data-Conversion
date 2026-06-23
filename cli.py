#!/usr/bin/env python3
"""
Command-line version of Excel to CSV Converter
Enhanced with service account support and auto-naming
Usage: python cli.py <excel_file> [sheet_name] [output_file]
"""

import sys
import pandas as pd
from pathlib import Path
from datetime import datetime
import json


def load_service_account(key_path):
    """Load and validate service account credentials
    
    Args:
        key_path (str): Path to service account JSON key file
        
    Returns:
        dict: Service account credentials if valid, None otherwise
    """
    try:
        with open(key_path, 'r') as f:
            creds = json.load(f)
        
        # Validate required fields
        required_fields = ['type', 'project_id']
        if all(field in creds for field in required_fields):
            print(f"✅ Service account loaded: {creds.get('project_id', 'Unknown')}")
            return creds
        else:
            print("❌ Invalid service account format")
            return None
    except FileNotFoundError:
        print(f"❌ Service account file not found: {key_path}")
        return None
    except json.JSONDecodeError:
        print(f"❌ Invalid JSON in service account file")
        return None


def log_conversion(output_file, input_file, rows, columns, service_account=None):
    """Log conversion details
    
    Args:
        output_file (str): Path to output CSV file
        input_file (str): Path to input Excel file
        rows (int): Number of rows converted
        columns (int): Number of columns converted
        service_account (dict, optional): Service account info for logging
    """
    log_entry = {
        "timestamp": datetime.now().isoformat(),
        "input_file": str(input_file),
        "output_file": str(output_file),
        "rows": rows,
        "columns": columns,
        "service_account": service_account.get('project_id') if service_account else None
    }
    
    # Write to log
    log_file = Path(output_file).parent / "conversion_log.jsonl"
    try:
        with open(log_file, "a", encoding='utf-8') as f:
            f.write(json.dumps(log_entry) + "\n")
        print(f"📋 Logged to: {log_file}")
    except Exception as e:
        print(f"⚠️  Warning: Could not write to log: {str(e)}")


def convert_excel_to_csv(excel_file, sheet_name=None, output_file=None, service_account_key=None):
    """Convert an Excel file to CSV format
    
    Args:
        excel_file (str): Path to the Excel file
        sheet_name (str, optional): Name of the sheet to convert. Defaults to first sheet.
        output_file (str, optional): Path for output CSV. Defaults to same name as input.
        service_account_key (str, optional): Path to service account JSON key
    
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        # Load service account if provided
        service_account = None
        if service_account_key:
            service_account = load_service_account(service_account_key)
        
        excel_path = Path(excel_file)
        
        # Validate input file
        if not excel_path.exists():
            print(f"❌ Error: File not found - {excel_file}")
            return False
        
        if excel_path.suffix.lower() not in ['.xlsx', '.xls']:
            print(f"❌ Error: Not an Excel file - {excel_file}")
            return False
        
        print(f"📂 Reading: {excel_file}")
        
        # Read Excel file
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
        
        # Determine output file path (same name as input by default)
        if output_file is None:
            csv_path = excel_path.with_suffix('.csv')
        else:
            csv_path = Path(output_file)
        
        # Handle duplicate filenames
        counter = 1
        original_csv = csv_path
        while csv_path.exists():
            csv_path = original_csv.with_stem(f"{original_csv.stem}_{counter}")
            counter += 1
        
        # Create output directory if needed
        csv_path.parent.mkdir(parents=True, exist_ok=True)
        
        print(f"💾 Writing: {csv_path}")
        print(f"📊 Data: {len(df)} rows × {len(df.columns)} columns")
        
        # Write to CSV
        df.to_csv(csv_path, index=False, encoding='utf-8')
        
        # Log the conversion
        log_conversion(csv_path, excel_path, len(df), len(df.columns), service_account)
        
        print(f"✅ Successfully converted to CSV!")
        return True
        
    except Exception as e:
        print(f"❌ Error: {str(e)}")
        return False


def main():
    """Main entry point for CLI"""
    if len(sys.argv) < 2:
        print("Excel to CSV Converter - Command Line Tool (Enhanced)")
        print("\nUsage:")
        print("  python cli.py <excel_file> [sheet_name] [output_file]")
        print("\nWith Service Account Authentication:")
        print("  python cli.py <excel_file> [sheet_name] [output_file] --service-account <key_file>")
        print("\nExamples:")
        print("  python cli.py data.xlsx")
        print("  python cli.py data.xlsx Sheet2")
        print("  python cli.py data.xlsx Sheet1 output.csv")
        print("  python cli.py data.xlsx Sheet1 output.csv --service-account credentials.json")
        sys.exit(1)
    
    excel_file = sys.argv[1]
    sheet_name = sys.argv[2] if len(sys.argv) > 2 else None
    output_file = sys.argv[3] if len(sys.argv) > 3 else None
    service_account_key = None
    
    # Check for service account flag
    if "--service-account" in sys.argv:
        idx = sys.argv.index("--service-account")
        if idx + 1 < len(sys.argv):
            service_account_key = sys.argv[idx + 1]
    
    success = convert_excel_to_csv(excel_file, sheet_name, output_file, service_account_key)
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
