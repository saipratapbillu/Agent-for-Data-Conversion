#!/usr/bin/env python3
"""
Verify batch conversion setup
Checks paths, permissions, and configuration before running batch conversion
"""

import os
import sys
from pathlib import Path


def verify_setup():
    """Verify all required paths and permissions"""
    
    print("=" * 70)
    print("BATCH CONVERSION SETUP VERIFICATION")
    print("=" * 70)
    print()
    
    # Configuration
    excel_path = r"\\10.210.32.5\IGS\SINYAR\HSES\PowerBIReports"
    csv_path = r"C:\Users\VemanasaiPratapBillu\OneDrive - Unify Group\Desktop\HSE CSV Files\CSV"
    
    results = {
        'excel_path_exists': False,
        'csv_path_exists': False,
        'csv_path_writable': False,
        'excel_files_found': 0,
        'python_available': False,
        'pandas_available': False,
        'openpyxl_available': False,
    }
    
    # 1. Check Python
    print("1. Checking Python Environment...")
    print("-" * 70)
    try:
        import pandas
        import openpyxl
        results['python_available'] = True
        results['pandas_available'] = True
        results['openpyxl_available'] = True
        print("   ✓ Python 3 available")
        print(f"   ✓ pandas {pandas.__version__} installed")
        print(f"   ✓ openpyxl {openpyxl.__version__} installed")
    except ImportError as e:
        print(f"   ❌ Missing module: {str(e)}")
        print("   Run: pip install -r requirements.txt")
    
    print()
    
    # 2. Check Excel source path
    print("2. Checking Excel Source Path...")
    print("-" * 70)
    print(f"   Path: {excel_path}")
    
    excel_path_obj = Path(excel_path)
    if excel_path_obj.exists():
        results['excel_path_exists'] = True
        print("   ✓ Path exists and is accessible")
        
        # Find Excel files
        try:
            xlsx_files = list(excel_path_obj.glob("**/*.xlsx"))
            xls_files = list(excel_path_obj.glob("**/*.xls"))
            total_files = len(xlsx_files) + len(xls_files)
            results['excel_files_found'] = total_files
            
            print(f"   ✓ Found {len(xlsx_files)} .xlsx files")
            print(f"   ✓ Found {len(xls_files)} .xls files")
            print(f"   ✓ Total: {total_files} Excel file(s)")
            
            if total_files > 0:
                print("\n   Excel Files Found:")
                for file in (xlsx_files + xls_files)[:5]:  # Show first 5
                    print(f"     - {file.name}")
                if total_files > 5:
                    print(f"     ... and {total_files - 5} more")
        except Exception as e:
            print(f"   ⚠️  Error reading files: {str(e)}")
    else:
        print("   ❌ Path does not exist or is not accessible!")
        print("   Check:")
        print("      - Network connectivity")
        print("      - Correct server address")
        print("      - User permissions")
    
    print()
    
    # 3. Check CSV output path
    print("3. Checking CSV Output Path...")
    print("-" * 70)
    print(f"   Path: {csv_path}")
    
    csv_path_obj = Path(csv_path)
    if csv_path_obj.exists():
        results['csv_path_exists'] = True
        print("   ✓ Path exists")
        
        # Check write permissions
        try:
            test_file = csv_path_obj / ".write_test"
            test_file.write_text("test")
            test_file.unlink()
            results['csv_path_writable'] = True
            print("   ✓ Directory is writable")
        except PermissionError:
            print("   ❌ Directory is NOT writable!")
            print("   Fix: Grant write permissions to your user account")
        except Exception as e:
            print(f"   ⚠️  Error testing write: {str(e)}")
        
        # List existing CSV files
        csv_files = list(csv_path_obj.glob("*.csv"))
        if csv_files:
            print(f"   ⚠️  Found {len(csv_files)} existing CSV file(s)")
            print("   Note: These will be deleted during batch conversion")
    else:
        print("   ⚠️  Path does not exist - will be created automatically")
    
    print()
    
    # 4. Summary
    print("4. Summary...")
    print("-" * 70)
    
    checks = [
        ("Excel source path accessible", results['excel_path_exists']),
        ("Excel files found", results['excel_files_found'] > 0),
        ("CSV output path exists/writable", results['csv_path_exists']),
        ("CSV output is writable", results['csv_path_writable']),
        ("Python environment ready", results['python_available']),
        ("pandas installed", results['pandas_available']),
        ("openpyxl installed", results['openpyxl_available']),
    ]
    
    all_ready = True
    for check_name, check_result in checks:
        status = "✓" if check_result else "❌"
        print(f"   {status} {check_name}")
        if not check_result:
            all_ready = False
    
    print()
    print("=" * 70)
    
    if all_ready and results['excel_files_found'] > 0:
        print("✅ Setup is READY! You can run batch conversion:")
        print()
        print("   Option 1: Double-click batch_run.bat")
        print("   Option 2: Run: .\\batch_convert.ps1")
        print("   Option 3: Run: python batch_convert.py <excel> <csv>")
        print()
        return 0
    else:
        print("❌ Setup is NOT READY - fix issues above")
        print()
        if not results['excel_path_exists']:
            print("   Fix Excel path access:")
            print("      - Verify network path is correct")
            print("      - Check network connectivity")
            print("      - Verify user credentials")
        if not results['csv_path_writable']:
            print("   Fix CSV output path:")
            print("      - Create directory if needed")
            print("      - Grant write permissions")
        if not (results['pandas_available'] and results['openpyxl_available']):
            print("   Install Python dependencies:")
            print("      - Run: pip install -r requirements.txt")
        print()
        return 1


def main():
    """Main entry point"""
    try:
        exit_code = verify_setup()
        sys.exit(exit_code)
    except Exception as e:
        print(f"\n❌ Verification failed: {str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    main()
