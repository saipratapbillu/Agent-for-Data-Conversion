#!/usr/bin/env python3
"""
Batch Excel to CSV Converter
Converts all Excel files from source directory, deletes existing CSVs, and updates them
"""

import sys
import json
from pathlib import Path
from datetime import datetime
import pandas as pd
from config import EXCEL_SOURCE_PATH, DEFAULT_OUTPUT_PATH


class BatchConverter:
    def __init__(self, excel_path, csv_path, delete_existing=True):
        """
        Initialize batch converter
        
        Args:
            excel_path (str): Source directory with Excel files
            csv_path (str): Destination directory for CSV files
            delete_existing (bool): Delete existing CSV files before conversion
        """
        self.excel_path = Path(excel_path)
        self.csv_path = Path(csv_path)
        self.delete_existing = delete_existing
        self.results = {
            'total': 0,
            'converted': 0,
            'failed': 0,
            'deleted': 0,
            'errors': []
        }
        
    def validate_paths(self):
        """Validate source and destination paths"""
        print("=" * 70)
        print("BATCH EXCEL TO CSV CONVERTER")
        print("=" * 70)
        print(f"\n📂 Excel Source: {self.excel_path}")
        print(f"💾 CSV Output:   {self.csv_path}")
        
        # Check if source exists
        if not self.excel_path.exists():
            print(f"❌ Error: Source path does not exist!")
            return False
        
        # Create output directory if it doesn't exist
        try:
            self.csv_path.mkdir(parents=True, exist_ok=True)
            print(f"✓ Output directory ready")
        except Exception as e:
            print(f"❌ Error creating output directory: {str(e)}")
            return False
        
        return True
    
    def find_excel_files(self):
        """Find all Excel files in source directory"""
        excel_files = []
        
        # Search for xlsx and xls files; ignore Excel temp files
        for ext in ['*.xlsx', '*.xls']:
            excel_files.extend(
                [f for f in self.excel_path.glob(f"**/{ext}") if not f.name.startswith("~$")]
            )
        
        return sorted(excel_files)
    
    def delete_existing_csvs(self):
        """Delete existing CSV files in output directory"""
        if not self.delete_existing:
            return
        
        csv_files = list(self.csv_path.glob("*.csv"))
        
        if not csv_files:
            print(f"\n✓ No existing CSV files to delete")
            return
        
        print(f"\n🗑️  Deleting {len(csv_files)} existing CSV files...")
        
        for csv_file in csv_files:
            try:
                csv_file.unlink()
                self.results['deleted'] += 1
                print(f"   ✓ Deleted: {csv_file.name}")
            except Exception as e:
                print(f"   ❌ Failed to delete {csv_file.name}: {str(e)}")
        
        print(f"✓ Deleted {self.results['deleted']} CSV files")
    
    def convert_files(self):
        """Convert all Excel files to CSV"""
        excel_files = self.find_excel_files()
        
        if not excel_files:
            print(f"\n❌ No Excel files found in: {self.excel_path}")
            return False
        
        print(f"\n📊 Found {len(excel_files)} Excel file(s) to convert")
        print("\n" + "-" * 70)
        print("CONVERSION PROGRESS:")
        print("-" * 70)
        
        self.results['total'] = len(excel_files)
        
        for idx, excel_file in enumerate(excel_files, 1):
            print(f"\n[{idx}/{len(excel_files)}] Converting: {excel_file.name}")
            
            try:
                # Read Excel file
                df = pd.read_excel(excel_file)
                
                # Generate output filename (same name as input)
                csv_file = self.csv_path / f"{excel_file.stem}.csv"
                
                # Write to CSV
                df.to_csv(csv_file, index=False, encoding='utf-8')
                
                print(f"    ✓ Success: {csv_file.name}")
                print(f"      Rows: {len(df)} | Columns: {len(df.columns)}")
                
                self.results['converted'] += 1
                
            except Exception as e:
                error_msg = f"Failed to convert {excel_file.name}: {str(e)}"
                print(f"    ❌ {error_msg}")
                self.results['failed'] += 1
                self.results['errors'].append(error_msg)
        
        return True
    
    def log_results(self):
        """Log conversion results"""
        log_file = self.csv_path / "batch_conversion_log.json"
        
        log_entry = {
            "timestamp": datetime.now().isoformat(),
            "excel_source": str(self.excel_path),
            "csv_output": str(self.csv_path),
            "total_files": self.results['total'],
            "converted": self.results['converted'],
            "failed": self.results['failed'],
            "deleted_csvs": self.results['deleted'],
            "errors": self.results['errors']
        }
        
        try:
            with open(log_file, 'w', encoding='utf-8') as f:
                json.dump(log_entry, f, indent=2, ensure_ascii=False)
            print(f"\n📋 Log saved: {log_file.name}")
        except Exception as e:
            print(f"\n⚠️  Could not save log: {str(e)}")
    
    def print_summary(self):
        """Print conversion summary"""
        print("\n" + "=" * 70)
        print("CONVERSION SUMMARY")
        print("=" * 70)
        print(f"Total Files:      {self.results['total']}")
        print(f"Converted:        {self.results['converted']} ✓")
        print(f"Failed:           {self.results['failed']} ❌")
        print(f"CSV Files Deleted: {self.results['deleted']} 🗑️")
        
        if self.results['errors']:
            print(f"\nErrors encountered:")
            for error in self.results['errors']:
                print(f"  - {error}")
        
        success_rate = (self.results['converted'] / self.results['total'] * 100) if self.results['total'] > 0 else 0
        print(f"\nSuccess Rate:     {success_rate:.1f}%")
        print("=" * 70 + "\n")
        
        return self.results['failed'] == 0
    
    def run(self):
        """Execute batch conversion"""
        # Validate paths
        if not self.validate_paths():
            return False
        
        # Delete existing CSVs
        self.delete_existing_csvs()
        
        # Convert files
        self.convert_files()
        
        # Log results
        self.log_results()
        
        # Print summary
        success = self.print_summary()
        
        return success


def main():
    """Main entry point"""
    excel_path = EXCEL_SOURCE_PATH
    csv_path = DEFAULT_OUTPUT_PATH
    delete_existing = True

    if len(sys.argv) >= 2:
        excel_path = sys.argv[1]
    if len(sys.argv) >= 3:
        csv_path = sys.argv[2]
    if "--keep-existing" in sys.argv:
        delete_existing = False

    if not excel_path or not csv_path:
        print("Batch Excel to CSV Converter")
        print("\nUsage:")
        print("  python batch_convert.py <excel_source> <csv_output> [--keep-existing]")
        print("\nIf no arguments are provided, defaults from config.py are used:")
        print(f"  source = {EXCEL_SOURCE_PATH}")
        print(f"  destination = {DEFAULT_OUTPUT_PATH}")
        print("\nExamples:")
        print("  python batch_convert.py")
        print("  python batch_convert.py \\\\10.210.32.5\\IGS\\SINYAR\\HSES\\PowerBIReports")
        print("  python batch_convert.py \\\\10.210.32.5\\IGS\\SINYAR\\HSES\\PowerBIReports C:\\CSV\\Output")
        print("  python batch_convert.py C:\\Excel\\Files C:\\CSV\\Output --keep-existing")
        print("\nOptions:")
        print("  --keep-existing : Keep existing CSV files (don't delete)")
        sys.exit(1)

    converter = BatchConverter(excel_path, csv_path, delete_existing)
    success = converter.run()

    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
