# Excel to CSV Converter Agent

A powerful Python application that converts Excel files (.xlsx, .xls) to CSV format with a graphical interface and command-line tools.

## ✨ Features

✅ **Easy File Selection** - Browse and select Excel files from your local drive  
✅ **Sheet Selection** - Choose specific sheets from multi-sheet Excel files  
✅ **Batch Conversion** - Convert any Excel format to CSV  
✅ **Auto-naming** - Automatically handles duplicate filenames  
✅ **GUI Interface** - No command line needed  
✅ **CLI Tools** - For automation and scripting  
✅ **Default Path Configuration** - Set preferred folder for file selection and output  
✅ **Service Account Authentication** - Authenticate with cloud services  
✅ **Auto-Save Feature** - Automatically save files to default path with same name  
✅ **Conversion Logging** - Track all conversions with timestamps  
✅ **File Size Support** - Handles large Excel files efficiently

## Installation

### Prerequisites
- Python 3.7 or higher

### Setup

1. Navigate to the Agent folder:
```bash
cd "c:\Users\YourUsername\OneDrive - Unify Group\Desktop\HSE CSV Files\Agent"
```

2. Install required dependencies:
```bash
pip install -r requirements.txt
```

Or use the virtual environment:
```bash
& ".\.venv\Scripts\python.exe" -m pip install -r requirements.txt
```

## Usage

### GUI Application (Easiest)

**Basic Usage:**
```bash
python main.py
```

**With Default Path:**
```bash
python main.py "C:\Your\Output\Path"
```

**With Service Account Authentication:**
```bash
python main.py "C:\Your\Output\Path" "C:\path\to\service_account.json"
```

### Command-Line Interface

**Convert file with default naming:**
```bash
python cli.py "data.xlsx"
```

**Convert specific sheet:**
```bash
python cli.py "data.xlsx" "Sheet2"
```

**Custom output filename:**
```bash
python cli.py "data.xlsx" "Sheet1" "output.csv"
```

**With service account:**
```bash
python cli.py "data.xlsx" "Sheet1" "output.csv" --service-account "service_account.json"
```

## Configuration

Edit `config.py` to set permanent defaults:

```python
# Default output directory
DEFAULT_OUTPUT_PATH = r"C:\Users\YourName\Desktop"

# Service account key (optional)
SERVICE_ACCOUNT_KEY_PATH = r"C:\keys\service_account.json"

# Auto-save settings
AUTO_SAVE_ENABLED = True
CREATE_LOG_FILE = True

# Advanced settings
MAX_FILE_SIZE_MB = 500  # None for unlimited
DEBUG_MODE = False
```

## Features in Detail

### 🎯 Default Path Configuration
Set a default directory that will be used for:
- File selection starting point
- Default location for CSV output
- Conversion log location

### 🔐 Service Account Authentication
Authenticate with cloud services using JSON credentials:
1. Create service account in cloud provider
2. Download JSON key file
3. Pass to agent: `main.py "path" "key.json"`
4. Agent validates and uses for operations

### 📝 Automatic Naming
CSV files automatically get the same filename as Excel files:
- `sales_report.xlsx` → `sales_report.csv`
- `Q1_data.xls` → `Q1_data.csv`

### 💾 Auto-Save Feature
When enabled:
- Files automatically save to default path
- Uses same filename as Excel file
- Prevents manual filename entry
- Toggle in GUI or config file

### 📊 Conversion Logging
Automatic log creation tracks:
- Timestamp of conversion
- Source and output filenames
- Number of rows and columns converted
- Service account used (if applicable)

Log file: `conversion_log.jsonl`

## Supported Formats

### Input
- `.xlsx` (Excel 2007+)
- `.xls` (Excel 97-2003)

### Output
- `.csv` (Comma-Separated Values)

## Technical Details

- **pandas**: Handles Excel file reading and CSV writing
- **openpyxl**: Supports modern Excel file formats
- **tkinter**: Built-in Python GUI toolkit (no additional installation needed)

## Troubleshooting

### "No module named 'pandas'"
Run: `pip install -r requirements.txt`

### "Sheet name not found"
Check available sheets in the status bar when you select a file

### "Permission denied" error
Make sure the output CSV file isn't already open in another program

## Example Workflow

```
1. Open command prompt in Agent folder
2. Run: python main.py
3. Click "Select Excel File"
4. Browse to your Excel file (e.g., "c:\Data\report.xlsx")
5. Leave "Sheet Name" empty for default sheet
6. Click "Convert to CSV"
7. Find your CSV file in the same location
```

## Support

If you encounter any issues:
1. Check that your Excel file is not corrupted
2. Ensure Python and all dependencies are properly installed
3. Make sure you have write permissions to the output folder

---

**Created**: June 2026
**Version**: 1.0
