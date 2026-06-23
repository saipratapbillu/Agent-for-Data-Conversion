# Excel to CSV Converter Agent - Enhanced Usage Guide

## 🆕 New Features

### ✨ Feature 1: Default Path Configuration
Set a default directory for file selection and automatic CSV output.

### ✨ Feature 2: Service Account Authentication
Authenticate with cloud services using service account credentials.

### ✨ Feature 3: Automatic CSV Naming
CSV files automatically get the same name as the Excel file.

### ✨ Feature 4: Auto-Save Feature
Automatically save converted files to the default output path.

---

## 🚀 GUI Usage with New Features

### Launch with Default Path

```powershell
# Launch with Desktop as default path
& ".\.venv\Scripts\python.exe" main.py "C:\Users\YourName\Desktop"

# Launch with custom path
& ".\.venv\Scripts\python.exe" main.py "C:\Data\Exports"

# Launch with service account authentication
& ".\.venv\Scripts\python.exe" main.py "C:\Data\Exports" "C:\keys\service_account.json"
```

### GUI Interface Features

**Default Output Path Display**
- Shows at top of application
- All converted files save to this location
- Can be set via command-line argument

**Auto-Save Toggle**
- Checkbox: "✓ Auto-save CSV (same name as Excel file)"
- When enabled: Files save to default path with same name
- When disabled: Choose location for each file

**File Selection**
- Opens file browser at default path
- Automatically filters Excel files
- Shows available sheets in status bar

---

## 💻 Command-Line Usage (Enhanced)

### Basic Usage

```powershell
cd "c:\Users\YourName\OneDrive - Unify Group\Desktop\HSE CSV Files\Agent"

# Convert with default settings
& ".\.venv\Scripts\python.exe" cli.py "data.xlsx"

# Specify sheet name
& ".\.venv\Scripts\python.exe" cli.py "data.xlsx" "Sheet2"

# Custom output filename
& ".\.venv\Scripts\python.exe" cli.py "data.xlsx" "Sheet1" "report.csv"
```

### With Service Account Authentication

```powershell
# Authenticate and convert
& ".\.venv\Scripts\python.exe" cli.py "data.xlsx" "" "" --service-account "C:\keys\service_account.json"

# With sheet name and service account
& ".\.venv\Scripts\python.exe" cli.py "data.xlsx" "Sales" "output.csv" --service-account "credentials.json"
```

---

## 🔐 Service Account Setup

### Step 1: Create Service Account

**For Google Cloud:**
1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project
3. Go to "Service Accounts" section
4. Click "Create Service Account"
5. Download the JSON key file

**For Azure:**
1. Go to [Azure Portal](https://portal.azure.com/)
2. Navigate to "App registrations"
3. Create new registration
4. Generate client secret
5. Save credentials

### Step 2: Configure Agent

**Option A: Edit config.py**

```python
# In config.py
SERVICE_ACCOUNT_KEY_PATH = "C:\\keys\\service_account.json"
```

**Option B: Pass as Command-Line Argument**

```powershell
& ".\.venv\Scripts\python.exe" main.py "C:\Data" "C:\keys\service_account.json"
```

### Step 3: Verify Authentication

```powershell
# Test if service account is valid
python config.py
```

Expected output:
```
✅ Configuration is valid!
```

---

## 📋 Example Workflows

### Example 1: Daily Batch Processing

```powershell
# Set up default output path
$defaultPath = "C:\Data\CSV_Output"
$serviceAccount = "C:\keys\service_account.json"

# Process multiple files
& ".\.venv\Scripts\python.exe" main.py $defaultPath $serviceAccount
```

1. Application launches with default path set to `C:\Data\CSV_Output`
2. Select Excel files one by one
3. Auto-save converts them with original names
4. All CSVs saved to `C:\Data\CSV_Output`
5. Conversion log created automatically

---

### Example 2: Automated Script

Create file: `batch_convert.ps1`

```powershell
# Batch conversion script with service account
$agent_path = "c:\Users\YourName\OneDrive - Unify Group\Desktop\HSE CSV Files\Agent"
$output_path = "C:\Data\CSV_Output"
$service_account = "C:\keys\service_account.json"
$source_files = Get-ChildItem "C:\Data\Excel_Files" -Filter "*.xlsx"

cd $agent_path

foreach ($file in $source_files) {
    Write-Host "Converting: $($file.Name)"
    & ".\.venv\Scripts\python.exe" cli.py $file.FullName "" $output_path --service-account $service_account
    Write-Host "✅ Done"
}

Write-Host "All files converted!"
```

Run the script:
```powershell
.\batch_convert.ps1
```

---

### Example 3: Scheduled Task (Windows)

1. Create script: `convert_daily.ps1`
2. Open Task Scheduler
3. Create Basic Task: "Daily Excel to CSV Conversion"
4. Set trigger: Daily at 8:00 AM
5. Set action: Run PowerShell script
6. Configure with service account for unattended execution

---

## 🔍 Configuration File (config.py)

Edit `config.py` to customize behavior:

```python
# Default output directory
DEFAULT_OUTPUT_PATH = os.path.expanduser("~\Desktop")

# Service account JSON key path
SERVICE_ACCOUNT_KEY_PATH = "C:\\keys\\service_account.json"

# Enable auto-save by default
AUTO_SAVE_ENABLED = True

# Create log file of conversions
CREATE_LOG_FILE = True

# Maximum file size to process (MB)
MAX_FILE_SIZE_MB = 500  # None for unlimited
```

---

## 📊 Logging and Tracking

### Conversion Log

Automatically created as `conversion_log.jsonl` in output directory:

```json
{"timestamp": "2024-06-10T14:30:00", "input_file": "data.xlsx", "output_file": "data.csv", "rows": 1000, "columns": 15, "service_account": "project-123"}
```

### View Conversion History

```powershell
# Display last 10 conversions
Get-Content "C:\Data\CSV_Output\conversion_log.jsonl" -Tail 10
```

---

## 🔑 Authentication Methods

### Method 1: Configuration File (Permanent)

Edit `config.py`:
```python
SERVICE_ACCOUNT_KEY_PATH = "C:\keys\service_account.json"
```

### Method 2: Command-Line Argument (Temporary)

```powershell
& ".\.venv\Scripts\python.exe" main.py "C:\Data" "C:\keys\service_account.json"
```

### Method 3: Environment Variable (Advanced)

```powershell
$env:EXCEL_CSV_SERVICE_ACCOUNT = "C:\keys\service_account.json"
& ".\.venv\Scripts\python.exe" main.py "C:\Data"
```

---

## 📁 File Naming Convention

### Automatic Naming (Auto-Save Enabled)

| Input File | Output File |
|-----------|------------|
| `data.xlsx` | `data.csv` |
| `report_2024.xlsx` | `report_2024.csv` |
| `Q1_Sales.xls` | `Q1_Sales.csv` |

### Duplicate Handling

If `data.csv` already exists:
- First conversion: `data.csv`
- Second conversion: `data_1.csv`
- Third conversion: `data_2.csv`

---

## ✅ Verification Checklist

- [ ] Python 3.7+ installed
- [ ] Dependencies installed: `pip install -r requirements.txt`
- [ ] Default path exists and is writable
- [ ] Service account key (if using authentication)
- [ ] Service account JSON file is valid
- [ ] Sufficient disk space for output files

---

## 🛠️ Troubleshooting

### "Service account file not found"
```powershell
# Check file path is correct
Test-Path "C:\keys\service_account.json"
```

### "Permission denied" on output path
```powershell
# Check folder permissions
icacls "C:\Data\CSV_Output"

# Grant write permissions if needed
icacls "C:\Data\CSV_Output" /grant:r "%USERNAME%:M"
```

### "Auto-save failed"
- Ensure output directory exists
- Check write permissions
- Verify free disk space

---

## 📝 Command Reference

| Command | Purpose |
|---------|---------|
| `python main.py` | GUI with default Desktop path |
| `python main.py "path"` | GUI with custom path |
| `python main.py "path" "key.json"` | GUI with service account |
| `python cli.py file.xlsx` | CLI: Convert to same name |
| `python cli.py file.xlsx "Sheet2"` | CLI: Convert specific sheet |
| `python cli.py file.xlsx "" out.csv` | CLI: Custom output name |
| `python config.py` | Verify configuration |

---

## 🎯 Best Practices

1. **Always verify authentication** before processing critical files
2. **Test with sample files** before running batch operations
3. **Monitor conversion logs** for errors or issues
4. **Keep service account keys secure** - don't commit to version control
5. **Set reasonable file size limits** to prevent performance issues
6. **Schedule backups** of CSV output files

---

**Last Updated:** June 2026  
**Version:** 2.0 (Enhanced)
