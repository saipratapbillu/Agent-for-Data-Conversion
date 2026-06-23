# Quick Start - Enhanced Agent Features

## 🆕 What's New?

✅ **Default Path** - Set default folder for files  
✅ **Service Account Auth** - Authenticate with credentials  
✅ **Auto Naming** - CSV keeps Excel filename  
✅ **Auto Save** - Auto-save to default path  

---

## ⚡ 5-Minute Setup

### Step 1: Verify Installation

```powershell
cd "c:\Users\VemanasaiPratapBillu\OneDrive - Unify Group\Desktop\HSE CSV Files\Agent"
& ".\.venv\Scripts\python.exe" -c "import pandas; print('✅ Ready!')"
```

### Step 2: Configure Default Path (Optional)

Open `config.py` and change:
```python
DEFAULT_OUTPUT_PATH = r"C:\Your\Preferred\Path"
```

### Step 3: Launch Application

**With default settings:**
```powershell
& ".\.venv\Scripts\python.exe" main.py
```

**With custom path:**
```powershell
& ".\.venv\Scripts\python.exe" main.py "C:\MyData\Output"
```

**With service account:**
```powershell
& ".\.venv\Scripts\python.exe" main.py "C:\MyData\Output" "C:\keys\service_account.json"
```

---

## 🎯 Common Tasks

### Convert Single File

1. Launch: `& ".\.venv\Scripts\python.exe" main.py`
2. Click "📂 Select File"
3. Choose Excel file
4. Check "✓ Auto-save CSV..."
5. Click "⚙️ Convert to CSV"
6. ✅ Done! File saved with same name

### Convert with Default Path

```powershell
$path = "C:\Data\Reports"
& ".\.venv\Scripts\python.exe" main.py $path
```

Files automatically save to `C:\Data\Reports` with original names.

### Command-Line Conversion

```powershell
& ".\.venv\Scripts\python.exe" cli.py "C:\Data\sales.xlsx"
# Output: C:\Data\sales.csv
```

---

## 🔐 Service Account Setup (2 minutes)

### Quick Setup

1. Create JSON file with credentials
2. Save as `credentials.json`
3. Launch with: `& ".\.venv\Scripts\python.exe" main.py "." "credentials.json"`

### Example credentials.json

```json
{
  "type": "service_account",
  "project_id": "my-project",
  "private_key_id": "key123",
  "client_email": "agent@my-project.iam.gserviceaccount.com"
}
```

---

## 📊 Auto-Save Features

**What is Auto-Save?**
- Automatically saves CSV files
- Uses same filename as Excel file
- Saves to default path
- Prevents duplicate names

**Enable/Disable:**
- GUI: Toggle checkbox "✓ Auto-save CSV..."
- Config: Edit `AUTO_SAVE_ENABLED = True/False` in `config.py`

**Example:**
```
Input:  C:\Data\report_2024.xlsx
Output: C:\Output\report_2024.csv (auto-saved)
```

---

## 📝 Configuration

Edit `config.py` to customize:

```python
# Output directory
DEFAULT_OUTPUT_PATH = r"C:\Data\Output"

# Service account (optional)
SERVICE_ACCOUNT_KEY_PATH = r"C:\keys\service_account.json"

# Auto-save behavior
AUTO_SAVE_ENABLED = True

# Create conversion log
CREATE_LOG_FILE = True

# Max file size (MB)
MAX_FILE_SIZE_MB = 500  # Set to None for unlimited
```

---

## 🚀 Batch Processing

Create file `batch_convert.ps1`:

```powershell
$agent = "c:\...\Agent"
$output = "C:\Output"
$files = Get-ChildItem "C:\Input" -Filter "*.xlsx"

cd $agent

foreach ($file in $files) {
    "Converting $($file.Name)..."
    & ".\.venv\Scripts\python.exe" cli.py $file.FullName "" $output
}

"✅ All done!"
```

Run: `.\batch_convert.ps1`

---

## ✅ Verification

Check that everything works:

```powershell
# Verify configuration
python config.py

# Should show:
# ✅ Configuration is valid!
```

---

## 📋 Command Cheat Sheet

```powershell
# GUI with defaults
& ".\.venv\Scripts\python.exe" main.py

# GUI with custom path
& ".\.venv\Scripts\python.exe" main.py "C:\Data"

# GUI with authentication
& ".\.venv\Scripts\python.exe" main.py "C:\Data" "creds.json"

# CLI: Convert one file
& ".\.venv\Scripts\python.exe" cli.py "file.xlsx"

# CLI: Specific sheet
& ".\.venv\Scripts\python.exe" cli.py "file.xlsx" "Sheet2"

# CLI: Custom output
& ".\.venv\Scripts\python.exe" cli.py "file.xlsx" "Sheet1" "out.csv"

# CLI: With service account
& ".\.venv\Scripts\python.exe" cli.py "file.xlsx" "" "" --service-account "key.json"
```

---

## 🔍 Logging

Conversion log automatically created: `conversion_log.jsonl`

View recent conversions:
```powershell
Get-Content "C:\Output\conversion_log.jsonl" -Tail 5 | ConvertFrom-Json
```

---

## ❓ FAQ

**Q: Where do files save?**  
A: Default path shown at top of GUI, or use command-line argument

**Q: Does filename change?**  
A: No! Excel name → CSV name (e.g., `report.xlsx` → `report.csv`)

**Q: Can I disable auto-save?**  
A: Yes, uncheck "✓ Auto-save CSV..." in GUI

**Q: How to use service account?**  
A: Pass credentials file: `main.py "path" "credentials.json"`

**Q: What if CSV exists?**  
A: Auto-renamed: `file_1.csv`, `file_2.csv`, etc.

---

**Ready to go?** Start with: `& ".\.venv\Scripts\python.exe" main.py`

For detailed help, see `ENHANCED_USAGE.md` and `README.md`
