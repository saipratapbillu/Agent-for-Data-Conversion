# Batch Conversion Guide - Network Path Setup

## 🎯 Configuration Summary

**Excel Source Path:** `\\10.210.32.5\IGS\SINYAR\HSES\PowerBIReports`  
**CSV Output Path:** `C:\Users\VemanasaiPratapBillu\OneDrive - Unify Group\Desktop\HSE CSV Files\CSV`

**Features:**
- ✅ Converts ALL Excel files from network path
- ✅ Deletes existing CSVs before conversion (keeps data fresh)
- ✅ Saves with same filename as Excel file
- ✅ Creates conversion logs
- ✅ Auto-update capability

---

## 🚀 Quick Start

### Option 1: Batch File (Easiest for Windows Users)

Simply double-click:
```
batch_run.bat
```

This will:
1. Verify Python is installed
2. Scan all Excel files in network path
3. Delete existing CSV files
4. Convert all Excel files to CSV
5. Save to CSV folder with same names
6. Display completion status

---

### Option 2: PowerShell Script (Recommended)

Open PowerShell and run:

```powershell
cd "c:\Users\VemanasaiPratapBillu\OneDrive - Unify Group\Desktop\HSE CSV Files\Agent"
.\batch_convert.ps1
```

**With options:**
```powershell
# Keep existing CSV files (don't delete)
.\batch_convert.ps1 -KeepExisting

# Custom paths
.\batch_convert.ps1 -ExcelPath "\\server\share\files" -CsvPath "C:\Output"

# Keep existing + custom paths
.\batch_convert.ps1 -ExcelPath "\\server\share" -CsvPath "C:\Output" -KeepExisting
```

---

### Option 3: Python CLI (Advanced)

```powershell
cd "c:\Users\VemanasaiPratapBillu\OneDrive - Unify Group\Desktop\HSE CSV Files\Agent"

# Run with configured paths
& ".\.venv\Scripts\python.exe" batch_convert.py `
  "\\10.210.32.5\IGS\SINYAR\HSES\PowerBIReports" `
  "C:\Users\VemanasaiPratapBillu\OneDrive - Unify Group\Desktop\HSE CSV Files\CSV"

# Keep existing CSV files
& ".\.venv\Scripts\python.exe" batch_convert.py `
  "\\10.210.32.5\IGS\SINYAR\HSES\PowerBIReports" `
  "C:\Users\VemanasaiPratapBillu\OneDrive - Unify Group\Desktop\HSE CSV Files\CSV" `
  --keep-existing
```

---

## 📋 What Happens During Batch Conversion

### Step 1: Validation
```
✓ Verify Excel source path exists
✓ Create CSV output directory (if needed)
✓ Find all Excel files (.xlsx, .xls)
```

### Step 2: Cleanup
```
🗑️  Delete existing CSV files in output directory
   - This ensures data is always fresh
   - No duplicate files
```

### Step 3: Conversion
```
[1/N] Converting: file1.xlsx
    ✓ Success: file1.csv
    Rows: 1000 | Columns: 15

[2/N] Converting: file2.xlsx
    ✓ Success: file2.csv
    Rows: 500 | Columns: 8
```

### Step 4: Summary
```
Total Files:     2
Converted:       2 ✓
Failed:          0 ❌
Success Rate:    100%
```

---

## 🗂️ File Organization

After batch conversion:

```
Excel Source (\\10.210.32.5\IGS\SINYAR\HSES\PowerBIReports)
├── Report1.xlsx
├── Report2.xlsx
├── Report3.xls
└── Subfolder/
    └── Report4.xlsx

CSV Output (C:\...\Desktop\HSE CSV Files\CSV)
├── Report1.csv          (newly converted)
├── Report2.csv          (newly converted)
├── Report3.csv          (newly converted)
├── Report4.csv          (newly converted)
├── batch_conversion_log.json
└── batch_conversion_log.txt
```

---

## 📊 Conversion Logs

### JSON Log (Machine Readable)
**File:** `batch_conversion_log.json`

```json
{
  "timestamp": "2024-06-10T14:30:00",
  "excel_source": "\\10.210.32.5\IGS\SINYAR\HSES\PowerBIReports",
  "csv_output": "C:\Users\VemanasaiPratapBillu\OneDrive - Unify Group\Desktop\HSE CSV Files\CSV",
  "total_files": 4,
  "converted": 4,
  "failed": 0,
  "deleted_csvs": 0,
  "errors": []
}
```

### Text Log (Human Readable)
**File:** `batch_conversion_log.txt`

Shows:
- Source and output paths
- File list
- Success/failure count
- Any errors encountered

---

## 🔄 Scheduled Automation

### Windows Task Scheduler Setup

1. **Open Task Scheduler:**
   - Press `Win + R`
   - Type: `taskschd.msc`
   - Press Enter

2. **Create Basic Task:**
   - Name: "Daily Excel to CSV Conversion"
   - Description: "Batch convert all Excel files from network path"

3. **Set Trigger:**
   - Choose: "Daily"
   - Time: 8:00 AM (or preferred time)
   - Repeat every: 1 day

4. **Set Action:**
   - Action: "Start a program"
   - Program: `powershell.exe`
   - Add arguments: 
   ```
   -NoProfile -WindowStyle Hidden -File "C:\Path\To\batch_convert.ps1"
   ```

5. **Configure Run Settings:**
   - User account: Your account
   - Run with highest privileges: ✓
   - Run whether user is logged in: ✓

6. **Test:**
   - Right-click task → Run
   - Check CSV folder for new files

---

## 🔍 Monitoring Batch Conversions

### Check Latest Conversion
```powershell
# View last conversion log
Get-Content "C:\Users\VemanasaiPratapBillu\OneDrive - Unify Group\Desktop\HSE CSV Files\CSV\batch_conversion_log.txt"
```

### Monitor in Real-Time
```powershell
# Watch CSV folder for new files
Get-ChildItem "C:\Users\VemanasaiPratapBillu\OneDrive - Unify Group\Desktop\HSE CSV Files\CSV" -Filter "*.csv" | 
  Select-Object Name, LastWriteTime, @{Name="Size KB";Expression={[Math]::Round($_.Length/1KB,2)}}
```

### Count Converted Files
```powershell
(Get-ChildItem "C:\Users\VemanasaiPratapBillu\OneDrive - Unify Group\Desktop\HSE CSV Files\CSV" -Filter "*.csv").Count
```

---

## ⚙️ Advanced Options

### Keep Existing CSV Files

By default, old CSV files are deleted. To keep them:

**Batch File:** Edit `batch_run.bat`
```batch
python batch_convert.py "%EXCEL_SOURCE%" "%CSV_OUTPUT%" --keep-existing
```

**PowerShell:**
```powershell
.\batch_convert.ps1 -KeepExisting
```

### Custom Paths

**PowerShell with custom paths:**
```powershell
.\batch_convert.ps1 `
  -ExcelPath "\\10.210.32.5\IGS\SINYAR\HSES\PowerBIReports" `
  -CsvPath "C:\MyCustomOutput" `
  -KeepExisting
```

**Python with custom paths:**
```powershell
& ".\.venv\Scripts\python.exe" batch_convert.py `
  "\\your\excel\path" `
  "C:\your\csv\path" `
  --keep-existing
```

---

## 🛠️ Troubleshooting

### "Network path not found"
```powershell
# Verify network path is accessible
Test-Path "\\10.210.32.5\IGS\SINYAR\HSES\PowerBIReports"

# Try mapping network drive
net use Z: "\\10.210.32.5\IGS\SINYAR\HSES\PowerBIReports" /persistent:yes
```

### "Permission denied"
```powershell
# Run PowerShell as Administrator
# Or use network credentials
net use "\\10.210.32.5\IGS" /user:DOMAIN\USERNAME password
```

### "No Excel files found"
```powershell
# Check if files exist in source
Get-ChildItem "\\10.210.32.5\IGS\SINYAR\HSES\PowerBIReports" -Filter "*.xlsx" -Recurse
```

### "CSV output directory not writable"
```powershell
# Check permissions
icacls "C:\Users\VemanasaiPratapBillu\OneDrive - Unify Group\Desktop\HSE CSV Files\CSV"

# Grant write permissions if needed
icacls "C:\Users\VemanasaiPratapBillu\OneDrive - Unify Group\Desktop\HSE CSV Files\CSV" /grant:r "%USERNAME%:M"
```

---

## 📋 Command Reference

| Task | Command |
|------|---------|
| Quick batch (double-click) | `batch_run.bat` |
| PowerShell batch | `.\batch_convert.ps1` |
| PowerShell (keep existing) | `.\batch_convert.ps1 -KeepExisting` |
| Python batch | `python batch_convert.py <excel_path> <csv_path>` |
| Python (keep existing) | `python batch_convert.py <excel_path> <csv_path> --keep-existing` |
| View config | `python config.py` |

---

## ✅ Verification

### Before First Run
```powershell
# 1. Verify paths exist
Test-Path "\\10.210.32.5\IGS\SINYAR\HSES\PowerBIReports"
Test-Path "C:\Users\VemanasaiPratapBillu\OneDrive - Unify Group\Desktop\HSE CSV Files\CSV"

# 2. Check Excel files available
Get-ChildItem "\\10.210.32.5\IGS\SINYAR\HSES\PowerBIReports" -Filter "*.xlsx" -Recurse | Measure-Object

# 3. Verify permissions
icacls "C:\Users\VemanasaiPratapBillu\OneDrive - Unify Group\Desktop\HSE CSV Files\CSV"
```

### After Batch Run
```powershell
# 1. Check CSV files created
Get-ChildItem "C:\Users\VemanasaiPratapBillu\OneDrive - Unify Group\Desktop\HSE CSV Files\CSV" -Filter "*.csv"

# 2. View conversion log
Get-Content "C:\Users\VemanasaiPratapBillu\OneDrive - Unify Group\Desktop\HSE CSV Files\CSV\batch_conversion_log.txt"

# 3. Verify file counts match
$excelCount = (Get-ChildItem "\\10.210.32.5\IGS\SINYAR\HSES\PowerBIReports" -Filter "*.xlsx" -Recurse).Count
$csvCount = (Get-ChildItem "C:\Users\VemanasaiPratapBillu\OneDrive - Unify Group\Desktop\HSE CSV Files\CSV" -Filter "*.csv").Count
Write-Host "Excel files: $excelCount | CSV files: $csvCount"
```

---

## 🎯 Next Steps

1. **Test batch conversion:**
   ```
   Double-click batch_run.bat
   ```

2. **Verify CSV files created:**
   ```
   Open C:\Users\VemanasaiPratapBillu\OneDrive - Unify Group\Desktop\HSE CSV Files\CSV
   ```

3. **Set up scheduled task:**
   - Open Task Scheduler
   - Create task as described above
   - Test before scheduling

4. **Monitor conversions:**
   - Check conversion logs
   - Verify file counts
   - Review for errors

---

**Configuration Status:** ✅ Ready to Use  
**Paths Verified:** ✅ Configured  
**Auto-Delete:** ✅ Enabled  
**Auto-Update:** ✅ Enabled
