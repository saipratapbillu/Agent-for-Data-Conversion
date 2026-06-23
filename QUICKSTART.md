# Quick Start Guide - Excel to CSV Converter Agent

## For Windows Users - Easiest Way

### Option 1: Automatic Setup (Recommended)

1. **Open Command Prompt** in the Agent folder
   - Right-click in the Agent folder
   - Select "Open in Terminal" or "Open Command Prompt here"

2. **Run the installer**:
   ```bash
   run.bat
   ```

3. The application will:
   - Check if Python is installed
   - Install required libraries automatically
   - Launch the GUI agent

### Option 2: Manual Setup

1. **Install Python**
   - Download from https://www.python.org/
   - ✅ Check "Add Python to PATH" during installation
   - Verify: Open Command Prompt and type `python --version`

2. **Install Dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Run the Agent**
   ```bash
   python main.py
   ```

---

## Using the GUI Agent

### Once the Application Launches:

**Step 1: Select Excel File**
- Click the green "Select Excel File" button
- Browse to your Excel file (.xlsx or .xls)
- The file path will appear in the window

**Step 2: Choose Sheet (Optional)**
- If your Excel file has multiple sheets, type the sheet name
- Leave empty to use the first sheet
- Available sheets are shown in the status bar

**Step 3: Convert**
- Click the blue "Convert to CSV" button
- A dialog will show the conversion result
- The CSV file appears in the same folder as your Excel file

---

## Using the Command-Line Version

If you prefer command line:

```bash
python cli.py your_file.xlsx
```

With specific sheet:
```bash
python cli.py your_file.xlsx "Sheet2"
```

With custom output name:
```bash
python cli.py your_file.xlsx "Sheet1" output_data.csv
```

---

## Troubleshooting

### ❌ "Python not found" or "python: command not found"
- Python is not installed or not in PATH
- Solution: Install Python again, checking "Add Python to PATH"

### ❌ "ModuleNotFoundError: No module named 'pandas'"
- Dependencies not installed
- Solution: Run `pip install -r requirements.txt`

### ❌ "File not found" error in GUI
- The Excel file was deleted or moved
- Solution: Select the file again using the GUI

### ❌ "Sheet name not found"
- You typed an incorrect sheet name
- Solution: Leave empty for first sheet or check correct name

### ❌ "Permission denied" when saving
- The CSV file is open in another program
- Solution: Close Excel/other programs using the CSV file

---

## What Gets Converted

- ✅ All data from selected sheet
- ✅ Preserves data types (numbers, dates, text)
- ✅ Column headers preserved
- ✅ Handles large files efficiently

---

## File Locations

```
Agent/
├── main.py              (GUI version)
├── cli.py               (Command-line version)
├── run.bat              (Windows launcher)
├── requirements.txt     (Python dependencies)
├── README.md            (Full documentation)
├── QUICKSTART.md        (This file)
└── [Your converted CSV files will appear here]
```

---

## Next Steps

1. ✅ Install dependencies: `pip install -r requirements.txt`
2. ✅ Run the agent: `python main.py`
3. ✅ Select your Excel file
4. ✅ Click Convert
5. ✅ Find your CSV file in the same location!

---

**Questions?** Check README.md for more details
