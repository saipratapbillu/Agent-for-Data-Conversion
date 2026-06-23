# Batch Excel to CSV Converter - PowerShell Script
# Converts all Excel files from network path to local CSV

param(
    [string]$ExcelPath = "\\10.210.32.5\IGS\SINYAR\HSES\PowerBIReports",
    [string]$CsvPath = "C:\Users\VemanasaiPratapBillu\OneDrive - Unify Group\Desktop\HSE CSV Files\CSV",
    [switch]$KeepExisting = $false
)

# Change to the script directory to ensure relative paths resolve correctly
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
Set-Location -Path $scriptDir

# Clear screen
Clear-Host

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Batch Excel to CSV Converter" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Display configuration
Write-Host "📂 Excel Source: $ExcelPath" -ForegroundColor White
Write-Host "💾 CSV Output:   $CsvPath" -ForegroundColor White
Write-Host "🗑️  Delete Existing: $(if ($KeepExisting) { 'No' } else { 'Yes' })" -ForegroundColor White
Write-Host ""

# Validate paths
if (-not (Test-Path $ExcelPath)) {
    Write-Host "❌ Error: Excel source path not found!" -ForegroundColor Red
    Write-Host "   Path: $ExcelPath" -ForegroundColor Red
    exit 1
}

# Create CSV output directory
if (-not (Test-Path $CsvPath)) {
    Write-Host "📁 Creating CSV output directory..." -ForegroundColor Yellow
    New-Item -ItemType Directory -Path $CsvPath -Force | Out-Null
}

# Find all Excel files
$excelFiles = @()
$excelFiles += Get-ChildItem -Path $ExcelPath -Filter "*.xlsx" -Recurse -ErrorAction SilentlyContinue
$excelFiles += Get-ChildItem -Path $ExcelPath -Filter "*.xls" -Recurse -ErrorAction SilentlyContinue

if ($excelFiles.Count -eq 0) {
    Write-Host "❌ No Excel files found!" -ForegroundColor Red
    exit 1
}

Write-Host "📊 Found $($excelFiles.Count) Excel file(s)" -ForegroundColor Green
Write-Host ""

# Delete existing CSVs
if (-not $KeepExisting) {
    $existingCsvs = Get-ChildItem -Path $CsvPath -Filter "*.csv" -ErrorAction SilentlyContinue
    
    if ($existingCsvs.Count -gt 0) {
        Write-Host "🗑️  Deleting $($existingCsvs.Count) existing CSV file(s)..." -ForegroundColor Yellow
        foreach ($csv in $existingCsvs) {
            Remove-Item $csv.FullName -Force -ErrorAction SilentlyContinue
            Write-Host "   ✓ Deleted: $($csv.Name)" -ForegroundColor Gray
        }
        Write-Host ""
    }
}

# Conversion progress
Write-Host "=" * 70 -ForegroundColor Cyan
Write-Host "CONVERSION PROGRESS:" -ForegroundColor Cyan
Write-Host "=" * 70 -ForegroundColor Cyan
Write-Host ""

$converted = 0
$failed = 0
$errors = @()
$timestamp = Get-Date

foreach ($excelFile in $excelFiles) {
    $fileNum = $excelFiles.IndexOf($excelFile) + 1
    Write-Host "[$fileNum/$($excelFiles.Count)] Converting: $($excelFile.Name)" -ForegroundColor White
    
    try {
        # Generate output filename (same name as input)
        $csvFileName = [System.IO.Path]::GetFileNameWithoutExtension($excelFile.Name) + ".csv"
        $csvPath_Full = Join-Path $CsvPath $csvFileName
        
        # Use Python to convert with the workspace virtual environment
        $pythonPath = Join-Path $scriptDir ".venv\Scripts\python.exe"
        $cliPath = Join-Path $scriptDir "cli.py"
        
        # Run conversion
        & $pythonPath $cliPath $excelFile.FullName | Out-Null
        
        # Check if output was created (cli.py saves with same name by default)
        if (Test-Path $csvPath_Full) {
            Write-Host "    ✓ Success: $csvFileName" -ForegroundColor Green
            
            # Get file info
            $csvFile = Get-Item $csvPath_Full
            Write-Host "      Size: $([Math]::Round($csvFile.Length / 1KB, 2)) KB" -ForegroundColor Gray
            
            $converted++
        } else {
            Write-Host "    ❌ Output file not found" -ForegroundColor Red
            $failed++
            $errors += "Output file not created: $csvFileName"
        }
    }
    catch {
        Write-Host "    ❌ Error: $($_.Exception.Message)" -ForegroundColor Red
        $failed++
        $errors += "Exception converting $($excelFile.Name): $($_.Exception.Message)"
    }
}

# Summary
Write-Host ""
Write-Host "=" * 70 -ForegroundColor Cyan
Write-Host "CONVERSION SUMMARY" -ForegroundColor Cyan
Write-Host "=" * 70 -ForegroundColor Cyan

Write-Host "Total Files:     $($excelFiles.Count)" -ForegroundColor White
Write-Host "Converted:       $converted ✓" -ForegroundColor Green
Write-Host "Failed:          $failed ❌" -ForegroundColor $(if ($failed -gt 0) { "Red" } else { "Green" })

if ($errors.Count -gt 0) {
    Write-Host ""
    Write-Host "Errors:" -ForegroundColor Red
    foreach ($error in $errors) {
        Write-Host "  - $error" -ForegroundColor Red
    }
}

# Calculate success rate
$successRate = if ($excelFiles.Count -gt 0) { 
    [Math]::Round(($converted / $excelFiles.Count) * 100, 1) 
} else { 
    0 
}

Write-Host ""
Write-Host "Success Rate:    $successRate%" -ForegroundColor $(if ($successRate -eq 100) { "Green" } else { "Yellow" })
Write-Host "Timestamp:       $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Gray

Write-Host ""
Write-Host "=" * 70 -ForegroundColor Cyan

# Log results
$logFile = Join-Path $CsvPath "batch_conversion_log.txt"
$logContent = @"
Batch Conversion Log - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
================================================
Excel Source: $ExcelPath
CSV Output:   $CsvPath

Summary:
--------
Total Files:    $($excelFiles.Count)
Converted:      $converted
Failed:         $failed
Success Rate:   $successRate%

File List:
----------
"@

foreach ($file in $excelFiles) {
    $logContent += "`n  - $($file.Name)"
}

if ($errors.Count -gt 0) {
    $logContent += "`n`nErrors:`n-------`n"
    foreach ($error in $errors) {
        $logContent += "`n  - $error"
    }
}

try {
    $logContent | Out-File -FilePath $logFile -Encoding UTF8 -Force
    Write-Host ""
    Write-Host "📋 Log saved: $logFile" -ForegroundColor Gray
}
catch {
    Write-Host ""
    Write-Host "⚠️  Could not save log file" -ForegroundColor Yellow
}

Write-Host ""

# Exit with appropriate code
if ($failed -eq 0) {
    Write-Host "✅ Batch conversion completed successfully!" -ForegroundColor Green
    exit 0
} else {
    Write-Host "❌ Batch conversion completed with errors!" -ForegroundColor Red
    exit 1
}
