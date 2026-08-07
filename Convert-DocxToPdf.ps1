# PowerShell script to convert DOCX to PDF using Word
# This script opens Word once and processes all files

$ErrorActionPreference = "Continue"

# Configuration
$inputFolder = "C:\Users\Public\Generated_Admit_Cards_Girls"
$wdFormatPDF = 17

# Get only real DOCX files (not temp files starting with ~$)
$docxFiles = Get-ChildItem -Path $inputFolder -Filter "*.docx" | Where-Object { -not $_.Name.StartsWith("~$") }

if ($docxFiles.Count -eq 0) {
    Write-Host "No DOCX files found in $inputFolder"
    exit
}

Write-Host "Found $($docxFiles.Count) DOCX files to convert"
Write-Host "Starting Word application..."

# Create Word application
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    Write-Host "Word started successfully`n"
} catch {
    Write-Host "ERROR: Could not start Microsoft Word"
    Write-Host "Make sure Microsoft Word is installed"
    exit
}

$successCount = 0
$failCount = 0
$current = 0

foreach ($file in $docxFiles) {
    $current++
    $docxPath = $file.FullName
    $pdfPath = $docxPath -replace '\.docx$', '.pdf'
    
    # Skip if PDF already exists
    if (Test-Path $pdfPath) {
        Write-Host "[$current/$($docxFiles.Count)] SKIP: $($file.Name) (PDF exists)"
        $successCount++
        continue
    }
    
    try {
        # Open document
        $doc = $word.Documents.Open($docxPath, $false, $true)
        
        # Save as PDF
        $doc.SaveAs([ref]$pdfPath, [ref]$wdFormatPDF)
        
        # Close document
        $doc.Close($false)
        
        Write-Host "[$current/$($docxFiles.Count)] SUCCESS: $($file.Name)"
        $successCount++
        
    } catch {
        Write-Host "[$current/$($docxFiles.Count)] FAILED: $($file.Name) - $($_.Exception.Message)"
        $failCount++
        
        # Try to close document if it's open
        try {
            if ($doc) {
                $doc.Close($false)
            }
        } catch {}
    }
    
    # Small delay
    Start-Sleep -Milliseconds 100
}

# Quit Word
Write-Host "`nClosing Word..."
try {
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
} catch {}

# Summary
Write-Host "`n========================================="
Write-Host "CONVERSION COMPLETE"
Write-Host "========================================="
Write-Host "Total files: $($docxFiles.Count)"
Write-Host "Successful: $successCount"
Write-Host "Failed: $failCount"
Write-Host "========================================="
Write-Host "`nPDF files saved to: $inputFolder"
