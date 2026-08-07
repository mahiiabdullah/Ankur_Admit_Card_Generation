# PDF Conversion Solutions for Admit Cards

## Problem
Microsoft Word's COM automation is failing with "Open.SaveAs" error. This is a common issue with automated Word conversion.

## ✅ RECOMMENDED SOLUTIONS

### Option 1: Use Microsoft Print to PDF (Built into Windows)

**Steps:**
1. Open File Explorer and navigate to: `C:\Users\Lenovo\Generated_Admit_Cards_Fixed`
2. Select ALL .docx files (Ctrl+A)
3. Right-click → Print
4. Select "Microsoft Print to PDF" as printer
5. Choose output folder and print all

This will convert all 167 files at once!

### Option 2: Use an Online Batch Converter (Recommended for Large Batches)

**Best free online tools:**
1. **ilovepdf.com/word_to_pdf**
   - Upload up to 25 files at once
   - Download as ZIP
   - Repeat 7 times for all 167 files

2. **smallpdf.com/word-to-pdf**
   - Batch upload
   - Free with some limits

3. **zamzar.com**
   - Email results
   - No software needed

### Option 3: Install LibreOffice (FREE) and Use Command Line

**Download:** https://www.libreoffice.org/download/download/

**After installation, run this command:**
```powershell
cd "C:\Program Files\LibreOffice\program"
.\soffice.exe --headless --convert-to pdf --outdir "C:\Users\Lenovo\Generated_Admit_Cards_Fixed" "C:\Users\Lenovo\Generated_Admit_Cards_Fixed\*.docx"
```

### Option 4: PowerShell Script Using Word (If Word automation works)

Save this as `convert_pdf.ps1` and run it:

```powershell
$wordApp = New-Object -ComObject Word.Application
$wordApp.Visible = $false

Get-ChildItem "C:\Users\Lenovo\Generated_Admit_Cards_Fixed\*.docx" | ForEach-Object {
    $doc = $wordApp.Documents.Open($_.FullName)
    $pdfPath = $_.FullName -replace '\.docx$', '.pdf'
    $doc.SaveAs([ref]$pdfPath, [ref]17)
    $doc.Close()
}

$wordApp.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wordApp)
```

### Option 5: Use Adobe Acrobat (If you have it)

1. Open Adobe Acrobat
2. Tools → Create PDF → Multiple Files
3. Add all DOCX files from folder
4. Convert

## 📊 Current Status

- ✅ **167 DOCX files** ready in `C:\Users\Lenovo\Generated_Admit_Cards_Fixed`
- ✅ All data complete (including mobile numbers)
- ⏳ **Need PDF conversion** - Choose option above

## 🎯 My Recommendation

**For 167 files, I recommend Option 3 (LibreOffice):**

1. Download and install LibreOffice (free, 5 minutes)
2. Run one command in PowerShell
3. All 167 PDFs generated automatically
4. No file count limits, completely free

**Quick LibreOffice install and convert:**
```powershell
# After installing LibreOffice, run:
& "C:\Program Files\LibreOffice\program\soffice.exe" --headless --convert-to pdf --outdir "C:\Users\Lenovo\Generated_Admit_Cards_Fixed" "C:\Users\Lenovo\Generated_Admit_Cards_Fixed\*.docx"
```

This will take about 5-10 minutes for 167 files and create all PDFs in the same folder!
