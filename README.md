# Admit Card Generator

Automatically generates personalized admit cards from an Excel data file and a Word template.

## Features

- Reads student/exam data from Excel (.xlsx) files
- Replaces placeholders in Word (.docx) templates with data from Excel
- Generates individual admit cards (one per row of data)
- Automatically converts .docx files to PDF (requires Microsoft Word on Windows)
- Supports custom template and data file paths via command-line arguments

## Requirements

- Python 3.11 or higher
- Microsoft Word (for PDF conversion)
- Required Python packages (see `requirements.txt`)

## Installation

1. Install Python dependencies:
```powershell
python -m pip install -r requirements.txt
```

## Usage

### Default usage (using hardcoded paths in script):
```powershell
python new_file.py
```

### With custom paths:
```powershell
python new_file.py --template "path\to\template.docx" --data "path\to\data.xlsx" --out "output_folder"
```

### Command-line options:
- `--template`: Path to the Word template file (default: ASP Admit Card template in Downloads)
- `--data`: Path to the Excel file with student data (default: ASP Online Reg. Boy.xlsx)
- `--out`: Output directory for generated files (default: Generated_Admit_Cards)

## Template Format

Your Word template should contain placeholders in the format `{{ColumnName}}` where `ColumnName` matches the column headers in your Excel file.

**Example:**
- Excel column: `Student Name`
- Template placeholder: `{{Student Name}}`

## Output

The script generates:
- Individual .docx files for each row in the Excel file
- Individual .pdf files (if docx2pdf and Word are available)
- Files are named: `Admit_{ExamRoll}.docx` and `Admit_{ExamRoll}.pdf`

## Troubleshooting

### Error: Template file not found
Make sure the template path exists or provide the correct path using `--template`.

### Error: Data file not found
Make sure the data Excel file exists or provide the correct path using `--data`.

### PDF conversion fails
- Ensure Microsoft Word is installed on Windows
- The script will skip PDF conversion if docx2pdf is not available (prints a warning)

### Module not found errors
Install all required packages:
```powershell
python -m pip install -r requirements.txt
```

## Example Run

```powershell
PS C:\Users\Lenovo> python C:\Users\Public\new_file.py
100%|████████████████████████████████| 168/168 [02:20<00:00,  1.19it/s]
```

This example processed 168 records and generated 336 files (168 .docx + 168 .pdf).
