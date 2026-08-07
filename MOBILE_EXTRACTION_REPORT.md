# Mobile Number Extraction Report
**Date:** October 16, 2025  
**Source Directory:** `C:\Users\Public\Excel to Admit Card\Generated_Admit_Cards_Fixed`

## 📊 Extraction Summary

### Overall Statistics
- **Total DOCX Files Processed:** 221
- **Mobile Numbers Successfully Extracted:** 216 (97.7%)
- **Mobile Numbers NOT Found:** 5 (2.3%)

## ✅ Successful Extraction

The script successfully extracted mobile numbers from **216 out of 221 DOCX files** by parsing the XML content inside each DOCX file.

### Mobile Number Formats Found:
- Standard format: `01234567890` (11 digits)
- International format: `+8801234567890`
- Alternative format: `8801234567890`

## ❌ Files Missing Mobile Numbers

The following 5 files did NOT contain any extractable mobile numbers:

| # | Filename | Exam Roll |
|---|----------|-----------|
| 1 | Admit_51125102.docx | 51125102 |
| 2 | Admit_52125116.docx | 52125116 |
| 3 | Admit_61125080.docx | 61125080 |
| 4 | Admit_611251350.docx | 611251350 |
| 5 | Admit_72125060.docx | 72125060 |

**Possible Reasons:**
- Mobile number field was empty in source data
- Placeholder not replaced in template
- Invalid/missing data in Excel file

## 📁 Output Files

### Generated Files:
1. **CSV File:** `extracted_mobile_numbers.csv` (9.5 KB)
2. **Excel File:** `extracted_mobile_numbers.xlsx` (10.5 KB)

### File Location:
```
C:\Users\Public\Excel to Admit Card\Generated_Admit_Cards_Fixed\
```

### Column Structure:
| Column | Description |
|--------|-------------|
| Filename | Original DOCX filename |
| Exam Roll | Extracted exam roll number |
| Mobile Number | Extracted mobile number (or "NOT FOUND") |

## 📝 Sample Extracted Data

| Exam Roll | Mobile Number |
|-----------|---------------|
| 11125018 | 01819510955 |
| 11125019 | 01718815454 |
| 11125037 | +8801812977571 |
| 41125059 | 01920186550 |
| 51125103 | 01854148844 |
| 61125072 | 01608242350 |
| 71125069 | 01818671665 |
| 81125060 | 01539314488 |
| 91125023 | 01742690738 |

## 🔍 Technical Details

### Extraction Method:
- Used `python-docx` library to parse DOCX files
- DOCX files are ZIP archives containing XML files
- Extracted text from:
  - Document paragraphs
  - Table cells
  - All text runs

### Pattern Matching:
- Used regular expression: `(\+?88)?[-\s]?0?1[3-9]\d{8,9}`
- Matches Bangladesh mobile numbers in various formats
- Captures numbers with or without country code

### Processing:
- All 221 files processed sequentially
- Each file parsed individually
- Results compiled into DataFrame
- Exported to both CSV and Excel formats

## 💡 Recommendations

### For Missing Mobile Numbers:
1. Check the original Excel data file for these exam rolls
2. Verify if mobile numbers were present in source data
3. Manually add mobile numbers if needed
4. Regenerate specific admit cards if data is corrected

### Data Usage:
- Use CSV file for importing into other systems
- Use Excel file for manual review and editing
- Keep as backup reference for mobile numbers
- Can be used for SMS notifications or contact lists

## 📈 Success Rate

```
Success Rate: 97.7%
Missing Rate: 2.3%
Total Processed: 221 files
```

This is an excellent extraction rate, indicating that the vast majority of admit cards have properly formatted and accessible mobile numbers.

---

**Status:** ✅ Mobile number extraction completed successfully!  
**Files Generated:** 2 (CSV + Excel)  
**Quality:** 97.7% success rate
