"""
Extract mobile numbers from DOCX files (XML-based format)
"""
import os
import re
from docx import Document
from pathlib import Path
import pandas as pd

def extract_mobile_from_docx(docx_path):
    """Extract mobile number from a DOCX file"""
    try:
        doc = Document(docx_path)
        
        # Pattern to match mobile numbers (various formats)
        # Matches: +8801234567890, 01234567890, +88-01234567890, etc.
        mobile_pattern = r'(\+?88)?[-\s]?0?1[3-9]\d{8,9}'
        
        all_text = []
        
        # Extract from paragraphs
        for para in doc.paragraphs:
            all_text.append(para.text)
        
        # Extract from tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        all_text.append(para.text)
        
        # Join all text
        full_text = ' '.join(all_text)
        
        # Find all mobile numbers
        matches = re.findall(mobile_pattern, full_text)
        
        # Clean up the found numbers
        mobile_numbers = []
        for match in re.finditer(mobile_pattern, full_text):
            number = match.group(0).strip()
            mobile_numbers.append(number)
        
        # Return the first valid mobile number found
        if mobile_numbers:
            return mobile_numbers[0]
        
        return None
        
    except Exception as e:
        print(f"Error reading {docx_path}: {e}")
        return None


def extract_all_mobiles(directory):
    """Extract mobile numbers from all DOCX files in directory"""
    
    dir_path = Path(directory)
    
    if not dir_path.exists():
        print(f"Directory not found: {directory}")
        return
    
    # Get all DOCX files
    docx_files = list(dir_path.glob("*.docx"))
    
    if not docx_files:
        print(f"No DOCX files found in {directory}")
        return
    
    print(f"Found {len(docx_files)} DOCX files")
    print(f"Extracting mobile numbers...\n")
    
    results = []
    
    for docx_file in sorted(docx_files):
        # Extract exam roll from filename (e.g., Admit_12345.docx -> 12345)
        filename = docx_file.stem  # Gets filename without extension
        exam_roll = filename.replace('Admit_', '')
        
        mobile = extract_mobile_from_docx(docx_file)
        
        results.append({
            'Filename': docx_file.name,
            'Exam Roll': exam_roll,
            'Mobile Number': mobile if mobile else 'NOT FOUND'
        })
        
        if mobile:
            print(f"✅ {exam_roll}: {mobile}")
        else:
            print(f"❌ {exam_roll}: NOT FOUND")
    
    # Create DataFrame
    df = pd.DataFrame(results)
    
    # Save to CSV
    output_csv = os.path.join(directory, 'extracted_mobile_numbers.csv')
    df.to_csv(output_csv, index=False)
    
    # Save to Excel
    output_excel = os.path.join(directory, 'extracted_mobile_numbers.xlsx')
    df.to_excel(output_excel, index=False)
    
    # Print summary
    print(f"\n{'='*70}")
    print(f"EXTRACTION COMPLETE")
    print(f"{'='*70}")
    print(f"Total files processed: {len(docx_files)}")
    print(f"Mobile numbers found: {len([r for r in results if r['Mobile Number'] != 'NOT FOUND'])}")
    print(f"Mobile numbers NOT found: {len([r for r in results if r['Mobile Number'] == 'NOT FOUND'])}")
    
    # Show files without mobile numbers
    not_found = [r for r in results if r['Mobile Number'] == 'NOT FOUND']
    if not_found:
        print(f"\n⚠️  Files without mobile numbers:")
        for item in not_found[:20]:  # Show first 20
            print(f"  - {item['Filename']}")
        if len(not_found) > 20:
            print(f"  ... and {len(not_found) - 20} more")
    
    print(f"\n📁 Results saved to:")
    print(f"  - {output_csv}")
    print(f"  - {output_excel}")
    
    return df


if __name__ == "__main__":
    directory = r"C:\Users\Public\Excel to Admit Card\Generated_Admit_Cards_Fixed"
    df = extract_all_mobiles(directory)
    
    if df is not None:
        print(f"\n{'='*70}")
        print("SAMPLE DATA (First 10 rows):")
        print(f"{'='*70}")
        print(df.head(10).to_string(index=False))
