"""
Convert all DOCX files to PDF using Microsoft Word
"""
import os
import sys
from tqdm import tqdm
import time

try:
    from docx2pdf import convert
    import pythoncom
    from win32com import client as wc
except ImportError as e:
    print(f"Error: Required package not installed: {e}")
    print("Please run: pip install docx2pdf pywin32")
    sys.exit(1)

def convert_docx_to_pdf_batch(input_dir, close_word_after=True):
    """Convert all DOCX files in a directory to PDF"""
    
    # Get all docx files
    docx_files = [f for f in os.listdir(input_dir) if f.endswith('.docx')]
    
    if not docx_files:
        print(f"No DOCX files found in {input_dir}")
        return
    
    print(f"Found {len(docx_files)} DOCX files to convert")
    print("Starting conversion...\n")
    
    # Initialize COM for this thread
    pythoncom.CoInitialize()
    
    # Create Word application instance
    try:
        word = wc.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0  # Disable alerts
    except Exception as e:
        print(f"Error: Could not start Microsoft Word: {e}")
        print("Make sure Microsoft Word is installed.")
        pythoncom.CoUninitialize()
        return
    
    success_count = 0
    failed_count = 0
    failed_files = []
    
    try:
        for docx_file in tqdm(docx_files, desc="Converting"):
            docx_path = os.path.abspath(os.path.join(input_dir, docx_file))
            pdf_path = docx_path.replace('.docx', '.pdf')
            
            try:
                # Open document
                doc = word.Documents.Open(docx_path, ReadOnly=True)
                
                # Save as PDF (format 17 is PDF)
                doc.SaveAs(pdf_path, FileFormat=17)
                
                # Close document
                doc.Close(False)
                
                success_count += 1
                
            except Exception as e:
                failed_count += 1
                failed_files.append((docx_file, str(e)))
                # Try to close document if it's open
                try:
                    doc.Close(False)
                except:
                    pass
            
            # Small delay to prevent overwhelming Word
            time.sleep(0.1)
    
    finally:
        # Quit Word
        try:
            word.Quit()
        except:
            pass
        
        pythoncom.CoUninitialize()
    
    # Print summary
    print(f"\n{'='*70}")
    print(f"Conversion Summary:")
    print(f"{'='*70}")
    print(f"✅ Successfully converted: {success_count} files")
    print(f"❌ Failed: {failed_count} files")
    
    if failed_files:
        print(f"\nFailed files:")
        for filename, error in failed_files[:10]:  # Show first 10 failures
            print(f"  - {filename}: {error}")
        if len(failed_files) > 10:
            print(f"  ... and {len(failed_files) - 10} more")
    
    print(f"\n✅ PDF files saved to: {os.path.abspath(input_dir)}")


if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser(description="Convert DOCX files to PDF")
    parser.add_argument("--dir", default="Generated_Admit_Cards_Fixed", 
                       help="Directory containing DOCX files")
    args = parser.parse_args()
    
    if not os.path.exists(args.dir):
        print(f"Error: Directory '{args.dir}' not found")
        sys.exit(1)
    
    convert_docx_to_pdf_batch(args.dir)
