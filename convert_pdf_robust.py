"""
Robust DOCX to PDF converter using Word COM with better error handling
"""
import os
import sys
import time
import win32com.client
import pythoncom
from pathlib import Path

def convert_docx_to_pdf_robust(input_folder):
    """Convert all DOCX files to PDF using Word COM with retry logic"""
    
    input_path = Path(input_folder).absolute()
    
    # Get all docx files
    docx_files = list(input_path.glob("*.docx"))
    
    if not docx_files:
        print(f"No DOCX files found in {input_path}")
        return
    
    print(f"Found {len(docx_files)} DOCX files")
    print(f"Output directory: {input_path}")
    print("\nStarting conversion...\n")
    
    success = 0
    failed = 0
    failed_files = []
    
    # Process in smaller batches to avoid Word crashes
    batch_size = 20
    
    for batch_start in range(0, len(docx_files), batch_size):
        batch_end = min(batch_start + batch_size, len(docx_files))
        batch = docx_files[batch_start:batch_end]
        
        print(f"\nProcessing batch {batch_start//batch_size + 1} ({batch_start + 1}-{batch_end} of {len(docx_files)})...")
        
        # Initialize COM for this batch
        pythoncom.CoInitialize()
        
        try:
            # Create new Word instance for this batch
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            
            for i, docx_file in enumerate(batch, 1):
                doc = None
                try:
                    docx_path = str(docx_file)
                    pdf_path = str(docx_file.with_suffix('.pdf'))
                    
                    # Skip if PDF already exists
                    if os.path.exists(pdf_path):
                        print(f"  [{batch_start + i}/{len(docx_files)}] ⏭️  Skipping (PDF exists): {docx_file.name}")
                        success += 1
                        continue
                    
                    # Open document with minimal options
                    doc = word.Documents.Open(
                        FileName=docx_path,
                        ConfirmConversions=False,
                        ReadOnly=True,
                        AddToRecentFiles=False,
                        Visible=False
                    )
                    
                    # Export as PDF (better than SaveAs)
                    doc.ExportAsFixedFormat(
                        OutputFileName=pdf_path,
                        ExportFormat=17,  # wdExportFormatPDF
                        OpenAfterExport=False,
                        OptimizeFor=0,  # wdExportOptimizeForPrint
                        CreateBookmarks=0
                    )
                    
                    doc.Close(SaveChanges=False)
                    doc = None
                    
                    print(f"  [{batch_start + i}/{len(docx_files)}] ✅ {docx_file.name}")
                    success += 1
                    
                except Exception as e:
                    error_msg = str(e)
                    print(f"  [{batch_start + i}/{len(docx_files)}] ❌ {docx_file.name}: {error_msg}")
                    failed += 1
                    failed_files.append((docx_file.name, error_msg))
                    
                    # Try to close document if still open
                    if doc:
                        try:
                            doc.Close(SaveChanges=False)
                        except:
                            pass
                
                # Small delay between files
                time.sleep(0.05)
            
            # Close Word for this batch
            word.Quit()
            time.sleep(0.5)  # Give Word time to fully close
            
        except Exception as e:
            print(f"  Batch error: {e}")
        
        finally:
            pythoncom.CoUninitialize()
            time.sleep(1)  # Pause between batches
    
    # Print final summary
    print(f"\n{'='*70}")
    print(f"CONVERSION COMPLETE")
    print(f"{'='*70}")
    print(f"✅ Successfully converted: {success}/{len(docx_files)} files")
    print(f"❌ Failed: {failed}/{len(docx_files)} files")
    
    if failed_files and failed > 0:
        print(f"\nFailed files (first 20):")
        for filename, error in failed_files[:20]:
            print(f"  - {filename}")
    
    print(f"\n📁 PDF files location: {input_path}")
    
    return success, failed


if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser()
    parser.add_argument("--dir", default="Generated_Admit_Cards_Fixed",
                       help="Directory with DOCX files")
    args = parser.parse_args()
    
    if not os.path.exists(args.dir):
        print(f"Error: Directory not found: {args.dir}")
        sys.exit(1)
    
    success, failed = convert_docx_to_pdf_robust(args.dir)
    
    if failed == 0:
        print("\n🎉 All files converted successfully!")
    else:
        print(f"\n⚠️  {failed} files failed to convert.")
