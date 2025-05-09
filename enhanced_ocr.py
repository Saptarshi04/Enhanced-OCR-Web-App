# enhanced_ocr.py
import os
import sys
import argparse
import tempfile
import shutil
from pathlib import Path
import csv
import json

import ocrmypdf
import pytesseract
from PIL import Image
from pdf2image import convert_from_path
import fitz  # PyMuPDF
import pandas as pd
import numpy as np

# Import our custom modules
from table_extractors import extract_all_tables, export_tables_to_csv
from docx_styler import DocxStyler

# Configuration (update these paths based on your installation)
# If Tesseract is in PATH, you can comment this line
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'  # Update for your system

# For Windows, set poppler path if needed
POPPLER_PATH = None
# POPPLER_PATH = r'C:\path\to\poppler\bin'  # Uncomment and update for Windows

def setup_args():
    """Set up command line arguments"""
    parser = argparse.ArgumentParser(description='Convert scanned images to searchable text documents with enhanced table detection')
    parser.add_argument('input_file', help='Input file path (JPG, JPEG, PNG, TIFF, PDF)')
    parser.add_argument('--output', '-o', help='Output file path (PDF or DOCX)')
    parser.add_argument('--format', '-f', choices=['pdf', 'docx'], default='pdf', 
                      help='Output format: pdf or docx (default: pdf)')
    parser.add_argument('--lang', '-l', default='eng', 
                      help='OCR language (default: eng). Use + for multiple (e.g., eng+fra)')
    parser.add_argument('--dpi', type=int, default=300, 
                      help='DPI for image processing (default: 300)')
    parser.add_argument('--deskew', action='store_true', 
                      help='Deskew the scanned image')
    parser.add_argument('--clean', action='store_true', 
                      help='Clean the image before OCR')
    parser.add_argument('--table-detection', action='store_true', default=True,
                      help='Try to detect and preserve tables (default: enabled)')
    parser.add_argument('--table-extraction-method', choices=['pymupdf', 'camelot', 'tabula', 'all'], default='all',
                      help='Table extraction method to use (default: all)')
    parser.add_argument('--export-tables', 
                      help='Export tables to CSV files (provide directory path)')
    parser.add_argument('--table-style', choices=['basic', 'grid', 'light', 'fancy'], default='grid',
                      help='Style for tables in DOCX output (default: grid)')
    return parser.parse_args()

def image_to_pdf(image_path, output_path, dpi=300):
    """Convert a single image to PDF"""
    image = Image.open(image_path)
    # If DPI info isn't in the image, we'll set it
    if dpi:
        image.save(output_path, 'PDF', resolution=dpi)
    else:
        image.save(output_path, 'PDF')
    return output_path

def get_table_style_name(style_option):
    """Get the Word table style name based on the option"""
    style_map = {
        'basic': 'Table Normal',
        'grid': 'Table Grid',
        'light': 'Light List',
        'fancy': 'Medium Shading 1 Accent 1'
    }
    return style_map.get(style_option, 'Table Grid')

def image_to_searchable_pdf(image_path, output_path, lang='eng', dpi=300, deskew=False, clean=False,
                           table_detection=True, table_method='all', export_tables_dir=None):
    """Convert a single image to searchable PDF with enhanced table detection"""
    with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as temp_pdf:
        temp_pdf_path = temp_pdf.name
    
    # First convert the image to a plain PDF
    image_to_pdf(image_path, temp_pdf_path, dpi)
    
    # Then use OCRmyPDF to make it searchable
    ocr_options = {
        'language': lang,
        'deskew': deskew,
        'clean': clean,
        'optimize': 1,
        'output_type': 'pdfa',
        'progress_bar': True
    }
    
    try:
        ocrmypdf.ocr(temp_pdf_path, output_path, **ocr_options)
        print(f"Created searchable PDF: {output_path}")
        
        # Handle table detection if requested
        tables = []
        if table_detection:
            print("Detecting tables...")
            
            # Extract tables based on the method chosen
            if table_method == 'all':
                tables = extract_all_tables(output_path)
            else:
                # This would need to be extended to support individual methods
                tables = extract_all_tables(output_path)
            
            print(f"Found {len(tables)} tables in the document")
            
            # Export tables if requested
            if export_tables_dir and tables:
                base_filename = os.path.splitext(os.path.basename(image_path))[0]
                export_tables_to_csv(tables, export_tables_dir, base_filename)
        
        return tables
    except Exception as e:
        print(f"Error in OCR process: {e}")
        return []
    finally:
        # Clean up temp file
        if os.path.exists(temp_pdf_path):
            os.unlink(temp_pdf_path)

def pdf_to_searchable_pdf(pdf_path, output_path, lang='eng', deskew=False, clean=False, 
                         table_detection=True, table_method='all', export_tables_dir=None):
    """Convert a PDF to searchable PDF with enhanced table detection"""
    ocr_options = {
        'language': lang,
        'deskew': deskew,
        'clean': clean,
        'optimize': 1,
        'output_type': 'pdfa',
        'progress_bar': True
    }
    
    try:
        ocrmypdf.ocr(pdf_path, output_path, **ocr_options)
        print(f"Created searchable PDF: {output_path}")
        
        # Handle table detection if requested
        tables = []
        if table_detection:
            # Extract tables based on the method chosen
            if table_method == 'all':
                tables = extract_all_tables(output_path)
            else:
                # This would need to be extended to support individual methods
                tables = extract_all_tables(output_path)
            
            print(f"Found {len(tables)} tables in the document")
            
            # Export tables if requested
            if export_tables_dir and tables:
                base_filename = os.path.splitext(os.path.basename(pdf_path))[0]
                export_tables_to_csv(tables, export_tables_dir, base_filename)
        
        return tables
    except Exception as e:
        print(f"Error in OCR process: {e}")
        return []

def pdf_to_docx(pdf_path, docx_path, lang='eng', with_tables=None, table_style='grid'):
    """Convert PDF to DOCX maintaining structure and styling tables"""
    # Create a styler for the document
    styler = DocxStyler()
    doc = styler.get_document()
    
    # Open the PDF
    pdf = fitz.open(pdf_path)
    
    for page_num in range(len(pdf)):
        page = pdf[page_num]
        
        # Add a page heading
        styler.apply_heading_style(f"Page {page_num + 1}", level=1)
        
        # Get tables for this page
        page_tables = []
        if with_tables:
            page_tables = [table for table in with_tables if hasattr(table, 'attrs') and table.attrs.get('page') == page_num]
        
        if page_tables:
            # Extract text with blocks to maintain structure
            blocks = page.get_text("blocks")
            
            # Sort blocks by vertical position
            blocks.sort(key=lambda b: b[1])  # Sort by y0 coordinate
            
            # Get table positions to avoid duplicating content
            table_positions = []
            for table in page.find_tables():
                table_positions.append(table.rect)
            
            for block in blocks:
                # Check if this block overlaps with any table
                block_rect = fitz.Rect(block[:4])
                overlaps_table = any(block_rect.intersects(table_pos) for table_pos in table_positions)
                
                if not overlaps_table:
                    # Extract text from the block
                    text = block[4]
                    # Add the text to the document
                    styler.add_styled_paragraph(text)
            
            # Now add the tables with proper styling
            for table in page_tables:
                # Add a small heading for the table
                styler.add_styled_paragraph("Table", bold=True, size=12)
                
                # Add the table with styling
                style_name = get_table_style_name(table_style)
                styler.add_table_from_dataframe(table, style=style_name)
                
                # Add some space after the table
                styler.add_styled_paragraph("")
        else:
            # No tables, just process blocks normally
            blocks = page.get_text("blocks")
            blocks.sort(key=lambda b: b[1])  # Sort by y0 coordinate
            
            for block in blocks:
                # Extract text from the block
                text = block[4]
                # Add the text to the document
                styler.add_styled_paragraph(text)
        
        # Add a page break after each page unless it's the last page
        if page_num < len(pdf) - 1:
            styler.add_page_break()
    
    # Save the document
    doc.save(docx_path)
    print(f"Created Word document: {docx_path}")

def process_image_to_docx(image_path, output_path, lang='eng', dpi=300, deskew=False, clean=False,
                         table_detection=True, table_method='all', export_tables_dir=None, table_style='grid'):
    """Process an image to DOCX format via PDF with enhanced table handling"""
    # First make a searchable PDF
    with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as temp_pdf:
        temp_pdf_path = temp_pdf.name
    
    tables = image_to_searchable_pdf(
        image_path, 
        temp_pdf_path, 
        lang, 
        dpi, 
        deskew, 
        clean,
        table_detection,
        table_method,
        export_tables_dir
    )
    
    # Then convert the PDF to DOCX
    pdf_to_docx(temp_pdf_path, output_path, lang, with_tables=tables, table_style=table_style)
    
    # Clean up temp file
    if os.path.exists(temp_pdf_path):
        os.unlink(temp_pdf_path)

def main():
    """Main function to process the input file"""
    args = setup_args()
    
    # Check if input file exists
    if not os.path.exists(args.input_file):
        print(f"Error: Input file '{args.input_file}' does not exist.")
        sys.exit(1)
    
    # Determine input file type
    input_path = Path(args.input_file)
    input_type = input_path.suffix.lower()
    
    # Set output file path if not specified
    if args.output:
        output_path = args.output
    else:
        if args.format == 'pdf':
            output_path = str(input_path.with_suffix('.searchable.pdf'))
        else:
            output_path = str(input_path.with_suffix('.docx'))
    
    # Process based on input file type and desired output
    if input_type in ['.jpg', '.jpeg', '.png', '.tiff', '.tif']:
        print(f"Processing image file: {input_path}")
        
        if args.format == 'pdf':
            image_to_searchable_pdf(
                args.input_file, 
                output_path, 
                lang=args.lang, 
                dpi=args.dpi, 
                deskew=args.deskew, 
                clean=args.clean,
                table_detection=args.table_detection,
                table_method=args.table_extraction_method,
                export_tables_dir=args.export_tables
            )
        else:  # docx
            process_image_to_docx(
                args.input_file,
                output_path,
                lang=args.lang,
                dpi=args.dpi,
                deskew=args.deskew,
                clean=args.clean,
                table_detection=args.table_detection,
                table_method=args.table_extraction_method,
                export_tables_dir=args.export_tables,
                table_style=args.table_style
            )
    
    elif input_type == '.pdf':
        print(f"Processing PDF file: {input_path}")
        
        if args.format == 'pdf':
            pdf_to_searchable_pdf(
                args.input_file,
                output_path,
                lang=args.lang,
                deskew=args.deskew,
                clean=args.clean,
                table_detection=args.table_detection,
                table_method=args.table_extraction_method,
                export_tables_dir=args.export_tables
            )
        else:  # docx
            # First make sure the PDF is searchable
            with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as temp_pdf:
                temp_pdf_path = temp_pdf.name
            
            tables = pdf_to_searchable_pdf(
                args.input_file,
                temp_pdf_path,
                lang=args.lang,
                deskew=args.deskew,
                clean=args.clean,
                table_detection=args.table_detection,
                table_method=args.table_extraction_method,
                export_tables_dir=args.export_tables
            )
            
            # Then convert to DOCX
            pdf_to_docx(
                temp_pdf_path, 
                output_path, 
                lang=args.lang, 
                with_tables=tables, 
                table_style=args.table_style
            )
            
            # Clean up temp file
            if os.path.exists(temp_pdf_path):
                os.unlink(temp_pdf_path)
    
    else:
        print(f"Unsupported file type: {input_type}")
        print("Supported types: .jpg, .jpeg, .png, .tiff, .tif, .pdf")
        sys.exit(1)

if __name__ == "__main__":
    main()