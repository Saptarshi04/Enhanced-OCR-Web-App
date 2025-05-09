# table_extractors.py
import os
import fitz  # PyMuPDF
import camelot
import tabula
import pandas as pd
import numpy as np
import cv2
from PIL import Image

class TableExtractor:
    """Base class for table extraction"""
    def extract_tables(self, pdf_path, page_num=0):
        """Extract tables from a PDF page"""
        raise NotImplementedError("Subclasses must implement extract_tables method")
    
    def is_compatible(self, pdf_path):
        """Check if the extractor is compatible with the given PDF"""
        return True

class PyMuPDFTableExtractor(TableExtractor):
    """Table extraction using PyMuPDF"""
    def extract_tables(self, pdf_path, page_num=0):
        """Extract tables from a PDF page using PyMuPDF"""
        doc = fitz.open(pdf_path)
        page = doc[page_num]
        
        # Extract tables using PyMuPDF's built-in table detection
        tables = page.find_tables()
        
        extracted_tables = []
        for table in tables:
            # Extract cells as a list of lists
            rows = []
            for i in range(table.row_count):
                row = []
                for j in range(table.column_count):
                    cell = table.extract_cell(i, j)
                    row.append(cell.text if cell else "")
                rows.append(row)
            
            # Convert to pandas DataFrame
            if rows:
                df = pd.DataFrame(rows)
                # Use first row as header if it contains text
                if not df.empty and not df.iloc[0].isna().all():
                    df.columns = df.iloc[0]
                    df = df.iloc[1:]
                extracted_tables.append(df)
        
        return extracted_tables

class CamelotTableExtractor(TableExtractor):
    """Table extraction using Camelot"""
    def extract_tables(self, pdf_path, page_num=0):
        """Extract tables from a PDF page using Camelot"""
        try:
            # Try lattice mode first (for tables with borders)
            tables = camelot.read_pdf(
                pdf_path, 
                pages=str(page_num + 1),  # Camelot uses 1-indexed pages
                flavor='lattice'
            )
            
            # If no tables found or low accuracy, try stream mode
            if len(tables) == 0 or (len(tables) > 0 and tables[0].accuracy < 80):
                stream_tables = camelot.read_pdf(
                    pdf_path, 
                    pages=str(page_num + 1),
                    flavor='stream'
                )
                
                # If stream mode found tables with better accuracy, use those
                if len(stream_tables) > 0 and (len(tables) == 0 or stream_tables[0].accuracy > tables[0].accuracy):
                    tables = stream_tables
            
            return [table.df for table in tables]
        except Exception as e:
            print(f"Camelot extraction error: {e}")
            return []
    
    def is_compatible(self, pdf_path):
        """Check if PDF is compatible with Camelot (text-based, not scanned)"""
        try:
            # Open the PDF
            doc = fitz.open(pdf_path)
            page = doc[0]
            
            # Check if there's extractable text
            text = page.get_text()
            return len(text.strip()) > 0
        except:
            return False

class TabulaTableExtractor(TableExtractor):
    """Table extraction using Tabula"""
    def extract_tables(self, pdf_path, page_num=0):
        """Extract tables from a PDF page using Tabula"""
        try:
            # Try with lattice mode first
            tables = tabula.read_pdf(
                pdf_path, 
                pages=page_num + 1,  # Tabula uses 1-indexed pages
                lattice=True,
                multiple_tables=True
            )
            
            # If no tables found, try with stream mode
            if not tables:
                tables = tabula.read_pdf(
                    pdf_path, 
                    pages=page_num + 1,
                    stream=True,
                    multiple_tables=True
                )
            
            return tables
        except Exception as e:
            print(f"Tabula extraction error: {e}")
            return []
    
    def is_compatible(self, pdf_path):
        """Check if Java is available (required for Tabula)"""
        try:
            import jpype
            return True
        except ImportError:
            try:
                # Check if Java is in PATH
                import subprocess
                subprocess.check_output(['java', '-version'], stderr=subprocess.STDOUT)
                return True
            except:
                return False

class TableExtractorFactory:
    """Factory for creating table extractors"""
    @staticmethod
    def create_extractors():
        """Create all available table extractors"""
        return [
            PyMuPDFTableExtractor(),
            CamelotTableExtractor(),
            TabulaTableExtractor()
        ]
    
    @staticmethod
    def get_best_extractor(pdf_path):
        """Get the best extractor for the given PDF"""
        extractors = TableExtractorFactory.create_extractors()
        compatible_extractors = [e for e in extractors if e.is_compatible(pdf_path)]
        
        if not compatible_extractors:
            # Fallback to PyMuPDF if no compatible extractors
            return PyMuPDFTableExtractor()
        
        # For now, prioritize in this order: Camelot > Tabula > PyMuPDF
        for extractor in compatible_extractors:
            if isinstance(extractor, CamelotTableExtractor):
                return extractor
        
        for extractor in compatible_extractors:
            if isinstance(extractor, TabulaTableExtractor):
                return extractor
        
        return compatible_extractors[0]

def combine_tables(all_tables):
    """Combine and deduplicate tables from different extractors"""
    if not all_tables:
        return []
    
    # Flatten list of lists
    tables = [table for sublist in all_tables for table in sublist if not table.empty]
    
    # Simple deduplication based on shape and first few values
    unique_tables = []
    table_signatures = set()
    
    for table in tables:
        # Create a signature for the table
        if table.empty:
            continue
            
        shape_sig = f"{table.shape[0]}x{table.shape[1]}"
        data_sig = ""
        
        # Take first 3 rows, 3 cols (or less if smaller)
        sample_rows = min(3, table.shape[0])
        sample_cols = min(3, table.shape[1])
        
        sample = table.iloc[:sample_rows, :sample_cols]
        data_sig = str(sample.values.flatten())
        
        signature = f"{shape_sig}:{data_sig}"
        
        if signature not in table_signatures:
            table_signatures.add(signature)
            unique_tables.append(table)
    
    return unique_tables

def extract_all_tables(pdf_path, page_nums=None):
    """Extract tables from all specified pages using all available extractors"""
    doc = fitz.open(pdf_path)
    
    if page_nums is None:
        page_nums = range(len(doc))
    
    all_tables = []
    
    for page_num in page_nums:
        if page_num >= len(doc):
            continue
        
        page_tables = []
        
        # Get the best extractor for this PDF
        extractor = TableExtractorFactory.get_best_extractor(pdf_path)
        
        # Extract tables using the best extractor
        tables = extractor.extract_tables(pdf_path, page_num)
        if tables:
            page_tables.append(tables)
        
        # If no tables found or explicitly requested, try all extractors
        if not tables:
            for extractor in TableExtractorFactory.create_extractors():
                if extractor.is_compatible(pdf_path):
                    tables = extractor.extract_tables(pdf_path, page_num)
                    if tables:
                        page_tables.append(tables)
        
        # Combine and deduplicate tables
        unique_tables = combine_tables(page_tables)
        
        # Add page number information to each table
        for table in unique_tables:
            table.attrs = {'page': page_num}
        
        all_tables.extend(unique_tables)
    
    return all_tables

def export_tables_to_csv(tables, output_dir, base_filename):
    """Export extracted tables to CSV files"""
    os.makedirs(output_dir, exist_ok=True)
    
    for i, table in enumerate(tables):
        page_num = table.attrs.get('page', 0) if hasattr(table, 'attrs') else 0
        csv_filename = os.path.join(output_dir, f"{base_filename}_page{page_num+1}_table{i+1}.csv")
        
        table.to_csv(csv_filename, index=False)
        print(f"Exported table to: {csv_filename}")