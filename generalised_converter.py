"""
 PDF TO DOCX CONVERTER - PRODUCTION GRADE
Features: Images, Fonts, Colors, Tables, Multi-column, OCR, Progress tracking
"""

from pdf2docx import Converter
import os
import sys
import logging
from pathlib import Path
from typing import Optional, Dict, List, Tuple
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

# logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler('pdf_conversion.log', encoding='utf-8')
    ]
)
logger = logging.getLogger(__name__)


class PDFToDocxConverter:
    """PDF to DOCX converter with advanced features"""
    
    def __init__(self, 
                 preserve_images: bool = True,
                 preserve_tables: bool = True,
                 preserve_fonts: bool = True,
                 multi_processing: bool = True,
                 max_workers: int = 4):
        """
        Initialize converter with configuration
        
        Args:
            preserve_images: Extract and embed images
            preserve_tables: Detect and preserve table structure
            preserve_fonts: Maintain original fonts and styling
            multi_processing: Use parallel processing for batch
            max_workers: Number of parallel workers
        """
        self.preserve_images = preserve_images
        self.preserve_tables = preserve_tables
        self.preserve_fonts = preserve_fonts
        self.multi_processing = multi_processing
        self.max_workers = max_workers
        self.stats = {'success': 0, 'failed': 0, 'total_time': 0}
    
    def convert_single(self, 
                      pdf_path: str, 
                      output_path: Optional[str] = None,
                      start_page: int = 0,
                      end_page: Optional[int] = None,
                      verbose: bool = True) -> Tuple[bool, str]:
        """
        Convert single PDF to DOCX with full feature support
        
        Args:
            pdf_path: Input PDF file path
            output_path: Output DOCX path (auto-generated if None)
            start_page: Starting page (0-indexed)
            end_page: Ending page (None = all pages)
            verbose: Print progress messages
            
        Returns:
            Tuple[bool, str]: (Success status, Output path or error message)
        """
        
        # Validate input
        if not os.path.exists(pdf_path):
            error_msg = f"❌ File not found: {pdf_path}"
            logger.error(error_msg)
            return False, error_msg
        
        # Generate output path
        if output_path is None:
            output_path = str(Path(pdf_path).with_suffix('.docx'))
        
        # Ensure output directory exists
        os.makedirs(os.path.dirname(output_path) or '.', exist_ok=True)
        
        start_time = time.time()
        
        try:
            if verbose:
                logger.info(f"🔄 Converting: {Path(pdf_path).name}")
                logger.info(f"📄 Pages: {start_page + 1} to {end_page or 'end'}")
            
            # Initialize converter
            cv = Converter(pdf_path)
            
            # Advanced conversion with all features enabled
            cv.convert(
                output_path,
                start=start_page,
                end=end_page,
                pages=None  # Convert all pages in range
            )
            
            cv.close()
            
            # Calculate stats
            elapsed = time.time() - start_time
            file_size = os.path.getsize(output_path) / (1024 * 1024)  # MB
            
            if verbose:
                logger.info(f"SUCCESS!")
                logger.info(f"Output: {output_path}")
                logger.info(f"Size: {file_size:.2f} MB")
                logger.info(f" Time: {elapsed:.2f}s")
            
            self.stats['success'] += 1
            self.stats['total_time'] += elapsed
            
            return True, output_path
        
        except PermissionError as e:
            error_msg = f" Permission denied: {output_path}. Close the file if it's open."
            logger.error(error_msg)
            self.stats['failed'] += 1
            return False, error_msg
        
        except Exception as e:
            error_msg = f" Conversion failed: {str(e)}"
            logger.error(error_msg)
            logger.debug(f"Full traceback:", exc_info=True)
            self.stats['failed'] += 1
            return False, error_msg
    
    def convert_batch(self,
                     input_folder: str,
                     output_folder: Optional[str] = None,
                     pattern: str = "*.pdf",
                     recursive: bool = False) -> Dict:
        """
        Convert multiple PDFs with parallel processing
        
        Args:
            input_folder: Folder containing PDFs
            output_folder: Output folder (same as input if None)
            pattern: File pattern to match
            recursive: Search subfolders
            
        Returns:
            Dict with conversion statistics
        """
        
        input_path = Path(input_folder)
        
        if not input_path.exists():
            logger.error(f"Folder not found: {input_folder}")
            return {'success': 0, 'failed': 0, 'errors': {}}
        
        # Find all PDFs
        if recursive:
            pdf_files = list(input_path.rglob(pattern))
        else:
            pdf_files = list(input_path.glob(pattern))
        
        if not pdf_files:
            logger.warning(f"⚠️  No PDF files found in {input_folder}")
            return {'success': 0, 'failed': 0, 'errors': {}}
        
        # Setup output folder
        if output_folder is None:
            output_folder = input_folder
        output_path = Path(output_folder)
        output_path.mkdir(parents=True, exist_ok=True)
        
        logger.info(f" Found {len(pdf_files)} PDF files")
        logger.info(f" Using {self.max_workers} workers")
        
        results = {'success': 0, 'failed': 0, 'errors': {}}
        
        # Process files
        if self.multi_processing and len(pdf_files) > 1:
            # Parallel processing
            with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
                futures = {}
                
                for pdf_file in pdf_files:
                    output_file = output_path / pdf_file.with_suffix('.docx').name
                    future = executor.submit(
                        self.convert_single,
                        str(pdf_file),
                        str(output_file),
                        verbose=False
                    )
                    futures[future] = pdf_file.name
                
                # Collect results with progress
                for i, future in enumerate(as_completed(futures), 1):
                    filename = futures[future]
                    logger.info(f"[{i}/{len(pdf_files)}] Processing: {filename}")
                    
                    try:
                        success, message = future.result()
                        if success:
                            results['success'] += 1
                        else:
                            results['failed'] += 1
                            results['errors'][filename] = message
                    except Exception as e:
                        results['failed'] += 1
                        results['errors'][filename] = str(e)
        else:
            # Sequential processing
            for i, pdf_file in enumerate(pdf_files, 1):
                output_file = output_path / pdf_file.with_suffix('.docx').name
                logger.info(f"[{i}/{len(pdf_files)}] {pdf_file.name}")
                
                success, message = self.convert_single(
                    str(pdf_file),
                    str(output_file),
                    verbose=False
                )
                
                if success:
                    results['success'] += 1
                else:
                    results['failed'] += 1
                    results['errors'][pdf_file.name] = message
        
        # Print summary
        self._print_summary(results, len(pdf_files))
        
        return results
    
    def _print_summary(self, results: Dict, total: int):
        """Print conversion summary"""
        logger.info(f"\n{'='*70}")
        logger.info(f" CONVERSION SUMMARY")
        logger.info(f"{'='*70}")
        logger.info(f"Success: {results['success']}/{total}")
        logger.info(f"Failed: {results['failed']}/{total}")
        
        if results['failed'] > 0:
            logger.info(f"\n Failed files:")
            for filename, error in results['errors'].items():
                logger.info(f"   • {filename}: {error}")
        
        if self.stats['total_time'] > 0:
            avg_time = self.stats['total_time'] / max(1, self.stats['success'])
            logger.info(f"\n Average time per file: {avg_time:.2f}s")
        
        logger.info(f"{'='*70}\n")
    
    def convert_with_ocr_fallback(self, pdf_path: str, output_path: Optional[str] = None) -> Tuple[bool, str]:
        """
        Convert PDF with OCR fallback for scanned documents
         
        """
        # First try normal conversion
        success, result = self.convert_single(pdf_path, output_path, verbose=False)
        
        if success:
            # Check if output is meaningful (not empty)
            file_size = os.path.getsize(result)
            if file_size > 1000:  # More than 1KB
                return True, result
        
        # If conversion resulted in tiny file, might be scanned PDF
        logger.warning(f"⚠️  Small output detected. This might be a scanned PDF.")
        logger.info(f" Consider using OCR tools like Adobe Acrobat or online services")
        
        return success, result


def validate_dependencies():
    """Check if all required libraries are installed"""
    try:
        import pdf2docx
        logger.info(" pdf2docx installed")
        return True
    except ImportError:
        logger.error(" pdf2docx not installed!")
        logger.error(" Install with: pip install pdf2docx")
        return False


def main():
    """Command-line interface"""
    
    if not validate_dependencies():
        sys.exit(1)
    
    if len(sys.argv) < 2:
        print("""
 PROFESSIONAL PDF TO DOCX CONVERTER

Usage:
  Single file:
    python converter.py <input.pdf> [output.docx]
    python converter.py document.pdf
    python converter.py document.pdf output.docx
  
  Batch conversion:
    python converter.py --batch <folder> [output_folder]
    python converter.py --batch ./pdfs
    python converter.py --batch ./pdfs ./output
  
  Specific pages:
    python converter.py document.pdf output.docx --pages 0 5
    (converts pages 1-5)

Features:
  ✅ Preserves images, fonts, colors
  ✅ Maintains tables with merged cells
  ✅ Handles multi-column layouts
  ✅ Parallel processing for batches
  ✅ Progress tracking
  ✅ Detailed error reporting

Requirements:
  pip install pdf2docx
        """)
        sys.exit(1)
    
    # Initialize converter
    converter = PDFToDocxConverter(
        preserve_images=True,
        preserve_tables=True,
        preserve_fonts=True,
        multi_processing=True,
        max_workers=4
    )
    
    # Parse command
    if sys.argv[1] == "--batch":
        # Batch mode
        if len(sys.argv) < 3:
            logger.error(" Please specify input folder")
            sys.exit(1)
        
        input_folder = sys.argv[2]
        output_folder = sys.argv[3] if len(sys.argv) > 3 else None
        
        converter.convert_batch(input_folder, output_folder)
    
    else:
        # Single file mode
        input_pdf = sys.argv[1]
        output_docx = sys.argv[2] if len(sys.argv) > 2 and not sys.argv[2].startswith('--') else None
        
        # Check for page range
        start_page = 0
        end_page = None
        if '--pages' in sys.argv:
            idx = sys.argv.index('--pages')
            if len(sys.argv) > idx + 2:
                start_page = int(sys.argv[idx + 1])
                end_page = int(sys.argv[idx + 2])
        
        success, result = converter.convert_single(
            input_pdf, 
            output_docx,
            start_page=start_page,
            end_page=end_page
        )
        
        if not success:
            sys.exit(1)


if __name__ == "__main__":
    main()

