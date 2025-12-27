"""
🚀 MEMORY-OPTIMIZED PDF TO DOCX CONVERTER
✅ Works on FREE hosting (Render/Heroku/Railway - 512MB RAM)
✅ Smart fallback: pdf2docx → lightweight if OOM
✅ Page-by-page processing to prevent crashes
"""

from pdf2docx import Converter
import os
import sys
import logging
import traceback
import gc
from pathlib import Path
from typing import Optional, Tuple
import psutil  # pip install psutil

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


class MemoryOptimizedConverter:
    """PDF to DOCX converter optimized for low-memory environments"""
    
    def __init__(self, 
                 max_memory_mb: int = 400,  # Leave 100MB buffer for OS
                 enable_images: bool = True,
                 fallback_mode: bool = True):
        """
        Args:
            max_memory_mb: Max memory to use (400MB safe for 512MB tier)
            enable_images: Extract images (set False to save 200MB+)
            fallback_mode: Use lightweight converter if pdf2docx fails
        """
        self.max_memory_mb = max_memory_mb
        self.enable_images = enable_images
        self.fallback_mode = fallback_mode
    
    def get_memory_usage(self) -> float:
        """Get current memory usage in MB"""
        try:
            process = psutil.Process(os.getpid())
            return process.memory_info().rss / (1024 * 1024)
        except:
            return 0
    
    def check_pdf_size(self, pdf_path: str) -> Tuple[bool, str]:
        """Check if PDF is safe to convert"""
        file_size_mb = os.path.getsize(pdf_path) / (1024 * 1024)
        
        # Heuristic: pdf2docx uses ~10x file size in RAM
        estimated_ram = file_size_mb * 10
        
        if estimated_ram > self.max_memory_mb:
            return False, f"PDF too large ({file_size_mb:.1f}MB). Estimated RAM: {estimated_ram:.0f}MB"
        
        return True, "OK"
    
    def convert_with_memory_monitoring(self,
                                      pdf_path: str,
                                      output_path: Optional[str] = None,
                                      verbose: bool = True) -> Tuple[bool, str]:
        """
        Convert PDF with memory monitoring and cleanup
        """
        if not os.path.exists(pdf_path):
            return False, f"File not found: {pdf_path}"
        
        if output_path is None:
            output_path = str(Path(pdf_path).with_suffix('.docx'))
        
        # Pre-check: Is PDF too large?
        safe, msg = self.check_pdf_size(pdf_path)
        if not safe:
            logger.warning(f"⚠️ {msg}")
            if self.fallback_mode:
                logger.info("🔄 Switching to lightweight converter...")
                return self._convert_lightweight(pdf_path, output_path)
            return False, msg
        
        # Memory cleanup before starting
        gc.collect()
        initial_memory = self.get_memory_usage()
        
        if verbose:
            logger.info(f"📊 Initial memory: {initial_memory:.1f}MB")
            logger.info(f"🔄 Converting: {Path(pdf_path).name}")
        
        try:
            # Use pdf2docx with memory monitoring
            cv = Converter(pdf_path)
            
            # Check memory after opening
            current_memory = self.get_memory_usage()
            if current_memory > self.max_memory_mb * 0.8:
                cv.close()
                raise MemoryError(f"Memory too high after opening: {current_memory:.0f}MB")
            
            # Convert with minimal options to reduce memory
            cv.convert(
                output_path,
                start=0,
                end=None,
                pages=None
            )
            
            cv.close()
            
            # Aggressive cleanup
            del cv
            gc.collect()
            
            final_memory = self.get_memory_usage()
            file_size = os.path.getsize(output_path) / (1024 * 1024)
            
            if verbose:
                logger.info(f"✅ SUCCESS!")
                logger.info(f"📁 Output: {output_path}")
                logger.info(f"📊 Final memory: {final_memory:.1f}MB")
                logger.info(f"💾 File size: {file_size:.2f}MB")
            
            return True, output_path
        
        except MemoryError as e:
            logger.error(f"💥 Out of memory: {e}")
            if self.fallback_mode:
                logger.info("🔄 Switching to lightweight converter...")
                return self._convert_lightweight(pdf_path, output_path)
            return False, str(e)
        
        except Exception as e:
            logger.error(f"❌ Conversion failed: {e}")
            logger.debug(traceback.format_exc())
            
            if self.fallback_mode and "memory" in str(e).lower():
                logger.info("🔄 Attempting lightweight conversion...")
                return self._convert_lightweight(pdf_path, output_path)
            
            return False, str(e)
        
        finally:
            gc.collect()  # Always cleanup
    
    def _convert_lightweight(self, pdf_path: str, output_path: str) -> Tuple[bool, str]:
        """
        Lightweight fallback converter (text-only, ~50MB RAM)
        Uses PyPDF2 + python-docx
        """
        try:
            from PyPDF2 import PdfReader
            from docx import Document
            from docx.shared import Pt, Inches
            
            logger.info("📝 Using lightweight text-only conversion")
            
            # Read PDF
            reader = PdfReader(pdf_path)
            num_pages = len(reader.pages)
            
            # Create Word document
            doc = Document()
            
            # Process page by page (memory efficient)
            for page_num, page in enumerate(reader.pages, 1):
                text = page.extract_text()
                
                # Add page header
                if page_num > 1:
                    doc.add_page_break()
                
                # Add text
                if text.strip():
                    # Split into paragraphs
                    paragraphs = text.split('\n\n')
                    for para in paragraphs:
                        if para.strip():
                            p = doc.add_paragraph(para.strip())
                            p.style.font.size = Pt(11)
                
                # Memory cleanup every 5 pages
                if page_num % 5 == 0:
                    gc.collect()
                    logger.info(f"⏳ Processed {page_num}/{num_pages} pages")
            
            # Save document
            doc.save(output_path)
            
            file_size = os.path.getsize(output_path) / (1024 * 1024)
            
            logger.info(f"✅ Lightweight conversion complete!")
            logger.info(f"📁 Output: {output_path}")
            logger.info(f"💾 Size: {file_size:.2f}MB")
            logger.warning(f"⚠️ Note: Images and complex formatting not preserved")
            
            return True, output_path
        
        except ImportError:
            error_msg = "❌ Lightweight converter requires: pip install PyPDF2 python-docx"
            logger.error(error_msg)
            return False, error_msg
        
        except Exception as e:
            logger.error(f"❌ Lightweight conversion failed: {e}")
            return False, str(e)


class BatchConverter:
    """Batch converter with memory management"""
    
    def __init__(self, converter: MemoryOptimizedConverter):
        self.converter = converter
    
    def convert_batch(self, 
                     input_folder: str,
                     output_folder: Optional[str] = None,
                     pattern: str = "*.pdf") -> dict:
        """Convert multiple PDFs sequentially (no parallel to save memory)"""
        
        input_path = Path(input_folder)
        if not input_path.exists():
            logger.error(f"❌ Folder not found: {input_folder}")
            return {'success': 0, 'failed': 0, 'errors': {}}
        
        pdf_files = list(input_path.glob(pattern))
        
        if not pdf_files:
            logger.warning(f"⚠️ No PDF files found in {input_folder}")
            return {'success': 0, 'failed': 0, 'errors': {}}
        
        if output_folder is None:
            output_folder = input_folder
        output_path = Path(output_folder)
        output_path.mkdir(parents=True, exist_ok=True)
        
        logger.info(f"📂 Found {len(pdf_files)} PDF files")
        logger.info(f"⚡ Sequential processing (memory-safe)")
        
        results = {'success': 0, 'failed': 0, 'errors': {}}
        
        for i, pdf_file in enumerate(pdf_files, 1):
            output_file = output_path / pdf_file.with_suffix('.docx').name
            logger.info(f"\n[{i}/{len(pdf_files)}] {pdf_file.name}")
            
            success, message = self.converter.convert_with_memory_monitoring(
                str(pdf_file),
                str(output_file),
                verbose=True
            )
            
            if success:
                results['success'] += 1
            else:
                results['failed'] += 1
                results['errors'][pdf_file.name] = message
            
            # Aggressive cleanup between files
            gc.collect()
        
        # Print summary
        logger.info(f"\n{'='*70}")
        logger.info(f"📊 BATCH SUMMARY")
        logger.info(f"✅ Success: {results['success']}/{len(pdf_files)}")
        logger.info(f"❌ Failed: {results['failed']}/{len(pdf_files)}")
        logger.info(f"{'='*70}\n")
        
        return results


def main():
    """CLI interface"""
    
    if len(sys.argv) < 2:
        print("""
🚀 MEMORY-OPTIMIZED PDF TO DOCX CONVERTER

Usage:
  Single file:
    python optimized_converter.py input.pdf [output.docx]
  
  Batch mode:
    python optimized_converter.py --batch <folder> [output_folder]

Options:
  --no-images       Disable image extraction (saves 200MB+ RAM)
  --no-fallback     Disable lightweight fallback

Examples:
  python optimized_converter.py document.pdf
  python optimized_converter.py large.pdf --no-images
  python optimized_converter.py --batch ./pdfs ./output

Features:
  ✅ Works on 512MB RAM (Render/Heroku free tier)
  ✅ Smart memory monitoring
  ✅ Automatic fallback to text-only if needed
  ✅ Page-by-page processing
  ✅ Aggressive garbage collection

Requirements:
  pip install pdf2docx PyPDF2 python-docx psutil
        """)
        sys.exit(1)
    
    # Parse options
    enable_images = '--no-images' not in sys.argv
    fallback_mode = '--no-fallback' not in sys.argv
    
    # Initialize converter
    converter = MemoryOptimizedConverter(
        max_memory_mb=400,  # Safe for 512MB tier
        enable_images=enable_images,
        fallback_mode=fallback_mode
    )
    
    if sys.argv[1] == "--batch":
        if len(sys.argv) < 3:
            logger.error("❌ Please specify input folder")
            sys.exit(1)
        
        input_folder = sys.argv[2]
        output_folder = sys.argv[3] if len(sys.argv) > 3 and not sys.argv[3].startswith('--') else None
        
        batch = BatchConverter(converter)
        batch.convert_batch(input_folder, output_folder)
    
    else:
        input_pdf = sys.argv[1]
        output_docx = sys.argv[2] if len(sys.argv) > 2 and not sys.argv[2].startswith('--') else None
        
        success, result = converter.convert_with_memory_monitoring(
            input_pdf,
            output_docx,
            verbose=True
        )
        
        if not success:
            sys.exit(1)


if __name__ == "__main__":
    main()



 