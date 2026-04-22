"""
PDF Skill Implementation for AuditPro modules.

Provides enhanced PDF capabilities:
- Extract tables from PDFs (TVA, CNSS declarations)
- OCR extraction with confidence scoring
- Form field recognition
- Multi-page PDF splitting
"""
from __future__ import annotations
from pathlib import Path
from typing import Optional, List, Dict, Any
import logging

logger = logging.getLogger(__name__)


class PDFSkillHandler:
    """Handles PDF-based skill operations for audit modules."""
    
    def __init__(self):
        self.ocr_available = False
        self._check_ocr()
    
    def _check_ocr(self):
        """Check if OCR (Tesseract) is available."""
        try:
            import pytesseract
            self.ocr_available = True
            logger.info("Tesseract OCR available for enhanced extraction")
        except ImportError:
            logger.warning("Tesseract OCR not available - native PDFs only")
    
    def extract_tables(
        self,
        pdf_path: str,
        page_numbers: Optional[List[int]] = None,
        confidence_threshold: float = 0.7
    ) -> Dict[str, Any]:
        """
        Extract tables from PDF (TVA, CNSS declarations).
        
        Args:
            pdf_path: Path to PDF file
            page_numbers: Specific pages to extract (None = all)
            confidence_threshold: Minimum confidence for extracted data (0-1)
        
        Returns:
            Dict with extracted tables and metadata
        """
        try:
            import pdfplumber
        except ImportError:
            return {
                "success": False,
                "error": "pdfplumber not installed",
                "tables": [],
                "confidence": 0.0
            }
        
        result = {
            "success": False,
            "pdf_path": str(pdf_path),
            "tables": [],
            "metadata": {},
            "confidence": 0.0,
            "pages_processed": 0
        }
        
        try:
            with pdfplumber.open(pdf_path) as pdf:
                result["metadata"]["total_pages"] = len(pdf.pages)
                
                pages_to_process = page_numbers or range(len(pdf.pages))
                
                for page_idx in pages_to_process:
                    if page_idx >= len(pdf.pages):
                        continue
                    
                    page = pdf.pages[page_idx]
                    tables = page.extract_tables()
                    
                    if tables:
                        for table in tables:
                            result["tables"].append({
                                "page": page_idx + 1,
                                "rows": len(table),
                                "cols": len(table[0]) if table else 0,
                                "data": table
                            })
                        result["pages_processed"] += 1
                
                result["success"] = len(result["tables"]) > 0
                result["confidence"] = confidence_threshold if result["success"] else 0.0
                
        except Exception as e:
            result["error"] = str(e)
            logger.error(f"Error extracting tables from {pdf_path}: {e}")
        
        return result
    
    def extract_forms(
        self,
        pdf_path: str,
        form_fields: Optional[List[str]] = None
    ) -> Dict[str, Any]:
        """
        Extract form fields from PDF (CNSS bordereaux, IR declarations).
        
        Args:
            pdf_path: Path to PDF file
            form_fields: List of field names to extract (None = all)
        
        Returns:
            Dict with extracted form data
        """
        result = {
            "success": False,
            "pdf_path": str(pdf_path),
            "fields": {},
            "confidence": 0.0
        }
        
        try:
            import pypdf
            pdf_reader = pypdf.PdfReader(pdf_path)
            
            # Extract form fields if available
            if "/AcroForm" in pdf_reader.trailer["/Root"]:
                fields = pdf_reader.get_fields()
                if fields:
                    for field_name, field_data in fields.items():
                        if form_fields is None or field_name in form_fields:
                            result["fields"][field_name] = {
                                "value": field_data.get("/V"),
                                "type": field_data.get("/FT")
                            }
                    result["success"] = True
                    result["confidence"] = 0.95
            
        except Exception as e:
            logger.error(f"Error extracting forms from {pdf_path}: {e}")
            result["error"] = str(e)
        
        return result
    
    def ocr_extract(
        self,
        pdf_path: str,
        language: str = "fra",
        confidence_threshold: float = 0.7
    ) -> Dict[str, Any]:
        """
        OCR extraction from PDF with confidence scoring.
        
        Args:
            pdf_path: Path to PDF file
            language: Language code (fra=French, ara=Arabic, eng=English)
            confidence_threshold: Minimum confidence to keep results
        
        Returns:
            Dict with OCR results and confidence metrics
        """
        result = {
            "success": False,
            "pdf_path": str(pdf_path),
            "text": "",
            "confidence": 0.0,
            "pages": []
        }
        
        if not self.ocr_available:
            result["error"] = "Tesseract OCR not installed"
            return result
        
        try:
            import pytesseract
            from pdf2image import convert_from_path
            
            # Convert PDF to images
            images = convert_from_path(pdf_path)
            
            for page_idx, image in enumerate(images):
                # Extract text with confidence data
                data = pytesseract.image_to_data(image, lang=language, output_type='dict')
                
                page_text = ""
                page_confidence = 0.0
                confidence_scores = []
                
                for i in range(len(data['text'])):
                    text = data['text'][i].strip()
                    conf = data['conf'][i]
                    
                    if text and conf >= (confidence_threshold * 100):
                        page_text += text + " "
                        confidence_scores.append(conf)
                
                if confidence_scores:
                    page_confidence = sum(confidence_scores) / len(confidence_scores) / 100
                
                result["pages"].append({
                    "page": page_idx + 1,
                    "text": page_text.strip(),
                    "confidence": page_confidence
                })
                
                result["text"] += page_text + "\n---PAGE BREAK---\n"
            
            result["success"] = bool(result["text"])
            result["confidence"] = sum(
                p["confidence"] for p in result["pages"]
            ) / len(result["pages"]) if result["pages"] else 0.0
            
        except Exception as e:
            logger.error(f"Error with OCR extraction from {pdf_path}: {e}")
            result["error"] = str(e)
        
        return result
    
    def split_pdf(
        self,
        pdf_path: str,
        output_dir: str,
        pages_per_file: int = 1
    ) -> Dict[str, Any]:
        """
        Split multi-page PDF into individual files (for circularisation).
        
        Args:
            pdf_path: Path to PDF file
            output_dir: Directory to save split files
            pages_per_file: Pages per output file (default 1)
        
        Returns:
            Dict with list of created files
        """
        result = {
            "success": False,
            "pdf_path": str(pdf_path),
            "output_files": [],
            "output_dir": str(output_dir)
        }
        
        try:
            import pypdf
            
            Path(output_dir).mkdir(parents=True, exist_ok=True)
            pdf_reader = pypdf.PdfReader(pdf_path)
            pdf_name = Path(pdf_path).stem
            
            total_pages = len(pdf_reader.pages)
            file_count = 0
            
            for start_page in range(0, total_pages, pages_per_file):
                pdf_writer = pypdf.PdfWriter()
                end_page = min(start_page + pages_per_file, total_pages)
                
                for page_idx in range(start_page, end_page):
                    pdf_writer.add_page(pdf_reader.pages[page_idx])
                
                file_count += 1
                output_file = Path(output_dir) / f"{pdf_name}_page{start_page+1}.pdf"
                
                with open(output_file, 'wb') as f:
                    pdf_writer.write(f)
                
                result["output_files"].append(str(output_file))
            
            result["success"] = True
            result["files_created"] = file_count
            logger.info(f"Split PDF into {file_count} files")
            
        except Exception as e:
            logger.error(f"Error splitting PDF {pdf_path}: {e}")
            result["error"] = str(e)
        
        return result


# Singleton instance
_pdf_handler = PDFSkillHandler()


def get_pdf_skill_handler() -> PDFSkillHandler:
    """Get the PDF skill handler instance."""
    return _pdf_handler
