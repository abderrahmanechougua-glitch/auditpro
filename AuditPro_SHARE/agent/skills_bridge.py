"""
Skills Integration Bridge — Unified interface for Claude Code skills in AuditPro.

This module provides:
1. Unified skill access for modules
2. Tool wrappers for agent integration
3. Configuration and skill initialization
"""
from __future__ import annotations
from typing import Any, Callable, Dict, Optional
import logging
from pathlib import Path

logger = logging.getLogger(__name__)

# Import skill handlers
from agent.skills_registry import SkillsRegistry, get_skills_registry
from agent.skills_pdf import get_pdf_skill_handler
from agent.skills_excel import get_excel_skill_handler
from agent.skills_xlsx import get_xlsx_skill_handler
from agent.skills_visualization import get_visualization_skill_handler
from agent.skills_docx import get_docx_skill_handler


class SkillsIntegrationBridge:
    """
    Unified interface to skills for audit modules.
    
    Provides a single API for modules to access skill capabilities
    without worrying about implementation details.
    """
    
    def __init__(self):
        self.registry = get_skills_registry()
        self.pdf = get_pdf_skill_handler()
        self.excel = get_excel_skill_handler()
        self.xlsx = get_xlsx_skill_handler()
        self.visualization = get_visualization_skill_handler()
        self.docx = get_docx_skill_handler()
        
        self._register_handlers()
        logger.info("Skills integration bridge initialized")
    
    def _register_handlers(self):
        """Register all skill handlers with the registry."""
        
        # PDF Skill Handlers
        self.registry.register_skill_handler(
            "pdf:extract_tables",
            lambda **kw: self.pdf.extract_tables(**kw)
        )
        self.registry.register_skill_handler(
            "pdf:extract_forms",
            lambda **kw: self.pdf.extract_forms(**kw)
        )
        self.registry.register_skill_handler(
            "pdf:ocr_extract",
            lambda **kw: self.pdf.ocr_extract(**kw)
        )
        self.registry.register_skill_handler(
            "pdf:split",
            lambda **kw: self.pdf.split_pdf(**kw)
        )
        
        # Excel Skill Handlers
        self.registry.register_skill_handler(
            "excel:validate",
            lambda **kw: self.excel.validate_data(**kw)
        )
        self.registry.register_skill_handler(
            "excel:pivot_table",
            lambda **kw: self.excel.create_pivot_table(**kw)
        )
        self.registry.register_skill_handler(
            "excel:detect_format",
            lambda **kw: self.excel.detect_format(**kw)
        )
        self.registry.register_skill_handler(
            "excel:normalize",
            lambda **kw: self.excel.normalize_gl(**kw)
        )
        self.registry.register_skill_handler(
            "excel:detect_anomalies",
            lambda **kw: self.excel.detect_anomalies(**kw)
        )

        # XLSX Skill Handlers
        self.registry.register_skill_handler(
            "xlsx:profile",
            lambda **kw: self.xlsx.profile_workbook(**kw)
        )
        self.registry.register_skill_handler(
            "xlsx:normalize",
            lambda **kw: self.xlsx.normalize_sheet(**kw)
        )
        
        # Visualization Handlers
        self.registry.register_skill_handler(
            "viz:bar_chart",
            lambda **kw: self.visualization.create_bar_chart(**kw)
        )
        self.registry.register_skill_handler(
            "viz:line_chart",
            lambda **kw: self.visualization.create_line_chart(**kw)
        )
        self.registry.register_skill_handler(
            "viz:pie_chart",
            lambda **kw: self.visualization.create_pie_chart(**kw)
        )
        self.registry.register_skill_handler(
            "viz:comparison",
            lambda **kw: self.visualization.create_comparison_chart(**kw)
        )
        self.registry.register_skill_handler(
            "viz:summary_stats",
            lambda **kw: self.visualization.generate_summary_stats(**kw)
        )
        
        # DOCX Handlers
        self.registry.register_skill_handler(
            "docx:create",
            lambda **kw: self.docx.create_document(**kw)
        )
        self.registry.register_skill_handler(
            "docx:add_table",
            lambda **kw: self.docx.add_table(**kw)
        )
        self.registry.register_skill_handler(
            "docx:add_image",
            lambda **kw: self.docx.add_image(**kw)
        )
        self.registry.register_skill_handler(
            "docx:mail_merge",
            lambda **kw: self.docx.mail_merge(**kw)
        )
        self.registry.register_skill_handler(
            "docx:add_comment",
            lambda **kw: self.docx.add_comment(**kw)
        )
    
    def extract_pdf_tables(self, pdf_path: str, **kwargs) -> Dict[str, Any]:
        """Extract tables from PDF."""
        return self.pdf.extract_tables(pdf_path, **kwargs)
    
    def extract_pdf_forms(self, pdf_path: str, **kwargs) -> Dict[str, Any]:
        """Extract form fields from PDF."""
        return self.pdf.extract_forms(pdf_path, **kwargs)
    
    def ocr_extract_from_pdf(self, pdf_path: str, **kwargs) -> Dict[str, Any]:
        """Extract text via OCR from PDF."""
        return self.pdf.ocr_extract(pdf_path, **kwargs)
    
    def validate_excel_data(self, df, rules=None) -> Dict[str, Any]:
        """Validate Excel data."""
        return self.excel.validate_data(df, rules)
    
    def detect_accounting_format(self, file_path: str) -> Dict[str, Any]:
        """Auto-detect accounting file format."""
        return self.excel.detect_format(file_path)
    
    def normalize_general_ledger(self, file_path: str, **kwargs) -> Dict[str, Any]:
        """Normalize GL from various formats."""
        return self.excel.normalize_gl(file_path, **kwargs)
    
    def detect_data_anomalies(self, df, **kwargs) -> Dict[str, Any]:
        """Detect statistical anomalies in data."""
        return self.excel.detect_anomalies(df, **kwargs)

    def audit_accounting_data(self, df, key_columns=None) -> Dict[str, Any]:
        """Run accounting-focused checks on a ledger-like dataset."""
        return self.excel.analyze_accounting_data(df, key_columns=key_columns)

    def audit_srm_data(self, df) -> Dict[str, Any]:
        """Run SRM-focused checks on an SRM source table."""
        return self.excel.analyze_srm_data(df)

    def profile_xlsx(self, file_path: str) -> Dict[str, Any]:
        """Return workbook-level profile for an XLSX file."""
        return self.xlsx.profile_workbook(file_path)

    def normalize_xlsx_sheet(
        self,
        file_path: str,
        output_path: str,
        sheet_name: Optional[str] = None,
        header_row: int = 0,
    ) -> Dict[str, Any]:
        """Normalize one workbook sheet and export it as XLSX."""
        return self.xlsx.normalize_sheet(file_path, output_path, sheet_name=sheet_name, header_row=header_row)
    
    def create_audit_chart(
        self,
        chart_type: str,
        data: Any,
        title: str,
        output_path: Optional[str] = None,
        **kwargs
    ) -> Dict[str, Any]:
        """Create visualization chart."""
        if chart_type == "bar":
            return self.visualization.create_bar_chart(data, title, output_path=output_path, **kwargs)
        elif chart_type == "line":
            return self.visualization.create_line_chart(data, title, output_path=output_path, **kwargs)
        elif chart_type == "pie":
            return self.visualization.create_pie_chart(data, title, output_path=output_path, **kwargs)
        elif chart_type == "comparison":
            return self.visualization.create_comparison_chart(data[0], data[1], title, output_path=output_path, **kwargs)
        else:
            return {"error": f"Unknown chart type: {chart_type}"}
    
    def create_word_document(self, title: str, output_path: str, **kwargs) -> Dict[str, Any]:
        """Create Word document."""
        return self.docx.create_document(title, output_path, **kwargs)
    
    def generate_personalized_letters(
        self,
        template_path: str,
        merge_data: list,
        output_dir: str,
        **kwargs
    ) -> Dict[str, Any]:
        """Generate personalized documents via mail merge."""
        return self.docx.mail_merge(template_path, merge_data, output_dir, **kwargs)
    
    def get_module_enhancements(self, module_name: str) -> Dict[str, Any]:
        """Get available skill enhancements for a module."""
        return self.registry.get_module_skill_summary(module_name)
    
    def get_all_enhancements(self) -> Dict[str, list]:
        """Get all module-skill mappings."""
        return self.registry.get_all_modules_with_skills()


# Global bridge instance
_bridge = None


def initialize_skills_bridge() -> SkillsIntegrationBridge:
    """Initialize the skills integration bridge."""
    global _bridge
    if _bridge is None:
        _bridge = SkillsIntegrationBridge()
    return _bridge


def get_skills_bridge() -> SkillsIntegrationBridge:
    """Get the skills integration bridge instance."""
    global _bridge
    if _bridge is None:
        _bridge = SkillsIntegrationBridge()
    return _bridge


# Tool wrappers for agent integration
def skill_tool_pdf_extract(
    pdf_path: str,
    extraction_type: str = "tables",
    **kwargs
) -> Dict[str, Any]:
    """
    Tool wrapper: Extract data from PDF.
    
    Args:
        pdf_path: Path to PDF file
        extraction_type: "tables", "forms", or "ocr"
        **kwargs: Additional parameters for extraction
    """
    bridge = get_skills_bridge()
    
    if extraction_type == "tables":
        return bridge.extract_pdf_tables(pdf_path, **kwargs)
    elif extraction_type == "forms":
        return bridge.extract_pdf_forms(pdf_path, **kwargs)
    elif extraction_type == "ocr":
        return bridge.ocr_extract_from_pdf(pdf_path, **kwargs)
    else:
        return {"error": f"Unknown extraction type: {extraction_type}"}


def skill_tool_excel_analyze(
    file_path: str,
    analysis_type: str = "validate",
    **kwargs
) -> Dict[str, Any]:
    """
    Tool wrapper: Analyze Excel file.
    
    Args:
        file_path: Path to Excel file
        analysis_type: "validate", "format_detect", or "normalize"
        **kwargs: Additional parameters
    """
    bridge = get_skills_bridge()
    
    if analysis_type == "validate":
        import pandas as pd
        df = pd.read_excel(file_path)
        return bridge.validate_excel_data(df, **kwargs)
    elif analysis_type == "format_detect":
        return bridge.detect_accounting_format(file_path)
    elif analysis_type == "normalize":
        return bridge.normalize_general_ledger(file_path, **kwargs)
    else:
        return {"error": f"Unknown analysis type: {analysis_type}"}


def skill_tool_generate_chart(
    chart_type: str,
    data_source: str | Any,
    title: str,
    output_path: Optional[str] = None,
    **kwargs
) -> Dict[str, Any]:
    """
    Tool wrapper: Generate audit chart.
    
    Args:
        chart_type: "bar", "line", "pie", "comparison"
        data_source: File path or data object
        title: Chart title
        output_path: Output image path
    """
    bridge = get_skills_bridge()
    
    # Load data if file path provided
    if isinstance(data_source, str) and Path(data_source).exists():
        import pandas as pd
        data = pd.read_excel(data_source) if data_source.endswith('.xlsx') else data_source
    else:
        data = data_source
    
    return bridge.create_audit_chart(chart_type, data, title, output_path, **kwargs)


def skill_tool_generate_srm(
    srm_data: str | Any,
    output_path: str,
    include_charts: bool = True,
    **kwargs
) -> Dict[str, Any]:
    """
    Tool wrapper: Generate SRM document.
    
    Args:
        srm_data: SRM data (file path or dict)
        output_path: Output document path
        include_charts: Include visualization charts
    """
    bridge = get_skills_bridge()
    
    # Create SRM document
    result = bridge.create_word_document("SRM (Summary of Recorded Misstatements)", output_path, **kwargs)
    
    return result


def skill_tool_xlsx_process(
    file_path: str,
    operation: str = "profile",
    output_path: Optional[str] = None,
    **kwargs
) -> Dict[str, Any]:
    """
    Tool wrapper: process XLSX workbook/sheet.

    Args:
        file_path: Path to XLSX file
        operation: "profile" or "normalize"
        output_path: Output XLSX path (required for normalize)
    """
    bridge = get_skills_bridge()

    if operation == "profile":
        return bridge.profile_xlsx(file_path)
    if operation == "normalize":
        if not output_path:
            return {"error": "output_path is required for operation='normalize'"}
        return bridge.normalize_xlsx_sheet(file_path, output_path, **kwargs)
    return {"error": f"Unknown XLSX operation: {operation}"}
