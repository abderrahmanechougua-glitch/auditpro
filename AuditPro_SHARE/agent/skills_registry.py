"""
Skills Registry — Maps Claude Code skills to AuditPro modules.

This module provides a unified interface to leverage Claude Code skills
to enhance the capabilities of each audit module.

Integrated Skills:
- pdf: PDF extraction, OCR, table parsing
- excel-analysis: Data analysis, pivot tables, validation
- data-visualization: Charts, graphs, report generation
- docx: Word document generation, formatting, mail merge
"""
from __future__ import annotations
from typing import Callable, Optional, Any
from pathlib import Path
from dataclasses import dataclass
import json


@dataclass
class SkillCapability:
    """Describes a skill capability and how it enhances a module."""
    skill_name: str
    capability: str
    description: str
    example_use: str


# ─── Skill Capability Mappings ──────────────────────────────────────

SKILL_ENHANCEMENTS = {
    "tva": [
        SkillCapability(
            skill_name="pdf",
            capability="pdf_extract_tables",
            description="Extract TVA declaration tables with OCR fallback",
            example_use="Extract SA, SG, D8, bases, rates from TVA PDF"
        ),
        SkillCapability(
            skill_name="excel-analysis",
            capability="excel_validate",
            description="Validate extracted TVA data against accounting rules",
            example_use="Cross-check TVA due = collectée - déductible"
        ),
        SkillCapability(
            skill_name="data-visualization",
            capability="generate_chart",
            description="Visualize TVA trends by month/rate",
            example_use="Bar chart: TVA by rate category (20%, 14%, 10%, 7%)"
        ),
    ],
    "cnss": [
        SkillCapability(
            skill_name="pdf",
            capability="pdf_extract_forms",
            description="Extract CNSS bordereau with form field recognition",
            example_use="Extract RG (AF, PS, TFP) + AMO from 2-page PDF"
        ),
        SkillCapability(
            skill_name="excel-analysis",
            capability="excel_validate",
            description="Validate CNSS rates and totals",
            example_use="Cross-check: AF 6.40% + PS 13.46% + TFP 1.60% + AMO"
        ),
    ],
    "extraction_factures": [
        SkillCapability(
            skill_name="pdf",
            capability="pdf_ocr_extract",
            description="Enhanced OCR for invoice extraction",
            example_use="Extract invoice number, date, amount, VAT from PDF"
        ),
        SkillCapability(
            skill_name="excel-analysis",
            capability="excel_create_report",
            description="Organize extracted invoices into structured Excel report",
            example_use="Create invoice register with validation formulas"
        ),
        SkillCapability(
            skill_name="data-visualization",
            capability="generate_summary_chart",
            description="Visualize invoice distribution by vendor/amount",
            example_use="Pie chart: expenses by supplier category"
        ),
    ],
    "lettrage": [
        SkillCapability(
            skill_name="excel-analysis",
            capability="excel_pivot_table",
            description="Create pivot tables for GL account analysis",
            example_use="Pivot: accounts x status (lettré/partiel/non-lettré)"
        ),
        SkillCapability(
            skill_name="xlsx",
            capability="xlsx_profile",
            description="Profile workbook structure and data quality",
            example_use="Inspect sheets, row counts, and duplicate ratios in GL workbook"
        ),
        SkillCapability(
            skill_name="data-visualization",
            capability="generate_chart",
            description="Visualize reconciliation status by account",
            example_use="Stacked bar: matched vs. unmatched by account"
        ),
    ],
    "retraitement": [
        SkillCapability(
            skill_name="excel-analysis",
            capability="excel_format_detect",
            description="Auto-detect accounting file format (Sage, D365, AX)",
            example_use="Normalize multi-format GL into standard format"
        ),
        SkillCapability(
            skill_name="excel-analysis",
            capability="excel_validate",
            description="Validate GL integrity after reprocessing",
            example_use="Check: sum(debits) = sum(credits)"
        ),
        SkillCapability(
            skill_name="xlsx",
            capability="xlsx_normalize",
            description="Normalize and clean spreadsheet sheets",
            example_use="Drop empty rows/columns and standardize headers in retraitement output"
        ),
        SkillCapability(
            skill_name="data-visualization",
            capability="generate_chart",
            description="Show GL balance distribution before/after",
            example_use="Comparison chart: original vs. reprocessed GL"
        ),
    ],
    "srm_generator": [
        SkillCapability(
            skill_name="data-visualization",
            capability="generate_professional_chart",
            description="Create publication-quality audit charts",
            example_use="Line chart: financial performance N vs N-1"
        ),
        SkillCapability(
            skill_name="docx",
            capability="docx_create_formatted",
            description="Generate SRM with professional formatting",
            example_use="SRM with embedded charts, formatted tables, annotations"
        ),
        SkillCapability(
            skill_name="excel-analysis",
            capability="excel_analyze_data",
            description="Analyze SRM data for anomalies and patterns",
            example_use="Detect unusual variations between N and N-1"
        ),
        SkillCapability(
            skill_name="xlsx",
            capability="xlsx_profile",
            description="Profile SRM workbook quality before generation",
            example_use="Verify SRM sheet dimensions and missing-cell hotspots"
        ),
    ],
    "circularisation": [
        SkillCapability(
            skill_name="docx",
            capability="docx_mail_merge",
            description="Generate personalized confirmation letters",
            example_use="Mail merge: letter template with tier-specific data"
        ),
        SkillCapability(
            skill_name="pdf",
            capability="pdf_split",
            description="Split multi-page PDF into individual files by tier",
            example_use="Split consolidated PDF into 1 file per customer/supplier"
        ),
        SkillCapability(
            skill_name="pdf",
            capability="pdf_extract_pages",
            description="Extract specific PDF pages for processing",
            example_use="Extract tier responses from signed PDF"
        ),
    ],
}


class SkillsRegistry:
    """Central registry for skill-module mappings and execution."""
    
    def __init__(self):
        self.skills_available = {}
        self.skill_handlers = {}
    
    def register_skill_handler(self, skill_name: str, handler: Callable):
        """Register a skill handler function."""
        self.skill_handlers[skill_name] = handler
    
    def get_enhancements_for_module(self, module_name: str) -> list[SkillCapability]:
        """Get list of available skill enhancements for a module."""
        return SKILL_ENHANCEMENTS.get(module_name, [])
    
    def get_all_modules_with_skills(self) -> dict[str, list[SkillCapability]]:
        """Get all module-skill mappings."""
        return SKILL_ENHANCEMENTS
    
    def execute_skill(
        self,
        module_name: str,
        skill_name: str,
        capability: str,
        **kwargs
    ) -> Any:
        """
        Execute a skill capability for a module.
        
        Args:
            module_name: Module to enhance (e.g., "tva", "lettrage")
            skill_name: Skill to use (e.g., "pdf", "excel-analysis")
            capability: Specific capability (e.g., "pdf_extract_tables")
            **kwargs: Skill-specific parameters
        
        Returns:
            Result from skill execution
        """
        # Verify the mapping exists
        enhancements = self.get_enhancements_for_module(module_name)
        if not any(e.skill_name == skill_name for e in enhancements):
            raise ValueError(
                f"Skill '{skill_name}' not registered for module '{module_name}'"
            )
        
        # Look for registered handler
        handler_key = f"{skill_name}:{capability}"
        if handler_key not in self.skill_handlers:
            raise NotImplementedError(
                f"Handler for '{handler_key}' not yet implemented"
            )
        
        # Execute the skill
        return self.skill_handlers[handler_key](**kwargs)
    
    def get_module_skill_summary(self, module_name: str) -> dict:
        """Get a summary of skills available for a module."""
        enhancements = self.get_enhancements_for_module(module_name)
        return {
            "module": module_name,
            "total_skills": len(enhancements),
            "skills": [
                {
                    "name": e.skill_name,
                    "capability": e.capability,
                    "description": e.description,
                }
                for e in enhancements
            ]
        }
    
    def export_mapping_as_json(self, output_path: str):
        """Export skill mappings to JSON for documentation."""
        mapping = {}
        for module, capabilities in SKILL_ENHANCEMENTS.items():
            mapping[module] = [
                {
                    "skill": cap.skill_name,
                    "capability": cap.capability,
                    "description": cap.description,
                    "example": cap.example_use,
                }
                for cap in capabilities
            ]
        
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(mapping, f, indent=2, ensure_ascii=False)


# Global registry instance
_registry = SkillsRegistry()


def get_skills_registry() -> SkillsRegistry:
    """Get the global skills registry."""
    return _registry
