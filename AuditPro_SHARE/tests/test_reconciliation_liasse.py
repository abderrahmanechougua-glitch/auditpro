"""
Test suite for reconciliation_bg_liasse module.

Tests cover:
- Format detection (_detect_liasse_format)
- Data extraction (_extract_liasse)
- Edge cases and error handling
- Data integrity and accuracy
"""

import pytest
import sys
import pandas as pd
from pathlib import Path
from typing import Tuple
import tempfile
import os

# Add internal modules path
sys.path.insert(0, r'AuditPro_SHARE\dist\AuditPro\_internal')

try:
    from modules.reconciliation_bg_liasse.reconciliation import (
        _detect_liasse_format,
        _extract_liasse
    )
except ImportError as e:
    pytest.skip(f"Cannot import reconciliation module: {e}", allow_module_level=True)


class TestFormatDetection:
    """Tests for _detect_liasse_format function."""

    def test_detect_xlsx_format(self, sample_xlsx_path):
        """Test detection of Excel 2007+ (.xlsx) format."""
        fmt = _detect_liasse_format(sample_xlsx_path)
        assert fmt in ['xlsx', 'excel', 'XLSX', 'EXCEL'], f"Expected xlsx format, got {fmt}"

    def test_detect_xls_format(self, sample_xls_path):
        """Test detection of Excel 97-2003 (.xls) format."""
        fmt = _detect_liasse_format(sample_xls_path)
        assert fmt in ['xls', 'excel', 'XLS', 'EXCEL'], f"Expected xls format, got {fmt}"

    def test_detect_pdf_format(self, sample_pdf_path):
        """Test detection of PDF format."""
        fmt = _detect_liasse_format(sample_pdf_path)
        assert fmt in ['pdf', 'PDF'], f"Expected pdf format, got {fmt}"

    def test_detect_invalid_format(self):
        """Test detection with invalid file raises appropriate error."""
        with pytest.raises((FileNotFoundError, ValueError, OSError)):
            _detect_liasse_format('/nonexistent/file.txt')

    def test_detect_with_none_path(self):
        """Test detection with None path raises error."""
        with pytest.raises((TypeError, AttributeError, ValueError)):
            _detect_liasse_format(None)

    def test_detect_with_empty_path(self):
        """Test detection with empty path string."""
        with pytest.raises((FileNotFoundError, ValueError, OSError)):
            _detect_liasse_format('')


class TestLiasseExtraction:
    """Tests for _extract_liasse function."""

    def test_extract_xlsx_returns_dataframe(self, sample_xlsx_path):
        """Test that extract_liasse returns a DataFrame for .xlsx files."""
        result = _extract_liasse(sample_xlsx_path)
        assert isinstance(result, pd.DataFrame), f"Expected DataFrame, got {type(result)}"

    def test_extract_xls_returns_dataframe(self, sample_xls_path):
        """Test that extract_liasse returns a DataFrame for .xls files."""
        result = _extract_liasse(sample_xls_path)
        assert isinstance(result, pd.DataFrame), f"Expected DataFrame, got {type(result)}"

    def test_extract_pdf_returns_dataframe(self, sample_pdf_path):
        """Test that extract_liasse returns a DataFrame for PDF files."""
        result = _extract_liasse(sample_pdf_path)
        assert isinstance(result, pd.DataFrame), f"Expected DataFrame, got {type(result)}"

    def test_extracted_dataframe_has_required_columns(self, sample_xlsx_path):
        """Test that extracted DataFrame contains required columns."""
        df = _extract_liasse(sample_xlsx_path)
        
        # Expected columns for liasse data
        expected_cols = ['Rubrique', 'Montant_Liasse']
        for col in expected_cols:
            assert col in df.columns, f"Missing expected column: {col}"

    def test_extracted_dataframe_not_empty(self, sample_xlsx_path):
        """Test that extracted data is not empty."""
        df = _extract_liasse(sample_xlsx_path)
        assert len(df) > 0, "Extracted DataFrame is empty"

    def test_montant_column_is_numeric(self, sample_xlsx_path):
        """Test that Montant_Liasse column contains numeric values."""
        df = _extract_liasse(sample_xlsx_path)
        
        # Check that values are numeric (convert to numeric, coerce errors)
        numeric_vals = pd.to_numeric(df['Montant_Liasse'], errors='coerce')
        assert numeric_vals.notna().sum() > 0, "No numeric values in Montant_Liasse"

    def test_extract_with_nonexistent_file(self):
        """Test extraction with nonexistent file raises error."""
        with pytest.raises((FileNotFoundError, ValueError, OSError)):
            _extract_liasse('/nonexistent/path/file.xlsx')

    def test_extract_with_corrupted_file(self, corrupted_xlsx_path):
        """Test extraction with corrupted Excel file."""
        with pytest.raises((ValueError, OSError, Exception)):
            _extract_liasse(corrupted_xlsx_path)


class TestDataIntegrity:
    """Tests for data integrity and accuracy."""

    def test_no_duplicate_rubriques(self, sample_xlsx_path):
        """Test that there are no duplicate line items in extracted data."""
        df = _extract_liasse(sample_xlsx_path)
        
        duplicates = df[df.duplicated(subset=['Rubrique'], keep=False)]
        assert len(duplicates) == 0, f"Found duplicate rubriques: {duplicates['Rubrique'].tolist()}"

    def test_rubriques_not_empty_or_nan(self, sample_xlsx_path):
        """Test that all rubriques are non-empty and not NaN."""
        df = _extract_liasse(sample_xlsx_path)
        
        empty_rubriques = df[df['Rubrique'].isna() | (df['Rubrique'] == '')]
        assert len(empty_rubriques) == 0, f"Found {len(empty_rubriques)} empty rubriques"

    def test_montant_not_null(self, sample_xlsx_path):
        """Test that montant values are not null."""
        df = _extract_liasse(sample_xlsx_path)
        
        null_montants = df[df['Montant_Liasse'].isna()]
        assert len(null_montants) == 0, f"Found {len(null_montants)} null montant values"

    def test_extract_consistency_multiple_reads(self, sample_xlsx_path):
        """Test that multiple extractions of same file produce consistent results."""
        df1 = _extract_liasse(sample_xlsx_path)
        df2 = _extract_liasse(sample_xlsx_path)
        
        # Sort by Rubrique for consistent comparison
        df1_sorted = df1.sort_values('Rubrique').reset_index(drop=True)
        df2_sorted = df2.sort_values('Rubrique').reset_index(drop=True)
        
        pd.testing.assert_frame_equal(df1_sorted, df2_sorted, check_exact=True)


# ============================================================================
# Fixtures for test data
# ============================================================================

@pytest.fixture
def sample_xlsx_path() -> str:
    """Path to sample XLSX liasse file."""
    # Try multiple potential locations
    candidates = [
        r'C:\Users\Abderrahmane.CHOUGUA\OneDrive - Fidaroc Grant Thornton\Fichiers de Omar ESSBAI - Issal Madina\1- Back up\Controle des comptes\Finalisation\ISSAL MADINA - Liasse Comptable 2024.xlsx',
        r'tests\fixtures\sample_liasse.xlsx',
        r'AuditPro_SHARE\fixtures\sample_liasse.xlsx',
    ]
    
    for path in candidates:
        if Path(path).exists():
            return path
    
    pytest.skip("Sample XLSX file not found")


@pytest.fixture
def sample_xls_path() -> str:
    """Path to sample XLS liasse file."""
    candidates = [
        r'C:\Users\Abderrahmane.CHOUGUA\OneDrive - Fidaroc Grant Thornton\Fichiers de Yassine MOUKRIM - GIM - CI - 2025\Finalisation\PBC\Liasse_Fiscale_GLOBAL-INTERNATIONAL-MOTORS_2025 (2).xls',
        r'tests\fixtures\sample_liasse.xls',
        r'AuditPro_SHARE\fixtures\sample_liasse.xls',
    ]
    
    for path in candidates:
        if Path(path).exists():
            return path
    
    pytest.skip("Sample XLS file not found")


@pytest.fixture
def sample_pdf_path() -> str:
    """Path to sample PDF liasse file."""
    candidates = [
        r'C:\Users\Abderrahmane.CHOUGUA\OneDrive - Fidaroc Grant Thornton\Fichiers de Yassine MOUKRIM - GIM - CI - 2025\Controle des comptes GIM\PBC\TR_ Acomptes IS 2025\Liasse fiscale SIMPL IS 2024.pdf',
        r'tests\fixtures\sample_liasse.pdf',
        r'AuditPro_SHARE\fixtures\sample_liasse.pdf',
    ]
    
    for path in candidates:
        if Path(path).exists():
            return path
    
    pytest.skip("Sample PDF file not found")


@pytest.fixture
def corrupted_xlsx_path() -> str:
    """Path to corrupted XLSX file for error testing."""
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
        # Write garbage data
        f.write(b'This is not a valid Excel file')
        temp_path = f.name
    
    yield temp_path
    
    # Cleanup
    try:
        os.unlink(temp_path)
    except:
        pass
