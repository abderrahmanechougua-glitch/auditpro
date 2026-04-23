---
name: auditpro-module-development
description: "Develop new AuditPro audit modules with full compliance validation, testing, and performance optimization. Use when: creating a new audit module, refactoring an existing module, or adding features to reconciliation/extraction workflows. Integrates BaseModule inheritance rules, pytest patterns, performance profiling, and pre-merge compliance checks."
---

# AuditPro Module Development Workflow

## Purpose

This skill guides developers through a structured, quality-assured workflow for creating and enhancing AuditPro audit modules. It ensures:
- All modules inherit from `BaseModule` (architecture consistency)
- Comprehensive test coverage (pytest with fixtures)
- Performance profiling and optimization (cProfile)
- Compliance with standing OCR/path rules
- Pre-merge validation against AuditPro compliance checklist

## When to Use This Skill

- **Creating a new audit module** (e.g., TVA, Circularisation, Reconciliation variant)
- **Refactoring an existing module** for performance or maintainability
- **Adding data extraction logic** with OCR, PDF, or Excel dependencies
- **Implementing business logic** that processes liasse, balance sheets, or tax data
- **Before committing/merging** module changes to production

## Quick Workflow (5 Steps)

```
1. SCAFFOLD    → Create module folder with BaseModule inheritance
2. IMPLEMENT   → Write module logic following OCR/path patterns
3. TEST        → Build comprehensive pytest suite (19+ tests)
4. PROFILE     → Identify and optimize performance bottlenecks
5. REVIEW      → Validate compliance before merge
```

---

## Step-by-Step Guide

### Step 1: Scaffold the Module

**Objective:** Create module structure with correct inheritance.

**Action:** Create `modules/<module_name>/module.py`:

```python
"""
<Module Name> audit module.

Inherits from BaseModule to ensure:
- Consistent lifecycle hooks (setup, execute, cleanup)
- Shared validation and error handling
- Integration with workspace and module registry
"""

import sys
from pathlib import Path

# Add internal modules path
sys.path.insert(0, r'AuditPro_SHARE\dist\AuditPro\_internal')

try:
    from core.base import BaseModule
except ImportError as e:
    raise ImportError(f"Cannot import BaseModule: {e}")


class <ModuleName>Module(BaseModule):
    """<Module description> audit module."""

    def __init__(self, *args, **kwargs):
        """Initialize <module name> module."""
        super().__init__(*args, **kwargs)
        self.module_name = "<module_name>"
        self.description = "<Human-readable description>"

    def setup(self) -> None:
        """Set up module resources."""
        super().setup()
        # Initialize OCR paths via core/ocr_paths.py (NOT hardcoded)
        # from core.ocr_paths import _load_ocr_paths
        # self.ocr_paths = _load_ocr_paths()

    def execute(self, inputs: dict) -> dict:
        """Execute the module logic.
        
        Args:
            inputs: Module input parameters from UI
            
        Returns:
            Dictionary with results and metadata
        """
        raise NotImplementedError("Subclasses must implement execute()")

    def cleanup(self) -> None:
        """Clean up module resources."""
        super().cleanup()


if __name__ == '__main__':
    # Test module instantiation
    module = <ModuleName>Module()
    print(f"✓ Module loaded: {module.module_name}")
```

**Compliance Check:**
- ✓ Imports from `core.base.BaseModule` (not custom classes)
- ✓ Inherits from `BaseModule`
- ✓ Implements `setup()`, `execute()`, `cleanup()`
- ✓ NO hardcoded paths or OCR parameters

---

### Step 2: Implement Module Logic

**Objective:** Write business logic following standing rules.

**Key Patterns:**

#### OCR/Path Handling (SELF-CONTAINED)
```python
from core.ocr_paths import _load_ocr_paths

def execute(self, inputs: dict) -> dict:
    """Execute with OCR paths resolved internally."""
    # ✓ CORRECT: OCR paths resolved in module, not exposed in UI
    ocr_paths = _load_ocr_paths()
    
    # Graceful degradation if OCR unavailable
    if not ocr_paths.get('tesseract_path'):
        print("WARNING: Tesseract OCR not available, skipping OCR step")
        # Continue without OCR, don't crash
    
    return {'status': 'success', 'rows': 100}

# ✗ WRONG: Never pass OCR paths as UI parameters
# ✗ WRONG: def execute(self, tesseract_path, poppler_path, ...)
```

#### Module Wrapper Rules
```python
# In module.py UI inputs definition:

# ✓ CORRECT: No mention of OCR paths
inputs = [
    {'name': 'file_path', 'type': 'file', 'label': 'Fichier Excel'},
    {'name': 'sheet_name', 'type': 'text', 'label': 'Nom feuille'},
]

# ✗ WRONG: Never expose OCR parameters
# inputs = [
#     {'name': 'tesseract_path', 'type': 'text', 'label': '...'},
# ]
```

#### UI Rules
```python
# In workspace.py or UI initialization:

# ✓ CORRECT: Keep output_frame hidden
self.output_frame.setVisible(False)

# ✗ WRONG: Never show output directory box
# self.output_frame.setVisible(True)
```

#### PDF Splitting (Circularisation)
```python
# In modules/circularisation/module.py _run_split():

# ✓ CORRECT: Use tiers name from data, with fallback
tiers_keywords = ['nom', 'name', 'tiers', 'denomination', 'dénomination', 'raison_sociale', 'intitule']
tiers_col = next((c for c in df.columns if c.lower() in tiers_keywords), df.columns[0])

# Clean pandas artifacts
filename = str(df[tiers_col].iloc[0]).replace('nan', 'tiers_001').replace('None', '')

# ✓ CORRECT: Clean naming with collision suffix
filename_safe = f"{filename}_{i:03d}.pdf"  # zero-padded

# ✗ WRONG: Never hardcode 'C:/Program Files'
# ✗ WRONG: Don't use substring-matching for PATH
```

---

### Step 3: Write Comprehensive Tests

**Objective:** Ensure module behavior and data integrity.

**Use Skill:** `python-testing-patterns`

**Create:** `tests/test_<module_name>.py` with:

```python
import pytest
import sys
from pathlib import Path

# Module import pattern
sys.path.insert(0, r'AuditPro_SHARE\dist\AuditPro\_internal')
from modules.<module_name>.module import <ModuleName>Module

class Test<ModuleName>Inheritance:
    """Verify BaseModule inheritance rules."""
    
    def test_inherits_from_base_module(self):
        """Module must inherit from BaseModule."""
        from core.base import BaseModule
        module = <ModuleName>Module()
        assert isinstance(module, BaseModule), "Must inherit from BaseModule"

class Test<ModuleName>Setup:
    """Test module lifecycle."""
    
    def test_setup_completes(self):
        """Setup hook must complete without error."""
        module = <ModuleName>Module()
        module.setup()  # Should not raise

    def test_cleanup_completes(self):
        """Cleanup hook must complete without error."""
        module = <ModuleName>Module()
        module.cleanup()  # Should not raise

class Test<ModuleName>Execution:
    """Test module business logic."""
    
    def test_execute_with_valid_inputs(self, sample_data):
        """Module executes with valid inputs."""
        module = <ModuleName>Module()
        result = module.execute(sample_data)
        
        assert isinstance(result, dict), "Must return dict"
        assert 'status' in result, "Must include status"

    def test_execute_with_invalid_inputs(self):
        """Module handles invalid inputs gracefully."""
        module = <ModuleName>Module()
        with pytest.raises((ValueError, TypeError, KeyError)):
            module.execute({})  # Empty inputs

class Test<ModuleName>DataIntegrity:
    """Test output data quality."""
    
    def test_output_no_nulls(self, sample_data):
        """Output should not contain unexpected nulls."""
        module = <ModuleName>Module()
        result = module.execute(sample_data)
        
        # Assert data quality based on module specifics
        assert result['rows'] > 0, "Output should have data"

@pytest.fixture
def sample_data():
    """Sample input data for testing."""
    return {
        'file_path': 'tests/fixtures/sample.xlsx',
        'sheet_name': 'Data',
    }
```

**Test Coverage Target:** 19+ tests across:
- ✓ Inheritance validation (BaseModule)
- ✓ Setup/execute/cleanup lifecycle
- ✓ Input validation (valid/invalid)
- ✓ Data integrity (no duplicates, non-null)
- ✓ Edge cases and error handling

**Run Tests:**
```bash
cd AuditPro_SHARE
pytest tests/test_<module_name>.py -v --tb=short
```

---

### Step 4: Profile and Optimize Performance

**Objective:** Identify and fix bottlenecks.

**Use Skill:** `python-performance-optimization`

**Create:** `profile_<module_name>.py` with cProfile analysis:

```python
import cProfile
import pstats
import io
import time
from modules.<module_name>.module import <ModuleName>Module

def profile_execution():
    """Profile module execution with sample data."""
    module = <ModuleName>Module()
    module.setup()
    
    # Profile with cProfile
    profiler = cProfile.Profile()
    profiler.enable()
    
    start_time = time.time()
    result = module.execute({'file_path': 'tests/fixtures/large_sample.xlsx'})
    elapsed = time.time() - start_time
    
    profiler.disable()
    
    # Print top functions by cumulative time
    s = io.StringIO()
    ps = pstats.Stats(profiler, stream=s).sort_stats('cumulative')
    ps.print_stats(10)
    
    print(f"Execution time: {elapsed*1000:.2f} ms")
    print(f"\nTop 10 functions:\n{s.getvalue()}")
    
    module.cleanup()

if __name__ == '__main__':
    profile_execution()
```

**Run Profiling:**
```bash
python profile_<module_name>.py
```

**Common Bottlenecks & Fixes:**

| Issue | Root Cause | Solution |
|-------|-----------|----------|
| Slow file reads | Reading entire file repeatedly | Cache file format, use memory-mapped files |
| DataFrame operations | Inefficient dtypes | Pre-specify dtypes, use categorical for large strings |
| OCR extraction | Line-by-line processing | Use batch processing, parallel workers |
| Memory spikes | Large intermediate DataFrames | Drop columns early, use chunking |

---

### Step 5: Compliance Review & Merge

**Objective:** Validate all standing rules before production.

**Use Skill:** `code-review-and-quality` + `auditpro-compliance`

**Pre-Merge Checklist:**

```markdown
## AuditPro Compliance Checklist

- [ ] **Inheritance**: Module inherits from BaseModule (not custom classes)
- [ ] **OCR Paths**: All OCR logic self-contained in module, NOT exposed as UI parameters
- [ ] **No Hardcoded Paths**: No `C:/Program Files`, `C:/Users`, Windows-specific paths
- [ ] **UI Rules**: `output_frame` never set to `setVisible(True)`
- [ ] **Graceful Degradation**: Missing Tesseract/Poppler → warns + skips, doesn't crash
- [ ] **PDF Naming**: Split files named by tiers (not generic names)
- [ ] **Imports**: All imports at module level (not inside functions)
- [ ] **Dependencies**: Required tools bundled in `vendor/`, no user install prompts
- [ ] **Tests**: 19+ tests, all passing locally
- [ ] **Performance**: Profile completed, no obvious bottlenecks
- [ ] **Help Text**: No mention of "Tesseract OCR installé" as prerequisite
- [ ] **Error Messages**: Clear, actionable French/English error messages
```

**Code Review Prompts:**

```
@copilot-code-review-and-quality
Review this module for:
1. BaseModule inheritance correctness
2. OCR/path handling following core/ocr_paths.py
3. No hardcoded paths or user prompts
4. Test coverage (minimum 19 tests)
5. Any architectural drift from standing rules
```

---

## Templates & Assets

### Template: Module Scaffold
Location: `.github/templates/module_scaffold.py`

### Template: Test Suite
Location: `.github/templates/test_module_template.py`

### Template: Performance Profile
Location: `.github/templates/profile_template.py`

### Standing Rules Reference
Location: `auditpro-compliance` skill (`c:\Users\...\skills\auditpro-compliance\SKILL.md`)

---

## Common Mistakes to Avoid

| ❌ WRONG | ✓ CORRECT | Why |
|---------|-----------|-----|
| `from base_module import BaseModule` | `from core.base import BaseModule` | Single source of truth for lifecycle |
| Module exposes `tesseract_path` parameter in UI | OCR paths resolved internally in `execute()` | Simplifies UI, ensures self-contained processing |
| `output_frame.setVisible(True)` | `output_frame.setVisible(False)` | Compliance rule: hide output directory box |
| `import glob` inside function | `import glob` at module level | Faster imports, cleaner code |
| No test coverage | 19+ tests (format/input/data integrity) | Catch regressions early |
| No performance profiling | cProfile analysis + optimization recommendations | Identify 10x improvements |
| Hardcoded `C:\Program Files\Tesseract` | Use `core/ocr_paths.py` | Platform-agnostic, portable |
| Crash if OCR missing | Warn + continue without OCR | Graceful degradation |

---

## Related Skills

- **python-testing-patterns** — Write pytest fixtures, fixtures, parameterization
- **python-performance-optimization** — cProfile, memory profiling, optimization patterns
- **code-review-and-quality** — Multi-axis code review, architecture validation
- **auditpro-compliance** — Standing rules checker (UI, OCR, module wrappers, PDF naming)

---

## Example Prompts

Once this skill is installed, use these prompts:

### Create a New Module
```
@auditpro-module-development Create new TVA reconciliation module
Inherit from BaseModule, add Excel/PDF extraction, write 20+ tests, profile performance
```

### Refactor Existing Module
```
@auditpro-module-development Refactor circularisation module
- Keep BaseModule inheritance
- Optimize PDF split naming logic
- Add 15 tests for edge cases
- Profile and reduce memory usage
```

### Add Feature to Module
```
@auditpro-module-development Add CSV export feature to reconciliation module
- Don't expose OCR paths in UI
- Keep output_frame hidden
- Add 8+ tests for export scenarios
```

---

## Lifecycle Hooks

This skill integrates with:
- **Pre-Commit**: Runs AuditPro compliance checker
- **Pre-Merge**: Runs test suite + code review
- **Post-Merge**: Logs module changes to activity log

---

## FAQ

**Q: Can I skip the profiling step?**  
A: No. Performance profiling identifies 10x improvements early. Always profile new extraction/processing logic.

**Q: What if I don't inherit from BaseModule?**  
A: Module won't integrate with workspace lifecycle, won't have shared error handling, won't pass compliance review. Always use BaseModule.

**Q: How many tests are enough?**  
A: Minimum 19 tests (format detection, extraction, integrity, edge cases). More for complex modules.

**Q: What if OCR library isn't available?**  
A: Don't crash. Warn user, skip OCR step, continue processing. This is graceful degradation.

**Q: Can I hardcode paths for debugging?**  
A: No. Always use `core/ocr_paths.py`. If debugging, use temp env vars instead.

**Q: What's the difference between `module.py` and the actual script (`factextv19.py` etc.)?**  
A: `module.py` = wrapper (UI, lifecycle, BaseModule). Script = business logic (extraction, transformation). OCR paths used only in script, resolved via `core/ocr_paths.py`.
