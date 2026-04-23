# AuditPro Agent Guide

## Workspace Map

- `LANCER.bat` at the repo root delegates to `AuditPro_SHARE/LANCER.bat`.
- `AuditPro_SHARE/dist/AuditPro/_internal/` contains the runtime code that the packaged app loads in this workspace.
- `AuditPro_Agent/AuditPro/` contains the PyInstaller build inputs, including `AuditPro.spec`.
- `AuditPro_SHARE/tests/`, `AuditPro_SHARE/test_liasse.py`, and `AuditPro_SHARE/profile_liasse_performance.py` are the main validation and profiling entry points.

## Preferred Commands

- Launch the app from the repo root with `LANCER.bat`.
- Run focused tests from the repo root with `python -m pytest AuditPro_SHARE/tests/test_reconciliation_liasse.py -v`.
- Run the profiling script from the repo root with `python AuditPro_SHARE/profile_liasse_performance.py`.
- If you change PyInstaller inputs under `AuditPro_Agent/AuditPro/`, rebuild from that folder with `pyinstaller AuditPro.spec`.

## Project-Specific Rules

- All AuditPro modules must inherit from `modules.base_module.BaseModule`.
- New modules belong in `modules/<module_name>/module.py` and must provide the metadata the registry expects: `name`, `description`, `category`, `version`, `help_text`, `detection_keywords`, and `detection_threshold`.
- The module registry auto-discovers packages that contain `module.py`; keep module entry points simple and import-safe.
- Standalone scripts and tests in this repo commonly add `AuditPro_SHARE\\dist\\AuditPro\\_internal` to `sys.path` before importing app modules. Preserve that pattern when extending existing scripts.
- Keep `ui/workspace.py` compliant with the hidden output-directory rule: `output_frame` stays hidden.
- OCR and bundled binary resolution must go through `core/ocr_paths.py`. Do not hardcode Tesseract or Poppler paths.
- Do not expose OCR path inputs in `module.py` wrappers and do not pass OCR path parameters from wrappers into module logic.

## Working In This Repo

- `_internal` is ignored by `.gitignore`, but it is still the runtime tree present in this workspace. Treat edits there as surgical and verify them carefully.
- Avoid broad refactors in packaged runtime code unless the task explicitly requires it.
- Some existing tests and profiling scripts depend on local OneDrive sample files and may skip or degrade when those files are absent.
- When touching module behavior, prefer a narrow pytest run for the affected slice before wider validation.

## References

- Module workflow skill: [.github/skills/auditpro-module-development/SKILL.md](.github/skills/auditpro-module-development/SKILL.md)
- Bug-fix example and compliance notes: [AuditPro_SHARE/BUG_FIX_REPORT.md](AuditPro_SHARE/BUG_FIX_REPORT.md)
- Current reconciliation tests: [AuditPro_SHARE/tests/test_reconciliation_liasse.py](AuditPro_SHARE/tests/test_reconciliation_liasse.py)
- Current profiling entry point: [AuditPro_SHARE/profile_liasse_performance.py](AuditPro_SHARE/profile_liasse_performance.py)
