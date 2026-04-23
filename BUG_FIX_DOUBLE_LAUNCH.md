# Bug Fix: Double/Triple Instance Launch

**Date**: April 23, 2026  
**Status**: ✅ FIXED

---

## Problem 1: Two LANCER.bat Files

### Before
```
AuditPro_v1.1_1/
├── LANCER.bat                    ← Version 1 (at root)
└── AuditPro_SHARE/
    └── LANCER.bat                ← Version 2 (actual launcher)
```

### Root Cause
- **Root LANCER.bat**: Was a "proxy" that just called AuditPro_SHARE/LANCER.bat
- This created confusion about which one to use
- Users might accidentally have shortcuts to both

### Solution
- Keep only **AuditPro_SHARE/LANCER.bat** as the canonical launcher
- Root LANCER.bat now just delegates to it with a deprecation warning
- Users should bookmark/shortcut: `AuditPro_SHARE\LANCER.bat`

---

## Problem 2: Multiple Instances Opening Automatically

### Root Cause
The LANCER.bat script had a critical logic error:

```batch
REM Try editable mode
if exist "run_internal.py" (
    "%PYTHON_CMD%" "run_internal.py"
    set "RET=%ERRORLEVEL%"
    if "%RET%"=="0" exit /b 0
    REM ← Falls through if RET != 0
)

REM Try packaged exe (ALWAYS executed if first method failed!)
if exist "dist\AuditPro\AuditPro.exe" (
    start "" "dist\AuditPro\AuditPro.exe"  ← Second instance!
    exit /b 0
)
```

**The Bug**: If `run_internal.py` returned non-zero, the script would ALSO launch the .exe, creating 2+ instances.

### Scenarios That Triggered This:
1. **Fallback bug**: If run_internal.py had any error, both launcher methods would execute
2. **Duplicate shortcuts**: User had shortcuts to both LANCER.bat files
3. **Rapid clicking**: Launching LANCER.bat multiple times in quick succession
4. **Task scheduler**: If automated tasks launched it multiple times

---

## Solution Implemented

### Fixed LANCER.bat Strategy
```batch
REM ── NEW: Try ONE method only, not multiple ───────────────
if exist "run_internal.py" (
    "%PYTHON_CMD%" "run_internal.py"
    REM Terminate immediately on exit (success OR failure)
    exit /b %ERRORLEVEL%
)

REM Only try .exe if run_internal.py doesn't exist
if exist "dist\AuditPro\AuditPro.exe" (
    start "" "dist\AuditPro\AuditPro.exe"
    exit /b 0
)
```

### Key Changes
1. **Priority order**: Try editable mode (run_internal.py) FIRST
2. **No fallthrough**: Each method has its own exit, preventing double-launch
3. **Mutual exclusion**: Only executes one launcher method
4. **Clear error handling**: Reports which method failed and why

---

## Files Modified

```
✓ LANCER.bat (root)           → Marked as deprecated, delegates to AuditPro_SHARE
✓ AuditPro_SHARE/LANCER.bat   → FIXED: Prevents double launch
```

---

## Testing

### Test 1: Run from AuditPro_SHARE
```bash
cd AuditPro_SHARE
LANCER.bat
```
✓ Should launch only ONE instance

### Test 2: Run from root
```bash
LANCER.bat
```
✓ Should show deprecation warning, then launch only ONE instance from AuditPro_SHARE

### Test 3: Rapid clicks
```bash
# Click LANCER.bat multiple times
```
✓ Each click should launch a separate instance (expected behavior for multiple launches)  
✓ NOT multiple instances from a single LANCER.bat call

---

## Recommendations

### For Users
1. **Delete root LANCER.bat** (or keep it as deprecated)
2. **Use only**: `AuditPro_SHARE\LANCER.bat`
3. **Create shortcuts** to: `AuditPro_SHARE\LANCER.bat` (not root)
4. **If using Task Scheduler**: Only schedule ONE task pointing to `AuditPro_SHARE\LANCER.bat`

### For Developers
1. Remove or consolidate the root LANCER.bat eventually
2. Consider creating a `.cmd` wrapper for common use cases:
   - Debug mode (run_internal.py with logging)
   - Release mode (use .exe)
3. Add mutual process locking to prevent multiple instances:
   ```python
   # In run_internal.py
   import fcntl
   lock_file = Path.home() / ".auditpro.lock"
   with open(lock_file, 'w') as f:
       try:
           fcntl.flock(f.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
       except BlockingIOError:
           print("AuditPro is already running")
           sys.exit(1)
   ```

---

## Directory Structure (Cleaned)

```
AuditPro_v1.1_1/
├── LANCER.bat                          (deprecated: delegates to AuditPro_SHARE)
└── AuditPro_SHARE/
    ├── LANCER.bat                      (FIXED: canonical launcher)
    ├── run_internal.py                 (Python mode: editable)
    └── dist/AuditPro/
        ├── AuditPro.exe                (packaged executable)
        └── _internal/                  (extracted runtime)
```

**Recommendation**: Use `AuditPro_SHARE\LANCER.bat` exclusively.

---

## Summary

✅ **Eliminated dual-launch bug**  
✅ **Clarified launcher hierarchy**  
✅ **Prevented fallback conflicts**  
✅ **Only one LANCER.bat should be used going forward**

