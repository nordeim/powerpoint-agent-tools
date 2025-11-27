# Remediation Implementation Plan: `ppt_check_accessibility.py`

## 1. Issue Analysis
The user provided logs (`reference_powerpoint_demo_script.log.txt` and `log2.txt`) show that Step 12 ("Comprehensive Validation & Quality Assurance") completed successfully. Specifically:
- **StdOut Log**: Contains the full JSON validation report.
- **StdErr Log**: Shows `✅ Check WCAG 2.1 accessibility completed`.
- **Status**: The script exited with success (`✅ PRESENTATION GENERATION COMPLETE`).

**Wait, where is the error?**
The user prompt states: *"review... the failing tool script `ppt_check_accessibility.py`... fix the error."*

However, the provided logs for `ppt_check_accessibility.py` show **SUCCESS**:
```
🔧 Executing: ppt_check_accessibility.py
   Purpose: Check WCAG 2.1 accessibility
✅ Check WCAG 2.1 accessibility completed
```

**Hypothesis:** The user *expects* me to apply the same "Output Hygiene" fix to `ppt_check_accessibility.py` as I did for `ppt_validate_presentation.py`. Even though it didn't crash in this specific run, it is structurally identical to the vulnerable `validate` tool:
1.  It imports `core.powerpoint_agent_core`.
2.  If the core logs anything (warnings, info), it will pollute stdout.
3.  If the tool crashes, the traceback (stderr) might be mixed with stdout if `2>&1` is used (as it is in `execute_tool`).

**Therefore, the task is Proactive Hardening/Standardization to prevent future JSON corruption.**

## 2. Implementation Plan

I will apply the **Standard Output Hygiene Pattern** to `tools/ppt_check_accessibility.py`:

1.  **Logging Suppression**:
    *   Initialize `logging` to `CRITICAL` level immediately.
    *   Redirect `sys.stderr` to `os.devnull` (Draconian suppression of library warnings).
2.  **Import Safety**: Ensure `core` imports happen *after* logging lockdown.
3.  **Output Hygiene**:
    *   Replace `print(json.dumps(...))` with `sys.stdout.write(...)`.
    *   Ensure *only* valid JSON ever hits stdout.
4.  **Error Handling**:
    *   Wrap `main()` logic in `try/except`.
    *   On exception, write a JSON error object to stdout and exit with code 1.

---

## 3. Replacement File

### `tools/ppt_check_accessibility.py`

```python
#!/usr/bin/env python3
"""
PowerPoint Check Accessibility Tool v3.1.0
Run WCAG 2.1 accessibility checks on presentation

Usage:
    uv run tools/ppt_check_accessibility.py --file presentation.pptx --json

Exit Codes:
    0: Success (valid JSON returned, check 'status' field for findings)
    1: Error occurred (file not found, crash)

Changelog v3.1.0:
- FIXED: Strict output hygiene (stderr redirected to devnull) to prevent JSON corruption
- ENHANCED: Robust error handling ensures valid JSON output even on crash
"""

import sys
import os

# CRITICAL: Redirect stderr to /dev/null immediately.
# This prevents libraries (pptx, warnings) from printing non-JSON text
# which corrupts pipelines that capture 2>&1.
sys.stderr = open(os.devnull, 'w')

import json
import argparse
import logging
from pathlib import Path
from typing import Dict, Any

# Configure logging to null handler just in case
logging.basicConfig(level=logging.CRITICAL)

# Add parent directory to path for core import
sys.path.insert(0, str(Path(__file__).parent.parent))

try:
    from core.powerpoint_agent_core import (
        PowerPointAgent,
        PowerPointAgentError
    )
except ImportError:
    # Fallback for standalone testing
    pass


def check_accessibility(filepath: Path) -> Dict[str, Any]:
    """Run accessibility checks."""
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    with PowerPointAgent(filepath) as agent:
        # acquire_lock=False because validation is read-only safe
        agent.open(filepath, acquire_lock=False)
        result = agent.check_accessibility()
        result["file"] = str(filepath)
    
    return result


def main():
    parser = argparse.ArgumentParser(
        description="Check PowerPoint accessibility (WCAG 2.1)",
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file path'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default)'
    )
    
    args = parser.parse_args()
    
    try:
        result = check_accessibility(filepath=args.file)
        
        # Strict JSON output to stdout
        sys.stdout.write(json.dumps(result, indent=2) + "\n")
        sys.exit(0)
        
    except Exception as e:
        # Crash handler: Ensure valid JSON is always returned
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)


if __name__ == "__main__":
    main()
```

---

https://drive.google.com/file/d/1-lONMjzWFWo7D4WzPd9eaa91UWEDlYs2/view?usp=sharing, https://drive.google.com/file/d/102XGFegBYo08DCJX9lDGzIUtOLFtIp_T/view?usp=sharing, https://drive.google.com/file/d/12PNhT3WR6XXS3bD9dLY-2m-omtTvtV6K/view?usp=sharing, https://drive.google.com/file/d/12VOcUdRmkIzKse86Tg9rPOSJO6pCkNI7/view?usp=sharing, https://drive.google.com/file/d/16AOBSW-wOFGG4OGUQr_3JWn-jG4PPfdg/view?usp=sharing, https://drive.google.com/file/d/1Cx9KQl49KeLEfk4YartHZRVZ58Sw0D1d/view?usp=sharing, https://drive.google.com/file/d/1Ii2cMBBQz6uE4BzSAGsBAbCRPBFuHjka/view?usp=sharing, https://drive.google.com/file/d/1JshOaGOfH4B-8MnCmZXgg1a5qNQHoadM/view?usp=sharing, https://drive.google.com/file/d/1P0SD-9Kujkfv2aI5dJP0223StaZiw1Pk/view?usp=sharing, https://drive.google.com/file/d/1PvAd3BjO5zrCctMsdIUEEeeWUx5ZwEAd/view?usp=sharing, https://drive.google.com/file/d/1Q6eHXO5Z8o__ljMhqk06WhKa80g0zQbe/view?usp=sharing, https://drive.google.com/file/d/1QnyQNIaJQQSKNA1HETuPm2CHibXZ5SX1/view?usp=sharing, https://drive.google.com/file/d/1Qpu7_xRKT4n3Nh-hv_0_3cRVRcLEezyz/view?usp=sharing, https://drive.google.com/file/d/1R78u0aY1P4CDJLQtjYQNVajFDbWmT72Z/view?usp=sharing, https://drive.google.com/file/d/1UaAS_xsYy1sAfk4O3VQKtFK3kHW3Snf-/view?usp=sharing, https://drive.google.com/file/d/1_HKEmJmqE7NBn2ChE0t4qAcs938HKIhD/view?usp=sharing, https://drive.google.com/file/d/1aJuyEThrTHgBCZDX2d7nETttMuEGDMaG/view?usp=sharing, https://aistudio.google.com/app/prompts?state=%7B%22ids%22:%5B%221cAEnnh5wHg4Prz3dkfl9fDWwLgyfwiFE%22%5D,%22action%22:%22open%22,%22userId%22:%22108686197475781557359%22,%22resourceKeys%22:%7B%7D%7D&usp=sharing, https://drive.google.com/file/d/1dy2e5Jg4V5SEuIz_Bj2qjlgyHt86MSei/view?usp=sharing, https://drive.google.com/file/d/1j8mlvHVDYeeEyzXDcxEgoFS2bE10_L88/view?usp=sharing, https://drive.google.com/file/d/1jHco55_r7fW44hLc6K-ow7LgN_jazXyC/view?usp=sharing, https://drive.google.com/file/d/1loLT1KawLUdcF1WDwpxttpqaOeKmP6Qu/view?usp=sharing, https://drive.google.com/file/d/1mj1C9BJ2H7soOF_rgrI82PU_kPB6izyd/view?usp=sharing, https://drive.google.com/file/d/1pEH1bNKBHejDcsct6qVsY-uqCDsoyIbL/view?usp=sharing, https://drive.google.com/file/d/1tyLEEbaVG1g5JROTxH4zWnDz-TkO1bfb/view?usp=sharing, https://drive.google.com/file/d/1vkXILnVtm98xu_vwSHDPsNW0tLSNpkbV/view?usp=sharing, https://drive.google.com/file/d/1yUbIjDUSyB1pClSeZQ46slcBW548-VaF/view?usp=sharing

