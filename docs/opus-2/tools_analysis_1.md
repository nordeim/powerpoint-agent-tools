# Tool Validation Report
## Meticulous Review & Analysis of 5 PowerPoint Agent Tools

---

## Executive Summary

| Tool | Critical Issues | High Priority | Medium Priority | Overall Status |
|------|-----------------|---------------|-----------------|----------------|
| `ppt_create_from_template.py` | 3 | 4 | 2 | 🔴 **Needs Fixes** |
| `ppt_create_new.py` | 3 | 4 | 2 | 🔴 **Needs Fixes** |
| `ppt_delete_slide.py` | 4 | 3 | 2 | 🔴 **Critical - Governance Violation** |
| `ppt_duplicate_slide.py` | 2 | 3 | 2 | 🔴 **Needs Fixes** |
| `ppt_get_info.py` | 1 | 2 | 1 | 🟡 **Minor Fixes** |

**Critical Finding**: `ppt_delete_slide.py` is a **destructive operation** that **does not enforce approval tokens**, violating the governance framework.

---

## Validation Framework Applied

```
┌─────────────────────────────────────────────────────────────────────────┐
│                     VALIDATION CHECKLIST                                │
├─────────────────────────────────────────────────────────────────────────┤
│                                                                         │
│  GOLDEN RULES COMPLIANCE:                                               │
│  ☐ Rule 1: Output Hygiene (Hygiene Block present)                      │
│  ☐ Rule 2: Clone-Before-Edit awareness                                 │
│  ☐ Rule 3: Fail Safely with JSON                                       │
│  ☐ Rule 4: Version Tracking (mutations)                                │
│  ☐ Rule 5: Approval Tokens (destructive ops)                           │
│                                                                         │
│  v3.1.x COMPLIANCE:                                                     │
│  ☐ Core methods return Dict (not primitives)                           │
│  ☐ presentation_version tracking                                       │
│  ☐ Proper exception types                                              │
│                                                                         │
│  CODE QUALITY:                                                          │
│  ☐ Naming convention (ppt_<verb>_<noun>.py)                            │
│  ☐ Module docstring with version/usage/exit codes                      │
│  ☐ Type hints on function signatures                                   │
│  ☐ Docstrings with Args/Returns/Raises                                 │
│  ☐ Proper error handling                                               │
│                                                                         │
└─────────────────────────────────────────────────────────────────────────┘
```

---

## Detailed Tool Analysis

### Tool 1: `ppt_create_from_template.py`

#### Classification
- **Type**: Creation tool
- **Destructive**: No
- **Requires Approval Token**: No

#### Issues Found

| Severity | Issue | Location | Details |
|----------|-------|----------|---------|
| 🔴 **CRITICAL** | Missing Hygiene Block | Top of file | No `sys.stderr` suppression - library warnings will corrupt JSON output |
| 🔴 **CRITICAL** | Missing Version Tracking | `create_from_template()` | No `presentation_version_before/after` in return |
| 🔴 **CRITICAL** | v3.1.x Return Handling | Line ~52 | `idx = agent.add_slide(...)` assumes int return, but v3.1.x returns Dict |
| 🟡 HIGH | Missing Version Constant | File header | No `__version__ = "3.x.x"` constant |
| 🟡 HIGH | Incorrect Command Syntax | Docstring | Uses `uv python` instead of `uv run tools/` |
| 🟡 HIGH | Missing Type Hints | Function signature | `-> Dict[str, Any]` present but incomplete docstring |
| 🟡 HIGH | Missing Docstring Details | `create_from_template()` | No Args/Returns/Raises documentation |
| 🟠 MEDIUM | `--json` Not Default True | argparse | Inconsistent with tool template |
| 🟠 MEDIUM | Missing `os` Import | Top of file | Required for hygiene block |

#### Code Issues

```python
# ❌ WRONG: v3.1.x add_slide() returns Dict, not int
idx = agent.add_slide(layout_name=layout)
slide_indices.append(idx)

# ✅ CORRECT: Extract slide_index from Dict
result = agent.add_slide(layout_name=layout)
slide_indices.append(result["slide_index"])
```

```python
# ❌ MISSING: Hygiene Block
import sys
import json
# ... (no hygiene block)

# ✅ REQUIRED: Add at top after imports
import os
sys.stderr = open(os.devnull, 'w')
```

```python
# ❌ MISSING: Version tracking in return
return {
    "status": "success",
    "file": str(output),
    # ... no presentation_version fields
}

# ✅ REQUIRED: Include version tracking
info_before = agent.get_presentation_info()
version_before = info_before.get("presentation_version")
# ... operations ...
info_after = agent.get_presentation_info()
version_after = info_after.get("presentation_version")

return {
    "status": "success",
    "file": str(output),
    "presentation_version": version_after,  # For new files, only after matters
    # ...
}
```

---

### Tool 2: `ppt_create_new.py`

#### Classification
- **Type**: Creation tool
- **Destructive**: No
- **Requires Approval Token**: No

#### Issues Found

| Severity | Issue | Location | Details |
|----------|-------|----------|---------|
| 🔴 **CRITICAL** | Missing Hygiene Block | Top of file | Same as above |
| 🔴 **CRITICAL** | Missing Version Tracking | Return statement | No `presentation_version` in output |
| 🔴 **CRITICAL** | v3.1.x Return Handling | Line ~45 | `idx = agent.add_slide(...)` assumes int |
| 🟡 HIGH | Missing Version Constant | File header | No `__version__` |
| 🟡 HIGH | Incorrect Command Syntax | Docstring | `uv python` vs `uv run tools/` |
| 🟡 HIGH | Incomplete Docstring | `create_new_presentation()` | Missing Args/Returns/Raises |
| 🟡 HIGH | Dual-Purpose Confusion | Function logic | Both creates new AND can use template - consider separation |
| 🟠 MEDIUM | Overlaps with `ppt_create_from_template.py` | Design | `--template` option duplicates other tool's purpose |

#### Architectural Concern

The tool accepts `--template` parameter, which overlaps with `ppt_create_from_template.py`. This creates confusion about which tool to use. Consider:
- Remove `--template` from this tool, OR
- Document clear distinction between the two tools

---

### Tool 3: `ppt_delete_slide.py` 

#### Classification
- **Type**: Mutation tool
- **Destructive**: **YES** ⚠️
- **Requires Approval Token**: **YES** ⚠️

#### 🚨 GOVERNANCE VIOLATION ALERT

```
┌─────────────────────────────────────────────────────────────────────────┐
│  ⚠️  CRITICAL GOVERNANCE VIOLATION                                      │
├─────────────────────────────────────────────────────────────────────────┤
│                                                                         │
│  VIOLATION: Destructive operation without approval token enforcement   │
│                                                                         │
│  DOCUMENTED REQUIREMENT (from Core Handbook):                           │
│  > delete_slide(index, approval_token=None)                            │
│  > Security: REQUIRES valid approval_token matching scope 'delete:slide'│
│  > Throws: ApprovalTokenError if token is invalid/missing              │
│                                                                         │
│  ACTUAL IMPLEMENTATION:                                                 │
│  > def delete_slide(filepath: Path, index: int) -> Dict[str, Any]:     │
│  > agent.delete_slide(index)  # ❌ NO TOKEN PASSED                     │
│                                                                         │
│  RISK: Unauthorized slide deletion possible                             │
│                                                                         │
└─────────────────────────────────────────────────────────────────────────┘
```

#### Issues Found

| Severity | Issue | Location | Details |
|----------|-------|----------|---------|
| 🔴 **CRITICAL** | **Missing Approval Token** | Entire tool | No `--approval-token` arg, no token passed to core |
| 🔴 **CRITICAL** | Missing Hygiene Block | Top of file | No stderr suppression |
| 🔴 **CRITICAL** | Missing Version Tracking | Return statement | No `presentation_version_before/after` |
| 🔴 **CRITICAL** | No Clone-Before-Edit Check | Function logic | Modifies file in-place without safety check |
| 🟡 HIGH | Minimal Docstring | Module level | Missing author/version/license |
| 🟡 HIGH | Missing Suggestion Field | Error response | Should include remediation guidance |
| 🟡 HIGH | Incorrect Command Syntax | Docstring | `uv python` vs `uv run tools/` |
| 🟠 MEDIUM | Missing `--dry-run` Option | argparse | Should show what would be deleted |
| 🟠 MEDIUM | No Backup Warning | Output | Should warn user about irreversibility |

#### Required Fixes

```python
# ❌ CURRENT: No approval token
def delete_slide(filepath: Path, index: int) -> Dict[str, Any]:
    # ...
    agent.delete_slide(index)

# ✅ REQUIRED: With approval token
def delete_slide(
    filepath: Path, 
    index: int,
    approval_token: str = None
) -> Dict[str, Any]:
    """
    Delete slide at index.
    
    SECURITY: Requires valid approval_token with scope 'delete:slide'
    
    Args:
        filepath: Path to PowerPoint file
        index: Slide index to delete (0-based)
        approval_token: HMAC-SHA256 approval token with scope 'delete:slide'
        
    Returns:
        Dict with deletion results and version tracking
        
    Raises:
        ApprovalTokenError: If token is missing or invalid
        SlideNotFoundError: If index is out of range
    """
    if not approval_token:
        raise ApprovalTokenError(
            "Approval token required for slide deletion",
            details={
                "operation": "delete_slide",
                "slide_index": index,
                "required_scope": "delete:slide"
            }
        )
    
    # ... version tracking ...
    agent.delete_slide(index, approval_token=approval_token)
    # ...
```

```python
# CLI argument needed:
parser.add_argument(
    '--approval-token',
    required=True,
    type=str,
    help='Approval token with scope "delete:slide" (required for destructive operation)'
)
```

---

### Tool 4: `ppt_duplicate_slide.py`

#### Classification
- **Type**: Mutation tool
- **Destructive**: No (additive operation)
- **Requires Approval Token**: No

#### Issues Found

| Severity | Issue | Location | Details |
|----------|-------|----------|---------|
| 🔴 **CRITICAL** | Missing Hygiene Block | Top of file | No stderr suppression |
| 🔴 **CRITICAL** | Missing Version Tracking | Return statement | No `presentation_version_before/after` |
| 🟡 HIGH | Minimal Docstring | Module level | Missing author/version/license/exit codes |
| 🟡 HIGH | v3.1.x Return Handling | Line ~25 | `new_index = agent.duplicate_slide(index)` - may return Dict |
| 🟡 HIGH | Incorrect Command Syntax | Docstring | `uv python` vs `uv run tools/` |
| 🟠 MEDIUM | Missing Return Type | Function signature | No `-> Dict[str, Any]` annotation |
| 🟠 MEDIUM | No Clone-Before-Edit Awareness | Function logic | Should warn about in-place modification |

#### Code Issue

```python
# ⚠️ POTENTIAL ISSUE: Verify core return type
new_index = agent.duplicate_slide(index)

# If v3.1.x returns Dict:
result = agent.duplicate_slide(index)
new_index = result["slide_index"]  # or result["new_slide_index"]
```

---

### Tool 5: `ppt_get_info.py`

#### Classification
- **Type**: Inspection tool (read-only)
- **Destructive**: No
- **Requires Approval Token**: No

#### Issues Found

| Severity | Issue | Location | Details |
|----------|-------|----------|---------|
| 🔴 **CRITICAL** | Missing Hygiene Block | Top of file | No stderr suppression |
| 🟡 HIGH | Missing `presentation_version` | Return statement | Should include version for caller tracking |
| 🟡 HIGH | Incorrect Command Syntax | Docstring | `uv python` vs `uv run tools/` |
| 🟠 MEDIUM | Missing Version Constant | File header | No `__version__` |

#### Positive Observations
✅ Uses `acquire_lock=False` for read-only operation (correct)
✅ Comprehensive help text with examples
✅ Good error handling structure
✅ Returns useful metadata

#### Recommended Enhancement

```python
# ✅ Include presentation_version for tracking
return {
    "status": "success",
    "file": info["file"],
    "presentation_version": info.get("presentation_version"),  # ADD THIS
    "slide_count": info["slide_count"],
    # ...
}
```

---

## Common Issues Across All Tools

### Issue 1: Missing Hygiene Block (ALL TOOLS)

**Current State**: None of the tools have the hygiene block.

**Required Fix** (add to ALL tools after imports):

```python
#!/usr/bin/env python3
"""..."""

import sys
import os

# --- HYGIENE BLOCK START ---
# CRITICAL: Redirect stderr to /dev/null to prevent library noise from corrupting JSON output
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

import json
import argparse
# ... rest of imports
```

### Issue 2: Incorrect Command Syntax in Docstrings (ALL TOOLS)

**Current**: `uv python ppt_*.py`
**Required**: `uv run tools/ppt_*.py`

### Issue 3: Missing Version Constant (ALL TOOLS)

**Required**: Add after imports:

```python
__version__ = "3.1.0"
```

### Issue 4: v3.1.x Return Value Handling

Tools assume core methods return primitives, but v3.1.x returns Dicts:

| Method | v3.0.x Return | v3.1.x Return |
|--------|---------------|---------------|
| `add_slide()` | `int` | `Dict` with `slide_index` |
| `duplicate_slide()` | `int` | `Dict` with `new_slide_index` |
| `delete_slide()` | `None` | `Dict` with deletion info |

---

## Remediation Plan

### Priority 1: Critical Fixes (Immediate)

#### 1.1 Fix `ppt_delete_slide.py` Governance Violation

```python
#!/usr/bin/env python3
"""
PowerPoint Delete Slide Tool v3.1.0
Remove a slide from the presentation

⚠️ DESTRUCTIVE OPERATION - Requires approval token

Usage:
    uv run tools/ppt_delete_slide.py --file presentation.pptx --index 1 --approval-token "HMAC-SHA256:..." --json

Exit Codes:
    0: Success
    1: Error (check error_type in JSON)
    4: Permission Error (missing/invalid approval token)
"""

import sys
import os

# --- HYGIENE BLOCK START ---
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

import json
import argparse
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, 
    PowerPointAgentError, 
    SlideNotFoundError,
    ApprovalTokenError
)

__version__ = "3.1.0"


def delete_slide(
    filepath: Path, 
    index: int,
    approval_token: str
) -> Dict[str, Any]:
    """
    Delete slide at specified index.
    
    ⚠️ DESTRUCTIVE OPERATION - Requires approval token
    
    Args:
        filepath: Path to PowerPoint file
        index: Slide index to delete (0-based)
        approval_token: HMAC-SHA256 token with scope 'delete:slide'
        
    Returns:
        Dict containing:
            - deleted_index: Index of deleted slide
            - remaining_slides: New slide count
            - presentation_version_before: Version hash before deletion
            - presentation_version_after: Version hash after deletion
            
    Raises:
        FileNotFoundError: If file doesn't exist
        SlideNotFoundError: If index is out of range
        ApprovalTokenError: If token is missing or invalid
    """
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if not approval_token:
        raise ApprovalTokenError(
            "Approval token required for slide deletion",
            details={
                "operation": "delete_slide",
                "slide_index": index,
                "required_scope": "delete:slide",
                "suggestion": "Generate token using approval token generator with scope 'delete:slide'"
            }
        )
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Capture version before
        info_before = agent.get_presentation_info()
        version_before = info_before.get("presentation_version")
        
        total_slides = agent.get_slide_count()
        if not 0 <= index < total_slides:
            raise SlideNotFoundError(
                f"Index {index} out of range",
                details={
                    "requested": index,
                    "available": total_slides,
                    "valid_range": f"0-{total_slides-1}"
                }
            )
        
        # Pass approval token to core
        agent.delete_slide(index, approval_token=approval_token)
        agent.save()
        
        # Capture version after
        info_after = agent.get_presentation_info()
        version_after = info_after.get("presentation_version")
        new_count = info_after["slide_count"]
    
    return {
        "status": "success",
        "file": str(filepath.resolve()),
        "deleted_index": index,
        "remaining_slides": new_count,
        "presentation_version_before": version_before,
        "presentation_version_after": version_after,
        "tool_version": __version__
    }


def main():
    parser = argparse.ArgumentParser(
        description="Delete PowerPoint slide (⚠️ DESTRUCTIVE - requires approval token)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
⚠️ DESTRUCTIVE OPERATION

This tool permanently removes a slide from the presentation.
An approval token with scope 'delete:slide' is REQUIRED.

Examples:
    uv run tools/ppt_delete_slide.py --file deck.pptx --index 2 --approval-token "HMAC-SHA256:..." --json
    
Safety Recommendations:
    1. Clone the presentation first: ppt_clone_presentation.py
    2. Verify slide count: ppt_get_info.py --file deck.pptx
    3. Inspect slide content: ppt_get_slide_info.py --file deck.pptx --slide 2
    4. Then delete with approval token
        """
    )
    
    parser.add_argument('--file', required=True, type=Path, help='PowerPoint file path')
    parser.add_argument('--index', required=True, type=int, help='Slide index to delete (0-based)')
    parser.add_argument(
        '--approval-token',
        required=True,
        type=str,
        help='Approval token with scope "delete:slide" (REQUIRED)'
    )
    parser.add_argument('--json', action='store_true', default=True, help='Output JSON')
    
    args = parser.parse_args()
    
    try:
        result = delete_slide(
            filepath=args.file, 
            index=args.index,
            approval_token=args.approval_token
        )
        print(json.dumps(result, indent=2))
        sys.exit(0)
        
    except ApprovalTokenError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ApprovalTokenError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Generate approval token with scope 'delete:slide'"
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(4)  # Permission error exit code
        
    except SlideNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "SlideNotFoundError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Use ppt_get_info.py to check available slides"
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()
```

#### 1.2 Add Hygiene Block to All Tools

Add to all 5 tools immediately after the shebang and docstring.

#### 1.3 Fix v3.1.x Return Value Handling

Update all `agent.add_slide()` and `agent.duplicate_slide()` calls:

```python
# Before
idx = agent.add_slide(layout_name=layout)

# After
result = agent.add_slide(layout_name=layout)
idx = result["slide_index"] if isinstance(result, dict) else result
```

### Priority 2: High Priority Fixes

1. Add version tracking to all mutation tools
2. Add `__version__` constant to all tools
3. Fix command syntax in all docstrings
4. Add proper docstrings with Args/Returns/Raises

### Priority 3: Medium Priority

1. Add `--dry-run` option to `ppt_delete_slide.py`
2. Resolve overlap between `ppt_create_new.py` and `ppt_create_from_template.py`
3. Add clone-before-edit warnings to mutation tools

---

## Validation Checklist Summary

| Check | create_from_template | create_new | delete_slide | duplicate_slide | get_info |
|-------|---------------------|------------|--------------|-----------------|----------|
| Hygiene Block | ❌ | ❌ | ❌ | ❌ | ❌ |
| Version Constant | ❌ | ❌ | ❌ | ❌ | ❌ |
| Correct CLI Syntax | ❌ | ❌ | ❌ | ❌ | ❌ |
| Version Tracking | ❌ | ❌ | ❌ | ❌ | ⚠️ |
| Approval Token | N/A | N/A | ❌ **CRITICAL** | N/A | N/A |
| v3.1.x Returns | ❌ | ❌ | ⚠️ | ⚠️ | ✅ |
| Error Handling | ✅ | ✅ | ✅ | ✅ | ✅ |
| Type Hints | ⚠️ | ⚠️ | ⚠️ | ❌ | ✅ |
| Docstrings | ⚠️ | ⚠️ | ❌ | ❌ | ✅ |

---

## Recommendations

### Immediate Actions Required

1. **🚨 URGENT**: Fix `ppt_delete_slide.py` to require approval token
2. **🚨 URGENT**: Add hygiene block to all 5 tools
3. **HIGH**: Add version tracking to all mutation tools
4. **HIGH**: Fix v3.1.x return value handling

### Process Improvements

1. Create a **tool linting script** that validates against the checklist
2. Add **pre-commit hooks** to catch missing hygiene blocks
3. Create **tool template generator** to ensure compliance from start

---

**Would you like me to provide complete corrected implementations for any or all of these tools?**
