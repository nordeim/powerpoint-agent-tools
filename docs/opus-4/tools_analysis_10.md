# 🔍 Tool Validation & Analysis Report
## Meticulous Review of 4 DISCOVER/VALIDATE Phase Tools

---

## Executive Summary

| Tool | Phase | Status | Critical Issues | Warnings |
|------|-------|--------|-----------------|----------|
| `ppt_capability_probe.py` | DISCOVER | ⚠️ Needs Fix | 1 | 6 |
| `ppt_check_accessibility.py` | VALIDATE | ⚠️ Needs Fix | 2 | 4 |
| `ppt_json_adapter.py` | Utility | ❌ Broken | 2 | 5 |
| `ppt_validate_presentation.py` | VALIDATE | ❌ Broken | 2 | 8 |

---

# Tool 1: `ppt_capability_probe.py` v1.1.1

## Classification
- **Phase**: DISCOVER
- **Operation Type**: Read-only (atomic)
- **Resilience Pattern**: ✅ Implemented (3-layer)

## ✅ Strengths

| Aspect | Assessment |
|--------|------------|
| **Probe Resilience** | ✅ Excellent - implements all 3 layers (timeout, transient slides, graceful degradation) |
| **Transient Slide Pattern** | ✅ Correct `_add_transient_slide()` with finally cleanup |
| **Atomic Verification** | ✅ MD5 checksum before/after comparison |
| **Metadata Richness** | ✅ `operation_id`, `duration_ms`, `library_versions` |
| **Schema Validation** | ✅ Uses `validate_against_schema()` |
| **Type Hints** | ✅ Comprehensive |
| **Documentation** | ✅ Excellent docstrings and usage examples |
| **Timeout Protection** | ✅ Per-iteration checks in layout analysis |

## ❌ Critical Issues

### Issue 1: Missing Stderr Hygiene Block
**Severity**: 🔴 Critical  
**Location**: Top of file (missing)

```python
# MISSING - Required per PROGRAMMING_GUIDE.md Rule 1
# This should be at the very top after imports begin:

import sys
import os

# --- HYGIENE BLOCK START ---
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---
```

**Impact**: Library warnings from `python-pptx` or `lxml` can corrupt JSON output, breaking downstream pipelines using `jq` or similar parsers.

**Evidence**: Other validated tools (`ppt_check_accessibility.py`, `ppt_validate_presentation.py`) correctly implement this pattern.

## ⚠️ Warnings

### Warning 1: Version Number Inconsistency
**Current**: Tool reports `v1.1.1`  
**Expected**: Should align with project version `v3.1.x`

```python
# Current
"tool_version": "1.1.1"

# Should be
"tool_version": "3.1.0"
```

### Warning 2: Hardcoded Schema Path May Not Exist
```python
schema_path = Path(__file__).parent.parent / "schemas" / "capability_probe.v1.1.1.schema.json"
```
**Risk**: If schema file doesn't exist, validation fails silently (caught in try/except).

### Warning 3: Bare Except Clauses
Multiple instances of `except:` without exception type:

```python
# Lines ~67-72
try:
    versions["python-pptx"] = importlib.metadata.version("python-pptx")
except:  # ❌ Should specify exception type
    versions["python-pptx"] = "unknown"
```

**Fix**:
```python
except importlib.metadata.PackageNotFoundError:
    versions["python-pptx"] = "unknown"
except Exception:
    versions["python-pptx"] = "error"
```

### Warning 4: Timeout Default Mismatch
- **Argparse default**: 30 seconds
- **Documentation standard**: 15 seconds

```python
parser.add_argument(
    '--timeout',
    type=int,
    default=30,  # Should be 15 per system prompt
    help='Timeout in seconds for analysis (default: 30)'
)
```

### Warning 5: Missing `presentation_version` in Output
While this is a read-only tool, capturing the presentation version would enable manifest creation:

```python
# Should add to result
"presentation_version": agent.get_presentation_version()  # If using agent
```

### Warning 6: `master_map` Key Stability
Using `id(layout)` as fallback key is fragile:
```python
try:
    key = layout.part.partname
except:
    key = id(layout)  # Fragile - object id can change
```

---

# Tool 2: `ppt_check_accessibility.py` v3.1.0

## Classification
- **Phase**: VALIDATE
- **Operation Type**: Read-only
- **Hygiene Block**: ✅ Present

## ✅ Strengths

| Aspect | Assessment |
|--------|------------|
| **Hygiene Block** | ✅ Correct placement at top |
| **Context Manager** | ✅ Proper `with` usage |
| **Lock Behavior** | ✅ `acquire_lock=False` for read-only |
| **Error Handling** | ✅ Catch-all with JSON output |
| **Simplicity** | ✅ Single responsibility |

## ❌ Critical Issues

### Issue 1: Output Doesn't Follow Standard Format
**Severity**: 🔴 Critical

The result from `agent.check_accessibility()` is returned directly without wrapping:

```python
# Current
result = agent.check_accessibility()
result["file"] = str(filepath)
# Missing: status, tool_version, presentation_version
```

**Expected Output Structure**:
```python
return {
    "status": "success",
    "file": str(filepath),
    "tool_version": "3.1.0",
    "presentation_version": agent.get_presentation_version(),
    "accessibility": result,  # Wrapped
    "issues": result.get("issues", {}),
    "summary": {
        "total_issues": len(result.get("issues", {}).get("missing_alt_text", [])) + ...
    }
}
```

### Issue 2: Silent ImportError Handling
```python
try:
    from core.powerpoint_agent_core import (
        PowerPointAgent,
        PowerPointAgentError
    )
except ImportError:
    pass  # ❌ PowerPointAgent will be undefined!
```

**Impact**: If import fails, `check_accessibility()` will raise `NameError: name 'PowerPointAgent' is not defined` instead of a helpful error.

**Fix**:
```python
except ImportError as e:
    print(json.dumps({
        "status": "error",
        "error": f"Core module import failed: {e}",
        "error_type": "ImportError"
    }, indent=2))
    sys.exit(1)
```

## ⚠️ Warnings

### Warning 1: Missing Version Constant
```python
# Should add
__version__ = "3.1.0"
```

### Warning 2: No Operation Metadata
Missing `operation_id`, `duration_ms`, `validated_at` fields that other tools include.

### Warning 3: No Presentation Version Tracking
Should capture `presentation_version` for audit trail.

### Warning 4: Incomplete Docstring
Exit codes section doesn't document what "check 'status' field for findings" means.

---

# Tool 3: `ppt_json_adapter.py`

## Classification
- **Phase**: Utility (post-processing)
- **Operation Type**: JSON transformation
- **Hygiene Block**: ❌ Missing

## ✅ Strengths

| Aspect | Assessment |
|--------|------------|
| **Alias Mapping** | ✅ Good handling of key drift |
| **Version Computation** | ✅ Fallback hash generation |
| **Schema Validation** | ✅ Uses jsonschema library |

## ❌ Critical Issues

### Issue 1: Broken ERROR_TEMPLATE Usage
**Severity**: 🔴 Critical

```python
ERROR_TEMPLATE = {
    "error": {
        "error_code": None,
        "message": None,
        "details": None,
        "retryable": False
    }
}

# Usage (BROKEN):
print(json.dumps({**ERROR_TEMPLATE, "error": {"error_code": "...", ...}}))
```

**Problem**: The spread creates `{"error": {...}, "error": {...}}` - the second `error` key overwrites the first, making `ERROR_TEMPLATE` completely useless.

**Fix**:
```python
def make_error(error_code: str, message: str, details=None, retryable=False):
    return {
        "status": "error",
        "error": message,
        "error_type": error_code,
        "details": details,
        "retryable": retryable
    }

# Usage:
print(json.dumps(make_error("SCHEMA_LOAD_ERROR", str(e))))
```

### Issue 2: Inconsistent Error Format
**Severity**: 🔴 Critical

The error structure doesn't match the project standard:

```python
# Current (nested):
{"error": {"error_code": "...", "message": "..."}}

# Project Standard (flat):
{"status": "error", "error": "message", "error_type": "ErrorType"}
```

## ⚠️ Warnings

### Warning 1: Missing Hygiene Block
No stderr redirection:
```python
# MISSING
sys.stderr = open(os.devnull, 'w')
```

### Warning 2: Uses `print()` Instead of `sys.stdout.write()`
```python
# Current
print(json.dumps(normalized, indent=2))

# Should be
sys.stdout.write(json.dumps(normalized, indent=2) + "\n")
```

### Warning 3: Exit Code Mapping Issues
| Code | Current Use | Standard Meaning |
|------|-------------|------------------|
| 5 | Schema load error | Internal Error |
| 3 | Input load error | Transient Error |
| 2 | Validation error | ✅ Correct |

### Warning 4: No Tool Metadata in Output
Missing `tool_version`, `processed_at` fields.

### Warning 5: compute_presentation_version Logic
Uses `.get("id", .get("index"))` which may produce inconsistent hashes if slides have both.

---

# Tool 4: `ppt_validate_presentation.py` v3.1.0

## Classification
- **Phase**: VALIDATE
- **Operation Type**: Read-only
- **Hygiene Block**: ✅ Present

## ✅ Strengths

| Aspect | Assessment |
|--------|------------|
| **Hygiene Block** | ✅ Correct placement |
| **Policy System** | ✅ Excellent configurable thresholds |
| **Dataclasses** | ✅ Clean structured data |
| **Recommendations** | ✅ Actionable fix suggestions |
| **Output Format** | ✅ Uses `sys.stdout.write()` |
| **Lock Behavior** | ✅ `acquire_lock=False` |

## ❌ Critical Issues

### Issue 1: Calls Undefined Core Method
**Severity**: 🔴 Critical  
**Location**: Line ~130

```python
asset_val = agent.validate_assets()  # ❌ NOT IN CORE API!
```

**Evidence**: The Core API Cheatsheet (Section 6) lists these validation methods:
- `validate_presentation()` ✅
- `check_accessibility()` ✅
- `validate_assets()` ❌ **NOT LISTED**

**Impact**: This will raise `AttributeError: 'PowerPointAgent' object has no attribute 'validate_assets'`

**Fix Options**:
1. Remove the call and `_process_assets` function
2. Implement `validate_assets()` in core
3. Implement asset validation inline in the tool

### Issue 2: Missing Presentation Version Tracking
**Severity**: 🔴 Critical (per governance protocol)

```python
# Should capture
version_before = agent.get_presentation_version()
# ... operations ...
# Include in output
"presentation_version": version_before
```

## ⚠️ Warnings

### Warning 1: Unused `filepath` Parameter in Helpers
```python
def _process_core_validation(val, issues, summary, filepath):  # filepath unused
def _process_accessibility(acc, issues, summary, filepath):    # filepath unused
def _process_assets(assets, issues, summary, filepath):        # filepath unused
```

### Warning 2: Exit Code Doesn't Match Matrix
```python
sys.exit(1 if result["status"] in ("critical", "failed") else 0)
```
Per matrix, validation errors should be exit code 2, not 1.

### Warning 3: `CORE_VERSION` Imported But Never Used
```python
try:
    from core.powerpoint_agent_core import (
        ...,
        __version__ as CORE_VERSION  # Never used
    )
except ImportError:
    CORE_VERSION = "0.0.0"  # Never used
```

### Warning 4: Incomplete Design Rule Validation
```python
def _validate_design_rules(agent, issues, summary, policy, filepath):
    # Only checks font count
    if summary.fonts_used > policy.thresholds.get("max_fonts", 5):
        issues.append(...)
    # Missing: 6x6 rule, color count, other design checks
```

Despite having `enforce_6x6_rule` in policy, it's never checked.

### Warning 5: `colors_detected` Always 0
```python
@dataclass
class ValidationSummary:
    colors_detected: int = 0  # Never populated!
```

### Warning 6: Missing `operation_id` and `duration_ms`
Unlike probe tool, no operation tracking metadata.

### Warning 7: Incomplete Error Response
```python
error_res = {
    "status": "error", 
    "error": str(e), 
    "error_type": type(e).__name__
    # Missing: "suggestion" field per standard
}
```

### Warning 8: `_generate_recommendations` Logic Flawed
```python
if any(i.details.get("issue_type") == "empty_slide" for i in issues):
```
But `ValidationIssue` for empty slides doesn't set `details["issue_type"]`:
```python
issues.append(ValidationIssue("structure", "critical", "Empty slide", slide_index=idx))
# details is empty dict!
```

---

# Consolidated Fix Recommendations

## Priority 1: Critical Fixes (Blocking)

| Tool | Fix Required |
|------|--------------|
| `ppt_capability_probe.py` | Add stderr hygiene block at top |
| `ppt_validate_presentation.py` | Remove or implement `validate_assets()` call |
| `ppt_json_adapter.py` | Fix error template and output format |
| `ppt_check_accessibility.py` | Add standard output wrapper with status field |

## Priority 2: High (Governance Compliance)

| Tool | Fix Required |
|------|--------------|
| All 4 tools | Add `presentation_version` to output |
| All 4 tools | Standardize error response format |
| `ppt_check_accessibility.py` | Fix ImportError handling |

## Priority 3: Medium (Standards Alignment)

| Tool | Fix Required |
|------|--------------|
| `ppt_capability_probe.py` | Update version to 3.1.x, fix bare excepts |
| `ppt_json_adapter.py` | Add hygiene block, use `sys.stdout.write()` |
| `ppt_validate_presentation.py` | Implement full design rule validation |
| All tools | Add `operation_id`, `duration_ms` metadata |

---

# Proposed Fix: `ppt_capability_probe.py`

```python
#!/usr/bin/env python3
"""
PowerPoint Capability Probe Tool v3.1.0
...
"""

import sys
import os

# --- HYGIENE BLOCK START ---
# CRITICAL: Redirect stderr to /dev/null immediately to prevent library noise
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

import json
# ... rest of imports ...

__version__ = "3.1.0"

# Fix bare excepts
def get_library_versions() -> Dict[str, str]:
    versions = {}
    try:
        versions["python-pptx"] = importlib.metadata.version("python-pptx")
    except importlib.metadata.PackageNotFoundError:
        versions["python-pptx"] = "not_installed"
    except Exception as e:
        versions["python-pptx"] = f"error:{type(e).__name__}"
    # ... similar for Pillow
    return versions
```

---

# Proposed Fix: `ppt_check_accessibility.py`

```python
def check_accessibility(filepath: Path) -> Dict[str, Any]:
    """Run accessibility checks with proper output wrapping."""
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath, acquire_lock=False)
        
        # Capture version for audit trail
        version = agent.get_presentation_version()
        
        result = agent.check_accessibility()
    
    # Wrap in standard format
    return {
        "status": "success",
        "file": str(filepath),
        "tool_version": __version__,
        "presentation_version": version,
        "accessibility_issues": result.get("issues", {}),
        "summary": {
            "missing_alt_text_count": len(result.get("issues", {}).get("missing_alt_text", [])),
            "low_contrast_count": len(result.get("issues", {}).get("low_contrast", [])),
            "total_issues": sum(len(v) for v in result.get("issues", {}).values() if isinstance(v, list))
        },
        "passed": result.get("passed", True)
    }
```

---

## Validation Complete

**Summary**: 4 tools analyzed, 7 critical issues identified, 23 warnings documented. Primary blockers are:
1. Missing hygiene block in probe tool
2. Undefined method call in validation tool
3. Broken error handling in JSON adapter

**Ready to implement fixes upon approval.**
