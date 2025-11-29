# 🔍 Meticulous Tool Review & Validation Report

I have completed a comprehensive review of the three tools against the project documentation, core library handbook, and governance requirements.

---

## 📊 Executive Summary

| Tool | Compliance | Critical Issues | Important Issues | Status |
|------|------------|-----------------|------------------|--------|
| `ppt_capability_probe.py` | 55% | 3 | 6 | ⚠️ Needs Fixes |
| `ppt_check_accessibility.py` | 75% | 1 | 4 | ⚠️ Needs Fixes |
| `ppt_create_from_template.py` | 90% | 0 | 3 | ✅ Good |

---

## 1. ppt_capability_probe.py

### 1.1 Tool Classification
| Attribute | Value |
|-----------|-------|
| **Primary Phase** | DISCOVER |
| **Destructive** | No (read-only with atomic verification) |
| **Requires Approval Token** | No |
| **Invalidates Indices** | No |

### 1.2 Compliance Checklist

| Requirement | Status | Notes |
|-------------|--------|-------|
| Hygiene Block (stderr redirect) | ❌ **MISSING** | Critical for JSON pipeline safety |
| JSON-only stdout | ❌ **BROKEN** | Uses `print()` instead of `sys.stdout.write()` |
| Exit codes (0/1) | ✅ | Correct |
| File existence check | ✅ | Present |
| Path validation (pathlib) | ✅ | Uses `Path` |
| Context manager usage | ⚠️ N/A | Uses `Presentation()` directly (acceptable for read-only) |
| Version tracking | ⚠️ Misaligned | `tool_version: "1.1.1"` should be `"3.1.1"` |
| Tool version in output | ✅ | Present but wrong version |
| Error format with suggestion | ❌ **MISSING** | No `suggestion` field in errors |
| Timeout protection | ✅ | Has timeout handling |
| Atomic verification | ✅ | Has checksum verification |
| Transient slide pattern | ✅ | Correctly implemented |

### 1.3 Critical Issues

#### Issue 1: Missing Hygiene Block
```python
# ❌ CURRENT (lines 1-24): No stderr redirection
#!/usr/bin/env python3
"""
PowerPoint Capability Probe Tool
...
"""

import sys
import json
import argparse
...
```

**This is the highest priority fix.** Any library warnings will corrupt JSON output.

**Fix Required:**
```python
#!/usr/bin/env python3
"""
PowerPoint Capability Probe Tool
...
"""

import sys
import os

# --- HYGIENE BLOCK START ---
# CRITICAL: Redirect stderr to /dev/null immediately to prevent library noise.
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

import json
import argparse
...
```

#### Issue 2: Import Error Prints Before Hygiene
```python
# ❌ CURRENT (lines 25-30): Prints before hygiene could be established
try:
    from pptx import Presentation
    from pptx.enum.shapes import PP_PLACEHOLDER
except ImportError:
    print(json.dumps({  # ❌ This print pollutes stdout
        "status": "error",
        "error": "python-pptx not installed",
        "error_type": "ImportError"
    }, indent=2))
    sys.exit(1)
```

**Fix Required:** Move import error handling after hygiene block and use `sys.stdout.write()`.

#### Issue 3: Uses print() Instead of sys.stdout.write()
```python
# ❌ CURRENT (lines 826-828)
if args.summary:
    print(format_summary(result))
else:
    print(json.dumps(result, indent=2))
```

**Fix Required:**
```python
# ✅ FIXED
if args.summary:
    sys.stdout.write(format_summary(result) + "\n")
else:
    sys.stdout.write(json.dumps(result, indent=2) + "\n")
sys.stdout.flush()
```

### 1.4 Important Issues

#### Issue 4: Version Misalignment
```python
# ⚠️ CURRENT: Uses 1.1.1 but project is 3.1.x
"tool_version": "1.1.1",
"schema_version": "capability_probe.v1.1.1",
```

**Fix:** Align to project versioning scheme (`"3.1.1"`).

#### Issue 5: Missing Suggestion in Error Response
```python
# ⚠️ CURRENT (line 839)
error_result = {
    "status": "error",
    "error": str(e),
    "error_type": type(e).__name__,
    # Missing: "suggestion": "..."
}
```

#### Issue 6: Bare Except Clauses
Multiple locations use `except:` without specifying exception type:
- Line 61, 85, 108, 115, 132, 175, 238, 276, 296, 311, 425, 488, 551, 614

**Fix:** Replace with `except Exception:` at minimum, or specific exception types.

#### Issue 7: Schema Path Hardcoded and No Graceful Handling
```python
# ⚠️ Lines 813-821: Schema validation assumes file exists
schema_path = Path(__file__).parent.parent / "schemas" / "capability_probe.v1.1.1.schema.json"
validate_against_schema(result, str(schema_path))
```

**Fix:** Wrap in try/except with graceful fallback:
```python
try:
    schema_path = Path(__file__).parent.parent / "schemas" / "capability_probe.v1.1.1.schema.json"
    if schema_path.exists():
        validate_against_schema(result, str(schema_path))
except FileNotFoundError:
    warnings.append("Schema file not found - skipping strict validation")
except Exception as e:
    warnings.append(f"Schema validation skipped: {str(e)}")
```

#### Issue 8: Library Version Detection Uses Bare Except
```python
# ⚠️ Lines 45-56: Uses bare except
def get_library_versions() -> Dict[str, str]:
    versions = {}
    try:
        versions["python-pptx"] = importlib.metadata.version("python-pptx")
    except:  # ❌ Bare except
        versions["python-pptx"] = "unknown"
```

#### Issue 9: Missing sys.stdout.flush()
Output is written but not flushed, which can cause issues in pipelines.

---

## 2. ppt_check_accessibility.py

### 2.1 Tool Classification
| Attribute | Value |
|-----------|-------|
| **Primary Phase** | VALIDATE |
| **Destructive** | No (read-only) |
| **Requires Approval Token** | No |
| **Invalidates Indices** | No |

### 2.2 Compliance Checklist

| Requirement | Status | Notes |
|-------------|--------|-------|
| Hygiene Block (stderr redirect) | ✅ | Correctly implemented |
| JSON-only stdout | ✅ | Uses `sys.stdout.write()` |
| Exit codes (0/1) | ✅ | Correct |
| File existence check | ✅ | Present |
| Path validation (pathlib) | ✅ | Uses `Path` |
| Context manager usage | ✅ | Uses `with PowerPointAgent()` |
| Version tracking | ❌ **MISSING** | No `presentation_version` in output |
| Tool version in output | ❌ **MISSING** | No `__version__` constant |
| Error format with suggestion | ❌ **MISSING** | No `suggestion` field |
| acquire_lock documentation | ✅ | Has comment (line 47) |

### 2.3 Critical Issue

#### Issue 1: Missing __version__ and Version Tracking
```python
# ❌ CURRENT: No version constant, no version in output
def check_accessibility(filepath: Path) -> Dict[str, Any]:
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath, acquire_lock=False)
        result = agent.check_accessibility()
        result["file"] = str(filepath)
        # Missing: presentation_version, tool_version
    return result
```

**Fix Required:**
```python
__version__ = "3.1.1"

def check_accessibility(filepath: Path) -> Dict[str, Any]:
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath, acquire_lock=False)  # Read-only, no lock needed
        result = agent.check_accessibility()
        presentation_version = agent.get_presentation_version()
    
    result["file"] = str(filepath.resolve())
    result["presentation_version"] = presentation_version
    result["tool_version"] = __version__
    return result
```

### 2.4 Important Issues

#### Issue 2: Incomplete Error Response
```python
# ⚠️ CURRENT: Missing suggestion
error_result = {
    "status": "error",
    "error": str(e),
    "error_type": type(e).__name__
}
```

**Fix:**
```python
error_result = {
    "status": "error",
    "error": str(e),
    "error_type": type(e).__name__,
    "suggestion": "Verify file exists and is a valid .pptx file",
    "tool_version": __version__
}
```

#### Issue 3: No sys.stdout.flush()
```python
# ⚠️ CURRENT: No flush after write
sys.stdout.write(json.dumps(result, indent=2) + "\n")
# Missing: sys.stdout.flush()
```

#### Issue 4: Minimal Help Text
```python
# ⚠️ CURRENT: No epilog with examples
parser = argparse.ArgumentParser(
    description="Check PowerPoint accessibility (WCAG 2.1)",
    formatter_class=argparse.RawDescriptionHelpFormatter
)
```

**Recommendation:** Add epilog with usage examples and output format description.

#### Issue 5: Status Field Passthrough
The tool just passes through the core result. If the core returns a result without a `status` field, the output will be inconsistent.

**Fix:** Ensure status is always set:
```python
if "status" not in result:
    result["status"] = "success" if result.get("total_issues", 0) == 0 else "issues_found"
```

---

## 3. ppt_create_from_template.py

### 3.1 Tool Classification
| Attribute | Value |
|-----------|-------|
| **Primary Phase** | CREATE |
| **Destructive** | No (creates new file) |
| **Requires Approval Token** | No |
| **Invalidates Indices** | No (new presentation) |

### 3.2 Compliance Checklist

| Requirement | Status | Notes |
|-------------|--------|-------|
| Hygiene Block (stderr redirect) | ✅ | Correctly implemented |
| JSON-only stdout | ✅ | Uses `sys.stdout.write()` |
| Exit codes (0/1) | ✅ | Correct |
| File existence check | ✅ | Present |
| Path validation (pathlib) | ✅ | Uses `Path` |
| Context manager usage | ✅ | Uses `with PowerPointAgent()` |
| Version tracking | ✅ | Has `presentation_version` |
| Tool version in output | ✅ | Has `__version__` and in output |
| Error format with suggestion | ⚠️ Partial | Most have suggestion, some don't |
| Comprehensive docstrings | ✅ | Excellent documentation |
| Help text with examples | ✅ | Very thorough |

### 3.3 This Tool is Well-Implemented ✅

This tool demonstrates good practices and can serve as a reference. Only minor issues.

### 3.4 Important Issues (Minor)

#### Issue 1: Missing Suggestion in Some Error Handlers
```python
# ⚠️ Line 218: PowerPointAgentError missing suggestion
except PowerPointAgentError as e:
    error_result = {
        "status": "error",
        "error": str(e),
        "error_type": type(e).__name__,
        "details": getattr(e, 'details', {})
        # Missing: "suggestion": "..."
    }

# ⚠️ Line 226: Generic exception missing suggestion
except Exception as e:
    error_result = {
        "status": "error",
        "error": str(e),
        "error_type": type(e).__name__,
        "tool_version": __version__
        # Missing: "suggestion": "..."
    }
```

**Fix:**
```python
except PowerPointAgentError as e:
    error_result = {
        "status": "error",
        "error": str(e),
        "error_type": type(e).__name__,
        "details": getattr(e, 'details', {}),
        "suggestion": "Check template file integrity and available layouts",
        "tool_version": __version__
    }

except Exception as e:
    error_result = {
        "status": "error",
        "error": str(e),
        "error_type": type(e).__name__,
        "suggestion": "Check logs for detailed error information",
        "tool_version": __version__
    }
```

#### Issue 2: Missing sys.stdout.flush()
```python
# ⚠️ Multiple locations: No flush after write
sys.stdout.write(json.dumps(result, indent=2) + "\n")
# Should add: sys.stdout.flush()
```

#### Issue 3: Potential Method Availability Issue
```python
# ⚠️ Line 111: Assumes method exists
available_layouts = agent.get_available_layouts()
```

**Recommendation:** Wrap in try/except or verify method exists:
```python
try:
    available_layouts = agent.get_available_layouts()
except AttributeError:
    # Fallback: get from presentation info
    info = agent.get_presentation_info()
    available_layouts = info.get("layouts", [])
```

---

## 📋 Consolidated Fix Priority Matrix

### 🔴 Critical (Must Fix Before Production)

| Tool | Issue | Fix Effort |
|------|-------|------------|
| `ppt_capability_probe.py` | Missing hygiene block | 5 lines |
| `ppt_capability_probe.py` | Import error uses print() | 10 lines |
| `ppt_capability_probe.py` | Uses print() throughout | 10 lines |
| `ppt_check_accessibility.py` | Missing __version__ and version tracking | 10 lines |

### 🟡 Important (Should Fix)

| Tool | Issue | Fix Effort |
|------|-------|------------|
| `ppt_capability_probe.py` | Version misalignment (1.1.1 → 3.1.1) | 5 lines |
| `ppt_capability_probe.py` | Missing suggestion in errors | 5 lines |
| `ppt_capability_probe.py` | Bare except clauses | 20 lines |
| `ppt_check_accessibility.py` | Missing suggestion in errors | 5 lines |
| `ppt_check_accessibility.py` | Status field passthrough | 3 lines |
| `ppt_create_from_template.py` | Missing suggestion in some errors | 5 lines |
| All 3 tools | Missing sys.stdout.flush() | 1 line each |

### 🟢 Minor (Nice to Have)

| Tool | Issue | Fix Effort |
|------|-------|------------|
| `ppt_capability_probe.py` | Schema path hardcoded | 10 lines |
| `ppt_check_accessibility.py` | Minimal help text | 20 lines |
| `ppt_create_from_template.py` | Method availability check | 5 lines |

---

## ✅ Validation Summary

| Aspect | ppt_capability_probe | ppt_check_accessibility | ppt_create_from_template |
|--------|---------------------|------------------------|--------------------------|
| **Hygiene Block** | ❌ | ✅ | ✅ |
| **JSON Output** | ❌ Uses print() | ✅ | ✅ |
| **Exit Codes** | ✅ | ✅ | ✅ |
| **Version Tracking** | ⚠️ Wrong version | ❌ | ✅ |
| **Error Format** | ❌ Missing suggestion | ❌ Missing suggestion | ⚠️ Partial |
| **Context Manager** | N/A (read-only) | ✅ | ✅ |
| **Path Safety** | ✅ | ✅ | ✅ |
| **Documentation** | ✅ Excellent | ⚠️ Minimal | ✅ Excellent |
| **Timeout Handling** | ✅ | N/A | N/A |
| **Atomic Verification** | ✅ | N/A | N/A |

---

## 🎯 Recommendations

### For ppt_capability_probe.py:
1. **Priority 1:** Add hygiene block immediately after shebang
2. **Priority 2:** Replace all `print()` with `sys.stdout.write()` + flush
3. **Priority 3:** Update version to 3.1.1 for consistency
4. **Priority 4:** Add suggestion field to error responses
5. **Priority 5:** Replace bare `except:` with `except Exception:`

### For ppt_check_accessibility.py:
1. **Priority 1:** Add `__version__` constant
2. **Priority 2:** Add `presentation_version` and `tool_version` to output
3. **Priority 3:** Add suggestion field to error response
4. **Priority 4:** Ensure status field is always present
5. **Priority 5:** Add helpful epilog to argparser

### For ppt_create_from_template.py:
1. **Priority 1:** Add suggestion to remaining error handlers
2. **Priority 2:** Add sys.stdout.flush() after writes
3. **Priority 3:** Add method availability check for get_available_layouts

---

Would you like me to proceed with creating complete validated replacement files for these tools, following the same meticulous approach used for the previous batch?
