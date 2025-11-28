# Comprehensive Tool Review & Validation

I have meticulously reviewed all four tools against the established project standards. Below is my detailed analysis with specific findings, severity ratings, and remediation recommendations.

---

## Executive Summary

| Tool | Compliance Score | Critical Issues | Warnings | Status |
|------|------------------|-----------------|----------|--------|
| `ppt_add_notes.py` | 65% | 3 | 5 | ⚠️ Needs Fixes |
| `ppt_create_from_structure.py` | 60% | 4 | 6 | ⚠️ Needs Fixes |
| `ppt_extract_notes.py` | 55% | 3 | 4 | ⚠️ Needs Fixes |
| `ppt_set_background.py` | 50% | 5 | 5 | 🔴 Significant Issues |

**Common Issues Across All Tools:**
1. ❌ Missing Hygiene Block (stderr redirect)
2. ❌ Missing Version Tracking (presentation_version_before/after)
3. ❌ Missing `__version__` constant
4. ❌ Incomplete error response format (missing `suggestion` field)
5. ❌ Missing Clone-Before-Edit validation

---

# Tool 1: `ppt_add_notes.py`

## 1.1 Compliance Matrix

| Requirement | Status | Details |
|-------------|--------|---------|
| Hygiene Block (stderr redirect) | ❌ Missing | Must be FIRST before any imports |
| Context Manager Pattern | ✅ Pass | Correctly uses `with PowerPointAgent()` |
| Version Tracking | ❌ Missing | No `presentation_version_before/after` |
| JSON Output Only | ⚠️ Warning | Line 46 prints to stderr (acceptable but risky) |
| Exit Codes | ✅ Pass | Uses 0/1 correctly |
| Path Validation | ✅ Pass | Uses `pathlib.Path` |
| File Extension Check | ❌ Wrong | Accepts `.ppt` which python-pptx doesn't support |
| Clone-Before-Edit Check | ❌ Missing | No `/work/` directory validation |
| Error Response Format | ⚠️ Incomplete | Missing `suggestion` field |
| Type Hints | ⚠️ Partial | Return type could be more specific |
| Docstrings | ⚠️ Incomplete | Missing Returns/Raises sections |
| Tool Version | ❌ Missing | No `__version__` constant |

## 1.2 Specific Issues

### Issue 1: Missing Hygiene Block (CRITICAL)
**Location:** Top of file
**Problem:** No stderr redirect before imports
**Impact:** Library warnings could pollute JSON output

```python
# ❌ CURRENT (Line 1-17)
#!/usr/bin/env python3
"""..."""
import sys
import json
...

# ✅ REQUIRED (Must be FIRST)
#!/usr/bin/env python3
"""..."""
import sys
import os

# --- HYGIENE BLOCK START ---
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

import json
...
```

### Issue 2: Wrong File Extension Validation (CRITICAL)
**Location:** Line 30-31
**Problem:** Accepts `.ppt` format which python-pptx cannot handle

```python
# ❌ CURRENT
if not filepath.suffix.lower() in ['.pptx', '.ppt']:
    raise ValueError("Invalid PowerPoint file format (must be .pptx or .ppt)")

# ✅ CORRECT
if not filepath.suffix.lower() == '.pptx':
    raise ValueError("Invalid file format. Only .pptx files are supported (not .ppt)")
```

### Issue 3: Missing Version Tracking (CRITICAL)
**Location:** Return statement (line 71-79)
**Problem:** No version hashes for audit trail

```python
# ❌ CURRENT return
return {
    "status": "success",
    "file": str(filepath),
    ...
}

# ✅ REQUIRED
with PowerPointAgent(filepath) as agent:
    agent.open(filepath)
    
    # Capture version BEFORE
    version_before = agent.get_presentation_version()
    
    # ... existing logic ...
    
    agent.save()
    
    # Capture version AFTER
    version_after = agent.get_presentation_version()

return {
    "status": "success",
    "file": str(filepath),
    "presentation_version_before": version_before,
    "presentation_version_after": version_after,
    ...
}
```

### Issue 4: Direct Presentation Access (WARNING)
**Location:** Line 48
**Problem:** Uses `agent.prs.slides[slide_index]` instead of core method

```python
# ❌ CURRENT
slide = agent.prs.slides[slide_index]

# ✅ PREFERRED (if core method exists)
# Use agent.get_slide(slide_index) or at minimum validate bounds first
if not 0 <= slide_index < slide_count:
    raise SlideNotFoundError(...)
slide = agent.prs.slides[slide_index]
```

### Issue 5: Missing Clone-Before-Edit Check (WARNING)
**Location:** After filepath validation
**Problem:** Allows editing of source files directly

```python
# ✅ ADD after line 33
# Clone-before-edit check (optional but recommended)
filepath_str = str(filepath.resolve())
if not ('/work/' in filepath_str or 'work_' in filepath_str or '/tmp/' in filepath_str):
    import warnings
    warnings.warn("Editing file outside /work/ directory. Consider cloning first.", UserWarning)
```

## 1.3 Complete Remediated Version

```python
#!/usr/bin/env python3
"""
PowerPoint Add Speaker Notes Tool v3.1.0
Add, append, or overwrite speaker notes for a specific slide.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    # Append note (default)
    uv run tools/ppt_add_notes.py --file deck.pptx --slide 0 --text "Talk about Q4 growth." --json
    
    # Overwrite existing notes
    uv run tools/ppt_add_notes.py --file deck.pptx --slide 0 --text "New script." --mode overwrite --json

Exit Codes:
    0: Success
    1: Error occurred
"""

import sys
import os

# --- HYGIENE BLOCK START ---
# CRITICAL: Redirect stderr to /dev/null immediately to prevent library noise
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

import json
import argparse
from pathlib import Path
from typing import Dict, Any

# Add parent directory to path for core import
sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError, SlideNotFoundError
)

# ============================================================================
# CONSTANTS
# ============================================================================

__version__ = "3.1.0"


def add_notes(
    filepath: Path,
    slide_index: int,
    text: str,
    mode: str = "append"
) -> Dict[str, Any]:
    """
    Add speaker notes to a slide.
    
    Args:
        filepath: Path to PowerPoint file (must be .pptx)
        slide_index: Index of slide to modify (0-based)
        text: Text to add to notes
        mode: Insertion mode - 'append' (default), 'prepend', or 'overwrite'
        
    Returns:
        Dict containing:
            - status: 'success'
            - file: Absolute file path
            - slide_index: Target slide index
            - mode: Mode used
            - original_length: Length of original notes
            - new_length: Length of final notes
            - preview: First 100 chars of final notes
            - presentation_version_before: Version hash before changes
            - presentation_version_after: Version hash after changes
            - tool_version: Tool version string
            
    Raises:
        FileNotFoundError: If file doesn't exist
        ValueError: If file format invalid or text empty
        SlideNotFoundError: If slide index out of range
        PowerPointAgentError: If notes access fails
    """
    
    # Validate file extension (python-pptx only supports .pptx)
    if filepath.suffix.lower() != '.pptx':
        raise ValueError("Invalid file format. Only .pptx files are supported (not .ppt)")

    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if not text or not text.strip():
        raise ValueError("Notes text cannot be empty")
    
    if mode not in ('append', 'prepend', 'overwrite'):
        raise ValueError(f"Invalid mode '{mode}'. Must be 'append', 'prepend', or 'overwrite'")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Capture version BEFORE changes
        version_before = agent.get_presentation_version()
        
        # Validate slide index
        slide_count = agent.get_slide_count()
        if not 0 <= slide_index < slide_count:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range",
                details={"requested": slide_index, "available": slide_count}
            )
            
        slide = agent.prs.slides[slide_index]
        
        # Access or create notes slide
        try:
            notes_slide = slide.notes_slide
            text_frame = notes_slide.notes_text_frame
        except Exception as e:
            raise PowerPointAgentError(f"Failed to access notes slide: {str(e)}")
        
        original_text = text_frame.text if text_frame.text else ""
        
        # Apply mode
        if mode == "overwrite":
            final_text = text
        elif mode == "append":
            if original_text.strip():
                final_text = original_text + "\n" + text
            else:
                final_text = text
        elif mode == "prepend":
            if original_text.strip():
                final_text = text + "\n" + original_text
            else:
                final_text = text
        
        text_frame.text = final_text
        
        agent.save()
        
        # Capture version AFTER changes
        version_after = agent.get_presentation_version()
        
    return {
        "status": "success",
        "file": str(filepath.resolve()),
        "slide_index": slide_index,
        "mode": mode,
        "original_length": len(original_text),
        "new_length": len(final_text),
        "preview": final_text[:100] + "..." if len(final_text) > 100 else final_text,
        "presentation_version_before": version_before,
        "presentation_version_after": version_after,
        "tool_version": __version__
    }


def main():
    parser = argparse.ArgumentParser(
        description="Add speaker notes to PowerPoint slide",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=f"""
Examples:
    # Append notes to slide 0
    uv run tools/ppt_add_notes.py --file deck.pptx --slide 0 --text "Key point here" --json
    
    # Overwrite existing notes
    uv run tools/ppt_add_notes.py --file deck.pptx --slide 0 --text "New script" --mode overwrite --json
    
    # Prepend notes
    uv run tools/ppt_add_notes.py --file deck.pptx --slide 0 --text "IMPORTANT:" --mode prepend --json

Version: {__version__}
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file path (.pptx only)'
    )
    
    parser.add_argument(
        '--slide',
        required=True,
        type=int,
        help='Slide index (0-based)'
    )
    
    parser.add_argument(
        '--text',
        required=True,
        help='Notes content to add'
    )
    
    parser.add_argument(
        '--mode',
        choices=['append', 'overwrite', 'prepend'],
        default='append',
        help='Insertion mode (default: append)'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        result = add_notes(
            filepath=args.file,
            slide_index=args.slide,
            text=args.text,
            mode=args.mode
        )
        print(json.dumps(result, indent=2))
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify the file path exists and is accessible"
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
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
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Check input parameters match expected format"
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "PowerPointAgentError",
            "suggestion": "Check file is not corrupted and slide has valid structure"
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "tool_version": __version__
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()
```

---

# Tool 2: `ppt_create_from_structure.py`

## 2.1 Compliance Matrix

| Requirement | Status | Details |
|-------------|--------|---------|
| Hygiene Block | ❌ Missing | Must be FIRST before any imports |
| Context Manager Pattern | ✅ Pass | Correctly uses `with PowerPointAgent()` |
| Version Tracking | ❌ Missing | No `presentation_version` in response |
| JSON Output Only | ❌ Fail | Non-JSON mode prints to stdout |
| Exit Codes | ✅ Pass | Uses 0/1 correctly |
| Path Validation | ✅ Pass | Uses `pathlib.Path` |
| File Extension Check | ⚠️ Auto-correct | Silently corrects extension |
| Clone-Before-Edit Check | N/A | Creates new file |
| Error Response Format | ⚠️ Incomplete | Missing `suggestion` field |
| Type Hints | ✅ Pass | Properly typed |
| Docstrings | ⚠️ Minimal | Missing detailed docstrings |
| Tool Version | ❌ Missing | No `__version__` constant |
| Slide Count Limit | ⚠️ Weak | Validates 100 but no timeout handling |

## 2.2 Specific Issues

### Issue 1: Missing Hygiene Block (CRITICAL)
**Location:** Top of file
**Same as Tool 1**

### Issue 2: Missing Version Tracking (CRITICAL)
**Location:** Return statement (line 124-135)

```python
# ✅ ADD to return statement
return {
    "status": "success" if len(stats["errors"]) == 0 else "success_with_errors",
    "file": str(output),
    "presentation_version": agent.get_presentation_version(),  # ADD THIS
    "slides_created": stats["slides_created"],
    ...
}
```

### Issue 3: Non-JSON Output Mode (CRITICAL)
**Location:** Lines 211-218
**Problem:** Printing to stdout breaks JSON-only contract

```python
# ❌ CURRENT
if args.json:
    print(json.dumps(result, indent=2))
else:
    print(f"✅ Created presentation: {result['file']}")  # Pollutes stdout!
    
# ✅ FIX - Always output JSON, use stderr for human messages
print(json.dumps(result, indent=2))
# Human-readable mode could write to stderr if needed
```

### Issue 4: Silent Extension Correction (WARNING)
**Location:** Line 196-197

```python
# ❌ CURRENT - Silent correction
if not args.output.suffix.lower() == '.pptx':
    args.output = args.output.with_suffix('.pptx')

# ✅ BETTER - Warn or reject
if args.output.suffix.lower() != '.pptx':
    raise ValueError(f"Output must have .pptx extension, got: {args.output.suffix}")
```

### Issue 5: Index Refresh After Add Operations (WARNING)
**Location:** Content creation loop (lines 63-109)
**Problem:** Adding multiple items doesn't track shape indices

```python
# ⚠️ If subsequent operations reference shapes by index,
# indices should be refreshed. Currently acceptable since
# each item is independent, but worth documenting.
```

## 2.3 Key Remediation Points

```python
#!/usr/bin/env python3
"""..."""
import sys
import os

# --- HYGIENE BLOCK START ---
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

import json
...

__version__ = "3.1.0"

# In return statement, add:
# "presentation_version": version_after,
# "tool_version": __version__

# Remove non-JSON output mode or redirect to stderr
```

---

# Tool 3: `ppt_extract_notes.py`

## 3.1 Compliance Matrix

| Requirement | Status | Details |
|-------------|--------|---------|
| Hygiene Block | ❌ Missing | Must be FIRST |
| Context Manager Pattern | ✅ Pass | Correctly uses context manager |
| Version Tracking | ❌ Missing | Should include presentation_version |
| JSON Output Only | ✅ Pass | Only outputs JSON |
| Exit Codes | ✅ Pass | Uses 0/1 |
| Path Validation | ✅ Pass | Uses pathlib |
| File Extension Check | ❌ Missing | No validation |
| Read-Only Lock | ✅ Pass | Uses `acquire_lock=False` |
| Error Response Format | ⚠️ Minimal | No suggestion field |
| Type Hints | ✅ Pass | Properly typed |
| Docstrings | ❌ Minimal | Only usage comment |
| Tool Version | ❌ Missing | No `__version__` |

## 3.2 Specific Issues

### Issue 1: Missing Hygiene Block (CRITICAL)
**Same as others**

### Issue 2: Missing Version Tracking (MODERATE)
**Location:** Return statement

```python
# ✅ ADD
return {
    "status": "success",
    "file": str(filepath.resolve()),
    "presentation_version": agent.get_presentation_version(),  # ADD
    "notes_found": len(notes),
    "notes": notes,
    "tool_version": __version__
}
```

### Issue 3: Missing File Extension Validation (MODERATE)
**Location:** After filepath.exists() check

```python
# ✅ ADD
if filepath.suffix.lower() != '.pptx':
    raise ValueError("Only .pptx files are supported")
```

### Issue 4: Minimal Docstring (WARNING)

```python
# ✅ EXPAND
def extract_notes(filepath: Path) -> Dict[str, Any]:
    """
    Extract speaker notes from all slides in a presentation.
    
    Args:
        filepath: Path to PowerPoint file (.pptx)
        
    Returns:
        Dict containing:
            - status: 'success'
            - file: Absolute file path
            - presentation_version: State hash
            - notes_found: Count of slides with notes
            - notes: Dict mapping slide index to notes text
            - tool_version: Tool version
            
    Raises:
        FileNotFoundError: If file doesn't exist
        ValueError: If file format is invalid
    """
```

## 3.3 Complete Remediated Version

```python
#!/usr/bin/env python3
"""
PowerPoint Extract Notes Tool v3.1.0
Extract speaker notes from all slides in a presentation.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_extract_notes.py --file presentation.pptx --json

Exit Codes:
    0: Success
    1: Error occurred
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

from core.powerpoint_agent_core import PowerPointAgent

__version__ = "3.1.0"


def extract_notes(filepath: Path) -> Dict[str, Any]:
    """
    Extract speaker notes from all slides in a presentation.
    
    Args:
        filepath: Path to PowerPoint file (.pptx only)
        
    Returns:
        Dict containing:
            - status: 'success'
            - file: Absolute file path
            - presentation_version: State hash of presentation
            - notes_found: Count of slides with notes
            - notes: Dict mapping slide index (int) to notes text (str)
            - tool_version: Tool version string
            
    Raises:
        FileNotFoundError: If file doesn't exist
        ValueError: If file format is not .pptx
    """
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if filepath.suffix.lower() != '.pptx':
        raise ValueError("Only .pptx files are supported")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath, acquire_lock=False)  # Read-only, no lock needed
        
        version = agent.get_presentation_version()
        notes = agent.extract_notes()
        
    return {
        "status": "success",
        "file": str(filepath.resolve()),
        "presentation_version": version,
        "notes_found": len([v for v in notes.values() if v and v.strip()]),
        "total_slides": len(notes),
        "notes": notes,
        "tool_version": __version__
    }


def main():
    parser = argparse.ArgumentParser(
        description="Extract speaker notes from PowerPoint",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=f"""
Examples:
    uv run tools/ppt_extract_notes.py --file presentation.pptx --json

Version: {__version__}
        """
    )
    parser.add_argument('--file', required=True, type=Path, help='PowerPoint file path')
    parser.add_argument('--json', action='store_true', default=True, help='Output JSON')
    
    args = parser.parse_args()
    
    try:
        result = extract_notes(filepath=args.file)
        print(json.dumps(result, indent=2))
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify the file path exists and is accessible"
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Ensure file has .pptx extension"
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "tool_version": __version__
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()
```

---

# Tool 4: `ppt_set_background.py`

## 4.1 Compliance Matrix

| Requirement | Status | Details |
|-------------|--------|---------|
| Hygiene Block | ❌ Missing | Must be FIRST |
| Context Manager Pattern | ✅ Pass | Uses context manager |
| Version Tracking | ❌ Missing | No version hashes |
| JSON Output Only | ✅ Pass | Only JSON output |
| Exit Codes | ✅ Pass | Uses 0/1 |
| Path Validation | ✅ Pass | Uses pathlib |
| File Extension Check | ❌ Missing | No validation |
| Clone-Before-Edit Check | ❌ Missing | **CRITICAL** - modifies files |
| Slide Index Validation | ❌ Missing | Direct index access could crash |
| Error Response Format | ⚠️ Minimal | No suggestion field |
| Type Hints | ✅ Pass | Properly typed |
| Docstrings | ❌ Minimal | Only usage comment |
| Tool Version | ❌ Missing | No `__version__` |
| Approval Token | ⚠️ Consider | All-slides mode is potentially destructive |

## 4.2 Specific Issues

### Issue 1: Missing Slide Index Validation (CRITICAL)
**Location:** Line 31
**Problem:** Direct index access without bounds checking

```python
# ❌ CURRENT
if slide_index is not None:
    target_slides = [agent.prs.slides[slide_index]]  # Could crash!

# ✅ REQUIRED
if slide_index is not None:
    slide_count = agent.get_slide_count()
    if not 0 <= slide_index < slide_count:
        raise SlideNotFoundError(
            f"Slide index {slide_index} out of range",
            details={"requested": slide_index, "available": slide_count}
        )
    target_slides = [agent.prs.slides[slide_index]]
```

### Issue 2: Missing Clone-Before-Edit Check (CRITICAL)
**Location:** Before file operations
**Problem:** Allows modifying source files directly

```python
# ✅ ADD after filepath validation
# Clone-before-edit warning (for mutation tools)
filepath_resolved = str(filepath.resolve())
if not any(pattern in filepath_resolved for pattern in ['/work/', 'work_', '/tmp/', '\\work\\']):
    # Log warning to stderr (not stdout!)
    import sys
    original_stderr = sys.stderr
    sys.stderr = sys.__stderr__  # Temporarily restore for warning
    print(f"⚠️ WARNING: Modifying file outside work directory: {filepath}", file=sys.stderr)
    sys.stderr = original_stderr
```

### Issue 3: Potential Destructive Operation (WARNING)
**Location:** All-slides mode
**Problem:** Changing background on ALL slides without explicit confirmation

```python
# Consider adding for all-slides mode:
if slide_index is None:
    # This affects ALL slides - consider requiring confirmation
    # or an explicit --all-slides flag
    pass
```

### Issue 4: Missing Version Tracking (CRITICAL)

```python
# ✅ ADD version tracking
with PowerPointAgent(filepath) as agent:
    agent.open(filepath)
    
    version_before = agent.get_presentation_version()
    
    # ... existing logic ...
    
    agent.save()
    
    version_after = agent.get_presentation_version()
    
return {
    "status": "success",
    "file": str(filepath.resolve()),
    "presentation_version_before": version_before,
    "presentation_version_after": version_after,
    ...
}
```

## 4.3 Complete Remediated Version

```python
#!/usr/bin/env python3
"""
PowerPoint Set Background Tool v3.1.0
Set slide background to solid color or image.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    # Set single slide background color
    uv run tools/ppt_set_background.py --file deck.pptx --slide 0 --color "#FFFFFF" --json
    
    # Set all slides background
    uv run tools/ppt_set_background.py --file deck.pptx --color "#F5F5F5" --all-slides --json
    
    # Set background image
    uv run tools/ppt_set_background.py --file deck.pptx --slide 0 --image bg.jpg --json

Exit Codes:
    0: Success
    1: Error occurred
"""

import sys
import os

# --- HYGIENE BLOCK START ---
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

import json
import argparse
from pathlib import Path
from typing import Dict, Any, Optional

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError, SlideNotFoundError, ColorHelper
)

__version__ = "3.1.0"


def set_background(
    filepath: Path,
    color: Optional[str] = None,
    image: Optional[Path] = None,
    slide_index: Optional[int] = None,
    all_slides: bool = False
) -> Dict[str, Any]:
    """
    Set slide background to solid color or image.
    
    Args:
        filepath: Path to PowerPoint file (.pptx)
        color: Hex color code (e.g., "#FFFFFF")
        image: Path to background image file
        slide_index: Specific slide to modify (0-based), or None for default behavior
        all_slides: If True, apply to all slides (requires explicit flag)
        
    Returns:
        Dict containing:
            - status: 'success'
            - file: Absolute file path
            - slides_affected: Number of slides modified
            - type: 'color' or 'image'
            - presentation_version_before: Version hash before changes
            - presentation_version_after: Version hash after changes
            - tool_version: Tool version string
            
    Raises:
        FileNotFoundError: If file or image doesn't exist
        ValueError: If neither color nor image specified, or invalid parameters
        SlideNotFoundError: If slide index out of range
    """
    # Validate file exists
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    # Validate file extension
    if filepath.suffix.lower() != '.pptx':
        raise ValueError("Only .pptx files are supported")
    
    # Validate mutually exclusive options
    if not color and not image:
        raise ValueError("Must specify either --color or --image")
    
    if color and image:
        raise ValueError("Cannot specify both --color and --image")
    
    # Validate color format if provided
    if color:
        if not color.startswith('#') or len(color) != 7:
            raise ValueError(f"Invalid color format '{color}'. Use hex format: #RRGGBB")
    
    # Validate image exists if provided
    if image and not image.exists():
        raise FileNotFoundError(f"Image not found: {image}")
    
    # Validate slide_index / all_slides logic
    if slide_index is not None and all_slides:
        raise ValueError("Cannot specify both --slide and --all-slides")
    
    if slide_index is None and not all_slides:
        raise ValueError("Must specify --slide INDEX or --all-slides")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Capture version BEFORE changes
        version_before = agent.get_presentation_version()
        
        slide_count = agent.get_slide_count()
        
        # Determine target slides
        if slide_index is not None:
            if not 0 <= slide_index < slide_count:
                raise SlideNotFoundError(
                    f"Slide index {slide_index} out of range",
                    details={"requested": slide_index, "available": slide_count}
                )
            target_slides = [agent.prs.slides[slide_index]]
            target_indices = [slide_index]
        else:
            target_slides = list(agent.prs.slides)
            target_indices = list(range(slide_count))
        
        # Apply background
        for slide in target_slides:
            bg = slide.background
            fill = bg.fill
            
            if color:
                fill.solid()
                fill.fore_color.rgb = ColorHelper.from_hex(color)
            elif image:
                fill.user_picture(str(image.resolve()))
        
        agent.save()
        
        # Capture version AFTER changes
        version_after = agent.get_presentation_version()
    
    return {
        "status": "success",
        "file": str(filepath.resolve()),
        "slides_affected": len(target_slides),
        "slide_indices": target_indices,
        "type": "color" if color else "image",
        "value": color if color else str(image),
        "presentation_version_before": version_before,
        "presentation_version_after": version_after,
        "tool_version": __version__
    }


def main():
    parser = argparse.ArgumentParser(
        description="Set PowerPoint slide background",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=f"""
Examples:
    # Set single slide background
    uv run tools/ppt_set_background.py --file deck.pptx --slide 0 --color "#FFFFFF" --json
    
    # Set all slides (requires explicit --all-slides flag)
    uv run tools/ppt_set_background.py --file deck.pptx --all-slides --color "#F5F5F5" --json
    
    # Use background image
    uv run tools/ppt_set_background.py --file deck.pptx --slide 0 --image background.jpg --json

Version: {__version__}
        """
    )
    
    parser.add_argument('--file', required=True, type=Path, help='PowerPoint file path')
    parser.add_argument('--slide', type=int, help='Slide index (0-based)')
    parser.add_argument('--all-slides', action='store_true', help='Apply to all slides')
    parser.add_argument('--color', help='Hex color code (e.g., #FFFFFF)')
    parser.add_argument('--image', type=Path, help='Background image path')
    parser.add_argument('--json', action='store_true', default=True, help='Output JSON')
    
    args = parser.parse_args()
    
    try:
        result = set_background(
            filepath=args.file,
            slide_index=args.slide,
            all_slides=args.all_slides,
            color=args.color,
            image=args.image
        )
        print(json.dumps(result, indent=2))
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify file/image paths exist and are accessible"
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
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
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Check argument format and combinations"
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "tool_version": __version__
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()
```

---

# Summary: Consolidated Findings

## Critical Issues (Must Fix)

| # | Issue | Affected Tools | Impact |
|---|-------|----------------|--------|
| 1 | Missing Hygiene Block | All 4 | JSON output corruption risk |
| 2 | Missing Version Tracking | All 4 | Audit trail broken, conflict detection impossible |
| 3 | Missing `__version__` | All 4 | Tool identification in errors |
| 4 | Wrong file extension (`.ppt`) | `ppt_add_notes.py` | Runtime crash with .ppt files |
| 5 | Missing slide index validation | `ppt_set_background.py` | Crash on invalid index |
| 6 | Non-JSON output mode | `ppt_create_from_structure.py` | Breaks JSON-only contract |

## Warnings (Should Fix)

| # | Issue | Affected Tools |
|---|-------|----------------|
| 1 | Missing `suggestion` in error responses | All 4 |
| 2 | Incomplete docstrings | All 4 |
| 3 | No Clone-Before-Edit warning | `ppt_add_notes.py`, `ppt_set_background.py` |
| 4 | Silent extension correction | `ppt_create_from_structure.py` |
| 5 | Direct `agent.prs.slides[]` access | Multiple |

## Remediation Priority

```
PRIORITY 1 (Immediate):
├── Add Hygiene Block to all tools
├── Add Version Tracking to all tools
├── Fix file extension validation in ppt_add_notes.py
└── Add slide index validation to ppt_set_background.py

PRIORITY 2 (Next Sprint):
├── Add __version__ constants
├── Add suggestion fields to error responses
├── Expand docstrings
└── Remove non-JSON output mode

PRIORITY 3 (Future):
├── Add Clone-Before-Edit warnings
├── Consider approval tokens for all-slides operations
└── Add comprehensive test coverage
```

---

Would you like me to provide:
1. **Complete remediated versions** of all four tools as ready-to-commit files?
2. **Test cases** for the critical fixes?
3. **A PR checklist** for these changes?
