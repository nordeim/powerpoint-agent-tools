# Agent Initialization Document: PowerPoint Agent Tools

## Document Metadata

| Field | Value |
|-------|-------|
| **Document Version** | 1.0.0 |
| **Project Version** | 3.1.0 |
| **Last Updated** | 2025-11-25 |
| **Purpose** | Single source of truth for AI agents and human developers |
| **Audience** | AI coding agents, human contributors, maintainers |

---

## Table of Contents

1. [Project Identity](#1-project-identity)
2. [Architecture Overview](#2-architecture-overview)
3. [Design Philosophy](#3-design-philosophy)
4. [Programming Model](#4-programming-model)
5. [Code Standards](#5-code-standards)
6. [Critical Patterns & Gotchas](#6-critical-patterns--gotchas)
7. [Testing Requirements](#7-testing-requirements)
8. [Contribution Workflow](#8-contribution-workflow)
9. [Quick Reference](#9-quick-reference)

---

## 1. Project Identity

### 1.1 What is PowerPoint Agent Tools?

**PowerPoint Agent Tools** is a suite of **37+ stateless CLI utilities** designed for AI agents to programmatically create, modify, and validate PowerPoint (`.pptx`) files.

```
┌─────────────────────────────────────────────────────────────┐
│                  POWERPOINT AGENT TOOLS                      │
│                                                              │
│   "Enabling AI agents to engineer presentations with         │
│    precision, safety, and visual intelligence"               │
│                                                              │
│   • 37+ CLI tools for complete presentation lifecycle        │
│   • Stateless, atomic operations with JSON I/O               │
│   • Production-grade safety (clone-before-edit, validation)  │
│   • Accessibility-first (WCAG compliance checking)           │
│   • Design-intelligent (typography, color, layout systems)   │
└─────────────────────────────────────────────────────────────┘
```

### 1.2 Why Does This Project Exist?

| Problem | Solution |
|---------|----------|
| AI agents hallucinate slide structure | Deep probing extracts actual layouts/placeholders |
| python-pptx API is complex | Semantic wrapper with intuitive methods |
| Agents lose track of shape indices | Mandatory re-query after structural changes |
| Destructive operations are risky | Clone-before-edit, approval tokens, rollback commands |
| Presentations lack accessibility | Built-in WCAG validation and remediation |
| Output is unparseable | JSON-first I/O with structured errors |

### 1.3 Who Uses This?

1. **AI Presentation Architects** - LLM-based agents that generate/modify presentations
2. **Automation Pipelines** - CI/CD systems that produce reports as slides
3. **Human Developers** - Building presentation automation workflows

---

## 2. Architecture Overview

### 2.1 Hub-and-Spoke Model

```
                         ┌─────────────────────────┐
                         │   AI AGENT / HUMAN      │
                         │   (Orchestration Layer) │
                         └───────────┬─────────────┘
                                     │
                    ┌────────────────┼────────────────┐
                    │                │                │
                    ▼                ▼                ▼
           ┌───────────────┐ ┌───────────────┐ ┌───────────────┐
           │ ppt_add_      │ │ ppt_get_      │ │ ppt_validate_ │
           │ shape.py      │ │ slide_info.py │ │ presentation  │
           │               │ │               │ │ .py           │
           │  (SPOKE)      │ │   (SPOKE)     │ │   (SPOKE)     │
           └───────┬───────┘ └───────┬───────┘ └───────┬───────┘
                   │                 │                 │
                   └─────────────────┼─────────────────┘
                                     │
                                     ▼
                    ┌─────────────────────────────────┐
                    │                                 │
                    │   powerpoint_agent_core.py      │
                    │           (HUB)                 │
                    │                                 │
                    │   • PowerPointAgent class       │
                    │   • All XML manipulation        │
                    │   • File locking                │
                    │   • Position/Size resolution    │
                    │   • Color helpers               │
                    │   • Validation logic            │
                    │                                 │
                    └─────────────────────────────────┘
                                     │
                                     ▼
                    ┌─────────────────────────────────┐
                    │         python-pptx             │
                    │     (Underlying Library)        │
                    └─────────────────────────────────┘
```

### 2.2 Directory Structure

```
powerpoint-agent-tools/
├── core/
│   ├── __init__.py                    # Exports public API
│   ├── powerpoint_agent_core.py       # THE HUB - all logic lives here
│   └── strict_validator.py            # JSON Schema validation utilities
├── tools/                             # THE SPOKES - 37+ CLI scripts
│   ├── ppt_add_shape.py
│   ├── ppt_add_slide.py
│   ├── ppt_get_info.py
│   ├── ppt_capability_probe.py
│   ├── ... (34+ more tools)
│   └── ppt_validate_presentation.py
├── schemas/                           # JSON Schemas for validation
│   ├── manifest.schema.json
│   ├── ppt_get_info.schema.json
│   └── ...
├── tests/                             # pytest test suite
│   ├── test_core.py
│   ├── test_shape_opacity.py
│   └── ...
├── assets/                            # Test images, sample decks
├── AGENT_SYSTEM_PROMPT.md             # System prompt for AI agents
├── CONTRIBUTING_TOOLS.md              # Guide for adding new tools
└── requirements.txt                   # Dependencies
```

### 2.3 Key Components

| Component | Location | Responsibility |
|-----------|----------|----------------|
| **PowerPointAgent** | `core/powerpoint_agent_core.py` | Context manager class; all operations |
| **CLI Tools** | `tools/ppt_*.py` | Thin wrappers; argparse + JSON output |
| **Strict Validator** | `core/strict_validator.py` | JSON Schema validation with caching |
| **Position/Size** | `core/powerpoint_agent_core.py` | Resolve %, inches, anchor, grid to absolute |
| **ColorHelper** | `core/powerpoint_agent_core.py` | Hex parsing, contrast calculation |

---

## 3. Design Philosophy

### 3.1 The Four Pillars

```
┌─────────────────────────────────────────────────────────────┐
│                    DESIGN PILLARS                            │
├─────────────────┬─────────────────┬─────────────────────────┤
│                 │                 │                         │
│   STATELESS     │    ATOMIC       │    COMPOSABLE           │
│                 │                 │                         │
│ Each tool call  │ Open → Modify   │ Tools can be chained    │
│ is independent  │ → Save → Close  │ in any sequence         │
│                 │                 │                         │
│ No memory of    │ One action per  │ Output of one feeds     │
│ previous calls  │ invocation      │ input of another        │
│                 │                 │                         │
├─────────────────┴─────────────────┴─────────────────────────┤
│                                                              │
│                      VISUAL-AWARE                            │
│                                                              │
│   Tools understand design: typography scales, color theory,  │
│   content density rules (6×6), accessibility requirements    │
│                                                              │
└─────────────────────────────────────────────────────────────┘
```

### 3.2 Core Principles

| Principle | Implementation |
|-----------|----------------|
| **Clone Before Edit** | Never modify source files; always work on copies |
| **Probe Before Operate** | Always run capability probe to discover layouts/shapes |
| **JSON-First I/O** | All tools output JSON to stdout; errors are structured JSON |
| **Fail Safely** | Incomplete is better than corrupted; validate before delivery |
| **Refresh After Structural Changes** | Shape indices shift; re-query after add/remove/z-order |
| **Accessibility by Default** | Alt-text validation, contrast checking, WCAG compliance |

### 3.3 The Statelessness Contract

```python
# CORRECT: Each call is independent
with PowerPointAgent(filepath) as agent:
    agent.open(filepath)
    agent.add_shape(...)
    agent.save()
# File is closed, lock released, no state retained

# WRONG: Assuming state persists
agent.add_shape(...)  # Will fail - no file open
```

**Why Stateless?**
1. AI agents may lose context between calls
2. Prevents race conditions in parallel execution
3. Enables pipeline composition
4. Simplifies error recovery

---

## 4. Programming Model

### 4.1 Adding a New Tool

Every tool follows this exact pattern:

```python
#!/usr/bin/env python3
"""
PowerPoint [Action] [Object] Tool v3.x.x
[One-line description]

Author: PowerPoint Agent Team
License: MIT
Version: 3.x.x

Usage:
    uv run tools/ppt_[verb]_[noun].py --file deck.pptx [args] --json

Exit Codes:
    0: Success
    1: Error occurred
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any, Optional, List

# 1. PATH SETUP (required for imports without package install)
sys.path.insert(0, str(Path(__file__).parent.parent))

# 2. IMPORTS FROM CORE
from core.powerpoint_agent_core import (
    PowerPointAgent,
    PowerPointAgentError,
    SlideNotFoundError,
    ShapeNotFoundError,
    # ... other needed imports
)

# ============================================================================
# CONSTANTS
# ============================================================================

__version__ = "3.x.x"

# ============================================================================
# MAIN LOGIC FUNCTION
# ============================================================================

def do_action(
    filepath: Path,
    # ... other typed parameters
) -> Dict[str, Any]:
    """
    Perform the action.
    
    Args:
        filepath: Path to PowerPoint file
        # ... document all args
        
    Returns:
        Dict with results
        
    Raises:
        FileNotFoundError: If file doesn't exist
        SlideNotFoundError: If slide index invalid
        # ... document all exceptions
    """
    # Validate file exists
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    # Open, operate, save
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # ... perform operations via agent methods ...
        
        agent.save()
    
    # Return structured result
    return {
        "status": "success",
        "file": str(filepath),
        # ... action-specific fields
    }

# ============================================================================
# CLI INTERFACE
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Tool description",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  uv run tools/ppt_xxx.py --file deck.pptx --json
        """
    )
    
    # Required arguments
    parser.add_argument('--file', required=True, type=Path, help='PowerPoint file')
    
    # Optional arguments
    parser.add_argument('--json', action='store_true', default=True, help='JSON output')
    
    args = parser.parse_args()
    
    try:
        result = do_action(filepath=args.file, ...)
        print(json.dumps(result, indent=2))
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError"
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

### 4.2 Adding a Core Method

When adding functionality to `powerpoint_agent_core.py`:

```python
def new_method(
    self,
    slide_index: int,
    # ... parameters with type hints
    optional_param: Optional[str] = None
) -> Dict[str, Any]:
    """
    One-line description.
    
    More detailed description if needed.
    
    Args:
        slide_index: Target slide index (0-based)
        optional_param: Description with default noted
        
    Returns:
        Dict with:
            - key1: Description
            - key2: Description
            
    Raises:
        SlideNotFoundError: If slide index is invalid
        ValueError: If parameters are invalid
        
    Example:
        result = agent.new_method(slide_index=0, ...)
    """
    # 1. Validate inputs
    slide = self._get_slide(slide_index)  # Raises SlideNotFoundError
    
    if optional_param is not None:
        # validate optional_param
        pass
    
    # 2. Perform operation
    # ... implementation ...
    
    # 3. Return structured result
    return {
        "slide_index": slide_index,
        "action": "new_method",
        # ... other relevant fields
    }
```

### 4.3 Data Structures

#### Position Dictionary

```python
# Percentage (Recommended)
{"left": "10%", "top": "20%"}

# Inches (Absolute)
{"left": 1.5, "top": 2.0}

# Anchor-based
{"anchor": "center", "offset_x": 0, "offset_y": -1.0}
# Anchors: top_left, top_center, top_right,
#          center_left, center, center_right,
#          bottom_left, bottom_center, bottom_right

# Grid-based (12-column)
{"grid_row": 2, "grid_col": 3, "grid_size": 12}
```

#### Size Dictionary

```python
# Percentage
{"width": "50%", "height": "40%"}

# Inches
{"width": 5.0, "height": 3.0}

# Auto (preserve aspect ratio)
{"width": "50%", "height": "auto"}
```

#### Colors

```python
# Hex format (with or without #)
"#0070C0"
"0070C0"

# Preset names (if tool supports)
"primary"    # #0070C0
"accent"     # #ED7D31
"success"    # #70AD47
```

### 4.4 Error Handling Pattern

```python
# In CLI tools - catch and convert to JSON
try:
    result = do_action(...)
    print(json.dumps(result, indent=2))
    sys.exit(0)

except FileNotFoundError as e:
    print(json.dumps({
        "status": "error",
        "error": str(e),
        "error_type": "FileNotFoundError",
        "suggestion": "Verify the file path exists"
    }, indent=2))
    sys.exit(1)

except SlideNotFoundError as e:
    print(json.dumps({
        "status": "error",
        "error": e.message,
        "error_type": "SlideNotFoundError",
        "details": e.details,
        "suggestion": "Use ppt_get_info.py to check available slides"
    }, indent=2))
    sys.exit(1)

except PowerPointAgentError as e:
    print(json.dumps({
        "status": "error",
        "error": e.message,
        "error_type": type(e).__name__,
        "details": e.details
    }, indent=2))
    sys.exit(1)

except Exception as e:
    print(json.dumps({
        "status": "error",
        "error": str(e),
        "error_type": type(e).__name__
    }, indent=2))
    sys.exit(1)
```

---

## 5. Code Standards

### 5.1 Style Requirements

| Aspect | Requirement |
|--------|-------------|
| **Python Version** | 3.8+ |
| **Type Hints** | Mandatory for all function signatures |
| **Docstrings** | Required for all modules, classes, and functions |
| **Line Length** | 100 characters (soft limit) |
| **Formatting** | `black` formatter |
| **Linting** | `ruff` linter |
| **Imports** | Grouped: stdlib → third-party → local; alphabetized |

### 5.2 Naming Conventions

| Element | Convention | Example |
|---------|------------|---------|
| **Tool files** | `ppt_<verb>_<noun>.py` | `ppt_add_shape.py` |
| **Core methods** | `snake_case` | `add_shape()`, `get_slide_info()` |
| **Private methods** | `_snake_case` | `_set_fill_opacity()` |
| **Constants** | `UPPER_SNAKE_CASE` | `AVAILABLE_SHAPES` |
| **Classes** | `PascalCase` | `PowerPointAgent` |
| **Type aliases** | `PascalCase` | `PositionDict`, `SizeDict` |

### 5.3 Documentation Standards

```python
def method_name(
    self,
    required_param: str,
    optional_param: Optional[int] = None
) -> Dict[str, Any]:
    """
    Short one-line description ending with period.
    
    Longer description if needed, explaining behavior,
    edge cases, and important notes.
    
    Args:
        required_param: Description of parameter
        optional_param: Description with default noted.
            Multi-line descriptions indented like this.
            
    Returns:
        Dict with the following keys:
            - key1 (str): Description
            - key2 (int): Description
            
    Raises:
        ValueError: When required_param is empty
        SlideNotFoundError: When slide doesn't exist
        
    Example:
        >>> result = agent.method_name("value", optional_param=42)
        >>> print(result["key1"])
        'expected output'
        
    Note:
        Any important caveats or warnings.
        
    See Also:
        related_method: For similar functionality
    """
```

### 5.4 JSON Output Standards

Every tool must output exactly one of:

**Success Response:**
```json
{
  "status": "success",
  "file": "/absolute/path/to/file.pptx",
  "action_specific_key": "value",
  "tool_version": "3.1.0"
}
```

**Warning Response (success with issues):**
```json
{
  "status": "warning",
  "file": "/absolute/path/to/file.pptx",
  "warnings": ["Warning message 1", "Warning message 2"],
  "result": { ... }
}
```

**Error Response:**
```json
{
  "status": "error",
  "error": "Human-readable error message",
  "error_type": "ExceptionClassName",
  "details": { "optional": "context" },
  "suggestion": "How to fix this"
}
```

---

## 6. Critical Patterns & Gotchas

### 6.1 The Shape Index Problem

**Problem:** Shape indices are **positional** and **shift** after structural operations.

```python
# WRONG - indices become stale
shape1 = agent.add_shape(slide_index=0, ...)  # Returns shape_index: 5
shape2 = agent.add_shape(slide_index=0, ...)  # Returns shape_index: 6
agent.remove_shape(slide_index=0, shape_index=5)
agent.format_shape(slide_index=0, shape_index=6, ...)  # WRONG! It's now index 5!

# CORRECT - re-query after structural changes
shape1 = agent.add_shape(slide_index=0, ...)
shape2 = agent.add_shape(slide_index=0, ...)
agent.remove_shape(slide_index=0, shape_index=shape1["shape_index"])
slide_info = agent.get_slide_info(slide_index=0)  # Refresh indices
# Now find shape2 by name or other identifier
```

**Operations that invalidate indices:**
- `add_shape()` - adds new index
- `remove_shape()` - shifts indices down
- `set_z_order()` - reorders indices
- `delete_slide()` - invalidates all indices on that slide

### 6.2 The Probe-First Pattern

**Problem:** Template layouts are unpredictable. Placeholder geometry is zero until instantiated.

```python
# WRONG - guessing layout names
agent.add_slide(layout_name="Title and Content")  # Might not exist!

# CORRECT - probe first
probe_result = agent.capability_probe(deep=True)
available_layouts = probe_result["layouts"]
# Pick from available_layouts
```

**The Deep Probe Innovation:**
The `ppt_capability_probe.py --deep` creates a **transient slide in memory** to measure actual placeholder geometry, then discards it. This is the only reliable way to know exact positioning.

### 6.3 The Overlay Pattern

When creating semi-transparent overlays for text readability:

```python
# 1. Add overlay shape with opacity
result = agent.add_shape(
    slide_index=0,
    shape_type="rectangle",
    position={"left": "0%", "top": "0%"},
    size={"width": "100%", "height": "100%"},
    fill_color="#FFFFFF",
    fill_opacity=0.15  # Subtle overlay (15% opaque)
)

# 2. IMMEDIATELY refresh indices
slide_info = agent.get_slide_info(slide_index=0)

# 3. Send overlay to back
agent.set_z_order(
    slide_index=0,
    shape_index=result["shape_index"],
    action="send_to_back"
)

# 4. IMMEDIATELY refresh indices again
slide_info = agent.get_slide_info(slide_index=0)
```

### 6.4 Opacity vs Transparency

```
OPACITY:     0.0 ←─────────────────────────→ 1.0
             Invisible                    Fully visible
             (see-through)                (solid)

TRANSPARENCY: 1.0 ←─────────────────────────→ 0.0
              Invisible                    Fully visible
              (see-through)                (solid)

CONVERSION: opacity = 1.0 - transparency
```

**The `transparency` parameter is DEPRECATED.** Use `fill_opacity` instead.

### 6.5 File Handling Safety

```python
# ALWAYS use absolute paths
filepath = Path(filepath).resolve()  # Convert to absolute

# ALWAYS validate existence
if not filepath.exists():
    raise FileNotFoundError(f"File not found: {filepath}")

# ALWAYS use context manager
with PowerPointAgent(filepath) as agent:
    agent.open(filepath)
    # ... operations ...
    agent.save()  # Or agent.save(new_path) for save-as
```

### 6.6 XML Manipulation (Advanced)

When python-pptx doesn't expose a feature (like opacity):

```python
from lxml import etree
from pptx.oxml.ns import qn

# Access shape XML
spPr = shape._sp.spPr

# Navigate XML tree
solidFill = spPr.find(qn('a:solidFill'))
color_elem = solidFill.find(qn('a:srgbClr'))

# Create new elements
alpha_elem = etree.SubElement(color_elem, qn('a:alpha'))
alpha_elem.set('val', str(15000))  # 15% opacity (15000/100000)
```

**OOXML Alpha Scale:** 0 = invisible, 100000 = fully opaque

---

## 7. Testing Requirements

### 7.1 Test Structure

```
tests/
├── test_core.py                  # Core library tests
├── test_shape_opacity.py         # Feature-specific tests
├── test_tools/                   # CLI tool tests
│   ├── test_ppt_add_shape.py
│   └── ...
├── conftest.py                   # Shared fixtures
└── assets/                       # Test files
    ├── sample.pptx
    └── template.pptx
```

### 7.2 Required Test Coverage

For any new feature, provide tests for:

| Category | What to Test |
|----------|--------------|
| **Happy Path** | Normal usage succeeds |
| **Edge Cases** | Boundary values (0, 1, max, empty) |
| **Error Cases** | Invalid inputs raise correct exceptions |
| **Validation** | Invalid ranges/formats rejected |
| **Backward Compatibility** | Old code still works |
| **CLI Integration** | Tool runs and produces valid JSON |

### 7.3 Test Pattern

```python
import pytest
import tempfile
from pathlib import Path

@pytest.fixture
def test_presentation(tmp_path):
    """Create a test presentation."""
    pptx_path = tmp_path / "test.pptx"
    with PowerPointAgent() as agent:
        agent.create_new()
        agent.add_slide(layout_name="Blank")
        agent.save(pptx_path)
    return pptx_path

class TestFeatureName:
    """Tests for feature_name functionality."""
    
    def test_happy_path(self, test_presentation):
        """Test normal usage."""
        with PowerPointAgent(test_presentation) as agent:
            agent.open(test_presentation)
            result = agent.feature_method(...)
            agent.save()
        
        assert result["status"] == "success"
        assert "expected_key" in result
    
    def test_edge_case_zero(self, test_presentation):
        """Test with zero value."""
        # ...
    
    def test_error_invalid_input(self, test_presentation):
        """Test that invalid input raises ValueError."""
        with PowerPointAgent(test_presentation) as agent:
            agent.open(test_presentation)
            
            with pytest.raises(ValueError) as excinfo:
                agent.feature_method(invalid_param=-1)
            
            assert "must be between" in str(excinfo.value)
```

### 7.4 Running Tests

```bash
# Run all tests
pytest tests/ -v

# Run specific test file
pytest tests/test_shape_opacity.py -v

# Run specific test class
pytest tests/test_shape_opacity.py::TestAddShapeOpacity -v

# Run with coverage
pytest tests/ --cov=core --cov-report=html

# Run only fast tests (skip slow integration tests)
pytest tests/ -v -m "not slow"
```

---

## 8. Contribution Workflow

### 8.1 Before Starting Work

1. **Understand the architecture** - Read this document fully
2. **Check existing tools** - Don't duplicate functionality
3. **Review the system prompt** - Understand how AI agents will use your code
4. **Set up environment:**
   ```bash
   uv pip install -r requirements.txt
   uv pip install -r requirements-dev.txt
   ```

### 8.2 PR Checklist

Before submitting a PR, verify:

#### Code Quality
- [ ] Type hints on all function signatures
- [ ] Docstrings on all public functions/methods
- [ ] Follows naming conventions
- [ ] `black` formatted
- [ ] `ruff` passes with no errors

#### For New Tools
- [ ] File named `ppt_<verb>_<noun>.py`
- [ ] Uses standard template structure
- [ ] Outputs valid JSON to stdout
- [ ] Uses exit code 0 for success, 1 for error
- [ ] Validates file paths with `pathlib.Path`
- [ ] Catches all exceptions and converts to JSON errors

#### For Core Changes
- [ ] Method has complete docstring with example
- [ ] Raises appropriate typed exceptions
- [ ] Returns Dict with documented structure
- [ ] Backward compatible (or deprecation path provided)

#### Testing
- [ ] Tests cover happy path
- [ ] Tests cover edge cases
- [ ] Tests cover error cases
- [ ] All tests pass: `pytest tests/ -v`

#### Documentation
- [ ] Updated relevant docstrings
- [ ] Updated CHANGELOG if applicable
- [ ] Updated system prompt if new capability added

### 8.3 Common PR Mistakes to Avoid

| Mistake | Why It's Wrong | Correct Approach |
|---------|----------------|------------------|
| Printing non-JSON to stdout | Breaks parsing | Use stderr for logs |
| Assuming shape indices persist | They shift | Re-query after changes |
| Not validating inputs | Cryptic errors | Validate early, fail fast |
| Forgetting context manager | File lock issues | Always use `with` |
| Hardcoding paths | Platform issues | Use `pathlib.Path` |
| Swallowing exceptions | Silent failures | Log warning, return status |

---

## 9. Quick Reference

### 9.1 Tool Catalog (37 Tools)

| Domain | Tools |
|--------|-------|
| **Creation** | `ppt_create_new`, `ppt_create_from_template`, `ppt_create_from_structure`, `ppt_clone_presentation` |
| **Slides** | `ppt_add_slide`, `ppt_delete_slide`, `ppt_duplicate_slide`, `ppt_reorder_slides`, `ppt_set_slide_layout`, `ppt_set_footer` |
| **Content** | `ppt_set_title`, `ppt_add_text_box`, `ppt_add_bullet_list`, `ppt_format_text`, `ppt_replace_text`, `ppt_add_notes` |
| **Images** | `ppt_insert_image`, `ppt_replace_image`, `ppt_crop_image`, `ppt_set_image_properties` |
| **Shapes** | `ppt_add_shape`, `ppt_format_shape`, `ppt_add_connector`, `ppt_set_background`, `ppt_set_z_order`, `ppt_remove_shape` |
| **Data Viz** | `ppt_add_chart`, `ppt_update_chart_data`, `ppt_format_chart`, `ppt_add_table` |
| **Inspection** | `ppt_get_info`, `ppt_get_slide_info`, `ppt_extract_notes`, `ppt_capability_probe` |
| **Validation** | `ppt_validate_presentation`, `ppt_check_accessibility`, `ppt_export_images`, `ppt_export_pdf` |

### 9.2 Core Exceptions

| Exception | When Raised |
|-----------|-------------|
| `PowerPointAgentError` | Base exception for all agent errors |
| `SlideNotFoundError` | Invalid slide index |
| `ShapeNotFoundError` | Invalid shape index |
| `LayoutNotFoundError` | Layout name doesn't exist |
| `ValidationError` | Schema or content validation failed |

### 9.3 Key Constants

```python
# Slide dimensions (default)
SLIDE_WIDTH_INCHES = 10.0
SLIDE_HEIGHT_INCHES = 7.5

# Content density limits
MAX_BULLETS_PER_SLIDE = 6
MAX_WORDS_PER_BULLET = 6

# Accessibility thresholds
MIN_CONTRAST_RATIO = 4.5  # WCAG AA
MIN_FONT_SIZE_PT = 10

# Overlay defaults
OVERLAY_OPACITY = 0.15
```

### 9.4 Common Commands

```bash
# Run a tool
uv run tools/ppt_add_shape.py --file deck.pptx --slide 0 --shape rectangle \
  --position '{"left":"10%","top":"10%"}' --size '{"width":"20%","height":"20%"}' --json

# Run tests
pytest tests/ -v

# Format code
black core/ tools/ tests/

# Lint code
ruff check core/ tools/ tests/

# Check types (if using mypy)
mypy core/ tools/
```

---

## Document End

**This document should be your first read before contributing.** It provides the mental model needed to work effectively on this project. When in doubt, refer back to the Design Philosophy and Critical Patterns sections.

For questions not covered here, consult:
1. `AGENT_SYSTEM_PROMPT.md` - How AI agents use these tools
2. `CONTRIBUTING_TOOLS.md` - Detailed tool creation guide
3. Existing tools in `tools/` - Working examples
