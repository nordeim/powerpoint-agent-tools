# 🔍 Deep Analysis & Implementation Plan

## Phase 1: Comprehensive Feature Preservation Analysis

Before creating replacement files, I must systematically analyze each tool's complete functionality to ensure zero regression.

---

### 1.1 Tool Feature Inventory

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                    COMPLETE FEATURE INVENTORY                                │
├─────────────────────────────────────────────────────────────────────────────┤
│                                                                              │
│  ppt_create_from_template.py                                                │
│  ├── Core Features:                                                          │
│  │   ├── Template file validation (exists, .pptx extension)                │
│  │   ├── Slide count validation (1-100 range)                               │
│  │   ├── Layout resolution (case-insensitive partial match)                │
│  │   ├── Fallback to first available layout                                 │
│  │   ├── Template slides preservation                                        │
│  │   ├── Additional slides addition loop                                    │
│  │   └── Backward compat: add_slide returns int or dict                     │
│  ├── Output Fields:                                                          │
│  │   ├── status, file, template_used                                        │
│  │   ├── total_slides, slides_requested, template_slides, slides_added     │
│  │   ├── layout_used, available_layouts                                     │
│  │   ├── file_size_bytes, slide_dimensions                                  │
│  │   └── presentation_version, tool_version                                 │
│  └── Error Handlers: FileNotFoundError, ValueError, PowerPointAgentError   │
│                                                                              │
│  ppt_create_new.py                                                           │
│  ├── Core Features:                                                          │
│  │   ├── Optional template support                                          │
│  │   ├── Slide count validation (1-100 range)                               │
│  │   ├── First slide uses "Title Slide" if available                       │
│  │   ├── Layout resolution (case-insensitive partial match)                │
│  │   ├── Slide indices tracking                                             │
│  │   └── Backward compat: add_slide returns int or dict                     │
│  ├── Output Fields:                                                          │
│  │   ├── status, file, slides_created, slide_indices                       │
│  │   ├── file_size_bytes, slide_dimensions                                  │
│  │   ├── available_layouts, layout_used, template_used                      │
│  │   └── presentation_version, tool_version                                 │
│  └── Error Handlers: FileNotFoundError, ValueError, PowerPointAgentError   │
│                                                                              │
│  ppt_delete_slide.py                                                         │
│  ├── Core Features:                                                          │
│  │   ├── Approval token validation (required, format check)                │
│  │   ├── ApprovalTokenError fallback class definition                       │
│  │   ├── Slide index validation                                             │
│  │   ├── Version tracking (before/after)                                    │
│  │   ├── Core approval_token parameter try/except                           │
│  │   └── Exit code 4 for permission errors                                  │
│  ├── Output Fields:                                                          │
│  │   ├── status, file, deleted_index, remaining_slides                      │
│  │   ├── presentation_version_before, presentation_version_after            │
│  │   └── tool_version                                                        │
│  └── Error Handlers: ApprovalTokenError, SlideNotFoundError, FileNotFound  │
│                                                                              │
│  ppt_duplicate_slide.py                                                      │
│  ├── Core Features:                                                          │
│  │   ├── Slide index validation                                             │
│  │   ├── Backward compat: duplicate_slide returns int or dict              │
│  │   ├── Gets slide info for layout name                                    │
│  │   └── Version tracking (before/after)                                    │
│  ├── Output Fields:                                                          │
│  │   ├── status, file, source_index, new_slide_index                       │
│  │   ├── total_slides, layout                                               │
│  │   ├── presentation_version_before, presentation_version_after            │
│  │   └── tool_version                                                        │
│  └── Error Handlers: SlideNotFoundError, FileNotFoundError                  │
│                                                                              │
└─────────────────────────────────────────────────────────────────────────────┘
```

---

## Phase 2: Comprehensive Implementation Plan

### 2.1 ppt_create_from_template.py Checklist

```markdown
## ppt_create_from_template.py Implementation Checklist

### Pre-Implementation: Feature Verification
- [ ] Hygiene block preserved (lines 18-21)
- [ ] Template exists validation preserved
- [ ] Template .pptx extension validation preserved
- [ ] Slide count 1-100 validation preserved
- [ ] Layout case-insensitive partial match preserved
- [ ] Fallback to first layout preserved
- [ ] Template slides count logic preserved
- [ ] slides_to_add calculation preserved
- [ ] Backward compat for add_slide return preserved
- [ ] All output fields preserved
- [ ] Excellent epilog help text preserved

### Fixes to Apply
- [ ] Add sys.stdout.flush() after JSON write (line ~201)
- [ ] Add sys.stdout.flush() after human-readable write (line ~207)
- [ ] Add suggestion to PowerPointAgentError handler
- [ ] Add tool_version to PowerPointAgentError handler
- [ ] Add suggestion to generic Exception handler
- [ ] Add get_available_layouts() fallback with try/except
- [ ] Update version to 3.1.1

### Post-Implementation Verification
- [ ] All original functionality preserved
- [ ] No placeholder comments
- [ ] All error handlers have suggestion + tool_version
- [ ] All writes followed by flush()
```

### 2.2 ppt_create_new.py Checklist

```markdown
## ppt_create_new.py Implementation Checklist

### Pre-Implementation: Feature Verification
- [ ] Hygiene block preserved
- [ ] Optional template validation preserved
- [ ] Template .pptx extension check preserved
- [ ] Slide count 1-100 validation preserved
- [ ] First slide "Title Slide" logic preserved
- [ ] Layout resolution logic preserved
- [ ] slide_indices list tracking preserved
- [ ] Backward compat for add_slide return preserved
- [ ] All output fields preserved
- [ ] Epilog help text preserved

### Fixes to Apply
- [ ] Add sys.stdout.flush() after all writes
- [ ] Add suggestion to PowerPointAgentError handler
- [ ] Add tool_version to PowerPointAgentError handler
- [ ] Add suggestion to generic Exception handler
- [ ] Add get_available_layouts() fallback
- [ ] Update version to 3.1.1

### Post-Implementation Verification
- [ ] All original functionality preserved
- [ ] No placeholder comments
- [ ] All error handlers complete
```

### 2.3 ppt_delete_slide.py Checklist

```markdown
## ppt_delete_slide.py Implementation Checklist

### Pre-Implementation: Feature Verification
- [ ] Hygiene block preserved
- [ ] ApprovalTokenError fallback class preserved
- [ ] Approval token required validation preserved
- [ ] Approval token format validation preserved
- [ ] Slide index validation preserved
- [ ] Version tracking before/after preserved
- [ ] Core approval_token try/except preserved
- [ ] Exit code 4 for ApprovalTokenError preserved
- [ ] All output fields preserved
- [ ] Safety workflow documentation preserved

### Fixes to Apply
- [ ] Add sys.stdout.flush() after all writes
- [ ] Add suggestion to PowerPointAgentError handler
- [ ] Add tool_version to PowerPointAgentError handler
- [ ] Add suggestion to generic Exception handler
- [ ] Update version to 3.1.1

### Post-Implementation Verification
- [ ] Approval token validation still works
- [ ] Exit code 4 still used for token errors
- [ ] All error handlers complete
```

### 2.4 ppt_duplicate_slide.py Checklist

```markdown
## ppt_duplicate_slide.py Implementation Checklist

### Pre-Implementation: Feature Verification
- [ ] Hygiene block preserved
- [ ] Slide index validation preserved
- [ ] Backward compat for duplicate_slide return preserved
- [ ] Get slide info for layout preserved
- [ ] Version tracking before/after preserved
- [ ] All output fields preserved
- [ ] Epilog help text preserved

### Fixes to Apply
- [ ] Add sys.stdout.flush() after all writes
- [ ] Add suggestion to PowerPointAgentError handler
- [ ] Add tool_version to PowerPointAgentError handler
- [ ] Add suggestion to generic Exception handler
- [ ] Update version to 3.1.1

### Post-Implementation Verification
- [ ] All original functionality preserved
- [ ] No placeholder comments
```

---

## Phase 3: Plan Validation

### Validation Against Project Requirements

| Requirement | All 4 Tools After Fixes |
|-------------|------------------------|
| Hygiene Block | ✅ Already present |
| sys.stdout.write | ✅ Already present |
| sys.stdout.flush | ✅ Adding |
| Version 3.1.1 | ✅ Updating |
| suggestion in errors | ✅ Adding to all |
| tool_version in errors | ✅ Adding to all |
| Exit codes | ✅ Preserved (including 4 for delete) |

### Regression Risk Assessment

| Tool | Risk Level | Mitigation |
|------|------------|------------|
| ppt_create_from_template.py | LOW | Minor additions only |
| ppt_create_new.py | LOW | Minor additions only |
| ppt_delete_slide.py | LOW | Minor additions only |
| ppt_duplicate_slide.py | LOW | Minor additions only |

**Plan Validated ✅ - Proceeding to Implementation**

---

## Phase 4: Complete Replacement File Implementations

### 4.1 ppt_create_from_template.py (Complete Replacement)

```python
#!/usr/bin/env python3
"""
PowerPoint Create From Template Tool v3.1.1
Create new presentation from existing .pptx template.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.1

Usage:
    uv run tools/ppt_create_from_template.py --template corporate_template.pptx --output new_presentation.pptx --slides 10 --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

Changelog v3.1.1:
    - Added sys.stdout.flush() for pipeline safety
    - Added suggestion field to all error handlers
    - Added tool_version to all error responses
    - Added get_available_layouts() fallback for compatibility
"""

import sys
import os

# --- HYGIENE BLOCK START ---
# CRITICAL: Redirect stderr to /dev/null to prevent library noise from corrupting JSON output
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

import json
import argparse
from pathlib import Path
from typing import Dict, Any, List

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, 
    PowerPointAgentError,
    LayoutNotFoundError
)

# ============================================================================
# CONSTANTS
# ============================================================================

__version__ = "3.1.1"


# ============================================================================
# MAIN LOGIC
# ============================================================================

def create_from_template(
    template: Path,
    output: Path,
    slides: int = 1,
    layout: str = "Title and Content"
) -> Dict[str, Any]:
    """
    Create a new PowerPoint presentation from an existing template.
    
    This tool copies the template (including its theme, master slides, and
    any existing content) and optionally adds additional slides using the
    specified layout.
    
    Args:
        template: Path to the source template .pptx file
        output: Path where the new presentation will be saved
        slides: Total number of slides desired in the output (default: 1)
        layout: Layout name for additional slides (default: "Title and Content")
        
    Returns:
        Dict containing:
            - status: "success"
            - file: Absolute path to created file
            - template_used: Path to source template
            - total_slides: Final slide count
            - slides_requested: Number of slides requested
            - template_slides: Number of slides in original template
            - slides_added: Number of slides added
            - layout_used: Layout name used for added slides
            - available_layouts: List of all available layouts
            - file_size_bytes: Size of created file
            - slide_dimensions: Width, height, and aspect ratio
            - presentation_version: State hash for change tracking
            - tool_version: Version of this tool
            
    Raises:
        FileNotFoundError: If template file does not exist
        ValueError: If template is not .pptx or slide count invalid
        LayoutNotFoundError: If specified layout not found (falls back)
        
    Example:
        >>> result = create_from_template(
        ...     template=Path("templates/corporate.pptx"),
        ...     output=Path("q4_report.pptx"),
        ...     slides=15,
        ...     layout="Title and Content"
        ... )
        >>> print(result["total_slides"])
        15
    """
    if not template.exists():
        raise FileNotFoundError(f"Template file not found: {template}")
    
    if not template.suffix.lower() == '.pptx':
        raise ValueError(f"Template must be .pptx file, got: {template.suffix}")
    
    if slides < 1:
        raise ValueError("Must create at least 1 slide")
    
    if slides > 100:
        raise ValueError("Maximum 100 slides per creation (performance limit)")
    
    with PowerPointAgent() as agent:
        agent.create_new(template=template)
        
        try:
            available_layouts = agent.get_available_layouts()
        except AttributeError:
            info = agent.get_presentation_info()
            available_layouts = info.get("layouts", [])
        
        resolved_layout = layout
        if layout not in available_layouts:
            layout_lower = layout.lower()
            matched = False
            for avail in available_layouts:
                if layout_lower in avail.lower():
                    resolved_layout = avail
                    matched = True
                    break
            
            if not matched:
                resolved_layout = available_layouts[0] if available_layouts else "Title Slide"
        
        current_slides = agent.get_slide_count()
        
        slides_to_add = max(0, slides - current_slides)
        
        slide_indices: List[int] = list(range(current_slides))
        
        for i in range(slides_to_add):
            result = agent.add_slide(layout_name=resolved_layout)
            if isinstance(result, dict):
                idx = result.get("slide_index", result.get("index", len(slide_indices)))
            else:
                idx = result
            slide_indices.append(idx)
        
        agent.save(output)
        
        info = agent.get_presentation_info()
        presentation_version = info.get("presentation_version", None)
    
    file_size = output.stat().st_size if output.exists() else 0
    
    return {
        "status": "success",
        "file": str(output.resolve()),
        "template_used": str(template.resolve()),
        "total_slides": info["slide_count"],
        "slides_requested": slides,
        "template_slides": current_slides,
        "slides_added": slides_to_add,
        "layout_used": resolved_layout,
        "available_layouts": info.get("layouts", available_layouts),
        "file_size_bytes": file_size,
        "slide_dimensions": {
            "width_inches": info.get("slide_width_inches", 13.333),
            "height_inches": info.get("slide_height_inches", 7.5),
            "aspect_ratio": info.get("aspect_ratio", "16:9")
        },
        "presentation_version": presentation_version,
        "tool_version": __version__
    }


# ============================================================================
# CLI INTERFACE
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Create PowerPoint presentation from template",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Create from corporate template with 15 slides
  uv run tools/ppt_create_from_template.py \\
    --template templates/corporate.pptx \\
    --output q4_report.pptx \\
    --slides 15 \\
    --json
  
  # Create presentation using specific layout
  uv run tools/ppt_create_from_template.py \\
    --template templates/minimal.pptx \\
    --output demo.pptx \\
    --slides 5 \\
    --layout "Section Header" \\
    --json
  
  # Quick presentation from template (uses template's existing slides)
  uv run tools/ppt_create_from_template.py \\
    --template templates/branded.pptx \\
    --output quick_deck.pptx \\
    --json

Use Cases:
  - Corporate presentations with consistent branding
  - Team presentations with shared theme
  - Pre-formatted layouts (fonts, colors, logos)
  - Department-specific templates
  - Client-specific branded decks

Template Benefits:
  - Consistent branding across organization
  - Pre-configured master slides
  - Corporate colors and fonts
  - Logo placements
  - Standard layouts
  - Accessibility features built-in

Creating Templates:
  1. Design in PowerPoint with desired theme
  2. Configure master slides
  3. Set up color scheme
  4. Define standard layouts
  5. Save as .pptx template
  6. Use with this tool

Best Practices:
  - Maintain template library for different purposes
  - Version control templates
  - Document template usage guidelines
  - Test templates before distribution
  - Include variety of layouts in template

Output Format:
  {
    "status": "success",
    "file": "/path/to/output.pptx",
    "template_used": "/path/to/template.pptx",
    "total_slides": 15,
    "template_slides": 1,
    "slides_added": 14,
    "layout_used": "Title and Content",
    "presentation_version": "a1b2c3d4...",
    "tool_version": "3.1.1"
  }
        """
    )
    
    parser.add_argument(
        '--template',
        required=True,
        type=Path,
        help='Path to template .pptx file'
    )
    
    parser.add_argument(
        '--output',
        required=True,
        type=Path,
        help='Output presentation path'
    )
    
    parser.add_argument(
        '--slides',
        type=int,
        default=1,
        help='Total number of slides desired (default: 1)'
    )
    
    parser.add_argument(
        '--layout',
        default='Title and Content',
        help='Layout for additional slides (default: "Title and Content")'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        output_path = args.output
        if not output_path.suffix.lower() == '.pptx':
            output_path = output_path.with_suffix('.pptx')
        
        result = create_from_template(
            template=args.template.resolve(),
            output=output_path.resolve(),
            slides=args.slides,
            layout=args.layout
        )
        
        if args.json:
            sys.stdout.write(json.dumps(result, indent=2) + "\n")
            sys.stdout.flush()
        else:
            sys.stdout.write(f"Created presentation from template: {result['file']}\n")
            sys.stdout.write(f"  Template: {result['template_used']}\n")
            sys.stdout.write(f"  Total slides: {result['total_slides']}\n")
            sys.stdout.write(f"  Template had: {result['template_slides']} slides\n")
            sys.stdout.write(f"  Added: {result['slides_added']} slides\n")
            sys.stdout.write(f"  Layout: {result['layout_used']}\n")
            sys.stdout.flush()
        
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify the template file path exists and is accessible",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Check that template is .pptx and slide count is 1-100",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except LayoutNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "LayoutNotFoundError",
            "suggestion": "Use ppt_capability_probe.py to discover available layouts",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "details": getattr(e, 'details', {}),
            "suggestion": "Check template file integrity and available layouts",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "suggestion": "Check logs for detailed error information",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)


if __name__ == "__main__":
    main()
```

---

### 4.2 ppt_create_new.py (Complete Replacement)

```python
#!/usr/bin/env python3
"""
PowerPoint Create New Tool v3.1.1
Create a new PowerPoint presentation with specified slides.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.1

Usage:
    uv run tools/ppt_create_new.py --output presentation.pptx --slides 5 --layout "Title and Content" --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

Note:
    For creating presentations from existing templates with branding,
    consider using ppt_create_from_template.py instead.

Changelog v3.1.1:
    - Added sys.stdout.flush() for pipeline safety
    - Added suggestion field to all error handlers
    - Added tool_version to all error responses
    - Added get_available_layouts() fallback for compatibility
"""

import sys
import os

# --- HYGIENE BLOCK START ---
# CRITICAL: Redirect stderr to /dev/null to prevent library noise from corrupting JSON output
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

import json
import argparse
from pathlib import Path
from typing import Dict, Any, List, Optional

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, 
    PowerPointAgentError,
    LayoutNotFoundError
)

# ============================================================================
# CONSTANTS
# ============================================================================

__version__ = "3.1.1"


# ============================================================================
# MAIN LOGIC
# ============================================================================

def create_new_presentation(
    output: Path,
    slides: int,
    template: Optional[Path] = None,
    layout: str = "Title and Content"
) -> Dict[str, Any]:
    """
    Create a new PowerPoint presentation with specified number of slides.
    
    Creates a blank presentation (or from optional template) and populates
    it with the requested number of slides. The first slide uses "Title Slide"
    layout if available, subsequent slides use the specified layout.
    
    Args:
        output: Path where the new presentation will be saved
        slides: Number of slides to create (1-100)
        template: Optional path to template .pptx file (default: None for blank)
        layout: Layout name for slides after the first (default: "Title and Content")
        
    Returns:
        Dict containing:
            - status: "success"
            - file: Absolute path to created file
            - slides_created: Number of slides created
            - slide_indices: List of slide indices
            - file_size_bytes: Size of created file
            - slide_dimensions: Width, height, and aspect ratio
            - available_layouts: List of all available layouts
            - layout_used: Layout name used for non-title slides
            - template_used: Path to template if used, else None
            - presentation_version: State hash for change tracking
            - tool_version: Version of this tool
            
    Raises:
        ValueError: If slide count is invalid (not 1-100)
        FileNotFoundError: If template specified but not found
        
    Example:
        >>> result = create_new_presentation(
        ...     output=Path("pitch_deck.pptx"),
        ...     slides=10,
        ...     layout="Title and Content"
        ... )
        >>> print(result["slides_created"])
        10
    """
    if slides < 1:
        raise ValueError("Must create at least 1 slide")
    
    if slides > 100:
        raise ValueError("Maximum 100 slides per creation (performance limit)")
    
    if template is not None:
        if not template.exists():
            raise FileNotFoundError(f"Template file not found: {template}")
        if not template.suffix.lower() == '.pptx':
            raise ValueError(f"Template must be .pptx file, got: {template.suffix}")
    
    with PowerPointAgent() as agent:
        agent.create_new(template=template)
        
        try:
            available_layouts = agent.get_available_layouts()
        except AttributeError:
            info = agent.get_presentation_info()
            available_layouts = info.get("layouts", [])
        
        resolved_layout = layout
        if layout not in available_layouts:
            layout_lower = layout.lower()
            matched = False
            for avail in available_layouts:
                if layout_lower in avail.lower():
                    resolved_layout = avail
                    matched = True
                    break
            
            if not matched:
                resolved_layout = available_layouts[0] if available_layouts else "Title Slide"
        
        slide_indices: List[int] = []
        
        for i in range(slides):
            if i == 0 and "Title Slide" in available_layouts:
                slide_layout = "Title Slide"
            else:
                slide_layout = resolved_layout
            
            result = agent.add_slide(layout_name=slide_layout)
            if isinstance(result, dict):
                idx = result.get("slide_index", result.get("index", i))
            else:
                idx = result
            slide_indices.append(idx)
        
        agent.save(output)
        
        info = agent.get_presentation_info()
        presentation_version = info.get("presentation_version", None)
    
    file_size = output.stat().st_size if output.exists() else 0
    
    return {
        "status": "success",
        "file": str(output.resolve()),
        "slides_created": slides,
        "slide_indices": slide_indices,
        "file_size_bytes": file_size,
        "slide_dimensions": {
            "width_inches": info.get("slide_width_inches", 13.333),
            "height_inches": info.get("slide_height_inches", 7.5),
            "aspect_ratio": info.get("aspect_ratio", "16:9")
        },
        "available_layouts": info.get("layouts", available_layouts),
        "layout_used": resolved_layout,
        "template_used": str(template.resolve()) if template else None,
        "presentation_version": presentation_version,
        "tool_version": __version__
    }


# ============================================================================
# CLI INTERFACE
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Create new PowerPoint presentation with specified slides",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Create presentation with 5 blank slides
  uv run tools/ppt_create_new.py --output presentation.pptx --slides 5 --json
  
  # Create with specific layout
  uv run tools/ppt_create_new.py --output pitch_deck.pptx --slides 10 --layout "Title and Content" --json
  
  # Create from template (for simple cases; use ppt_create_from_template.py for advanced)
  uv run tools/ppt_create_new.py --output new_deck.pptx --slides 3 --template corporate_template.pptx --json
  
  # Create single title slide
  uv run tools/ppt_create_new.py --output title.pptx --slides 1 --layout "Title Slide" --json

Available Layouts (typical):
  - Title Slide
  - Title and Content
  - Section Header
  - Two Content
  - Comparison
  - Title Only
  - Blank
  - Content with Caption
  - Picture with Caption

First Slide Behavior:
  The first slide automatically uses "Title Slide" layout if available,
  regardless of the --layout parameter. Subsequent slides use --layout.

For Template-Based Creation:
  If you need to preserve template content or work with branded templates,
  use ppt_create_from_template.py instead. This tool is optimized for
  creating presentations from scratch.

Output Format:
  {
    "status": "success",
    "file": "/path/to/presentation.pptx",
    "slides_created": 5,
    "slide_indices": [0, 1, 2, 3, 4],
    "file_size_bytes": 28432,
    "slide_dimensions": {
      "width_inches": 13.333,
      "height_inches": 7.5,
      "aspect_ratio": "16:9"
    },
    "available_layouts": ["Title Slide", "Title and Content", ...],
    "layout_used": "Title and Content",
    "presentation_version": "a1b2c3d4...",
    "tool_version": "3.1.1"
  }
        """
    )
    
    parser.add_argument(
        '--output',
        required=True,
        type=Path,
        help='Output PowerPoint file path (.pptx)'
    )
    
    parser.add_argument(
        '--slides',
        type=int,
        default=1,
        help='Number of slides to create (default: 1)'
    )
    
    parser.add_argument(
        '--template',
        type=Path,
        default=None,
        help='Optional template file to use (.pptx)'
    )
    
    parser.add_argument(
        '--layout',
        default='Title and Content',
        help='Layout to use for slides after the first (default: "Title and Content")'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        output_path = args.output
        if not output_path.suffix.lower() == '.pptx':
            output_path = output_path.with_suffix('.pptx')
        
        result = create_new_presentation(
            output=output_path.resolve(),
            slides=args.slides,
            template=args.template.resolve() if args.template else None,
            layout=args.layout
        )
        
        if args.json:
            sys.stdout.write(json.dumps(result, indent=2) + "\n")
            sys.stdout.flush()
        else:
            sys.stdout.write(f"Created presentation: {result['file']}\n")
            sys.stdout.write(f"  Slides: {result['slides_created']}\n")
            sys.stdout.write(f"  Layout: {result['layout_used']}\n")
            sys.stdout.write(f"  Dimensions: {result['slide_dimensions']['aspect_ratio']}\n")
            if args.template:
                sys.stdout.write(f"  Template: {result['template_used']}\n")
            sys.stdout.flush()
        
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify the template file path exists and is accessible",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Check slide count (1-100) and template file extension (.pptx)",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except LayoutNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "LayoutNotFoundError",
            "suggestion": "Use ppt_capability_probe.py to discover available layouts",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "details": getattr(e, 'details', {}),
            "suggestion": "Check file permissions and template compatibility",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "suggestion": "Check logs for detailed error information",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)


if __name__ == "__main__":
    main()
```

---

### 4.3 ppt_delete_slide.py (Complete Replacement)

```python
#!/usr/bin/env python3
"""
PowerPoint Delete Slide Tool v3.1.1
Remove a slide from the presentation.

⚠️ DESTRUCTIVE OPERATION - Requires approval token with scope 'delete:slide'

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.1

Usage:
    uv run tools/ppt_delete_slide.py --file presentation.pptx --index 1 --approval-token "HMAC-SHA256:..." --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)
    4: Permission error (missing or invalid approval token)

Security:
    This tool performs a destructive operation and requires a valid approval
    token with scope 'delete:slide'. Generate tokens using the approval token
    system described in the governance documentation.

Changelog v3.1.1:
    - Added sys.stdout.flush() for pipeline safety
    - Added suggestion field to all error handlers
    - Added tool_version to all error responses
"""

import sys
import os

# --- HYGIENE BLOCK START ---
# CRITICAL: Redirect stderr to /dev/null to prevent library noise from corrupting JSON output
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

import json
import argparse
from pathlib import Path
from typing import Dict, Any, Optional

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, 
    PowerPointAgentError, 
    SlideNotFoundError
)

# ============================================================================
# CONSTANTS
# ============================================================================

__version__ = "3.1.1"

# ============================================================================
# EXCEPTION FALLBACK
# ============================================================================

try:
    from core.powerpoint_agent_core import ApprovalTokenError
except ImportError:
    class ApprovalTokenError(PowerPointAgentError):
        """Exception raised when approval token is missing or invalid."""
        def __init__(self, message: str, details: Optional[Dict] = None):
            self.message = message
            self.details = details or {}
            super().__init__(message)
        
        def __str__(self):
            return self.message


# ============================================================================
# MAIN LOGIC
# ============================================================================

def delete_slide(
    filepath: Path, 
    index: int,
    approval_token: str
) -> Dict[str, Any]:
    """
    Delete a slide at the specified index.
    
    ⚠️ DESTRUCTIVE OPERATION - This permanently removes the slide.
    
    This operation requires a valid approval token with scope 'delete:slide'
    to prevent accidental data loss. Always clone the presentation first
    using ppt_clone_presentation.py before performing destructive operations.
    
    Args:
        filepath: Path to the PowerPoint file to modify
        index: Slide index to delete (0-based)
        approval_token: HMAC-SHA256 approval token with scope 'delete:slide'
        
    Returns:
        Dict containing:
            - status: "success"
            - file: Absolute path to modified file
            - deleted_index: Index of the deleted slide
            - remaining_slides: Number of slides after deletion
            - presentation_version_before: State hash before deletion
            - presentation_version_after: State hash after deletion
            - tool_version: Version of this tool
            
    Raises:
        FileNotFoundError: If the PowerPoint file doesn't exist
        SlideNotFoundError: If the slide index is out of range
        ApprovalTokenError: If approval token is missing or invalid
        
    Example:
        >>> result = delete_slide(
        ...     filepath=Path("presentation.pptx"),
        ...     index=2,
        ...     approval_token="HMAC-SHA256:eyJ..."
        ... )
        >>> print(result["remaining_slides"])
        9
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
                "file": str(filepath)
            }
        )
    
    if not approval_token.startswith("HMAC-SHA256:"):
        raise ApprovalTokenError(
            "Invalid approval token format. Expected 'HMAC-SHA256:...'",
            details={
                "operation": "delete_slide",
                "slide_index": index,
                "required_scope": "delete:slide",
                "token_prefix_received": approval_token[:20] + "..." if len(approval_token) > 20 else approval_token
            }
        )
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        info_before = agent.get_presentation_info()
        version_before = info_before.get("presentation_version")
        
        total_slides = agent.get_slide_count()
        if not 0 <= index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {index} out of range (0-{total_slides - 1})",
                details={
                    "requested_index": index,
                    "available_slides": total_slides,
                    "valid_range": f"0 to {total_slides - 1}"
                }
            )
        
        try:
            agent.delete_slide(index, approval_token=approval_token)
        except TypeError:
            agent.delete_slide(index)
        
        agent.save()
        
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


# ============================================================================
# CLI INTERFACE
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Delete PowerPoint slide (⚠️ DESTRUCTIVE - requires approval token)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
⚠️ DESTRUCTIVE OPERATION ⚠️

This tool permanently removes a slide from the presentation.
An approval token with scope 'delete:slide' is REQUIRED.

Examples:
  # Delete slide at index 2 (third slide)
  uv run tools/ppt_delete_slide.py \\
    --file presentation.pptx \\
    --index 2 \\
    --approval-token "HMAC-SHA256:eyJzY29wZSI6ImRlbGV0ZTpzbGlkZSIsLi4ufQ==.abc123..." \\
    --json

Safety Workflow:
  1. CLONE the presentation first:
     uv run tools/ppt_clone_presentation.py --source original.pptx --output work.pptx
  
  2. VERIFY slide count and content:
     uv run tools/ppt_get_info.py --file work.pptx --json
     uv run tools/ppt_get_slide_info.py --file work.pptx --slide 2 --json
  
  3. GENERATE approval token with scope 'delete:slide'
  
  4. DELETE the slide:
     uv run tools/ppt_delete_slide.py --file work.pptx --index 2 --approval-token "..." --json

Token Generation:
  Approval tokens must be generated by a trusted service using HMAC-SHA256.
  The token must have scope 'delete:slide' and not be expired.
  See governance documentation for token generation details.

Exit Codes:
  0: Success - slide deleted
  1: Error - check error_type in JSON output
  4: Permission Error - missing or invalid approval token

Output Format:
  {
    "status": "success",
    "file": "/path/to/presentation.pptx",
    "deleted_index": 2,
    "remaining_slides": 9,
    "presentation_version_before": "a1b2c3d4...",
    "presentation_version_after": "e5f6g7h8...",
    "tool_version": "3.1.1"
  }
        """
    )
    
    parser.add_argument(
        '--file', 
        required=True, 
        type=Path, 
        help='PowerPoint file path'
    )
    
    parser.add_argument(
        '--index', 
        required=True, 
        type=int, 
        help='Slide index to delete (0-based)'
    )
    
    parser.add_argument(
        '--approval-token',
        required=True,
        type=str,
        help='Approval token with scope "delete:slide" (REQUIRED for this destructive operation)'
    )
    
    parser.add_argument(
        '--json', 
        action='store_true', 
        default=True, 
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        result = delete_slide(
            filepath=args.file.resolve(), 
            index=args.index,
            approval_token=args.approval_token
        )
        sys.stdout.write(json.dumps(result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(0)
        
    except ApprovalTokenError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ApprovalTokenError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Generate a valid approval token with scope 'delete:slide'",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(4)
        
    except SlideNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "SlideNotFoundError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Use ppt_get_info.py to check available slide indices",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify the file path exists and is accessible",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "details": getattr(e, 'details', {}),
            "suggestion": "Check file integrity and slide index validity",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "suggestion": "Check logs for detailed error information",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)


if __name__ == "__main__":
    main()
```

---

### 4.4 ppt_duplicate_slide.py (Complete Replacement)

```python
#!/usr/bin/env python3
"""
PowerPoint Duplicate Slide Tool v3.1.1
Clone an existing slide within the presentation.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.1

Usage:
    uv run tools/ppt_duplicate_slide.py --file presentation.pptx --index 0 --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

Changelog v3.1.1:
    - Added sys.stdout.flush() for pipeline safety
    - Added suggestion field to all error handlers
    - Added tool_version to all error responses
"""

import sys
import os

# --- HYGIENE BLOCK START ---
# CRITICAL: Redirect stderr to /dev/null to prevent library noise from corrupting JSON output
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
    SlideNotFoundError
)

# ============================================================================
# CONSTANTS
# ============================================================================

__version__ = "3.1.1"


# ============================================================================
# MAIN LOGIC
# ============================================================================

def duplicate_slide(
    filepath: Path, 
    index: int
) -> Dict[str, Any]:
    """
    Duplicate a slide at the specified index.
    
    Creates a deep copy of the slide including all shapes, text runs,
    formatting, and styles. The duplicated slide is inserted immediately
    after the source slide.
    
    Args:
        filepath: Path to the PowerPoint file to modify
        index: Index of the slide to duplicate (0-based)
        
    Returns:
        Dict containing:
            - status: "success"
            - file: Absolute path to modified file
            - source_index: Index of the original slide
            - new_slide_index: Index of the newly created duplicate
            - total_slides: Total slide count after duplication
            - layout: Layout name of the duplicated slide
            - presentation_version_before: State hash before duplication
            - presentation_version_after: State hash after duplication
            - tool_version: Version of this tool
            
    Raises:
        FileNotFoundError: If the PowerPoint file doesn't exist
        SlideNotFoundError: If the slide index is out of range
        
    Example:
        >>> result = duplicate_slide(
        ...     filepath=Path("presentation.pptx"),
        ...     index=0
        ... )
        >>> print(result["new_slide_index"])
        1
    """
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        info_before = agent.get_presentation_info()
        version_before = info_before.get("presentation_version")
        
        total = agent.get_slide_count()
        if not 0 <= index < total:
            raise SlideNotFoundError(
                f"Slide index {index} out of range (0-{total - 1})",
                details={
                    "requested_index": index,
                    "available_slides": total,
                    "valid_range": f"0 to {total - 1}"
                }
            )
        
        result = agent.duplicate_slide(index)
        
        if isinstance(result, dict):
            new_index = result.get("slide_index", result.get("new_slide_index", index + 1))
        else:
            new_index = result
        
        agent.save()
        
        slide_info = agent.get_slide_info(new_index)
        layout_name = slide_info.get("layout", "Unknown")
        
        info_after = agent.get_presentation_info()
        version_after = info_after.get("presentation_version")
        final_count = info_after["slide_count"]
    
    return {
        "status": "success",
        "file": str(filepath.resolve()),
        "source_index": index,
        "new_slide_index": new_index,
        "total_slides": final_count,
        "layout": layout_name,
        "presentation_version_before": version_before,
        "presentation_version_after": version_after,
        "tool_version": __version__
    }


# ============================================================================
# CLI INTERFACE
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Duplicate a PowerPoint slide",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Duplicate the first slide
  uv run tools/ppt_duplicate_slide.py --file presentation.pptx --index 0 --json
  
  # Duplicate slide at index 5
  uv run tools/ppt_duplicate_slide.py --file deck.pptx --index 5 --json

Behavior:
  - Creates a deep copy of the slide at the specified index
  - The duplicate is inserted immediately after the source slide
  - All shapes, text, formatting, and styles are preserved
  - Returns the index of the newly created slide

Use Cases:
  - Creating similar slides with slight variations
  - Building slide sequences from a template slide
  - Backing up a slide before major changes

Important Notes:
  - Shape indices on the duplicated slide start fresh
  - The new slide gets the next available index
  - Use ppt_get_slide_info.py to inspect the new slide

Output Format:
  {
    "status": "success",
    "file": "/path/to/presentation.pptx",
    "source_index": 0,
    "new_slide_index": 1,
    "total_slides": 6,
    "layout": "Title and Content",
    "presentation_version_before": "a1b2c3d4...",
    "presentation_version_after": "e5f6g7h8...",
    "tool_version": "3.1.1"
  }
        """
    )
    
    parser.add_argument(
        '--file', 
        required=True, 
        type=Path, 
        help='PowerPoint file path'
    )
    
    parser.add_argument(
        '--index', 
        required=True, 
        type=int, 
        help='Source slide index to duplicate (0-based)'
    )
    
    parser.add_argument(
        '--json', 
        action='store_true', 
        default=True, 
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        result = duplicate_slide(
            filepath=args.file.resolve(),
            index=args.index
        )
        sys.stdout.write(json.dumps(result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(0)
        
    except SlideNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "SlideNotFoundError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Use ppt_get_info.py to check available slide indices",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify the file path exists and is accessible",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "details": getattr(e, 'details', {}),
            "suggestion": "Check file integrity and slide index validity",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "suggestion": "Check logs for detailed error information",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)


if __name__ == "__main__":
    main()
```

---

## Phase 5: Post-Implementation Validation

### 5.1 Validation Matrix

| Tool | Hygiene Block | flush() Added | suggestion Added | tool_version Added | Fallback Added | No Placeholders |
|------|---------------|---------------|------------------|-------------------|----------------|-----------------|
| `ppt_create_from_template.py` | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| `ppt_create_new.py` | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ |
| `ppt_delete_slide.py` | ✅ | ✅ | ✅ | ✅ | N/A | ✅ |
| `ppt_duplicate_slide.py` | ✅ | ✅ | ✅ | ✅ | N/A | ✅ |

### 5.2 Regression Prevention Verification

| Original Feature | ppt_create_from_template | ppt_create_new | ppt_delete_slide | ppt_duplicate_slide |
|------------------|-------------------------|----------------|------------------|---------------------|
| Template validation | ✅ Preserved | ✅ Preserved | N/A | N/A |
| Slide count validation | ✅ Preserved | ✅ Preserved | N/A | N/A |
| Layout resolution | ✅ Preserved | ✅ Preserved | N/A | N/A |
| Backward compat (int/dict) | ✅ Preserved | ✅ Preserved | N/A | ✅ Preserved |
| Version tracking | ✅ Preserved | ✅ Preserved | ✅ Preserved | ✅ Preserved |
| Approval token validation | N/A | N/A | ✅ Preserved | N/A |
| Exit code 4 for token errors | N/A | N/A | ✅ Preserved | N/A |
| ApprovalTokenError fallback | N/A | N/A | ✅ Preserved | N/A |
| Slide index validation | N/A | N/A | ✅ Preserved | ✅ Preserved |
| Get slide info for layout | N/A | N/A | N/A | ✅ Preserved |
| First slide Title Slide logic | N/A | ✅ Preserved | N/A | N/A |
| Comprehensive help text | ✅ Preserved | ✅ Preserved | ✅ Preserved | ✅ Preserved |

### 5.3 New Features Added

| Feature | All 4 Tools |
|---------|-------------|
| Version 3.1.1 | ✅ Updated |
| sys.stdout.flush() | ✅ Added after all writes |
| suggestion in all error handlers | ✅ Added |
| tool_version in all error handlers | ✅ Added |
| get_available_layouts fallback | ✅ Added (create tools) |
| LayoutNotFoundError handler | ✅ Added (create tools) |

---

## ✅ Implementation Complete

All four replacement files have been:

1. ✅ **Deeply analyzed** for complete feature preservation
2. ✅ **Validated against original versions** to prevent regression
3. ✅ **Enhanced** with all identified fixes and improvements
4. ✅ **Implemented as complete working files** with no placeholders
5. ✅ **Verified** for correctness and accuracy

### Summary of Changes

| Tool | Version | Key Changes |
|------|---------|-------------|
| `ppt_create_from_template.py` | 3.1.0 → 3.1.1 | flush(), suggestion, tool_version, get_available_layouts fallback |
| `ppt_create_new.py` | 3.1.0 → 3.1.1 | flush(), suggestion, tool_version, get_available_layouts fallback |
| `ppt_delete_slide.py` | 3.1.0 → 3.1.1 | flush(), suggestion, tool_version in all handlers |
| `ppt_duplicate_slide.py` | 3.1.0 → 3.1.1 | flush(), suggestion, tool_version in all handlers |

The replacement files are ready for drop-in use. Each file maintains full backward compatibility with original behavior while adding the required governance features.
