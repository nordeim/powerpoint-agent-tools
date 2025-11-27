Based on the analysis of the execution logs and the source code, I have identified two specific files requiring remediation to restore full functionality and compatibility with `PowerPointAgent` Core v3.1.0.

### 1. `tools/ppt_add_slide.py`
**Issue:** The tool fails with a `TypeError` because it expects `agent.add_slide()` to return an integer (slide index), but Core v3.1.0 returns a dictionary containing metadata and versioning. This causes the internal chaining logic (adding a title after creating the slide) to crash.
**Fix:** 
- Update the tool to extract `slide_index` from the returned dictionary.
- Capture and report the new versioning data (`presentation_version_before`/`after`).

### 2. `tools/ppt_set_footer.py`
**Issue:** The tool fails with an `ImportError` because it attempts to import `PP_PLACEHOLDER` from `core.powerpoint_agent_core`, but this constant is no longer exported in Core v3.1.0.
**Fix:** 
- Remove the invalid import from `core`.
- Import `PP_PLACEHOLDER` directly from `pptx.enum.shapes` (the standard library source).
- Retain the robust "Dual Strategy" logic (placeholders vs. text boxes).

---

Here are the complete, updated replacement files.

### File 1: `tools/ppt_add_slide.py`

```python
#!/usr/bin/env python3
"""
PowerPoint Add Slide Tool v3.1.0
Add new slide to existing presentation with specific layout.
Compatible with PowerPoint Agent Core v3.1.0 (Dictionary Returns).

Usage:
    uv run tools/ppt_add_slide.py --file presentation.pptx --layout "Title and Content" --index 2 --json

Exit Codes:
    0: Success
    1: Error occurred
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any

# Ensure we can import from core
sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError
)


def add_slide(
    filepath: Path,
    layout: str,
    index: int = None,
    set_title: str = None
) -> Dict[str, Any]:
    """
    Add slide to presentation.
    
    Handles the v3.1.0 Core API where add_slide returns a dictionary.
    """
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Get available layouts for validation
        available_layouts = agent.get_available_layouts()
        
        # Validate layout with fuzzy matching
        if layout not in available_layouts:
            layout_lower = layout.lower()
            match_found = False
            for avail in available_layouts:
                if layout_lower in avail.lower():
                    layout = avail
                    match_found = True
                    break
            
            if not match_found:
                raise ValueError(
                    f"Layout '{layout}' not found. "
                    f"Available: {available_layouts}"
                )
        
        # Add slide (Core v3.1.0 returns a dict)
        add_result = agent.add_slide(layout_name=layout, index=index)
        
        # CRITICAL FIX: Extract the integer index from the returned dictionary
        # The error "TypeError: '<=' not supported between 'int' and 'dict'" 
        # occurred because the previous tool assumed add_result was an int.
        slide_index = add_result["slide_index"]
        version_before = add_result.get("presentation_version_before")
        
        # Set title if provided
        title_set_result = None
        if set_title:
            # Pass the extracted integer index
            title_set_result = agent.set_title(slide_index, set_title)
        
        # Get slide info before saving (for verification)
        slide_info = agent.get_slide_info(slide_index)
        
        # Save the file
        agent.save()
        
        # Get updated presentation info (includes final version hash)
        prs_info = agent.get_presentation_info()
    
    # Construct comprehensive response
    return {
        "status": "success",
        "file": str(filepath),
        "slide_index": slide_index,
        "layout": layout,
        "title_set": set_title,
        "title_set_success": title_set_result.get("title_set") if title_set_result else False,
        "total_slides": prs_info["slide_count"],
        "slide_info": {
            "shape_count": slide_info["shape_count"],
            "has_notes": slide_info["has_notes"]
        },
        # Version Tracking (Observability)
        "presentation_version_before": version_before,
        "presentation_version_after": prs_info["presentation_version"]
    }


def main():
    parser = argparse.ArgumentParser(
        description="Add new slide to PowerPoint presentation",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Add slide at end
  uv run tools/ppt_add_slide.py --file presentation.pptx --layout "Title and Content" --json
  
  # Add slide at specific position
  uv run tools/ppt_add_slide.py --file deck.pptx --layout "Section Header" --index 2 --json
  
  # Add slide with title
  uv run tools/ppt_add_slide.py --file presentation.pptx --layout "Title Slide" --title "Q4 Results" --json

Common Layouts:
  - Title Slide
  - Title and Content
  - Section Header
  - Two Content
  - Blank
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file path'
    )
    
    parser.add_argument(
        '--layout',
        required=True,
        help='Layout name for new slide'
    )
    
    parser.add_argument(
        '--index',
        type=int,
        help='Position to insert slide (0-based, default: end)'
    )
    
    parser.add_argument(
        '--title',
        help='Optional title text to set'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        help='Output JSON response'
    )
    
    args = parser.parse_args()
    
    try:
        result = add_slide(
            filepath=args.file,
            layout=args.layout,
            index=args.index,
            set_title=args.title
        )
        
        if args.json:
            print(json.dumps(result, indent=2))
        else:
            print(f"✅ Added slide to: {result['file']}")
            print(f"   Layout: {result['layout']}")
            print(f"   Position: {result['slide_index']}")
            print(f"   Total slides: {result['total_slides']}")
            if args.title:
                print(f"   Title set: {result['title_set_success']}")
            print(f"   Version: {result['presentation_version_after']}")
        
        sys.exit(0)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__
        }
        
        if args.json:
            print(json.dumps(error_result, indent=2))
        else:
            print(f"❌ Error: {e}", file=sys.stderr)
        
        sys.exit(1)


if __name__ == "__main__":
    main()
```

### File 2: `tools/ppt_set_footer.py`

```python
#!/usr/bin/env python3
"""
PowerPoint Set Footer Tool v3.1.0
Configure slide footer, date, and slide number.

Dual-Strategy:
1. Primary: Attempts to use native placeholder footers (if template supports them)
2. Fallback: Creates text box overlays for presentations without footer placeholders

Changes in v3.1.0:
- Fixed: Removed dependency on core.PP_PLACEHOLDER (imports directly from pptx)
- Verified: Compatible with Core v3.1.0 dictionary return values

Usage:
    uv run tools/ppt_set_footer.py --file deck.pptx --text "Confidential" --json
    uv run tools/ppt_set_footer.py --file deck.pptx --text "Company" --show-number --json

Exit Codes:
    0: Success
    1: Error
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any, List

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import PowerPointAgent

# Import PP_PLACEHOLDER directly from the underlying library
# This ensures independence from Core exports which may change
try:
    from pptx.enum.shapes import PP_PLACEHOLDER
except ImportError:
    # Fallback constants if pptx is missing (unlikely) or structure changes
    class PP_PLACEHOLDER:
        FOOTER = 15
        SLIDE_NUMBER = 13
        DATE = 16


def set_footer(
    filepath: Path,
    text: str = None,
    show_number: bool = False,
    show_date: bool = False,
    apply_to_all: bool = True
) -> Dict[str, Any]:
    """
    Set footer on presentation slides using dual-strategy approach.
    """
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")

    warnings = []
    slide_indices_updated = set()
    method_used = None
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # ====================================================================
        # STRATEGY 1: Try native placeholder approach
        # ====================================================================
        
        placeholder_count = 0
        
        if text:
            # Try to set footer text on master slide layouts
            try:
                for master in agent.prs.slide_masters:
                    for layout in master.slide_layouts:
                        for shape in layout.placeholders:
                            if shape.placeholder_format.type == PP_PLACEHOLDER.FOOTER:
                                try:
                                    shape.text = text
                                    placeholder_count += 1
                                except:
                                    pass
            except Exception as e:
                warnings.append(f"Could not access master slide layouts: {str(e)}")
            
            # Try to set footer text on individual slides
            for slide_idx, slide in enumerate(agent.prs.slides):
                for shape in slide.placeholders:
                    if shape.placeholder_format.type == PP_PLACEHOLDER.FOOTER:
                        try:
                            shape.text = text
                            slide_indices_updated.add(slide_idx)
                            placeholder_count += 1
                        except:
                            pass
        
        # ====================================================================
        # STRATEGY 2: Fallback to text box overlay if no placeholders found
        # ====================================================================
        
        textbox_count = 0
        
        if placeholder_count == 0:
            if text:
                warnings.append(
                    "No footer placeholders found. Using text box overlay strategy."
                )
            
            # Skip title slide (slide 0) for footer overlays
            for slide_idx in range(1, len(agent.prs.slides)):
                try:
                    # Add footer text box (bottom-left)
                    if text:
                        agent.add_text_box(
                            slide_index=slide_idx,
                            text=text,
                            position={"left": "5%", "top": "92%"},
                            size={"width": "60%", "height": "5%"},
                            font_size=10,
                            color="#595959",
                            alignment="left"
                        )
                        textbox_count += 1
                        slide_indices_updated.add(slide_idx)
                    
                    # Add slide number box (bottom-right)
                    if show_number:
                        # Display number is 1-indexed
                        display_number = slide_idx + 1
                        agent.add_text_box(
                            slide_index=slide_idx,
                            text=str(display_number),
                            position={"left": "92%", "top": "92%"},
                            size={"width": "5%", "height": "5%"},
                            font_size=10,
                            color="#595959",
                            alignment="left"
                        )
                        textbox_count += 1
                        slide_indices_updated.add(slide_idx)
                        
                except Exception as e:
                    warnings.append(f"Failed to add text box to slide {slide_idx}: {str(e)}")
        
        # Determine which method was used
        if placeholder_count > 0 and textbox_count == 0:
            method_used = "placeholder"
        elif textbox_count > 0 and placeholder_count == 0:
            method_used = "text_box"
        elif placeholder_count > 0 and textbox_count > 0:
            method_used = "hybrid"
        else:
            method_used = "none"
            if text or show_number:
                warnings.append("No footer elements were added.")
        
        # ====================================================================
        # Handle --show-date flag
        # ====================================================================
        
        if show_date:
            warnings.append(
                "Date placeholder activation not yet implemented via python-pptx."
            )
        
        # Save changes
        agent.save()
        
        # Capture new version
        prs_info = agent.get_presentation_info()
    
    # ====================================================================
    # Build comprehensive response
    # ====================================================================
    
    status = "success" if len(slide_indices_updated) > 0 else "warning"
    
    result = {
        "status": status,
        "file": str(filepath),
        "method_used": method_used,
        "footer_text": text,
        "settings": {
            "show_number": show_number,
            "show_date": show_date
        },
        "slides_updated": len(slide_indices_updated),
        "slide_indices": sorted(list(slide_indices_updated)),
        "details": {
            "placeholder_count": placeholder_count,
            "textbox_count": textbox_count
        },
        "presentation_version_after": prs_info["presentation_version"]
    }
    
    if warnings:
        result["warnings"] = warnings
        
    if method_used == "text_box":
        result["recommendation"] = (
            "Template lacks footer placeholders. Text box overlays were used."
        )
    
    return result


def main():
    parser = argparse.ArgumentParser(
        description="Set PowerPoint footer with dual-strategy approach",
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    
    parser.add_argument('--file', required=True, type=Path, help='PowerPoint file path')
    parser.add_argument('--text', help='Footer text content')
    parser.add_argument('--show-number', action='store_true', help='Show slide numbers (text box overlay)')
    parser.add_argument('--show-date', action='store_true', help='Show date (placeholder)')
    parser.add_argument('--json', action='store_true', default=True, help='Output JSON response')
    
    args = parser.parse_args()
    
    try:
        result = set_footer(
            filepath=args.file, 
            text=args.text, 
            show_number=args.show_number,
            show_date=args.show_date
        )
        
        print(json.dumps(result, indent=2))
        sys.exit(0)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "file": str(args.file) if args.file else None
        }
        
        print(json.dumps(error_result, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()
```

