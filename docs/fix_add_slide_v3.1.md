# Meticulous Implementation Plan

## 1. Analysis & Root Cause Verification
The execution log explicitly points to a `TypeError: '<=' not supported between instances of 'int' and 'dict'` during **Step 3: Create Executive Summary**, which executes `ppt_add_slide.py`.

### 🔍 Root Cause
1.  **API Mismatch**: The `PowerPointAgent` Core Library was updated to v3.1.0. In this version, `add_slide()` returns a **Dictionary** containing metadata (e.g., `{'slide_index': 1, 'presentation_version_before': ...}`).
2.  **Legacy Tool Logic**: The tool `ppt_add_slide.py` (likely v2.0) expects `add_slide()` to return an **Integer** (the slide index).
3.  **Failure Point**: The tool captures the dictionary into `slide_index`, then attempts to pass this dictionary to `agent.set_title(slide_index, ...)` or `agent.get_slide_info(slide_index)`. Inside the core, `_get_slide(index)` tries to compare `0 <= index`, causing the crash.

## 2. Implementation Plan: `tools/ppt_add_slide.py` (v3.1.0)

I will create a complete replacement for `tools/ppt_add_slide.py` that is fully compatible with Core v3.1.0.

### 🛠️ Remediation Checklist
- [x] **Extract Index**: specific logic to extract `slide_index` from the dictionary returned by `agent.add_slide()`.
- [x] **Preserve Versioning**: Capture `presentation_version_before` (from `add_slide` result) and `presentation_version_after` (from `get_presentation_info` after save) to maintain the audit trail required by the project.
- [x] **Layout Validation**: Retain the existing "fuzzy match" logic for layout names as it is user-friendly.
- [x] **Chain Operations**: Ensure `set_title` uses the extracted integer index.
- [x] **JSON Output**: Ensure the final JSON response structure aligns with the demo script's expectations (e.g., includes `slide_info`).

### 📜 Code Structure Definition
```python
def add_slide(...):
    # ... setup ...
    with PowerPointAgent(filepath) as agent:
        # ...
        # 1. Execute & Extract
        result_dict = agent.add_slide(layout_name=layout, index=index)
        slide_index = result_dict["slide_index"]  # <--- CRITICAL FIX
        
        # 2. Chain Operations
        if set_title:
            agent.set_title(slide_index, set_title) # Uses int
            
        # ... get info & save ...
```

## 3. Validation Plan
After generating the code, I will verify:
1.  **Imports**: Correctly references `core.powerpoint_agent_core`.
2.  **Logic**: `slide_index` variable is strictly an `int` before being used in subsequent calls.
3.  **Output**: Returns a valid JSON object with `status: success`.

---

I am now ready to generate the corrected file.

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

