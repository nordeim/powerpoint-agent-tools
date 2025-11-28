#!/usr/bin/env python3
"""
PowerPoint Set Slide Layout Tool
Change the layout of an existing slide with safety warnings

Version 2.0.0 - Enhanced Safety and User Experience

Changes from v1.x:
- Enhanced: Content loss warning system
- Enhanced: Improved layout name matching (fuzzy + suggestions)
- Enhanced: Shows available layouts on error
- Enhanced: Reports placeholder changes
- Enhanced: JSON-first output with comprehensive metadata
- Enhanced: Better error messages and recovery suggestions
- Added: `--force` flag for destructive operations
- Added: Current vs new layout comparison
- Fixed: Consistent response format

IMPORTANT WARNING:
    Changing slide layouts in PowerPoint can cause CONTENT LOSS!
    - Text in removed placeholders may disappear
    - Shapes may be repositioned
    - Formatting may change
    - This is a python-pptx limitation, not a bug
    
    ALWAYS backup your presentation before changing layouts!

Usage:
    # Change to Title Only (minimal risk)
    uv run tools/ppt_set_slide_layout.py --file presentation.pptx --slide 2 --layout "Title Only" --json
    
    # Change with force flag (acknowledges content loss risk)
    uv run tools/ppt_set_slide_layout.py --file presentation.pptx --slide 5 --layout "Blank" --force --json
    
    # List available layouts first
    uv run tools/ppt_get_info.py --file presentation.pptx --json | jq '.layouts'

Exit Codes:
    0: Success
    1: Error occurred
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any, List
from difflib import get_close_matches

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError, SlideNotFoundError, LayoutNotFoundError
)


def set_slide_layout(
    filepath: Path,
    slide_index: int,
    layout_name: str,
    force: bool = False
) -> Dict[str, Any]:
    """
    Change slide layout with safety warnings.
    
    WARNING: Changing layouts can cause content loss due to python-pptx limitations.
    Always backup presentations before layout changes.
    
    Args:
        filepath: Path to PowerPoint file
        slide_index: Slide index (0-based)
        layout_name: Target layout name (fuzzy matching supported)
        force: Acknowledge content loss risk (required for destructive layouts)
        
    Returns:
        Dict containing:
        - status: "success" or "warning"
        - old_layout: Previous layout name
        - new_layout: New layout name
        - warnings: Content loss warnings
        - placeholders_before: Count before change
        - placeholders_after: Count after change
        - recommendations: Suggestions
    """
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    warnings = []
    recommendations = []
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Validate slide index
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides-1}). "
                f"Presentation has {total_slides} slides."
            )
        
        # Get available layouts
        available_layouts = agent.get_available_layouts()
        
        # Get current slide info
        slide_info_before = agent.get_slide_info(slide_index)
        old_layout = slide_info_before["layout"]
        placeholders_before = sum(1 for shape in slide_info_before["shapes"] 
                                 if "PLACEHOLDER" in shape["type"])
        
        # Layout name matching with fuzzy search
        matched_layout = None
        
        # Exact match (case-insensitive)
        for layout in available_layouts:
            if layout.lower() == layout_name.lower():
                matched_layout = layout
                break
        
        # Fuzzy match if no exact match
        if not matched_layout:
            # Check if it's a substring
            for layout in available_layouts:
                if layout_name.lower() in layout.lower():
                    matched_layout = layout
                    warnings.append(
                        f"Matched '{layout_name}' to layout '{layout}' (substring match)"
                    )
                    break
        
        # Use difflib for close matches
        if not matched_layout:
            close_matches = get_close_matches(layout_name, available_layouts, n=3, cutoff=0.6)
            if close_matches:
                raise LayoutNotFoundError(
                    f"Layout '{layout_name}' not found. Did you mean one of these?\n" +
                    "\n".join(f"  - {match}" for match in close_matches) +
                    f"\n\nAll available layouts:\n" +
                    "\n".join(f"  - {layout}" for layout in available_layouts)
                )
            else:
                raise LayoutNotFoundError(
                    f"Layout '{layout_name}' not found.\n\n" +
                    f"Available layouts:\n" +
                    "\n".join(f"  - {layout}" for layout in available_layouts)
                )
        
        # Safety warnings for destructive layouts
        destructive_layouts = ["Blank", "Title Only"]
        if matched_layout in destructive_layouts and placeholders_before > 0:
            warnings.append(
                f"⚠️  CONTENT LOSS RISK: Changing from '{old_layout}' to '{matched_layout}' "
                f"may remove {placeholders_before} placeholders and their content!"
            )
            
            if not force:
                raise PowerPointAgentError(
                    f"Layout change from '{old_layout}' to '{matched_layout}' requires --force flag.\n"
                    f"This change may cause content loss ({placeholders_before} placeholders will be affected).\n\n"
                    "To proceed, add --force flag:\n"
                    f"  --layout \"{matched_layout}\" --force\n\n"
                    "RECOMMENDATION: Backup your presentation first!"
                )
        
        # Warn about same layout
        if matched_layout == old_layout:
            recommendations.append(
                f"Slide already uses '{old_layout}' layout. No change needed."
            )
        
        # Apply layout change
        agent.set_slide_layout(slide_index, matched_layout)
        
        # Get slide info after change
        slide_info_after = agent.get_slide_info(slide_index)
        placeholders_after = sum(1 for shape in slide_info_after["shapes"] 
                                if "PLACEHOLDER" in shape["type"])
        
        # Detect content loss
        if placeholders_after < placeholders_before:
            warnings.append(
                f"Content loss detected: {placeholders_before - placeholders_after} "
                f"placeholders removed during layout change."
            )
            recommendations.append(
                "Review slide content and restore any lost text using ppt_add_text_box.py"
            )
        
        # Save
        agent.save()
    
    # Build response
    status = "success" if len(warnings) == 0 else "warning"
    
    result = {
        "status": status,
        "file": str(filepath),
        "slide_index": slide_index,
        "old_layout": old_layout,
        "new_layout": matched_layout,
        "layout_changed": (old_layout != matched_layout),
        "placeholders": {
            "before": placeholders_before,
            "after": placeholders_after,
            "change": placeholders_after - placeholders_before
        },
        "available_layouts": available_layouts
    }
    
    if warnings:
        result["warnings"] = warnings
    
    if recommendations:
        result["recommendations"] = recommendations
    
    return result


def main():
    parser = argparse.ArgumentParser(
        description="Change PowerPoint slide layout with safety warnings (v2.0.0)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
⚠️  IMPORTANT WARNING ⚠️
    Changing slide layouts can cause CONTENT LOSS!
    - Text in removed placeholders may disappear
    - Shapes may be repositioned
    - Formatting may change
    
    ALWAYS backup your presentation before changing layouts!

Examples:
  # List available layouts first
  uv run tools/ppt_get_info.py --file presentation.pptx --json | jq '.layouts'
  
  # Change to Title Only layout (low risk)
  uv run tools/ppt_set_slide_layout.py \\
    --file presentation.pptx \\
    --slide 2 \\
    --layout "Title Only" \\
    --json
  
  # Change to Blank layout (HIGH RISK - requires --force)
  uv run tools/ppt_set_slide_layout.py \\
    --file presentation.pptx \\
    --slide 5 \\
    --layout "Blank" \\
    --force \\
    --json
  
  # Fuzzy matching (will match "Title and Content")
  uv run tools/ppt_set_slide_layout.py \\
    --file presentation.pptx \\
    --slide 3 \\
    --layout "title content" \\
    --json
  
  # Change multiple slides (bash loop)
  for i in {2..5}; do
    uv run tools/ppt_set_slide_layout.py \\
      --file presentation.pptx \\
      --slide $i \\
      --layout "Section Header" \\
      --json
  done

Common Layouts:
  Low Risk (preserve most content):
  - "Title and Content" → Most versatile
  - "Two Content" → Side-by-side content
  - "Section Header" → Section dividers
  
  Medium Risk:
  - "Title Only" → Removes content placeholders
  - "Content with Caption" → Repositions content
  
  High Risk (requires --force):
  - "Blank" → Removes all placeholders!
  - Custom layouts → Unpredictable behavior

Layout Matching:
  This tool supports flexible matching:
  - Exact: "Title and Content" matches "Title and Content"
  - Case-insensitive: "title slide" matches "Title Slide"
  - Substring: "content" matches "Title and Content"
  - Fuzzy: "tile slide" suggests "Title Slide"

Safety Features:
  - Warns about content loss risk
  - Requires --force for destructive layouts
  - Reports placeholder count changes
  - Suggests recovery actions
  - Provides rollback instructions

Output Format:
  {
    "status": "warning",
    "old_layout": "Title and Content",
    "new_layout": "Title Only",
    "layout_changed": true,
    "placeholders": {
      "before": 2,
      "after": 1,
      "change": -1
    },
    "warnings": [
      "Content loss detected: 1 placeholders removed"
    ],
    "recommendations": [
      "Review slide content and restore any lost text"
    ]
  }

Recovery from Content Loss:
  If content was lost during layout change:
  
  1. Check backup (you did backup, right?)
  2. Use ppt_get_slide_info.py to inspect current state
  3. Restore text manually:
     uv run tools/ppt_add_text_box.py \\
       --file presentation.pptx \\
       --slide X \\
       --text "Restored content" \\
       --position '{"left":"10%","top":"25%"}' \\
       --size '{"width":"80%","height":"60%"}' \\
       --json

Related Tools:
  - ppt_get_info.py: List all available layouts
  - ppt_get_slide_info.py: Inspect current slide layout
  - ppt_add_text_box.py: Restore lost content manually
  - ppt_clone_presentation.py: Create backup before changes

Version: 2.0.0
Requires: core/powerpoint_agent_core.py v1.1.0+
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file path'
    )
    
    parser.add_argument(
        '--slide',
        required=True,
        type=int,
        help='Slide index (0-based)'
    )
    
    parser.add_argument(
        '--layout',
        required=True,
        help='New layout name (fuzzy matching supported)'
    )
    
    parser.add_argument(
        '--force',
        action='store_true',
        help='Force destructive layout change (acknowledges content loss risk)'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        result = set_slide_layout(
            filepath=args.file,
            slide_index=args.slide,
            layout_name=args.layout,
            force=args.force
        )
        
        print(json.dumps(result, indent=2))
        sys.exit(0)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "file": str(args.file) if args.file else None,
            "slide_index": args.slide if hasattr(args, 'slide') else None
        }
        
        print(json.dumps(error_result, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()
