#!/usr/bin/env python3
"""
PowerPoint Add Slide Tool
Add new slide to existing presentation with specific layout

Usage:
    uv python ppt_add_slide.py --file presentation.pptx --layout "Title and Content" --index 2 --json

Exit Codes:
    0: Success
    1: Error occurred
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any

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
    """Add slide to presentation."""
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Get available layouts
        available_layouts = agent.get_available_layouts()
        
        # Validate layout
        if layout not in available_layouts:
            # Try fuzzy match
            layout_lower = layout.lower()
            for avail in available_layouts:
                if layout_lower in avail.lower():
                    layout = avail
                    break
            else:
                raise ValueError(
                    f"Layout '{layout}' not found. "
                    f"Available: {available_layouts}"
                )
        
        # Add slide
        slide_index = agent.add_slide(layout_name=layout, index=index)
        
        # Set title if provided
        if set_title:
            agent.set_title(slide_index, set_title)
        
        # Get slide info before saving
        slide_info = agent.get_slide_info(slide_index)
        
        # Save
        agent.save()
        
        # Get updated presentation info
        prs_info = agent.get_presentation_info()
    
    return {
        "status": "success",
        "file": str(filepath),
        "slide_index": slide_index,
        "layout": layout,
        "title_set": set_title,
        "total_slides": prs_info["slide_count"],
        "slide_info": {
            "shape_count": slide_info["shape_count"],
            "has_notes": slide_info["has_notes"]
        }
    }


def main():
    parser = argparse.ArgumentParser(
        description="Add new slide to PowerPoint presentation",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Add slide at end
  uv python ppt_add_slide.py --file presentation.pptx --layout "Title and Content" --json
  
  # Add slide at specific position
  uv python ppt_add_slide.py --file deck.pptx --layout "Section Header" --index 2 --json
  
  # Add slide with title
  uv python ppt_add_slide.py --file presentation.pptx --layout "Title Slide" --title "Q4 Results" --json
  
  # Add blank slide
  uv python ppt_add_slide.py --file presentation.pptx --layout "Blank" --json

Common Layouts:
  - Title Slide: For presentation opening
  - Title and Content: Most common layout
  - Section Header: For section breaks
  - Two Content: Side-by-side content
  - Comparison: Compare two items
  - Title Only: Maximum content space
  - Blank: Complete freedom
  - Picture with Caption: Image-focused

Tips:
  - Use --index 0 to insert at beginning
  - Omit --index to add at end
  - Get available layouts: uv python ppt_get_info.py --file your.pptx --json
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
                print(f"   Title: {result['title_set']}")
        
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
