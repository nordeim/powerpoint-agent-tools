#!/usr/bin/env python3
"""
PowerPoint Create New Tool
Create a new PowerPoint presentation with specified slides

Usage:
    uv python ppt_create_new.py --output presentation.pptx --slides 5 --layout "Title and Content" --json

Exit Codes:
    0: Success
    1: Error occurred
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any

# Add parent directory to path for core import
sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError
)


def create_new_presentation(
    output: Path,
    slides: int,
    template: Path = None,
    layout: str = "Title and Content"
) -> Dict[str, Any]:
    """Create new PowerPoint presentation."""
    
    if slides < 1:
        raise ValueError("Must create at least 1 slide")
    
    if slides > 100:
        raise ValueError("Maximum 100 slides per creation (performance limit)")
    
    with PowerPointAgent() as agent:
        # Create from template or blank
        agent.create_new(template=template)
        
        # Get available layouts
        available_layouts = agent.get_available_layouts()
        
        # Validate layout
        if layout not in available_layouts:
            # Try to find closest match
            layout_lower = layout.lower()
            for avail in available_layouts:
                if layout_lower in avail.lower():
                    layout = avail
                    break
            else:
                # Use first available layout as fallback
                layout = available_layouts[0] if available_layouts else "Title Slide"
        
        # Add requested number of slides
        slide_indices = []
        for i in range(slides):
            # First slide uses "Title Slide" if available, others use specified layout
            if i == 0 and "Title Slide" in available_layouts:
                slide_layout = "Title Slide"
            else:
                slide_layout = layout
            
            idx = agent.add_slide(layout_name=slide_layout)
            slide_indices.append(idx)
        
        # Save
        agent.save(output)
        
        # Get presentation info
        info = agent.get_presentation_info()
    
    return {
        "status": "success",
        "file": str(output),
        "slides_created": slides,
        "slide_indices": slide_indices,
        "file_size_bytes": output.stat().st_size if output.exists() else 0,
        "slide_dimensions": {
            "width_inches": info["slide_width_inches"],
            "height_inches": info["slide_height_inches"],
            "aspect_ratio": info["aspect_ratio"]
        },
        "available_layouts": info["layouts"],
        "layout_used": layout,
        "template_used": str(template) if template else None
    }


def main():
    parser = argparse.ArgumentParser(
        description="Create new PowerPoint presentation with specified slides",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Create presentation with 5 blank slides
  uv python ppt_create_new.py --output presentation.pptx --slides 5 --json
  
  # Create with specific layout
  uv python ppt_create_new.py --output pitch_deck.pptx --slides 10 --layout "Title and Content" --json
  
  # Create from template
  uv python ppt_create_new.py --output new_deck.pptx --slides 3 --template corporate_template.pptx --json
  
  # Create single title slide
  uv python ppt_create_new.py --output title.pptx --slides 1 --layout "Title Slide" --json

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

Output Format:
  {
    "status": "success",
    "file": "presentation.pptx",
    "slides_created": 5,
    "file_size_bytes": 28432,
    "slide_dimensions": {
      "width_inches": 10.0,
      "height_inches": 7.5,
      "aspect_ratio": "16:9"
    },
    "available_layouts": ["Title Slide", "Title and Content", ...],
    "layout_used": "Title and Content"
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
        help='Optional template file to use (.pptx)'
    )
    
    parser.add_argument(
        '--layout',
        default='Title and Content',
        help='Layout to use for slides (default: "Title and Content")'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        help='Output JSON response'
    )
    
    args = parser.parse_args()
    
    try:
        # Validate template if specified
        if args.template:
            if not args.template.exists():
                raise FileNotFoundError(f"Template file not found: {args.template}")
            if not args.template.suffix.lower() == '.pptx':
                raise ValueError(f"Template must be .pptx file, got: {args.template.suffix}")
        
        # Validate output path
        if not args.output.suffix.lower() == '.pptx':
            args.output = args.output.with_suffix('.pptx')
        
        # Create presentation
        result = create_new_presentation(
            output=args.output,
            slides=args.slides,
            template=args.template,
            layout=args.layout
        )
        
        if args.json:
            print(json.dumps(result, indent=2))
        else:
            print(f"✅ Created presentation: {result['file']}")
            print(f"   Slides: {result['slides_created']}")
            print(f"   Layout: {result['layout_used']}")
            print(f"   Dimensions: {result['slide_dimensions']['aspect_ratio']}")
            if args.template:
                print(f"   Template: {result['template_used']}")
        
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
