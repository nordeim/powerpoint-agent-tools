#!/usr/bin/env python3
"""
PowerPoint Add Shape Tool
Add shape (rectangle, circle, arrow, etc.) to slide

Usage:
    uv python ppt_add_shape.py --file presentation.pptx --slide 0 --shape rectangle --position '{"left":"20%","top":"30%"}' --size '{"width":"60%","height":"40%"}' --fill-color "#0070C0" --json

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
    PowerPointAgent, PowerPointAgentError, SlideNotFoundError
)


def add_shape(
    filepath: Path,
    slide_index: int,
    shape_type: str,
    position: Dict[str, Any],
    size: Dict[str, Any],
    fill_color: str = None,
    line_color: str = None,
    line_width: float = 1.0
) -> Dict[str, Any]:
    """Add shape to slide."""
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Validate slide index
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides-1})"
            )
        
        # Add shape
        agent.add_shape(
            slide_index=slide_index,
            shape_type=shape_type,
            position=position,
            size=size,
            fill_color=fill_color,
            line_color=line_color,
            line_width=line_width
        )
        
        # Save
        agent.save()
    
    return {
        "status": "success",
        "file": str(filepath),
        "slide_index": slide_index,
        "shape_type": shape_type,
        "position": position,
        "size": size,
        "styling": {
            "fill_color": fill_color,
            "line_color": line_color,
            "line_width": line_width
        }
    }


def main():
    parser = argparse.ArgumentParser(
        description="Add shape to PowerPoint slide",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Available Shapes:
  - rectangle: Standard rectangle/box
  - rounded_rectangle: Rectangle with rounded corners
  - ellipse: Circle or oval
  - triangle: Triangle (isosceles)
  - arrow_right: Right-pointing arrow
  - arrow_left: Left-pointing arrow
  - arrow_up: Up-pointing arrow
  - arrow_down: Down-pointing arrow
  - star: 5-point star
  - heart: Heart shape

Common Uses:
  - rectangle: Callout boxes, containers, dividers
  - rounded_rectangle: Buttons, soft containers
  - ellipse: Emphasis, icons, venn diagrams
  - arrows: Process flows, directional indicators
  - star: Highlights, ratings, attention
  - triangle: Warning indicators, play buttons

Examples:
  # Blue callout box
  uv python ppt_add_shape.py \\
    --file presentation.pptx \\
    --slide 1 \\
    --shape rounded_rectangle \\
    --position '{"left":"10%","top":"15%"}' \\
    --size '{"width":"30%","height":"15%"}' \\
    --fill-color "#0070C0" \\
    --line-color "#FFFFFF" \\
    --line-width 2 \\
    --json
  
  # Process flow arrows
  uv python ppt_add_shape.py \\
    --file presentation.pptx \\
    --slide 2 \\
    --shape arrow_right \\
    --position '{"left":"30%","top":"40%"}' \\
    --size '{"width":"15%","height":"8%"}' \\
    --fill-color "#00B050" \\
    --json
  
  # Emphasis circle
  uv python ppt_add_shape.py \\
    --file presentation.pptx \\
    --slide 3 \\
    --shape ellipse \\
    --position '{"anchor":"center"}' \\
    --size '{"width":"20%","height":"20%"}' \\
    --fill-color "#FFC000" \\
    --line-color "#C65911" \\
    --line-width 3 \\
    --json
  
  # Warning triangle
  uv python ppt_add_shape.py \\
    --file presentation.pptx \\
    --slide 4 \\
    --shape triangle \\
    --position '{"left":"5%","top":"5%"}' \\
    --size '{"width":"8%","height":"8%"}' \\
    --fill-color "#FF0000" \\
    --json
  
  # Transparent overlay (no fill)
  uv python ppt_add_shape.py \\
    --file presentation.pptx \\
    --slide 5 \\
    --shape rectangle \\
    --position '{"left":"0%","top":"0%"}' \\
    --size '{"width":"100%","height":"100%"}' \\
    --line-color "#0070C0" \\
    --line-width 5 \\
    --json

Design Tips:
  - Use shapes to organize content visually
  - Consistent colors across shapes
  - Align shapes to grid for professional look
  - Use subtle colors for backgrounds
  - Bold colors for emphasis
  - Combine shapes to create diagrams
  - Layer shapes for depth (background first)

Color Palette (Corporate):
  - Primary Blue: #0070C0
  - Secondary Gray: #595959
  - Accent Orange: #ED7D31
  - Success Green: #70AD47
  - Warning Yellow: #FFC000
  - Danger Red: #C00000
  - White: #FFFFFF
  - Black: #000000

Shape Layering:
  - Shapes are added in order (first = back, last = front)
  - Use transparent shapes for overlays
  - Group related shapes together
  - Send to back: create shape first
  - Bring to front: create shape last
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
        '--shape',
        required=True,
        choices=['rectangle', 'rounded_rectangle', 'ellipse', 'triangle',
                'arrow_right', 'arrow_left', 'arrow_up', 'arrow_down',
                'star', 'heart'],
        help='Shape type'
    )
    
    parser.add_argument(
        '--position',
        required=True,
        type=json.loads,
        help='Position dict (JSON string)'
    )
    
    parser.add_argument(
        '--size',
        required=True,
        type=json.loads,
        help='Size dict (JSON string)'
    )
    
    parser.add_argument(
        '--fill-color',
        help='Fill color (hex, e.g., #0070C0)'
    )
    
    parser.add_argument(
        '--line-color',
        help='Line/border color (hex)'
    )
    
    parser.add_argument(
        '--line-width',
        type=float,
        default=1.0,
        help='Line width in points (default: 1.0)'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        help='Output JSON response'
    )
    
    args = parser.parse_args()
    
    try:
        result = add_shape(
            filepath=args.file,
            slide_index=args.slide,
            shape_type=args.shape,
            position=args.position,
            size=args.size,
            fill_color=args.fill_color,
            line_color=args.line_color,
            line_width=args.line_width
        )
        
        if args.json:
            print(json.dumps(result, indent=2))
        else:
            print(f"✅ Added {result['shape_type']} to slide {result['slide_index']}")
            if args.fill_color:
                print(f"   Fill: {result['styling']['fill_color']}")
            if args.line_color:
                print(f"   Line: {result['styling']['line_color']} ({result['styling']['line_width']}pt)")
        
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
