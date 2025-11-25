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
    PowerPointAgent, PowerPointAgentError, SlideNotFoundError,
    ColorHelper, RGBColor
)


def validate_shape_params(
    position: Dict[str, Any],
    size: Dict[str, Any],
    fill_color: str = None,
    allow_offslide: bool = False
) -> Dict[str, Any]:
    """
    Validate shape parameters and return warnings/recommendations.
    """
    warnings = []
    recommendations = []
    validation_results = {}
    
    # Position validation (if percentage-based)
    if position:
        try:
            if "left" in position:
                left_str = str(position["left"])
                if left_str.endswith('%'):
                    left_pct = float(left_str.rstrip('%'))
                    if (left_pct < 0 or left_pct > 100) and not allow_offslide:
                        warnings.append(
                            f"Left position {left_pct}% is outside slide bounds (0-100%). "
                            "Shape may not be visible. Use --allow-offslide if intentional."
                        )
            
            if "top" in position:
                top_str = str(position["top"])
                if top_str.endswith('%'):
                    top_pct = float(top_str.rstrip('%'))
                    if (top_pct < 0 or top_pct > 100) and not allow_offslide:
                        warnings.append(
                            f"Top position {top_pct}% is outside slide bounds (0-100%). "
                            "Shape may not be visible. Use --allow-offslide if intentional."
                        )
        except:
            pass
    
    # Size validation
    if size:
        try:
            if "width" in size:
                width_str = str(size["width"])
                if width_str.endswith('%'):
                    width_pct = float(width_str.rstrip('%'))
                    if width_pct < 1:
                        warnings.append(f"Width {width_pct}% is extremely small (<1%). Shape may be invisible.")
            
            if "height" in size:
                height_str = str(size["height"])
                if height_str.endswith('%'):
                    height_pct = float(height_str.rstrip('%'))
                    if height_pct < 1:
                        warnings.append(f"Height {height_pct}% is extremely small (<1%). Shape may be invisible.")
        except:
            pass
            
    # Color contrast validation (if fill color provided)
    if fill_color:
        try:
            # Check contrast against white background
            shape_color = ColorHelper.from_hex(fill_color)
            bg_color = RGBColor(255, 255, 255)
            contrast_ratio = ColorHelper.contrast_ratio(shape_color, bg_color)
            
            validation_results["contrast_ratio"] = round(contrast_ratio, 2)
            
            # Warn if very low contrast (e.g. white shape on white bg)
            if contrast_ratio < 1.1:
                warnings.append(
                    f"Shape fill color has very low contrast ({contrast_ratio:.2f}:1) against white background. "
                    "It may be invisible."
                )
        except Exception as e:
            validation_results["contrast_error"] = str(e)
            
    return {
        "warnings": warnings,
        "recommendations": recommendations,
        "validation_results": validation_results
    }


def add_shape(
    filepath: Path,
    slide_index: int,
    shape_type: str,
    position: Dict[str, Any],
    size: Dict[str, Any],
    fill_color: str = None,
    line_color: str = None,
    line_width: float = 1.0,
    allow_offslide: bool = False
) -> Dict[str, Any]:
    """Add shape to slide with validation."""
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
        
    # Validate parameters
    validation = validate_shape_params(position, size, fill_color, allow_offslide)
    
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
    
    result = {
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
        },
        "validation": validation["validation_results"]
    }
    
    if validation["warnings"]:
        result["warnings"] = validation["warnings"]
        result["status"] = "warning"
        
    if validation["recommendations"]:
        result["recommendations"] = validation["recommendations"]
        
    return result


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
        '--allow-offslide',
        action='store_true',
        help='Allow positioning outside slide bounds (disables off-slide warnings)'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        help='Output JSON response'
    )
    
    args = parser.parse_args()
    
    try:
        # Handle optional size and merge from position
        size = args.size if args.size else {}
        position = args.position
        
        if "width" in position and "width" not in size:
            size["width"] = position["width"]
        if "height" in position and "height" not in size:
            size["height"] = position["height"]

        # Apply defaults if still missing
        if "width" not in size:
            size["width"] = "20%"
        if "height" not in size:
            size["height"] = "20%"

        result = add_shape(
            filepath=args.file,
            slide_index=args.slide,
            shape_type=args.shape,
            position=position,
            size=size,
            fill_color=args.fill_color,
            line_color=args.line_color,
            line_width=args.line_width,
            allow_offslide=args.allow_offslide
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
