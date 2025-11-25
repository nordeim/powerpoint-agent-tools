#!/usr/bin/env python3
"""
PowerPoint Add Shape Tool v3.0
Add shapes (rectangle, circle, arrow, etc.) to slides with comprehensive styling options.

Fully aligned with PowerPoint Agent Core v3.0 and System Prompt v3.0.

Author: PowerPoint Agent Team
License: MIT
Version: 3.0.0

Usage:
    uv run tools/ppt_add_shape.py --file presentation.pptx --slide 0 \\
        --shape rectangle --position '{"left":"20%","top":"30%"}' \\
        --size '{"width":"60%","height":"40%"}' --fill-color "#0070C0" --json

Exit Codes:
    0: Success
    1: Error occurred

Changelog v3.0.0:
- Aligned with PowerPoint Agent Core v3.0
- Added text inside shapes support
- Enhanced validation with contrast warnings
- Captures and reports shape_index from core
- Added all shape types from v3.0 core
- Improved error handling and messages
- Added z-order awareness documentation
- Consistent JSON output structure
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any, List, Optional

# Add parent directory to path for core import
sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent,
    PowerPointAgentError,
    SlideNotFoundError,
    ShapeNotFoundError,
    ColorHelper,
    Position,
    Size,
    SLIDE_WIDTH_INCHES,
    SLIDE_HEIGHT_INCHES,
    CORPORATE_COLORS,
    __version__ as CORE_VERSION
)

# ============================================================================
# CONSTANTS
# ============================================================================

__version__ = "3.0.0"

# Available shape types (aligned with v3.0 core)
AVAILABLE_SHAPES = [
    "rectangle",
    "rounded_rectangle",
    "ellipse",
    "oval",
    "triangle",
    "arrow_right",
    "arrow_left",
    "arrow_up",
    "arrow_down",
    "diamond",
    "pentagon",
    "hexagon",
    "star",
    "heart",
    "lightning",
    "sun",
    "moon",
    "cloud",
]

# Shape type aliases for user convenience
SHAPE_ALIASES = {
    "rect": "rectangle",
    "round_rect": "rounded_rectangle",
    "circle": "ellipse",
    "arrow": "arrow_right",
    "right_arrow": "arrow_right",
    "left_arrow": "arrow_left",
    "up_arrow": "arrow_up",
    "down_arrow": "arrow_down",
    "5_point_star": "star",
}

# Default corporate colors for quick reference
COLOR_PRESETS = {
    "primary": "#0070C0",
    "secondary": "#595959",
    "accent": "#ED7D31",
    "success": "#70AD47",
    "warning": "#FFC000",
    "danger": "#C00000",
    "white": "#FFFFFF",
    "black": "#000000",
    "light_gray": "#D9D9D9",
    "dark_gray": "#404040",
}


# ============================================================================
# VALIDATION FUNCTIONS
# ============================================================================

def resolve_shape_type(shape_type: str) -> str:
    """
    Resolve shape type, handling aliases.
    
    Args:
        shape_type: User-provided shape type
        
    Returns:
        Canonical shape type name
    """
    shape_lower = shape_type.lower().strip()
    
    # Check aliases first
    if shape_lower in SHAPE_ALIASES:
        return SHAPE_ALIASES[shape_lower]
    
    # Check direct match
    if shape_lower in AVAILABLE_SHAPES:
        return shape_lower
    
    # Try to find partial match
    for available in AVAILABLE_SHAPES:
        if shape_lower in available or available in shape_lower:
            return available
    
    return shape_lower  # Let core handle unknown types


def resolve_color(color: Optional[str]) -> Optional[str]:
    """
    Resolve color, handling presets and validation.
    
    Args:
        color: User-provided color (hex or preset name)
        
    Returns:
        Resolved hex color or None
    """
    if color is None:
        return None
    
    color_lower = color.lower().strip()
    
    # Check presets
    if color_lower in COLOR_PRESETS:
        return COLOR_PRESETS[color_lower]
    
    # Ensure # prefix for hex colors
    if not color.startswith('#') and len(color) == 6:
        try:
            int(color, 16)
            return f"#{color}"
        except ValueError:
            pass
    
    return color


def validate_shape_params(
    position: Dict[str, Any],
    size: Dict[str, Any],
    fill_color: Optional[str] = None,
    line_color: Optional[str] = None,
    text: Optional[str] = None,
    allow_offslide: bool = False
) -> Dict[str, Any]:
    """
    Validate shape parameters and return warnings/recommendations.
    
    Args:
        position: Position specification
        size: Size specification
        fill_color: Fill color hex
        line_color: Line color hex
        text: Text to add inside shape
        allow_offslide: Allow off-slide positioning
        
    Returns:
        Dict with warnings, recommendations, and validation results
    """
    warnings: List[str] = []
    recommendations: List[str] = []
    validation_results: Dict[str, Any] = {}
    
    # Position validation
    if position:
        _validate_position(position, warnings, allow_offslide)
    
    # Size validation
    if size:
        _validate_size(size, warnings)
    
    # Color contrast validation
    if fill_color:
        _validate_color_contrast(
            fill_color, line_color, text,
            warnings, recommendations, validation_results
        )
    
    # Text validation
    if text:
        _validate_text(text, warnings, recommendations)
    
    return {
        "warnings": warnings,
        "recommendations": recommendations,
        "validation_results": validation_results,
        "has_warnings": len(warnings) > 0
    }


def _validate_position(
    position: Dict[str, Any],
    warnings: List[str],
    allow_offslide: bool
) -> None:
    """Validate position values."""
    try:
        for key in ["left", "top"]:
            if key in position:
                value_str = str(position[key])
                if value_str.endswith('%'):
                    pct = float(value_str.rstrip('%'))
                    if not allow_offslide and (pct < 0 or pct > 100):
                        warnings.append(
                            f"Position '{key}' is {pct}% which is outside slide bounds (0-100%). "
                            f"Shape may not be visible. Use --allow-offslide if intentional."
                        )
    except (ValueError, TypeError):
        pass


def _validate_size(size: Dict[str, Any], warnings: List[str]) -> None:
    """Validate size values."""
    try:
        for key in ["width", "height"]:
            if key in size:
                value_str = str(size[key])
                if value_str.endswith('%'):
                    pct = float(value_str.rstrip('%'))
                    if pct <= 0:
                        warnings.append(
                            f"Size '{key}' is {pct}% which is invalid (must be > 0%)."
                        )
                    elif pct < 1:
                        warnings.append(
                            f"Size '{key}' is {pct}% which is extremely small (<1%). "
                            f"Shape may be invisible."
                        )
                    elif pct > 100:
                        warnings.append(
                            f"Size '{key}' is {pct}% which exceeds slide dimensions."
                        )
    except (ValueError, TypeError):
        pass


def _validate_color_contrast(
    fill_color: str,
    line_color: Optional[str],
    text: Optional[str],
    warnings: List[str],
    recommendations: List[str],
    validation_results: Dict[str, Any]
) -> None:
    """Validate color contrast for visibility and accessibility."""
    try:
        # Parse fill color
        shape_rgb = ColorHelper.from_hex(fill_color)
        
        # Check contrast against white background
        from pptx.dml.color import RGBColor
        white_bg = RGBColor(255, 255, 255)
        bg_contrast = ColorHelper.contrast_ratio(shape_rgb, white_bg)
        validation_results["fill_contrast_vs_white"] = round(bg_contrast, 2)
        
        # Warn if very low contrast (shape may be invisible)
        if bg_contrast < 1.1:
            warnings.append(
                f"Fill color {fill_color} has very low contrast ({bg_contrast:.2f}:1) "
                f"against white background. Shape may be invisible on light slides."
            )
        
        # Check contrast against black background
        black_bg = RGBColor(0, 0, 0)
        dark_contrast = ColorHelper.contrast_ratio(shape_rgb, black_bg)
        validation_results["fill_contrast_vs_black"] = round(dark_contrast, 2)
        
        if dark_contrast < 1.1:
            warnings.append(
                f"Fill color {fill_color} has very low contrast ({dark_contrast:.2f}:1) "
                f"against black background. Shape may be invisible on dark slides."
            )
        
        # If shape has text, check text readability
        if text:
            # Assume white text by default
            text_rgb = RGBColor(255, 255, 255)
            text_contrast = ColorHelper.contrast_ratio(text_rgb, shape_rgb)
            validation_results["text_contrast_white"] = round(text_contrast, 2)
            
            # Also check black text
            text_rgb_black = RGBColor(0, 0, 0)
            text_contrast_black = ColorHelper.contrast_ratio(text_rgb_black, shape_rgb)
            validation_results["text_contrast_black"] = round(text_contrast_black, 2)
            
            # Recommend better text color
            if text_contrast < 4.5 and text_contrast_black >= 4.5:
                recommendations.append(
                    f"Consider using dark text on this fill color for better readability "
                    f"(black text contrast: {text_contrast_black:.2f}:1)."
                )
            elif text_contrast_black < 4.5 and text_contrast >= 4.5:
                recommendations.append(
                    f"White text provides good contrast on this fill color "
                    f"({text_contrast:.2f}:1 meets WCAG AA)."
                )
            elif text_contrast < 4.5 and text_contrast_black < 4.5:
                warnings.append(
                    f"Neither white nor black text has sufficient contrast on fill color "
                    f"{fill_color}. Text may be hard to read."
                )
        
        # Line color validation
        if line_color:
            line_rgb = ColorHelper.from_hex(line_color)
            line_fill_contrast = ColorHelper.contrast_ratio(line_rgb, shape_rgb)
            validation_results["line_fill_contrast"] = round(line_fill_contrast, 2)
            
            if line_fill_contrast < 1.5:
                warnings.append(
                    f"Line color {line_color} has low contrast ({line_fill_contrast:.2f}:1) "
                    f"against fill color {fill_color}. Border may not be visible."
                )
    
    except Exception as e:
        validation_results["color_validation_error"] = str(e)


def _validate_text(
    text: str,
    warnings: List[str],
    recommendations: List[str]
) -> None:
    """Validate text content."""
    if len(text) > 500:
        warnings.append(
            f"Text content is very long ({len(text)} characters). "
            f"Consider using a text box instead."
        )
    
    if '\n' in text and text.count('\n') > 5:
        recommendations.append(
            "Text has multiple lines. Consider using bullet points or a text box "
            "for better formatting control."
        )


# ============================================================================
# MAIN FUNCTION
# ============================================================================

def add_shape(
    filepath: Path,
    slide_index: int,
    shape_type: str,
    position: Dict[str, Any],
    size: Dict[str, Any],
    fill_color: Optional[str] = None,
    line_color: Optional[str] = None,
    line_width: float = 1.0,
    text: Optional[str] = None,
    allow_offslide: bool = False
) -> Dict[str, Any]:
    """
    Add shape to slide with comprehensive validation.
    
    Args:
        filepath: Path to PowerPoint file
        slide_index: Target slide index (0-based)
        shape_type: Type of shape to add
        position: Position specification dict
        size: Size specification dict
        fill_color: Fill color (hex or preset name)
        line_color: Line/border color (hex or preset name)
        line_width: Line width in points
        text: Optional text to add inside shape
        allow_offslide: Allow positioning outside slide bounds
        
    Returns:
        Result dict with shape details and validation info
        
    Raises:
        FileNotFoundError: If file doesn't exist
        SlideNotFoundError: If slide index is invalid
        PowerPointAgentError: If shape creation fails
    """
    # Validate file exists
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    # Resolve shape type
    resolved_shape = resolve_shape_type(shape_type)
    
    # Resolve colors
    resolved_fill = resolve_color(fill_color)
    resolved_line = resolve_color(line_color)
    
    # Validate parameters
    validation = validate_shape_params(
        position=position,
        size=size,
        fill_color=resolved_fill,
        line_color=resolved_line,
        text=text,
        allow_offslide=allow_offslide
    )
    
    # Open presentation and add shape
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Validate slide index
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range",
                details={
                    "slide_index": slide_index,
                    "total_slides": total_slides,
                    "valid_range": f"0-{total_slides-1}"
                }
            )
        
        # Get presentation version before change
        version_before = agent.get_presentation_version()
        
        # Add shape using v3.0 core (returns dict with shape_index)
        add_result = agent.add_shape(
            slide_index=slide_index,
            shape_type=resolved_shape,
            position=position,
            size=size,
            fill_color=resolved_fill,
            line_color=resolved_line,
            line_width=line_width,
            text=text
        )
        
        # Save changes
        agent.save()
        
        # Get presentation version after change
        version_after = agent.get_presentation_version()
    
    # Build result
    result = {
        "status": "success" if not validation["has_warnings"] else "warning",
        "file": str(filepath),
        "slide_index": slide_index,
        "shape_type": resolved_shape,
        "shape_type_requested": shape_type,
        "shape_index": add_result.get("shape_index"),
        "position": add_result.get("position", position),
        "size": add_result.get("size", size),
        "styling": {
            "fill_color": resolved_fill,
            "line_color": resolved_line,
            "line_width": line_width
        },
        "text": text,
        "presentation_version": {
            "before": version_before,
            "after": version_after
        },
        "core_version": CORE_VERSION,
        "tool_version": __version__
    }
    
    # Add validation results
    if validation["validation_results"]:
        result["validation"] = validation["validation_results"]
    
    if validation["warnings"]:
        result["warnings"] = validation["warnings"]
    
    if validation["recommendations"]:
        result["recommendations"] = validation["recommendations"]
    
    # Add z-order note
    result["notes"] = [
        "Shape added to top of z-order (in front of existing shapes).",
        "Use ppt_set_z_order.py to change layering if needed.",
        f"Shape index {add_result.get('shape_index')} may change if other shapes are added/removed."
    ]
    
    return result


# ============================================================================
# CLI INTERFACE
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Add shape to PowerPoint slide (v3.0)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

AVAILABLE SHAPES:
  Basic:        rectangle, rounded_rectangle, ellipse/oval, triangle, diamond
  Arrows:       arrow_right, arrow_left, arrow_up, arrow_down
  Polygons:     pentagon, hexagon
  Decorative:   star, heart, lightning, sun, moon, cloud

SHAPE ALIASES:
  rect → rectangle, circle → ellipse, arrow → arrow_right

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

POSITION FORMATS:
  Percentage:   {"left": "20%", "top": "30%"}
  Inches:       {"left": 2.0, "top": 3.0}
  Anchor:       {"anchor": "center", "offset_x": 0, "offset_y": -1.0}
  Grid:         {"grid_row": 2, "grid_col": 3, "grid_size": 12}

ANCHOR POINTS:
  top_left, top_center, top_right
  center_left, center, center_right
  bottom_left, bottom_center, bottom_right

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

COLOR PRESETS:
  primary (#0070C0)    secondary (#595959)    accent (#ED7D31)
  success (#70AD47)    warning (#FFC000)      danger (#C00000)
  white (#FFFFFF)      black (#000000)
  light_gray (#D9D9D9) dark_gray (#404040)

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

EXAMPLES:

  # Blue callout box with text
  uv run tools/ppt_add_shape.py \\
    --file presentation.pptx --slide 0 --shape rounded_rectangle \\
    --position '{"left":"10%","top":"15%"}' \\
    --size '{"width":"30%","height":"15%"}' \\
    --fill-color primary --text "Key Point" --json

  # Centered circle with border
  uv run tools/ppt_add_shape.py \\
    --file presentation.pptx --slide 1 --shape ellipse \\
    --position '{"anchor":"center"}' \\
    --size '{"width":"20%","height":"20%"}' \\
    --fill-color "#FFC000" --line-color "#000000" --line-width 2 --json

  # Process flow arrow
  uv run tools/ppt_add_shape.py \\
    --file presentation.pptx --slide 2 --shape arrow_right \\
    --position '{"left":"30%","top":"40%"}' \\
    --size '{"width":"15%","height":"8%"}' \\
    --fill-color success --json

  # Full-slide overlay (for backgrounds)
  uv run tools/ppt_add_shape.py \\
    --file presentation.pptx --slide 3 --shape rectangle \\
    --position '{"left":"0%","top":"0%"}' \\
    --size '{"width":"100%","height":"100%"}' \\
    --fill-color "#000000" --json
    # Note: Use ppt_set_z_order.py --action send_to_back after adding

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Z-ORDER (LAYERING):
  Shapes are added on TOP of existing shapes by default.
  To create background shapes:
    1. Add the shape
    2. Run: ppt_set_z_order.py --file FILE --slide N --shape INDEX --action send_to_back
  
  Shape indices change after z-order operations - always re-query with
  ppt_get_slide_info.py before referencing shapes by index.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        """
    )
    
    # Required arguments
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file path (must exist)'
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
        help=f'Shape type: {", ".join(AVAILABLE_SHAPES[:8])}... (see --help for full list)'
    )
    
    parser.add_argument(
        '--position',
        required=True,
        type=json.loads,
        help='Position dict as JSON string (e.g., \'{"left":"20%%","top":"30%%"}\')'
    )
    
    # Optional arguments
    parser.add_argument(
        '--size',
        type=json.loads,
        help='Size dict as JSON string (e.g., \'{"width":"40%%","height":"30%%"}\'). '
             'Defaults to 20%% x 20%% if not specified.'
    )
    
    parser.add_argument(
        '--fill-color',
        help='Fill color: hex (#0070C0) or preset name (primary, success, etc.)'
    )
    
    parser.add_argument(
        '--line-color',
        help='Line/border color: hex or preset name'
    )
    
    parser.add_argument(
        '--line-width',
        type=float,
        default=1.0,
        help='Line width in points (default: 1.0)'
    )
    
    parser.add_argument(
        '--text',
        help='Text to add inside the shape'
    )
    
    parser.add_argument(
        '--allow-offslide',
        action='store_true',
        help='Allow positioning outside slide bounds (suppresses warnings)'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default: true)'
    )
    
    parser.add_argument(
        '--version',
        action='version',
        version=f'%(prog)s {__version__} (core: {CORE_VERSION})'
    )
    
    args = parser.parse_args()
    
    try:
        # Process size - use defaults or merge from position
        size = args.size if args.size else {}
        position = args.position
        
        # Allow size in position dict for convenience
        if "width" in position and "width" not in size:
            size["width"] = position.pop("width")
        if "height" in position and "height" not in size:
            size["height"] = position.pop("height")
        
        # Apply defaults if still missing
        if "width" not in size:
            size["width"] = "20%"
        if "height" not in size:
            size["height"] = "20%"
        
        # Execute
        result = add_shape(
            filepath=args.file,
            slide_index=args.slide,
            shape_type=args.shape,
            position=position,
            size=size,
            fill_color=args.fill_color,
            line_color=args.line_color,
            line_width=args.line_width,
            text=args.text,
            allow_offslide=args.allow_offslide
        )
        
        # Output
        if args.json:
            print(json.dumps(result, indent=2))
        else:
            status_icon = "✅" if result["status"] == "success" else "⚠️"
            print(f"{status_icon} Added {result['shape_type']} to slide {result['slide_index']}")
            print(f"   Shape index: {result['shape_index']}")
            if args.fill_color:
                print(f"   Fill: {result['styling']['fill_color']}")
            if args.line_color:
                print(f"   Line: {result['styling']['line_color']} ({args.line_width}pt)")
            if args.text:
                print(f"   Text: {args.text[:50]}{'...' if len(args.text) > 50 else ''}")
            if result.get("warnings"):
                print("\n   Warnings:")
                for warning in result["warnings"]:
                    print(f"   ⚠️  {warning}")
        
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify the file path exists and is accessible."
        }
        if args.json:
            print(json.dumps(error_result, indent=2))
        else:
            print(f"❌ Error: {e}", file=sys.stderr)
        sys.exit(1)
        
    except SlideNotFoundError as e:
        error_result = {
            "status": "error",
            "error": e.message,
            "error_type": "SlideNotFoundError",
            "details": e.details,
            "suggestion": "Use ppt_get_info.py to check available slide indices."
        }
        if args.json:
            print(json.dumps(error_result, indent=2))
        else:
            print(f"❌ Error: {e.message}", file=sys.stderr)
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": e.message,
            "error_type": type(e).__name__,
            "details": e.details
        }
        if args.json:
            print(json.dumps(error_result, indent=2))
        else:
            print(f"❌ Error: {e.message}", file=sys.stderr)
        sys.exit(1)
        
    except json.JSONDecodeError as e:
        error_result = {
            "status": "error",
            "error": f"Invalid JSON in argument: {e}",
            "error_type": "JSONDecodeError",
            "suggestion": "Ensure --position and --size are valid JSON strings."
        }
        if args.json:
            print(json.dumps(error_result, indent=2))
        else:
            print(f"❌ Error: Invalid JSON - {e}", file=sys.stderr)
        sys.exit(1)
        
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
