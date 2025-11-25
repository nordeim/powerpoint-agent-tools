#!/usr/bin/env python3
"""
PowerPoint Format Shape Tool v3.0
Update styling of existing shapes including fill, line, transparency, and text formatting.

Fully aligned with PowerPoint Agent Core v3.0 and System Prompt v3.0.

Author: PowerPoint Agent Team
License: MIT
Version: 3.0.0

Usage:
    uv run tools/ppt_format_shape.py --file presentation.pptx --slide 0 --shape 1 \\
        --fill-color "#FF0000" --transparency 0.3 --json

Exit Codes:
    0: Success
    1: Error occurred

Changelog v3.0.0:
- Aligned with PowerPoint Agent Core v3.0
- Added transparency/opacity support
- Added color presets (primary, secondary, accent, etc.)
- Added shape info before/after formatting
- Added presentation version tracking
- Added color contrast validation
- Added text formatting within shapes
- Improved error handling with ShapeNotFoundError
- Enhanced CLI help with examples
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
    __version__ as CORE_VERSION
)

# ============================================================================
# CONSTANTS
# ============================================================================

__version__ = "3.0.0"

# Color presets (matching ppt_add_shape.py for consistency)
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
    "transparent": None,  # Special case for no fill
}

# Transparency presets for common use cases
TRANSPARENCY_PRESETS = {
    "opaque": 0.0,
    "subtle": 0.15,
    "light": 0.3,
    "medium": 0.5,
    "heavy": 0.7,
    "very_light": 0.85,
    "nearly_invisible": 0.95,
}


# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def resolve_color(color: Optional[str]) -> Optional[str]:
    """
    Resolve color value, handling presets and hex formats.
    
    Args:
        color: Color specification (hex, preset name, or None)
        
    Returns:
        Resolved hex color or None
    """
    if color is None:
        return None
    
    color_lower = color.lower().strip()
    
    # Check presets
    if color_lower in COLOR_PRESETS:
        return COLOR_PRESETS[color_lower]
    
    # Handle "none" or "transparent" to clear fill
    if color_lower in ("none", "transparent", "clear"):
        return None
    
    # Ensure # prefix for hex colors
    if not color.startswith('#'):
        # Check if it's a valid hex without #
        if len(color) == 6:
            try:
                int(color, 16)
                return f"#{color}"
            except ValueError:
                pass
    
    return color


def resolve_transparency(value: Optional[str]) -> Optional[float]:
    """
    Resolve transparency value, handling presets and numeric values.
    
    Args:
        value: Transparency specification (float, preset name, or percentage string)
        
    Returns:
        Transparency as float (0.0 = opaque, 1.0 = invisible) or None
    """
    if value is None:
        return None
    
    # Handle string presets
    if isinstance(value, str):
        value_lower = value.lower().strip()
        
        # Check presets
        if value_lower in TRANSPARENCY_PRESETS:
            return TRANSPARENCY_PRESETS[value_lower]
        
        # Handle percentage string (e.g., "30%")
        if value_lower.endswith('%'):
            try:
                return float(value_lower[:-1]) / 100.0
            except ValueError:
                pass
        
        # Try to parse as float
        try:
            return float(value_lower)
        except ValueError:
            raise ValueError(f"Invalid transparency value: {value}")
    
    # Handle numeric
    return float(value)


def validate_formatting_params(
    fill_color: Optional[str],
    line_color: Optional[str],
    transparency: Optional[float],
    current_shape_info: Optional[Dict[str, Any]] = None
) -> Dict[str, Any]:
    """
    Validate formatting parameters and generate warnings.
    
    Args:
        fill_color: Fill color hex
        line_color: Line color hex
        transparency: Transparency value
        current_shape_info: Current shape information for context
        
    Returns:
        Dict with warnings, recommendations, and validation results
    """
    warnings: List[str] = []
    recommendations: List[str] = []
    validation_results: Dict[str, Any] = {}
    
    # Validate transparency range
    if transparency is not None:
        if not 0.0 <= transparency <= 1.0:
            warnings.append(
                f"Transparency {transparency} is outside valid range (0.0-1.0). "
                f"Will be clamped."
            )
            validation_results["transparency_clamped"] = True
        
        if transparency > 0.9:
            warnings.append(
                f"Transparency {transparency} is very high. Shape may be nearly invisible."
            )
    
    # Validate and check fill color
    if fill_color:
        try:
            fill_rgb = ColorHelper.from_hex(fill_color)
            validation_results["fill_color_valid"] = True
            
            # Check contrast against common backgrounds
            from pptx.dml.color import RGBColor
            
            # White background contrast
            white_bg = RGBColor(255, 255, 255)
            white_contrast = ColorHelper.contrast_ratio(fill_rgb, white_bg)
            validation_results["fill_contrast_vs_white"] = round(white_contrast, 2)
            
            if white_contrast < 1.1:
                warnings.append(
                    f"Fill color {fill_color} has very low contrast against white backgrounds."
                )
            
            # Black background contrast
            black_bg = RGBColor(0, 0, 0)
            black_contrast = ColorHelper.contrast_ratio(fill_rgb, black_bg)
            validation_results["fill_contrast_vs_black"] = round(black_contrast, 2)
            
        except Exception as e:
            validation_results["fill_color_valid"] = False
            validation_results["fill_color_error"] = str(e)
            warnings.append(f"Invalid fill color format: {fill_color}")
    
    # Validate line color
    if line_color:
        try:
            ColorHelper.from_hex(line_color)
            validation_results["line_color_valid"] = True
            
            # Check line/fill contrast if both specified
            if fill_color:
                fill_rgb = ColorHelper.from_hex(fill_color)
                line_rgb = ColorHelper.from_hex(line_color)
                line_fill_contrast = ColorHelper.contrast_ratio(fill_rgb, line_rgb)
                validation_results["line_fill_contrast"] = round(line_fill_contrast, 2)
                
                if line_fill_contrast < 1.5:
                    warnings.append(
                        f"Line color has low contrast against fill color. "
                        f"Border may not be visible."
                    )
        except Exception as e:
            validation_results["line_color_valid"] = False
            validation_results["line_color_error"] = str(e)
            warnings.append(f"Invalid line color format: {line_color}")
    
    return {
        "warnings": warnings,
        "recommendations": recommendations,
        "validation_results": validation_results,
        "has_warnings": len(warnings) > 0
    }


def get_shape_info_summary(agent: PowerPointAgent, slide_index: int, shape_index: int) -> Dict[str, Any]:
    """
    Get summary information about a shape for before/after comparison.
    
    Args:
        agent: PowerPointAgent instance
        slide_index: Slide index
        shape_index: Shape index
        
    Returns:
        Dict with shape summary info
    """
    try:
        slide_info = agent.get_slide_info(slide_index)
        shapes = slide_info.get("shapes", [])
        
        if 0 <= shape_index < len(shapes):
            shape = shapes[shape_index]
            return {
                "index": shape_index,
                "type": shape.get("type", "unknown"),
                "name": shape.get("name", ""),
                "has_text": shape.get("has_text", False),
                "text_preview": shape.get("text", "")[:50] if shape.get("text") else None,
                "position": shape.get("position", {}),
                "size": shape.get("size", {})
            }
    except Exception:
        pass
    
    return {"index": shape_index, "type": "unknown"}


# ============================================================================
# MAIN FUNCTION
# ============================================================================

def format_shape(
    filepath: Path,
    slide_index: int,
    shape_index: int,
    fill_color: Optional[str] = None,
    line_color: Optional[str] = None,
    line_width: Optional[float] = None,
    transparency: Optional[float] = None,
    text_color: Optional[str] = None,
    text_size: Optional[int] = None,
    text_bold: Optional[bool] = None
) -> Dict[str, Any]:
    """
    Format existing shape with comprehensive styling options.
    
    Args:
        filepath: Path to PowerPoint file
        slide_index: Target slide index (0-based)
        shape_index: Target shape index (0-based)
        fill_color: Fill color (hex or preset name)
        line_color: Line/border color (hex or preset name)
        line_width: Line width in points
        transparency: Fill transparency (0.0=opaque to 1.0=invisible)
        text_color: Text color within shape (hex or preset name)
        text_size: Text size in points
        text_bold: Text bold setting
        
    Returns:
        Result dict with formatting details and validation info
        
    Raises:
        FileNotFoundError: If file doesn't exist
        SlideNotFoundError: If slide index is invalid
        ShapeNotFoundError: If shape index is invalid
        ValueError: If no formatting options provided
    """
    # Validate file exists
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    # Check at least one formatting option provided
    formatting_options = [
        fill_color, line_color, line_width, transparency,
        text_color, text_size, text_bold
    ]
    if all(v is None for v in formatting_options):
        raise ValueError(
            "At least one formatting option required. "
            "Use --fill-color, --line-color, --line-width, --transparency, "
            "--text-color, --text-size, or --text-bold."
        )
    
    # Resolve colors
    resolved_fill = resolve_color(fill_color)
    resolved_line = resolve_color(line_color)
    resolved_text_color = resolve_color(text_color)
    resolved_transparency = resolve_transparency(transparency) if transparency is not None else None
    
    # Clamp transparency to valid range
    if resolved_transparency is not None:
        resolved_transparency = max(0.0, min(1.0, resolved_transparency))
    
    # Validate parameters
    validation = validate_formatting_params(
        fill_color=resolved_fill,
        line_color=resolved_line,
        transparency=resolved_transparency
    )
    
    # Open presentation and format shape
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
        
        # Get shape info before formatting
        shape_before = get_shape_info_summary(agent, slide_index, shape_index)
        
        # Get presentation version before change
        version_before = agent.get_presentation_version()
        
        # Format shape using v3.0 core
        format_result = agent.format_shape(
            slide_index=slide_index,
            shape_index=shape_index,
            fill_color=resolved_fill,
            line_color=resolved_line,
            line_width=line_width,
            transparency=resolved_transparency
        )
        
        # Format text within shape if text options provided
        text_formatted = False
        if any(v is not None for v in [text_color, text_size, text_bold]):
            try:
                agent.format_text(
                    slide_index=slide_index,
                    shape_index=shape_index,
                    color=resolved_text_color,
                    font_size=text_size,
                    bold=text_bold
                )
                text_formatted = True
            except Exception as e:
                validation["warnings"].append(
                    f"Could not format text: {e}. Shape may not contain text."
                )
        
        # Save changes
        agent.save()
        
        # Get presentation version after change
        version_after = agent.get_presentation_version()
        
        # Get shape info after formatting
        shape_after = get_shape_info_summary(agent, slide_index, shape_index)
    
    # Build result
    result = {
        "status": "success" if not validation["has_warnings"] else "warning",
        "file": str(filepath),
        "slide_index": slide_index,
        "shape_index": shape_index,
        "shape_info": shape_before,
        "formatting_applied": {
            "fill_color": resolved_fill,
            "line_color": resolved_line,
            "line_width": line_width,
            "transparency": resolved_transparency,
            "text_color": resolved_text_color if text_formatted else None,
            "text_size": text_size if text_formatted else None,
            "text_bold": text_bold if text_formatted else None
        },
        "changes_from_core": format_result.get("changes_applied", []),
        "text_formatted": text_formatted,
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
    
    return result


# ============================================================================
# CLI INTERFACE
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Format existing PowerPoint shape (v3.0)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

FORMATTING OPTIONS:
  --fill-color     Shape fill color (hex or preset)
  --line-color     Border/line color (hex or preset)
  --line-width     Border width in points
  --transparency   Fill transparency (0.0=opaque to 1.0=invisible)
  --text-color     Text color within shape
  --text-size      Text size in points
  --text-bold      Make text bold

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

COLOR PRESETS:
  primary (#0070C0)    secondary (#595959)    accent (#ED7D31)
  success (#70AD47)    warning (#FFC000)      danger (#C00000)
  white (#FFFFFF)      black (#000000)
  light_gray (#D9D9D9) dark_gray (#404040)
  transparent / none   (removes fill)

TRANSPARENCY PRESETS:
  opaque (0.0)         subtle (0.15)          light (0.3)
  medium (0.5)         heavy (0.7)            very_light (0.85)
  Or use decimal: 0.0 to 1.0
  Or use percentage: "30%"

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

EXAMPLES:

  # Change fill color to red
  uv run tools/ppt_format_shape.py \\
    --file presentation.pptx --slide 0 --shape 1 \\
    --fill-color "#FF0000" --json

  # Use color preset
  uv run tools/ppt_format_shape.py \\
    --file presentation.pptx --slide 0 --shape 2 \\
    --fill-color primary --line-color black --line-width 2 --json

  # Create semi-transparent overlay
  uv run tools/ppt_format_shape.py \\
    --file presentation.pptx --slide 1 --shape 0 \\
    --fill-color black --transparency 0.5 --json

  # Use transparency preset
  uv run tools/ppt_format_shape.py \\
    --file presentation.pptx --slide 1 --shape 0 \\
    --fill-color white --transparency subtle --json

  # Format text within shape
  uv run tools/ppt_format_shape.py \\
    --file presentation.pptx --slide 0 --shape 3 \\
    --fill-color primary --text-color white --text-size 24 --text-bold --json

  # Remove fill (make transparent)
  uv run tools/ppt_format_shape.py \\
    --file presentation.pptx --slide 0 --shape 1 \\
    --fill-color transparent --line-color primary --line-width 3 --json

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

FINDING SHAPE INDEX:
  Use ppt_get_slide_info.py to find shape indices:
  
  uv run tools/ppt_get_slide_info.py --file presentation.pptx --slide 0 --json
  
  The output will list all shapes with their indices.

COMMON USE CASES:
  • Highlighting: Change fill to accent color
  • Overlay backgrounds: Set transparency to 0.15-0.3
  • Callout boxes: Set fill + contrasting border
  • Text emphasis: Format text color and size within shapes

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        """
    )
    
    # Required arguments
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
        type=int,
        help='Shape index (0-based). Use ppt_get_slide_info.py to find indices.'
    )
    
    # Formatting options
    parser.add_argument(
        '--fill-color',
        help='Fill color: hex (#FF0000) or preset (primary, danger, etc.)'
    )
    
    parser.add_argument(
        '--line-color',
        help='Line/border color: hex or preset'
    )
    
    parser.add_argument(
        '--line-width',
        type=float,
        help='Line width in points'
    )
    
    parser.add_argument(
        '--transparency',
        help='Fill transparency: 0.0 (opaque) to 1.0 (invisible), or preset (subtle, light, medium)'
    )
    
    # Text formatting options
    parser.add_argument(
        '--text-color',
        help='Text color within shape: hex or preset'
    )
    
    parser.add_argument(
        '--text-size',
        type=int,
        help='Text size in points'
    )
    
    parser.add_argument(
        '--text-bold',
        action='store_true',
        help='Make text bold'
    )
    
    # Output options
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
        result = format_shape(
            filepath=args.file,
            slide_index=args.slide,
            shape_index=args.shape,
            fill_color=args.fill_color,
            line_color=args.line_color,
            line_width=args.line_width,
            transparency=args.transparency,
            text_color=args.text_color,
            text_size=args.text_size,
            text_bold=args.text_bold if args.text_bold else None
        )
        
        if args.json:
            print(json.dumps(result, indent=2))
        else:
            status_icon = "✅" if result["status"] == "success" else "⚠️"
            print(f"{status_icon} Formatted shape {result['shape_index']} on slide {result['slide_index']}")
            print(f"   Changes: {', '.join(result.get('changes_from_core', []))}")
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
        
    except ShapeNotFoundError as e:
        error_result = {
            "status": "error",
            "error": e.message,
            "error_type": "ShapeNotFoundError",
            "details": e.details,
            "suggestion": "Use ppt_get_slide_info.py to check available shape indices."
        }
        if args.json:
            print(json.dumps(error_result, indent=2))
        else:
            print(f"❌ Error: {e.message}", file=sys.stderr)
        sys.exit(1)
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Provide at least one formatting option (--fill-color, --line-color, etc.)"
        }
        if args.json:
            print(json.dumps(error_result, indent=2))
        else:
            print(f"❌ Error: {e}", file=sys.stderr)
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
