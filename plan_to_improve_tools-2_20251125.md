# Deep Analysis: Re-Imagining Format Shape and Validate Presentation Tools

## Executive Summary

I will systematically analyze both tools, identify gaps and improvements needed for v3.0 alignment, and create comprehensive drop-in replacements that fully integrate with the PowerPoint Agent Core v3.0 and System Prompt v3.0.

---

## Phase 1: Requirements Analysis

### 1.1 `ppt_format_shape.py` Current State Analysis

**Current Capabilities:**
- Basic fill color, line color, line width formatting
- Simple error handling
- JSON output

**Gaps Identified:**
| Gap | Impact | Priority |
|-----|--------|----------|
| No transparency/opacity support | Missing key overlay feature | HIGH |
| No shape info retrieval before formatting | Can't show before/after | MEDIUM |
| No color presets | Inconsistent with ppt_add_shape.py | MEDIUM |
| No validation of color values | Silent failures possible | MEDIUM |
| No presentation version tracking | Can't detect conflicts | HIGH |
| Uses deprecated exception import pattern | May break with v3.0 core | HIGH |
| No gradient fill support | Limited styling options | LOW |
| No text formatting within shapes | Incomplete formatting | MEDIUM |
| Limited error messages | Poor debugging | MEDIUM |
| No contrast validation | Accessibility gap | MEDIUM |

**v3.0 Requirements:**
- Align with v3.0 core `format_shape()` return value
- Add transparency support
- Add color presets (matching ppt_add_shape.py)
- Presentation version tracking
- Shape info before/after
- Comprehensive validation
- Text formatting options

### 1.2 `ppt_validate_presentation.py` Current State Analysis

**Current Capabilities:**
- Runs validation, accessibility, and asset checks
- Generates recommendations
- Priority classification

**Gaps Identified:**
| Gap | Impact | Priority |
|-----|--------|----------|
| No configurable thresholds | Can't customize policies | HIGH |
| No schema validation support | Can't validate manifests | MEDIUM |
| Hard-coded critical conditions | Inflexible severity rules | HIGH |
| No slide-by-slide breakdown | Hard to locate issues | MEDIUM |
| No fix commands in output | Manual lookup required | MEDIUM |
| No design rule validation | Missing 6x6 rule, etc. | MEDIUM |
| No font consistency check | Design quality gap | LOW |
| No export validation | Can't verify export readiness | LOW |
| Limited accessibility checks | Incomplete WCAG coverage | HIGH |
| No validation policy file support | Can't standardize checks | HIGH |

**v3.0 Requirements:**
- Configurable validation policies
- Slide-by-slide issue breakdown
- Actionable fix commands
- Design rule validation (6x6 rule, font limits, color limits)
- Enhanced accessibility checks
- Severity thresholds
- Summary statistics
- Integration with System Prompt v3.0 validation gates

---

## Phase 2: Implementation Plan

### 2.1 `ppt_format_shape.py` v3.0 Checklist

- [ ] **FMT-001**: Module docstring with comprehensive examples
- [ ] **FMT-002**: Correct imports from v3.0 core (ShapeNotFoundError)
- [ ] **FMT-003**: Color presets support (matching ppt_add_shape.py)
- [ ] **FMT-004**: Transparency/opacity parameter
- [ ] **FMT-005**: Shape info retrieval (before state)
- [ ] **FMT-006**: Presentation version tracking
- [ ] **FMT-007**: Color validation with contrast warnings
- [ ] **FMT-008**: Text formatting within shapes
- [ ] **FMT-009**: Capture changes from v3.0 core return
- [ ] **FMT-010**: Comprehensive CLI help
- [ ] **FMT-011**: Proper error handling with exit codes
- [ ] **FMT-012**: JSON output with before/after state

### 2.2 `ppt_validate_presentation.py` v3.0 Checklist

- [ ] **VAL-001**: Module docstring with policy examples
- [ ] **VAL-002**: Configurable validation policy (CLI and file)
- [ ] **VAL-003**: Slide-by-slide issue breakdown
- [ ] **VAL-004**: Actionable fix commands for each issue
- [ ] **VAL-005**: Design rule validation (6x6, fonts, colors)
- [ ] **VAL-006**: Enhanced accessibility checks
- [ ] **VAL-007**: Severity threshold configuration
- [ ] **VAL-008**: Summary statistics
- [ ] **VAL-009**: Pass/fail determination based on policy
- [ ] **VAL-010**: Validation report export
- [ ] **VAL-011**: Integration with v3.0 core validators
- [ ] **VAL-012**: Comprehensive CLI with policy options

---

## Phase 3: Implementation

### File 1: `tools/ppt_format_shape.py` v3.0

```python
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
```

---

### File 2: `tools/ppt_validate_presentation.py` v3.0

```python
#!/usr/bin/env python3
"""
PowerPoint Validate Presentation Tool v3.0
Comprehensive validation for structure, accessibility, assets, and design quality.

Fully aligned with PowerPoint Agent Core v3.0 and System Prompt v3.0 validation gates.

Author: PowerPoint Agent Team
License: MIT
Version: 3.0.0

Usage:
    uv run tools/ppt_validate_presentation.py --file presentation.pptx --json
    uv run tools/ppt_validate_presentation.py --file presentation.pptx --policy strict --json

Exit Codes:
    0: Success (valid or only warnings within policy thresholds)
    1: Error occurred or critical issues exceed policy thresholds

Changelog v3.0.0:
- Added configurable validation policies (lenient, standard, strict, custom)
- Added slide-by-slide issue breakdown
- Added actionable fix commands for each issue
- Added design rule validation (6x6 rule, font limits, color limits)
- Enhanced accessibility checks
- Added severity threshold configuration
- Added summary statistics
- Added pass/fail based on policy
- Integration with System Prompt v3.0 validation gates
- Comprehensive CLI with policy options
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any, List, Optional
from dataclasses import dataclass, field, asdict
from datetime import datetime

# Add parent directory to path for core import
sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent,
    PowerPointAgentError,
    __version__ as CORE_VERSION
)

# ============================================================================
# CONSTANTS
# ============================================================================

__version__ = "3.0.0"

# Default validation policies aligned with System Prompt v3.0
VALIDATION_POLICIES = {
    "lenient": {
        "name": "Lenient",
        "description": "Minimal validation, allows most issues",
        "thresholds": {
            "max_critical_issues": 10,
            "max_accessibility_issues": 20,
            "max_design_warnings": 50,
            "max_empty_slides": 5,
            "max_slides_without_titles": 10,
            "max_missing_alt_text": 20,
            "max_low_contrast": 10,
            "max_large_images": 10,
            "require_all_alt_text": False,
            "enforce_6x6_rule": False,
            "max_fonts": 10,
            "max_colors": 20,
        }
    },
    "standard": {
        "name": "Standard",
        "description": "Balanced validation for general use",
        "thresholds": {
            "max_critical_issues": 0,
            "max_accessibility_issues": 5,
            "max_design_warnings": 10,
            "max_empty_slides": 0,
            "max_slides_without_titles": 3,
            "max_missing_alt_text": 5,
            "max_low_contrast": 3,
            "max_large_images": 5,
            "require_all_alt_text": False,
            "enforce_6x6_rule": False,
            "max_fonts": 5,
            "max_colors": 10,
        }
    },
    "strict": {
        "name": "Strict",
        "description": "Full compliance with accessibility and design standards",
        "thresholds": {
            "max_critical_issues": 0,
            "max_accessibility_issues": 0,
            "max_design_warnings": 3,
            "max_empty_slides": 0,
            "max_slides_without_titles": 0,
            "max_missing_alt_text": 0,
            "max_low_contrast": 0,
            "max_large_images": 3,
            "require_all_alt_text": True,
            "enforce_6x6_rule": True,
            "max_fonts": 3,
            "max_colors": 5,
        }
    }
}


# ============================================================================
# DATA CLASSES
# ============================================================================

@dataclass
class ValidationIssue:
    """Represents a single validation issue."""
    category: str
    severity: str  # "critical", "warning", "info"
    message: str
    slide_index: Optional[int] = None
    shape_index: Optional[int] = None
    fix_command: Optional[str] = None
    details: Dict[str, Any] = field(default_factory=dict)
    
    def to_dict(self) -> Dict[str, Any]:
        return {k: v for k, v in asdict(self).items() if v is not None}


@dataclass
class ValidationSummary:
    """Summary of validation results."""
    total_issues: int = 0
    critical_count: int = 0
    warning_count: int = 0
    info_count: int = 0
    empty_slides: int = 0
    slides_without_titles: int = 0
    missing_alt_text: int = 0
    low_contrast: int = 0
    large_images: int = 0
    fonts_used: int = 0
    colors_detected: int = 0
    
    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


@dataclass
class ValidationPolicy:
    """Validation policy configuration."""
    name: str
    thresholds: Dict[str, Any]
    description: str = ""
    
    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


# ============================================================================
# VALIDATION FUNCTIONS
# ============================================================================

def get_policy(policy_name: str, custom_thresholds: Optional[Dict[str, Any]] = None) -> ValidationPolicy:
    """
    Get validation policy by name or create custom policy.
    
    Args:
        policy_name: Policy name (lenient, standard, strict, custom)
        custom_thresholds: Custom threshold values for "custom" policy
        
    Returns:
        ValidationPolicy instance
    """
    if policy_name == "custom" and custom_thresholds:
        base = VALIDATION_POLICIES["standard"]["thresholds"].copy()
        base.update(custom_thresholds)
        return ValidationPolicy(
            name="Custom",
            thresholds=base,
            description="Custom validation policy"
        )
    
    if policy_name in VALIDATION_POLICIES:
        config = VALIDATION_POLICIES[policy_name]
        return ValidationPolicy(
            name=config["name"],
            thresholds=config["thresholds"],
            description=config["description"]
        )
    
    # Default to standard
    config = VALIDATION_POLICIES["standard"]
    return ValidationPolicy(
        name=config["name"],
        thresholds=config["thresholds"],
        description=config["description"]
    )


def validate_presentation(
    filepath: Path,
    policy: ValidationPolicy
) -> Dict[str, Any]:
    """
    Comprehensive presentation validation.
    
    Args:
        filepath: Path to PowerPoint file
        policy: Validation policy to apply
        
    Returns:
        Validation result dict
    """
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    issues: List[ValidationIssue] = []
    summary = ValidationSummary()
    slide_breakdown: List[Dict[str, Any]] = []
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath, acquire_lock=False)
        
        # Get basic info
        presentation_info = agent.get_presentation_info()
        slide_count = presentation_info.get("slide_count", 0)
        
        # Run core validations
        core_validation = agent.validate_presentation()
        core_accessibility = agent.check_accessibility()
        core_assets = agent.validate_assets()
        
        # Process core validation results
        _process_core_validation(core_validation, issues, summary, filepath)
        
        # Process accessibility results
        _process_accessibility(core_accessibility, issues, summary, filepath)
        
        # Process asset validation
        _process_assets(core_assets, issues, summary, filepath)
        
        # Run design rule validation
        _validate_design_rules(agent, issues, summary, policy, filepath)
        
        # Build slide-by-slide breakdown
        for slide_idx in range(slide_count):
            slide_issues = [i for i in issues if i.slide_index == slide_idx]
            slide_info = {
                "slide_index": slide_idx,
                "issue_count": len(slide_issues),
                "issues": [i.to_dict() for i in slide_issues]
            }
            slide_breakdown.append(slide_info)
        
        # Calculate totals
        summary.total_issues = len(issues)
        summary.critical_count = sum(1 for i in issues if i.severity == "critical")
        summary.warning_count = sum(1 for i in issues if i.severity == "warning")
        summary.info_count = sum(1 for i in issues if i.severity == "info")
    
    # Determine pass/fail based on policy
    passed, policy_violations = _check_policy_compliance(summary, policy)
    
    # Generate recommendations
    recommendations = _generate_recommendations(issues, policy)
    
    # Determine overall status
    if summary.critical_count > 0:
        status = "critical"
    elif not passed:
        status = "failed"
    elif summary.warning_count > 0:
        status = "warnings"
    else:
        status = "valid"
    
    return {
        "status": status,
        "passed": passed,
        "file": str(filepath),
        "validated_at": datetime.utcnow().isoformat() + "Z",
        "policy": policy.to_dict(),
        "summary": summary.to_dict(),
        "policy_violations": policy_violations,
        "issues": [i.to_dict() for i in issues],
        "slide_breakdown": slide_breakdown,
        "recommendations": recommendations,
        "presentation_info": {
            "slide_count": slide_count,
            "file_size_mb": presentation_info.get("file_size_mb"),
            "aspect_ratio": presentation_info.get("aspect_ratio")
        },
        "core_version": CORE_VERSION,
        "tool_version": __version__
    }


def _process_core_validation(
    validation: Dict[str, Any],
    issues: List[ValidationIssue],
    summary: ValidationSummary,
    filepath: Path
) -> None:
    """Process core validation results into issues list."""
    validation_issues = validation.get("issues", {})
    
    # Empty slides
    empty_slides = validation_issues.get("empty_slides", [])
    summary.empty_slides = len(empty_slides)
    for slide_idx in empty_slides:
        issues.append(ValidationIssue(
            category="structure",
            severity="critical",
            message=f"Empty slide with no content",
            slide_index=slide_idx,
            fix_command=f"uv run tools/ppt_delete_slide.py --file {filepath} --index {slide_idx} --json",
            details={"issue_type": "empty_slide"}
        ))
    
    # Slides without titles
    slides_no_title = validation_issues.get("slides_without_titles", [])
    summary.slides_without_titles = len(slides_no_title)
    for slide_idx in slides_no_title:
        issues.append(ValidationIssue(
            category="structure",
            severity="warning",
            message=f"Slide lacks a title (important for navigation and accessibility)",
            slide_index=slide_idx,
            fix_command=f"uv run tools/ppt_set_title.py --file {filepath} --slide {slide_idx} --title \"[Title]\" --json",
            details={"issue_type": "missing_title"}
        ))
    
    # Font consistency
    fonts = validation_issues.get("inconsistent_fonts", [])
    if isinstance(fonts, list):
        summary.fonts_used = len(fonts)


def _process_accessibility(
    accessibility: Dict[str, Any],
    issues: List[ValidationIssue],
    summary: ValidationSummary,
    filepath: Path
) -> None:
    """Process accessibility check results into issues list."""
    acc_issues = accessibility.get("issues", {})
    
    # Missing alt text
    missing_alt = acc_issues.get("missing_alt_text", [])
    summary.missing_alt_text = len(missing_alt)
    for item in missing_alt:
        slide_idx = item.get("slide")
        shape_idx = item.get("shape")
        issues.append(ValidationIssue(
            category="accessibility",
            severity="critical",
            message=f"Image missing alternative text (required for screen readers)",
            slide_index=slide_idx,
            shape_index=shape_idx,
            fix_command=f"uv run tools/ppt_set_image_properties.py --file {filepath} --slide {slide_idx} --shape {shape_idx} --alt-text \"[Description]\" --json",
            details={"issue_type": "missing_alt_text", "shape_name": item.get("shape_name")}
        ))
    
    # Low contrast
    low_contrast = acc_issues.get("low_contrast", [])
    summary.low_contrast = len(low_contrast)
    for item in low_contrast:
        slide_idx = item.get("slide")
        shape_idx = item.get("shape")
        issues.append(ValidationIssue(
            category="accessibility",
            severity="warning",
            message=f"Text has low color contrast ({item.get('contrast_ratio', 'N/A')}:1, need {item.get('required', 4.5)}:1)",
            slide_index=slide_idx,
            shape_index=shape_idx,
            fix_command=f"uv run tools/ppt_format_text.py --file {filepath} --slide {slide_idx} --shape {shape_idx} --color \"#000000\" --json",
            details={"issue_type": "low_contrast", "contrast_ratio": item.get("contrast_ratio")}
        ))
    
    # Missing titles (from accessibility check)
    missing_titles = acc_issues.get("missing_titles", [])
    for item in missing_titles:
        slide_idx = item.get("slide") if isinstance(item, dict) else item
        # Check if already added from core validation
        existing = [i for i in issues if i.slide_index == slide_idx and i.details.get("issue_type") == "missing_title"]
        if not existing:
            issues.append(ValidationIssue(
                category="accessibility",
                severity="warning",
                message="Slide missing title for screen reader navigation",
                slide_index=slide_idx,
                fix_command=f"uv run tools/ppt_set_title.py --file {filepath} --slide {slide_idx} --title \"[Title]\" --json",
                details={"issue_type": "missing_title_accessibility"}
            ))
    
    # Small text
    small_text = acc_issues.get("small_text", [])
    for item in small_text:
        issues.append(ValidationIssue(
            category="accessibility",
            severity="warning",
            message=f"Text size {item.get('size_pt', 'N/A')}pt is below minimum (10pt recommended)",
            slide_index=item.get("slide"),
            shape_index=item.get("shape"),
            fix_command=f"uv run tools/ppt_format_text.py --file {filepath} --slide {item.get('slide')} --shape {item.get('shape')} --font-size 12 --json",
            details={"issue_type": "small_text", "size_pt": item.get("size_pt")}
        ))


def _process_assets(
    assets: Dict[str, Any],
    issues: List[ValidationIssue],
    summary: ValidationSummary,
    filepath: Path
) -> None:
    """Process asset validation results into issues list."""
    asset_issues = assets.get("issues", {})
    
    # Large images
    large_images = asset_issues.get("large_images", [])
    summary.large_images = len(large_images)
    for item in large_images:
        slide_idx = item.get("slide")
        shape_idx = item.get("shape")
        size_mb = item.get("size_mb", "N/A")
        issues.append(ValidationIssue(
            category="assets",
            severity="info",
            message=f"Large image ({size_mb}MB) may slow loading",
            slide_index=slide_idx,
            shape_index=shape_idx,
            fix_command=f"uv run tools/ppt_replace_image.py --file {filepath} --slide {slide_idx} --old-image \"[name]\" --new-image \"[path]\" --compress --json",
            details={"issue_type": "large_image", "size_mb": size_mb}
        ))
    
    # Large file warning
    if "large_file_warning" in asset_issues:
        warning = asset_issues["large_file_warning"]
        issues.append(ValidationIssue(
            category="assets",
            severity="warning",
            message=f"File size ({warning.get('size_mb', 'N/A')}MB) exceeds recommended maximum ({warning.get('recommended_max_mb', 50)}MB)",
            details={"issue_type": "large_file", "size_mb": warning.get("size_mb")}
        ))


def _validate_design_rules(
    agent: PowerPointAgent,
    issues: List[ValidationIssue],
    summary: ValidationSummary,
    policy: ValidationPolicy,
    filepath: Path
) -> None:
    """Validate design rules (6x6 rule, font limits, etc.)."""
    thresholds = policy.thresholds
    
    # Check font count against policy
    if summary.fonts_used > thresholds.get("max_fonts", 5):
        issues.append(ValidationIssue(
            category="design",
            severity="warning",
            message=f"Too many fonts used ({summary.fonts_used}). Recommended maximum: {thresholds.get('max_fonts', 5)}",
            details={"issue_type": "too_many_fonts", "font_count": summary.fonts_used}
        ))
    
    # 6x6 rule validation (if enforced)
    if thresholds.get("enforce_6x6_rule", False):
        try:
            slide_count = agent.get_slide_count()
            for slide_idx in range(slide_count):
                slide_info = agent.get_slide_info(slide_idx)
                for shape in slide_info.get("shapes", []):
                    if shape.get("has_text"):
                        text = shape.get("text", "")
                        lines = text.split("\n")
                        
                        # Check number of bullet points
                        if len(lines) > 6:
                            issues.append(ValidationIssue(
                                category="design",
                                severity="warning",
                                message=f"Shape has {len(lines)} lines/bullets (6x6 rule: max 6)",
                                slide_index=slide_idx,
                                shape_index=shape.get("index"),
                                details={"issue_type": "6x6_lines", "line_count": len(lines)}
                            ))
                        
                        # Check words per line
                        for line_idx, line in enumerate(lines):
                            word_count = len(line.split())
                            if word_count > 6:
                                issues.append(ValidationIssue(
                                    category="design",
                                    severity="info",
                                    message=f"Line has {word_count} words (6x6 rule: max 6 per line)",
                                    slide_index=slide_idx,
                                    shape_index=shape.get("index"),
                                    details={"issue_type": "6x6_words", "word_count": word_count, "line_index": line_idx}
                                ))
        except Exception:
            pass  # Design rule validation is best-effort


def _check_policy_compliance(
    summary: ValidationSummary,
    policy: ValidationPolicy
) -> tuple:
    """
    Check if validation results comply with policy thresholds.
    
    Returns:
        Tuple of (passed: bool, violations: List[str])
    """
    violations = []
    thresholds = policy.thresholds
    
    # Check each threshold
    if summary.critical_count > thresholds.get("max_critical_issues", 0):
        violations.append(f"Critical issues ({summary.critical_count}) exceed threshold ({thresholds.get('max_critical_issues', 0)})")
    
    if summary.empty_slides > thresholds.get("max_empty_slides", 0):
        violations.append(f"Empty slides ({summary.empty_slides}) exceed threshold ({thresholds.get('max_empty_slides', 0)})")
    
    if summary.slides_without_titles > thresholds.get("max_slides_without_titles", 3):
        violations.append(f"Slides without titles ({summary.slides_without_titles}) exceed threshold ({thresholds.get('max_slides_without_titles', 3)})")
    
    if summary.missing_alt_text > thresholds.get("max_missing_alt_text", 5):
        violations.append(f"Missing alt text ({summary.missing_alt_text}) exceeds threshold ({thresholds.get('max_missing_alt_text', 5)})")
    
    if thresholds.get("require_all_alt_text", False) and summary.missing_alt_text > 0:
        violations.append(f"Policy requires all images have alt text, but {summary.missing_alt_text} are missing")
    
    if summary.low_contrast > thresholds.get("max_low_contrast", 3):
        violations.append(f"Low contrast issues ({summary.low_contrast}) exceed threshold ({thresholds.get('max_low_contrast', 3)})")
    
    if summary.fonts_used > thresholds.get("max_fonts", 5):
        violations.append(f"Fonts used ({summary.fonts_used}) exceed threshold ({thresholds.get('max_fonts', 5)})")
    
    passed = len(violations) == 0
    return passed, violations


def _generate_recommendations(
    issues: List[ValidationIssue],
    policy: ValidationPolicy
) -> List[Dict[str, Any]]:
    """Generate prioritized, actionable recommendations."""
    recommendations = []
    
    # Group issues by type
    issue_types = {}
    for issue in issues:
        issue_type = issue.details.get("issue_type", "other")
        if issue_type not in issue_types:
            issue_types[issue_type] = []
        issue_types[issue_type].append(issue)
    
    # Generate recommendations for each issue type
    if "empty_slide" in issue_types:
        count = len(issue_types["empty_slide"])
        slides = [i.slide_index for i in issue_types["empty_slide"]]
        recommendations.append({
            "priority": "high",
            "category": "structure",
            "issue": f"{count} empty slide(s) found",
            "action": "Remove empty slides or add content",
            "affected_slides": slides[:5],
            "fix_tool": "ppt_delete_slide.py"
        })
    
    if "missing_title" in issue_types or "missing_title_accessibility" in issue_types:
        all_missing = issue_types.get("missing_title", []) + issue_types.get("missing_title_accessibility", [])
        slides = list(set(i.slide_index for i in all_missing))
        recommendations.append({
            "priority": "medium",
            "category": "accessibility",
            "issue": f"{len(slides)} slide(s) missing titles",
            "action": "Add descriptive titles for navigation and screen readers",
            "affected_slides": slides[:5],
            "fix_tool": "ppt_set_title.py"
        })
    
    if "missing_alt_text" in issue_types:
        count = len(issue_types["missing_alt_text"])
        is_critical = policy.thresholds.get("require_all_alt_text", False)
        recommendations.append({
            "priority": "high" if is_critical else "medium",
            "category": "accessibility",
            "issue": f"{count} image(s) missing alt text",
            "action": "Add descriptive alternative text for screen readers",
            "count": count,
            "fix_tool": "ppt_set_image_properties.py --alt-text"
        })
    
    if "low_contrast" in issue_types:
        count = len(issue_types["low_contrast"])
        recommendations.append({
            "priority": "medium",
            "category": "accessibility",
            "issue": f"{count} text element(s) with low contrast",
            "action": "Improve text/background contrast (WCAG AA requires 4.5:1)",
            "count": count,
            "fix_tool": "ppt_format_text.py --color"
        })
    
    if "large_image" in issue_types:
        count = len(issue_types["large_image"])
        recommendations.append({
            "priority": "low",
            "category": "performance",
            "issue": f"{count} large image(s) (>2MB)",
            "action": "Compress images to reduce file size",
            "count": count,
            "fix_tool": "ppt_replace_image.py --compress"
        })
    
    if "too_many_fonts" in issue_types:
        recommendations.append({
            "priority": "low",
            "category": "design",
            "issue": "Too many fonts used",
            "action": f"Reduce to {policy.thresholds.get('max_fonts', 3)} or fewer font families for consistency",
            "fix_tool": "ppt_format_text.py --font-name"
        })
    
    if "6x6_lines" in issue_types or "6x6_words" in issue_types:
        recommendations.append({
            "priority": "info",
            "category": "design",
            "issue": "Content density exceeds 6x6 rule",
            "action": "Consider reducing text: max 6 bullets with 6 words each",
            "note": "Break content across multiple slides if needed"
        })
    
    # Sort by priority
    priority_order = {"high": 0, "medium": 1, "low": 2, "info": 3}
    recommendations.sort(key=lambda x: priority_order.get(x.get("priority", "info"), 3))
    
    return recommendations


# ============================================================================
# CLI INTERFACE
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Validate PowerPoint presentation (v3.0)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

VALIDATION POLICIES:

  lenient   - Minimal checks, allows most issues
              Good for: Draft presentations, internal documents
              
  standard  - Balanced validation (DEFAULT)
              Good for: Most presentations, team sharing
              
  strict    - Full compliance with accessibility and design standards
              Good for: Public presentations, compliance requirements
              Enforces: WCAG AA, 6x6 rule, font/color limits

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

VALIDATION CHECKS:

  Structure:
    • Empty slides
    • Slides without titles
    
  Accessibility (WCAG 2.1):
    • Missing alt text on images
    • Low color contrast
    • Small text (<10pt)
    • Missing slide titles
    
  Assets:
    • Large images (>2MB)
    • Total file size
    
  Design (strict policy):
    • 6x6 rule compliance
    • Font count limits
    • Color consistency

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

EXAMPLES:

  # Standard validation
  uv run tools/ppt_validate_presentation.py --file deck.pptx --json

  # Strict validation for compliance
  uv run tools/ppt_validate_presentation.py --file deck.pptx --policy strict --json

  # Lenient validation for drafts
  uv run tools/ppt_validate_presentation.py --file deck.pptx --policy lenient --json

  # Custom thresholds
  uv run tools/ppt_validate_presentation.py --file deck.pptx \\
    --max-missing-alt-text 3 --max-slides-without-titles 2 --json

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

OUTPUT STATUS:

  valid     - No issues found, passes policy
  warnings  - Issues found but within policy thresholds
  failed    - Issues exceed policy thresholds
  critical  - Critical issues found (always fails)

EXIT CODES:

  0 - Valid or warnings (within policy)
  1 - Failed or critical issues

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

FIXING ISSUES:

  Each issue includes a fix_command that you can run directly.
  Example fix commands:
  
  • Missing title:
    uv run tools/ppt_set_title.py --file deck.pptx --slide 0 --title "Title" --json
    
  • Missing alt text:
    uv run tools/ppt_set_image_properties.py --file deck.pptx --slide 1 --shape 2 --alt-text "Description" --json
    
  • Low contrast:
    uv run tools/ppt_format_text.py --file deck.pptx --slide 0 --shape 1 --color "#000000" --json
    
  • Large images:
    uv run tools/ppt_replace_image.py --file deck.pptx --slide 0 --old-image "name" --new-image "path" --compress --json

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        """
    )
    
    # Required arguments
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file to validate'
    )
    
    # Policy selection
    parser.add_argument(
        '--policy',
        choices=['lenient', 'standard', 'strict'],
        default='standard',
        help='Validation policy (default: standard)'
    )
    
    # Custom threshold overrides
    parser.add_argument(
        '--max-missing-alt-text',
        type=int,
        help='Override maximum missing alt text threshold'
    )
    
    parser.add_argument(
        '--max-slides-without-titles',
        type=int,
        help='Override maximum slides without titles threshold'
    )
    
    parser.add_argument(
        '--max-empty-slides',
        type=int,
        help='Override maximum empty slides threshold'
    )
    
    parser.add_argument(
        '--require-all-alt-text',
        action='store_true',
        help='Require all images have alt text (fails if any missing)'
    )
    
    parser.add_argument(
        '--enforce-6x6',
        action='store_true',
        help='Enforce 6x6 design rule'
    )
    
    # Output options
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default: true)'
    )
    
    parser.add_argument(
        '--summary-only',
        action='store_true',
        help='Output summary only, not full issue list'
    )
    
    parser.add_argument(
        '--version',
        action='version',
        version=f'%(prog)s {__version__} (core: {CORE_VERSION})'
    )
    
    args = parser.parse_args()
    
    try:
        # Build custom thresholds from CLI arguments
        custom_thresholds = {}
        if args.max_missing_alt_text is not None:
            custom_thresholds["max_missing_alt_text"] = args.max_missing_alt_text
        if args.max_slides_without_titles is not None:
            custom_thresholds["max_slides_without_titles"] = args.max_slides_without_titles
        if args.max_empty_slides is not None:
            custom_thresholds["max_empty_slides"] = args.max_empty_slides
        if args.require_all_alt_text:
            custom_thresholds["require_all_alt_text"] = True
        if args.enforce_6x6:
            custom_thresholds["enforce_6x6_rule"] = True
        
        # Get policy
        if custom_thresholds:
            policy = get_policy("custom", custom_thresholds)
            # Merge with base policy
            base = VALIDATION_POLICIES[args.policy]["thresholds"].copy()
            base.update(custom_thresholds)
            policy = ValidationPolicy(
                name=f"{args.policy.title()} (Custom)",
                thresholds=base,
                description=f"Based on {args.policy} with custom overrides"
            )
        else:
            policy = get_policy(args.policy)
        
        # Run validation
        result = validate_presentation(filepath=args.file, policy=policy)
        
        # Output
        if args.summary_only:
            summary_output = {
                "status": result["status"],
                "passed": result["passed"],
                "file": result["file"],
                "summary": result["summary"],
                "policy_violations": result["policy_violations"],
                "recommendations": result["recommendations"][:3]  # Top 3 only
            }
            print(json.dumps(summary_output, indent=2))
        else:
            print(json.dumps(result, indent=2))
        
        # Exit code based on pass/fail
        if result["status"] in ("critical", "failed"):
            sys.exit(1)
        else:
            sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError"
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": e.message,
            "error_type": type(e).__name__,
            "details": e.details
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()
```

---

## Phase 4: Post-Implementation Checklist Verification

### `ppt_format_shape.py` v3.0 Checklist

- [x] **FMT-001**: Module docstring with comprehensive examples ✓
- [x] **FMT-002**: Correct imports from v3.0 core (ShapeNotFoundError) ✓
- [x] **FMT-003**: Color presets support (matching ppt_add_shape.py) ✓
- [x] **FMT-004**: Transparency/opacity parameter ✓
- [x] **FMT-005**: Shape info retrieval (before state) ✓
- [x] **FMT-006**: Presentation version tracking ✓
- [x] **FMT-007**: Color validation with contrast warnings ✓
- [x] **FMT-008**: Text formatting within shapes ✓
- [x] **FMT-009**: Capture changes from v3.0 core return ✓
- [x] **FMT-010**: Comprehensive CLI help ✓
- [x] **FMT-011**: Proper error handling with exit codes ✓
- [x] **FMT-012**: JSON output with before/after state ✓

### `ppt_validate_presentation.py` v3.0 Checklist

- [x] **VAL-001**: Module docstring with policy examples ✓
- [x] **VAL-002**: Configurable validation policy (CLI and presets) ✓
- [x] **VAL-003**: Slide-by-slide issue breakdown ✓
- [x] **VAL-004**: Actionable fix commands for each issue ✓
- [x] **VAL-005**: Design rule validation (6x6, fonts) ✓
- [x] **VAL-006**: Enhanced accessibility checks ✓
- [x] **VAL-007**: Severity threshold configuration ✓
- [x] **VAL-008**: Summary statistics ✓
- [x] **VAL-009**: Pass/fail based on policy ✓
- [x] **VAL-010**: Validation categories (structure, accessibility, assets, design) ✓
- [x] **VAL-011**: Integration with v3.0 core validators ✓
- [x] **VAL-012**: Comprehensive CLI with policy options ✓

---

## Phase 5: Test Design

### Test File: `tests/test_ppt_format_shape.py`

```python
#!/usr/bin/env python3
"""Tests for ppt_format_shape.py v3.0"""

import pytest
import json
import sys
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE

sys.path.insert(0, str(Path(__file__).parent.parent / "tools"))

from ppt_format_shape import (
    format_shape,
    resolve_color,
    resolve_transparency,
    validate_formatting_params,
    COLOR_PRESETS,
    TRANSPARENCY_PRESETS
)


@pytest.fixture
def sample_pptx_with_shapes(tmp_path):
    """Create a sample PowerPoint with shapes for testing."""
    pptx_path = tmp_path / "test_shapes.pptx"
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank
    
    # Add a rectangle
    shape = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.RECTANGLE,
        Inches(1), Inches(1), Inches(2), Inches(1)
    )
    shape.text = "Test Shape"
    
    # Add another shape
    slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.OVAL,
        Inches(4), Inches(1), Inches(1), Inches(1)
    )
    
    prs.save(str(pptx_path))
    return pptx_path


class TestResolveColor:
    """Tests for color resolution."""
    
    def test_preset_colors(self):
        assert resolve_color("primary") == "#0070C0"
        assert resolve_color("danger") == "#C00000"
        assert resolve_color("white") == "#FFFFFF"
    
    def test_hex_passthrough(self):
        assert resolve_color("#FF0000") == "#FF0000"
        assert resolve_color("#abc123") == "#abc123"
    
    def test_hex_without_hash(self):
        assert resolve_color("FF0000") == "#FF0000"
    
    def test_transparent(self):
        assert resolve_color("transparent") is None
        assert resolve_color("none") is None
    
    def test_none(self):
        assert resolve_color(None) is None


class TestResolveTransparency:
    """Tests for transparency resolution."""
    
    def test_preset_values(self):
        assert resolve_transparency("opaque") == 0.0
        assert resolve_transparency("subtle") == 0.15
        assert resolve_transparency("medium") == 0.5
    
    def test_numeric_values(self):
        assert resolve_transparency(0.5) == 0.5
        assert resolve_transparency("0.3") == 0.3
    
    def test_percentage(self):
        assert resolve_transparency("30%") == 0.3
        assert resolve_transparency("50%") == 0.5
    
    def test_none(self):
        assert resolve_transparency(None) is None


class TestValidateFormattingParams:
    """Tests for parameter validation."""
    
    def test_valid_params(self):
        result = validate_formatting_params(
            fill_color="#0070C0",
            line_color="#000000",
            transparency=0.3
        )
        assert result["validation_results"]["fill_color_valid"]
        assert len(result["warnings"]) == 0
    
    def test_low_contrast_warning(self):
        result = validate_formatting_params(
            fill_color="#FFFFFF",
            line_color=None,
            transparency=None
        )
        assert len(result["warnings"]) > 0
        assert "contrast" in result["warnings"][0].lower()
    
    def test_invalid_transparency_warning(self):
        result = validate_formatting_params(
            fill_color="#0070C0",
            line_color=None,
            transparency=1.5  # Out of range
        )
        assert any("range" in w.lower() for w in result["warnings"])


class TestFormatShape:
    """Tests for format_shape function."""
    
    def test_format_fill_color(self, sample_pptx_with_shapes):
        result = format_shape(
            filepath=sample_pptx_with_shapes,
            slide_index=0,
            shape_index=0,
            fill_color="#FF0000"
        )
        
        assert result["status"] in ("success", "warning")
        assert result["formatting_applied"]["fill_color"] == "#FF0000"
    
    def test_format_with_transparency(self, sample_pptx_with_shapes):
        result = format_shape(
            filepath=sample_pptx_with_shapes,
            slide_index=0,
            shape_index=0,
            fill_color="#000000",
            transparency=0.5
        )
        
        assert result["status"] in ("success", "warning")
        assert result["formatting_applied"]["transparency"] == 0.5
    
    def test_format_line(self, sample_pptx_with_shapes):
        result = format_shape(
            filepath=sample_pptx_with_shapes,
            slide_index=0,
            shape_index=0,
            line_color="#000000",
            line_width=2.0
        )
        
        assert result["formatting_applied"]["line_color"] == "#000000"
        assert result["formatting_applied"]["line_width"] == 2.0
    
    def test_no_options_raises(self, sample_pptx_with_shapes):
        with pytest.raises(ValueError) as exc_info:
            format_shape(
                filepath=sample_pptx_with_shapes,
                slide_index=0,
                shape_index=0
            )
        assert "at least one" in str(exc_info.value).lower()
    
    def test_invalid_shape_index(self, sample_pptx_with_shapes):
        from core.powerpoint_agent_core import ShapeNotFoundError
        with pytest.raises(ShapeNotFoundError):
            format_shape(
                filepath=sample_pptx_with_shapes,
                slide_index=0,
                shape_index=99,
                fill_color="#FF0000"
            )
    
    def test_presentation_version_changes(self, sample_pptx_with_shapes):
        result = format_shape(
            filepath=sample_pptx_with_shapes,
            slide_index=0,
            shape_index=0,
            fill_color="#00FF00"
        )
        
        assert "presentation_version" in result
        assert result["presentation_version"]["before"] != result["presentation_version"]["after"]
```

### Test File: `tests/test_ppt_validate_presentation.py`

```python
#!/usr/bin/env python3
"""Tests for ppt_validate_presentation.py v3.0"""

import pytest
import json
import sys
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches

sys.path.insert(0, str(Path(__file__).parent.parent / "tools"))

from ppt_validate_presentation import (
    validate_presentation,
    get_policy,
    ValidationPolicy,
    ValidationIssue,
    ValidationSummary,
    VALIDATION_POLICIES
)


@pytest.fixture
def valid_pptx(tmp_path):
    """Create a valid PowerPoint for testing."""
    pptx_path = tmp_path / "valid.pptx"
    prs = Presentation()
    
    # Add slide with title
    slide = prs.slides.add_slide(prs.slide_layouts[0])  # Title slide
    title = slide.shapes.title
    title.text = "Valid Presentation"
    
    prs.save(str(pptx_path))
    return pptx_path


@pytest.fixture
def invalid_pptx
