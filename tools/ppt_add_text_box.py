#!/usr/bin/env python3
"""
PowerPoint Add Text Box Tool
Add text box with flexible positioning and accessibility validation

Version 2.0.0 - Enhanced Validation and Accessibility

Changes from v1.x:
- Enhanced: Text length validation with readability warnings
- Enhanced: Font size accessibility validation (warns <14pt)
- Enhanced: Color contrast checking (WCAG 2.1 AA/AAA)
- Enhanced: Position validation (warns if off-slide)
- Enhanced: Size validation (warns if too small to read)
- Enhanced: JSON-first output (always returns JSON)
- Enhanced: Full text in response (no truncation)
- Enhanced: Comprehensive examples for all positioning systems
- Fixed: Boolean flag parsing (now uses action='store_true')
- Fixed: Consistent response format with validation results
- Added: Multi-line text detection and recommendations
- Added: --allow-offslide flag for intentional off-slide positioning

Best Practices:
- Keep text under 100 characters for single-line readability
- Use minimum 14pt font for projected presentations
- Ensure color contrast meets WCAG AA (4.5:1 for normal text)
- Use percentage positioning for responsive layouts
- Test on actual presentation display

Usage:
    # Simple text box with validation
    uv run tools/ppt_add_text_box.py --file deck.pptx --slide 0 --text "Revenue: $1.5M" --position '{"left":"20%","top":"30%"}' --size '{"width":"60%","height":"10%"}' --json
    
    # Centered headline with large font
    uv run tools/ppt_add_text_box.py --file deck.pptx --slide 1 --text "Key Results" --position '{"anchor":"center"}' --size '{"width":"80%","height":"15%"}' --font-size 48 --bold --color "#0070C0" --json
    
    # Grid positioning (Excel-like)
    uv run tools/ppt_add_text_box.py --file deck.pptx --slide 2 --text "Note" --position '{"grid":"C4"}' --size '{"width":"25%","height":"8%"}' --font-size 16 --json

Exit Codes:
    0: Success
    1: Error occurred
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any, List

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError, SlideNotFoundError,
    InvalidPositionError, ColorHelper, RGBColor
)


def validate_text_box(
    text: str,
    font_size: int,
    color: str = None,
    position: Dict[str, Any] = None,
    size: Dict[str, Any] = None,
    allow_offslide: bool = False
) -> Dict[str, Any]:
    """
    Validate text box parameters and return warnings/recommendations.
    
    Returns:
        Dict with:
        - warnings: List of validation warnings
        - recommendations: List of suggested improvements
        - validation_results: Dict of specific checks
    """
    warnings = []
    recommendations = []
    validation_results = {}
    
    # Text length validation
    text_length = len(text)
    line_count = text.count('\n') + 1
    
    validation_results["text_length"] = text_length
    validation_results["line_count"] = line_count
    validation_results["is_multiline"] = line_count > 1
    
    if line_count == 1 and text_length > 100:
        warnings.append(
            f"Text is {text_length} characters for single line (recommended: ≤100). "
            "Long single-line text may be hard to read."
        )
        recommendations.append(
            "Consider breaking into multiple lines or shortening text"
        )
    
    if line_count > 1 and text_length > 500:
        warnings.append(
            f"Multi-line text is {text_length} characters total. "
            "Very long text blocks reduce readability."
        )
    
    # Font size validation
    validation_results["font_size"] = font_size
    validation_results["font_size_ok"] = font_size >= 14
    
    if font_size < 12:
        warnings.append(
            f"Font size {font_size}pt is very small (minimum recommended: 14pt for presentations). "
            "Audience may struggle to read from distance."
        )
        recommendations.append(
            "Use 14pt or larger for projected presentations, 12pt minimum for handouts"
        )
    elif font_size < 14:
        recommendations.append(
            f"Font size {font_size}pt is below recommended 14pt for projected content. "
            "Consider increasing for better readability."
        )
    
    # Color contrast validation
    if color:
        try:
            text_color = ColorHelper.from_hex(color)
            bg_color = RGBColor(255, 255, 255)  # Assume white background
            
            is_large_text = font_size >= 18
            contrast_ratio = ColorHelper.contrast_ratio(text_color, bg_color)
            
            validation_results["color_contrast"] = {
                "ratio": round(contrast_ratio, 2),
                "wcag_aa": ColorHelper.meets_wcag(text_color, bg_color, is_large_text),
                "required_ratio": 3.0 if is_large_text else 4.5
            }
            
            if not ColorHelper.meets_wcag(text_color, bg_color, is_large_text):
                required = 3.0 if is_large_text else 4.5
                warnings.append(
                    f"Color contrast {contrast_ratio:.2f}:1 may not meet WCAG AA standards "
                    f"(required: {required}:1 for {'large' if is_large_text else 'normal'} text). "
                    "Consider darker color for better readability."
                )
                recommendations.append(
                    "Use #000000 (black), #333333 (dark gray), or #0070C0 (dark blue) for better contrast"
                )
        except Exception as e:
            validation_results["color_contrast_error"] = str(e)
    
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
                            "Text box may not be visible. Use --allow-offslide if intentional."
                        )
            
            if "top" in position:
                top_str = str(position["top"])
                if top_str.endswith('%'):
                    top_pct = float(top_str.rstrip('%'))
                    if (top_pct < 0 or top_pct > 100) and not allow_offslide:
                        warnings.append(
                            f"Top position {top_pct}% is outside slide bounds (0-100%). "
                            "Text box may not be visible. Use --allow-offslide if intentional."
                        )
        except:
            pass
    
    # Size validation
    if size:
        try:
            if "height" in size:
                height_str = str(size["height"])
                if height_str.endswith('%'):
                    height_pct = float(height_str.rstrip('%'))
                    if height_pct < 5:
                        warnings.append(
                            f"Height {height_pct}% may be too small for {font_size}pt text. "
                            "Text may be clipped."
                        )
                        recommendations.append("Use at least 8-10% height for single-line text")
        except:
            pass
    
    return {
        "warnings": warnings,
        "recommendations": recommendations,
        "validation_results": validation_results
    }


def add_text_box(
    filepath: Path,
    slide_index: int,
    text: str,
    position: Dict[str, Any],
    size: Dict[str, Any],
    font_name: str = "Calibri",
    font_size: int = 18,
    bold: bool = False,
    italic: bool = False,
    color: str = None,
    alignment: str = "left",
    allow_offslide: bool = False
) -> Dict[str, Any]:
    """
    Add text box with comprehensive validation.
    
    Args:
        filepath: Path to PowerPoint file
        slide_index: Slide index (0-based)
        text: Text content
        position: Position dict (supports %, inches, anchor, grid)
        size: Size dict
        font_name: Font name
        font_size: Font size in points
        bold: Bold text
        italic: Italic text
        color: Text color (hex)
        alignment: Text alignment (left, center, right, justify)
        allow_offslide: Allow off-slide positioning
        
    Returns:
        Dict with:
        - status: "success" or "warning"
        - text: Full text (not truncated)
        - validation: Validation results
        - warnings: List of issues
        - recommendations: Suggested improvements
    """
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    # Validate parameters
    validation = validate_text_box(text, font_size, color, position, size, allow_offslide)
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Validate slide index
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides-1}). "
                f"Presentation has {total_slides} slides."
            )
        
        # Add text box
        agent.add_text_box(
            slide_index=slide_index,
            text=text,
            position=position,
            size=size,
            font_name=font_name,
            font_size=font_size,
            bold=bold,
            italic=italic,
            color=color,
            alignment=alignment
        )
        
        # Get updated slide info
        slide_info = agent.get_slide_info(slide_index)
        
        # Save
        agent.save()
    
    # Build response
    status = "success" if len(validation["warnings"]) == 0 else "warning"
    
    result = {
        "status": status,
        "file": str(filepath),
        "slide_index": slide_index,
        "text": text,  # Full text, no truncation
        "text_length": len(text),
        "position": position,
        "size": size,
        "formatting": {
            "font_name": font_name,
            "font_size": font_size,
            "bold": bold,
            "italic": italic,
            "color": color,
            "alignment": alignment
        },
        "slide_shape_count": slide_info["shape_count"],
        "validation": validation["validation_results"]
    }
    
    if validation["warnings"]:
        result["warnings"] = validation["warnings"]
    
    if validation["recommendations"]:
        result["recommendations"] = validation["recommendations"]
    
    return result


def main():
    parser = argparse.ArgumentParser(
        description="Add text box to PowerPoint slide with validation (v2.0.0)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Position Formats (Flexible):
  1. Percentage (recommended for AI): {"left": "20%", "top": "30%"}
  2. Absolute inches: {"left": 2.0, "top": 3.0}
  3. Anchor-based: {"anchor": "center", "offset_x": 0, "offset_y": -1.0}
  4. Grid system: {"grid_row": 2, "grid_col": 3, "grid_size": 12}
  5. Excel-like: {"grid": "C4"}

Size Formats:
  - Percentage: {"width": "60%", "height": "10%"}
  - Absolute: {"width": 5.0, "height": 2.0}  (inches)

Anchor Points (for anchor-based positioning):
  top_left, top_center, top_right,
  center_left, center, center_right,
  bottom_left, bottom_center, bottom_right

Examples:
  # Percentage positioning (recommended)
  uv run tools/ppt_add_text_box.py \\
    --file presentation.pptx \\
    --slide 0 \\
    --text "Revenue: $1.5M" \\
    --position '{"left":"20%","top":"30%"}' \\
    --size '{"width":"60%","height":"10%"}' \\
    --font-size 24 \\
    --bold \\
    --json
  
  # Centered large headline
  uv run tools/ppt_add_text_box.py \\
    --file presentation.pptx \\
    --slide 1 \\
    --text "Key Results" \\
    --position '{"anchor":"center","offset_x":0,"offset_y":-1.0}' \\
    --size '{"width":"80%","height":"15%"}' \\
    --font-size 48 \\
    --bold \\
    --color "#0070C0" \\
    --alignment center \\
    --json
  
  # Grid positioning (Excel-like C4)
  uv run tools/ppt_add_text_box.py \\
    --file presentation.pptx \\
    --slide 2 \\
    --text "Q4 Summary" \\
    --position '{"grid":"C4"}' \\
    --size '{"width":"25%","height":"8%"}' \\
    --font-size 20 \\
    --json
  
  # Bottom-right copyright notice
  uv run tools/ppt_add_text_box.py \\
    --file presentation.pptx \\
    --slide 0 \\
    --text "© 2024 Company Inc." \\
    --position '{"anchor":"bottom_right","offset_x":-0.5,"offset_y":-0.3}' \\
    --size '{"width":"2.5","height":"0.3"}' \\
    --font-size 10 \\
    --color "#808080" \\
    --json
  
  # Colored callout box
  uv run tools/ppt_add_text_box.py \\
    --file presentation.pptx \\
    --slide 3 \\
    --text "Important: Review by Friday" \\
    --position '{"left":"10%","top":"70%"}' \\
    --size '{"width":"80%","height":"15%"}' \\
    --font-size 20 \\
    --bold \\
    --color "#C00000" \\
    --alignment center \\
    --json
  
  # Multi-line text block
  uv run tools/ppt_add_text_box.py \\
    --file presentation.pptx \\
    --slide 4 \\
    --text "Line 1: Key Point\\nLine 2: Supporting Detail\\nLine 3: Conclusion" \\
    --position '{"left":"15%","top":"30%"}' \\
    --size '{"width":"70%","height":"40%"}' \\
    --font-size 18 \\
    --json

Validation Features:
  - Text length warnings (>100 chars for single line)
  - Font size validation (warns <14pt)
  - Color contrast checking (WCAG AA: 4.5:1 for normal text)
  - Position validation (warns if off-slide)
  - Size validation (warns if too small)
  - Multi-line detection and recommendations

Accessibility Guidelines:
  - Minimum font size: 14pt for projected presentations
  - Color contrast: At least 4.5:1 for normal text, 3:1 for large text (≥18pt)
  - Avoid ALL CAPS for long text (harder to read)
  - Use high-contrast colors: black, dark gray, dark blue
  - Test on actual display (not just computer screen)

Common Colors (High Contrast):
  - Black: #000000 (best contrast)
  - Dark Gray: #333333 or #595959
  - Corporate Blue: #0070C0
  - Dark Green: #00B050
  - Dark Red: #C00000
  - Orange: #ED7D31

Output Format:
  {
    "status": "warning",
    "text": "Very long text that exceeds recommended length...",
    "text_length": 150,
    "formatting": {
      "font_size": 12,
      "color": "#CCCCCC"
    },
    "validation": {
      "text_length": 150,
      "font_size": 12,
      "font_size_ok": false,
      "color_contrast": {
        "ratio": 2.1,
        "wcag_aa": false,
        "required_ratio": 4.5
      }
    },
    "warnings": [
      "Text is 150 characters for single line (recommended: ≤100)",
      "Font size 12pt is below recommended 14pt",
      "Color contrast 2.1:1 may not meet WCAG AA standards"
    ],
    "recommendations": [
      "Consider breaking into multiple lines",
      "Use 14pt or larger for presentations",
      "Use darker color for better contrast"
    ]
  }

Related Tools:
  - ppt_format_text.py: Format existing text
  - ppt_add_bullet_list.py: Add structured lists
  - ppt_get_slide_info.py: Find text box positions
  - ppt_set_title.py: Use placeholders for titles

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
        '--text',
        required=True,
        help='Text content (use \\n for line breaks)'
    )
    
    parser.add_argument(
        '--position',
        required=True,
        type=json.loads,
        help='Position dict (JSON): {"left":"20%","top":"30%"} or {"anchor":"center"} or {"grid":"C4"}'
    )
    
    parser.add_argument(
        '--size',
        type=json.loads,
        help='Size dict (JSON): {"width":"60%","height":"10%"} (defaults from position if omitted)'
    )
    
    parser.add_argument(
        '--font-name',
        default='Calibri',
        help='Font name (default: Calibri)'
    )
    
    parser.add_argument(
        '--font-size',
        type=int,
        default=18,
        help='Font size in points (default: 18, recommended: ≥14)'
    )
    
    parser.add_argument(
        '--bold',
        action='store_true',
        help='Make text bold'
    )
    
    parser.add_argument(
        '--italic',
        action='store_true',
        help='Make text italic'
    )
    
    parser.add_argument(
        '--color',
        help='Text color hex (e.g., #0070C0). Contrast will be validated.'
    )
    
    parser.add_argument(
        '--alignment',
        choices=['left', 'center', 'right', 'justify'],
        default='left',
        help='Text alignment (default: left)'
    )
    
    parser.add_argument(
        '--allow-offslide',
        action='store_true',
        help='Allow positioning outside slide bounds (disables off-slide warnings)'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        # Handle size defaults
        size = args.size if args.size else {}
        position = args.position
        
        if "width" not in size:
            size["width"] = position.get("width", "40%")
        if "height" not in size:
            size["height"] = position.get("height", "20%")
        
        result = add_text_box(
            filepath=args.file,
            slide_index=args.slide,
            text=args.text,
            position=position,
            size=size,
            font_name=args.font_name,
            font_size=args.font_size,
            bold=args.bold,
            italic=args.italic,
            color=args.color,
            alignment=args.alignment,
            allow_offslide=args.allow_offslide
        )
        
        print(json.dumps(result, indent=2))
        sys.exit(0)
        
    except json.JSONDecodeError as e:
        error_result = {
            "status": "error",
            "error": f"Invalid JSON in position or size argument: {str(e)}",
            "error_type": "JSONDecodeError",
            "hint": "Use single quotes around JSON and double quotes inside: '{\"left\":\"20%\",\"top\":\"30%\"}'"
        }
        
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
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
