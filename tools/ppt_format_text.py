#!/usr/bin/env python3
"""
PowerPoint Format Text Tool
Format existing text with accessibility validation and contrast checking

Version 2.0.0 - Enhanced Validation and Accessibility

Changes from v1.x:
- Enhanced: Shows before/after formatting preview
- Enhanced: Font size validation (warns <12pt)
- Enhanced: Color contrast checking (WCAG 2.1 AA/AAA)
- Enhanced: Shape type validation (ensures shape has text)
- Enhanced: JSON-first output (always returns JSON)
- Enhanced: Comprehensive documentation and examples
- Enhanced: Accessibility warnings and recommendations
- Enhanced: Font availability hints
- Added: Before/after comparison in response
- Added: Validation results with specific metrics
- Fixed: Consistent response format
- Fixed: Better error messages for non-text shapes

Best Practices:
- Use minimum 12pt font (14pt recommended for presentations)
- Ensure color contrast meets WCAG AA (4.5:1 for normal text)
- Test formatting on actual presentation display
- Avoid excessive bold/italic (reduces readability)
- Use standard fonts for compatibility

Usage:
    # Change font and size
    uv run tools/ppt_format_text.py --file deck.pptx --slide 0 --shape 0 --font-name "Arial" --font-size 24 --json
    
    # Make text bold and colored
    uv run tools/ppt_format_text.py --file deck.pptx --slide 1 --shape 2 --bold --color "#0070C0" --json
    
    # Comprehensive formatting with validation
    uv run tools/ppt_format_text.py --file deck.pptx --slide 0 --shape 1 --font-name "Calibri" --font-size 18 --bold --color "#000000" --json

Exit Codes:
    0: Success
    1: Error occurred
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any, List, Optional

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError, SlideNotFoundError,
    ColorHelper, RGBColor
)


def validate_formatting(
    font_size: Optional[int] = None,
    color: Optional[str] = None,
    current_font_size: Optional[int] = None
) -> Dict[str, Any]:
    """
    Validate formatting parameters.
    
    Returns:
        Dict with warnings, recommendations, and validation results
    """
    warnings = []
    recommendations = []
    validation_results = {}
    
    # Font size validation
    if font_size is not None:
        validation_results["font_size"] = font_size
        validation_results["font_size_ok"] = font_size >= 12
        
        if font_size < 10:
            warnings.append(
                f"Font size {font_size}pt is extremely small. "
                "Minimum recommended: 12pt for handouts, 14pt for presentations."
            )
        elif font_size < 12:
            warnings.append(
                f"Font size {font_size}pt is below minimum recommended 12pt. "
                "Audience may struggle to read."
            )
            recommendations.append("Use 12pt minimum, 14pt+ for projected content")
        elif font_size < 14:
            recommendations.append(
                f"Font size {font_size}pt is acceptable for handouts but consider 14pt+ for projected presentations"
            )
        
        # Check if decreasing size
        if current_font_size and font_size < current_font_size:
            diff = current_font_size - font_size
            recommendations.append(
                f"Decreasing font size by {diff}pt (from {current_font_size}pt to {font_size}pt). "
                "Verify readability on target display."
            )
    
    # Color contrast validation
    if color:
        try:
            text_color = ColorHelper.from_hex(color)
            bg_color = RGBColor(255, 255, 255)  # Assume white background
            
            # Determine if large text (use provided or assume 18pt if not specified)
            effective_font_size = font_size if font_size else current_font_size if current_font_size else 18
            is_large_text = effective_font_size >= 18
            
            contrast_ratio = ColorHelper.contrast_ratio(text_color, bg_color)
            wcag_aa = ColorHelper.meets_wcag(text_color, bg_color, is_large_text)
            
            validation_results["color_contrast"] = {
                "color": color,
                "ratio": round(contrast_ratio, 2),
                "wcag_aa": wcag_aa,
                "is_large_text": is_large_text,
                "required_ratio": 3.0 if is_large_text else 4.5
            }
            
            if not wcag_aa:
                required = 3.0 if is_large_text else 4.5
                warnings.append(
                    f"Color {color} has contrast ratio {contrast_ratio:.2f}:1 "
                    f"(WCAG AA requires {required}:1 for {'large' if is_large_text else 'normal'} text). "
                    "May not meet accessibility standards."
                )
                recommendations.append(
                    "Use high-contrast colors: #000000 (black), #333333 (dark gray), #0070C0 (dark blue)"
                )
            elif contrast_ratio < 7.0:
                recommendations.append(
                    f"Color contrast {contrast_ratio:.2f}:1 meets WCAG AA but not AAA (7:1). "
                    "Consider darker color for maximum accessibility."
                )
        except ValueError as e:
            validation_results["color_error"] = str(e)
            warnings.append(f"Invalid color format: {color}. Use hex format like #FF0000")
    
    return {
        "warnings": warnings,
        "recommendations": recommendations,
        "validation_results": validation_results
    }


def format_text(
    filepath: Path,
    slide_index: int,
    shape_index: int,
    font_name: str = None,
    font_size: int = None,
    color: str = None,
    bold: bool = None,
    italic: bool = None
) -> Dict[str, Any]:
    """
    Format text with validation and before/after reporting.
    
    Args:
        filepath: Path to PowerPoint file
        slide_index: Slide index (0-based)
        shape_index: Shape index (0-based)
        font_name: Optional font name
        font_size: Optional font size (pt)
        color: Optional text color (hex)
        bold: Optional bold setting
        italic: Optional italic setting
        
    Returns:
        Dict with:
        - status: "success" or "warning"
        - before: Original formatting
        - after: New formatting
        - validation: Validation results
        - warnings: List of issues
        - recommendations: Suggested improvements
    """
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    # Check that at least one formatting option is provided
    if all(v is None for v in [font_name, font_size, color, bold, italic]):
        raise ValueError(
            "At least one formatting option must be specified. "
            "Use --font-name, --font-size, --color, --bold, or --italic"
        )
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Validate slide index
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides-1}). "
                f"Presentation has {total_slides} slides."
            )
        
        # Get slide info to validate shape index
        slide_info = agent.get_slide_info(slide_index)
        if shape_index >= slide_info["shape_count"]:
            raise ValueError(
                f"Shape index {shape_index} out of range (0-{slide_info['shape_count']-1}). "
                f"Slide has {slide_info['shape_count']} shapes. "
                "Use ppt_get_slide_info.py to find valid shape indices."
            )
        
        # Check if shape has text
        shape_info = slide_info["shapes"][shape_index]
        if not shape_info["has_text"]:
            raise ValueError(
                f"Shape {shape_index} ({shape_info['type']}) does not contain text. "
                f"Cannot format non-text shape. "
                "Use ppt_get_slide_info.py to find text-containing shapes."
            )
        
        # Extract current formatting (basic preview)
        before_formatting = {
            "shape_type": shape_info["type"],
            "shape_name": shape_info["name"],
            "has_text": shape_info["has_text"]
        }
        
        # Get current font size if available (for validation)
        current_font_size = None
        try:
            slide = agent.get_slide(slide_index)
            shape = slide.shapes[shape_index]
            if hasattr(shape, 'text_frame') and shape.text_frame.paragraphs:
                first_para = shape.text_frame.paragraphs[0]
                if first_para.font.size:
                    current_font_size = int(first_para.font.size.pt)
                    before_formatting["font_size"] = current_font_size
        except:
            pass
        
        # Validate formatting
        validation = validate_formatting(font_size, color, current_font_size)
        
        # Apply formatting
        agent.format_text(
            slide_index=slide_index,
            shape_index=shape_index,
            font_name=font_name,
            font_size=font_size,
            bold=bold,
            italic=italic,
            color=color
        )
        
        # Save
        agent.save()
    
    # Build response
    status = "success" if len(validation["warnings"]) == 0 else "warning"
    
    after_formatting = {}
    if font_name is not None:
        after_formatting["font_name"] = font_name
    if font_size is not None:
        after_formatting["font_size"] = font_size
    if color is not None:
        after_formatting["color"] = color
    if bold is not None:
        after_formatting["bold"] = bold
    if italic is not None:
        after_formatting["italic"] = italic
    
    result = {
        "status": status,
        "file": str(filepath),
        "slide_index": slide_index,
        "shape_index": shape_index,
        "before": before_formatting,
        "after": after_formatting,
        "changes_applied": list(after_formatting.keys()),
        "validation": validation["validation_results"]
    }
    
    if validation["warnings"]:
        result["warnings"] = validation["warnings"]
    
    if validation["recommendations"]:
        result["recommendations"] = validation["recommendations"]
    
    return result


def main():
    parser = argparse.ArgumentParser(
        description="Format text in PowerPoint shape with validation (v2.0.0)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Change font and size
  uv run tools/ppt_format_text.py \\
    --file presentation.pptx \\
    --slide 0 \\
    --shape 0 \\
    --font-name "Arial" \\
    --font-size 24 \\
    --json
  
  # Make text bold and colored (with validation)
  uv run tools/ppt_format_text.py \\
    --file presentation.pptx \\
    --slide 1 \\
    --shape 2 \\
    --bold \\
    --color "#0070C0" \\
    --json
  
  # Comprehensive formatting
  uv run tools/ppt_format_text.py \\
    --file presentation.pptx \\
    --slide 0 \\
    --shape 1 \\
    --font-name "Calibri" \\
    --font-size 18 \\
    --bold \\
    --color "#000000" \\
    --json
  
  # Fix accessibility issue (increase size, darken color)
  uv run tools/ppt_format_text.py \\
    --file presentation.pptx \\
    --slide 2 \\
    --shape 3 \\
    --font-size 16 \\
    --color "#333333" \\
    --json
  
  # Remove bold/italic
  uv run tools/ppt_format_text.py \\
    --file presentation.pptx \\
    --slide 3 \\
    --shape 1 \\
    --no-bold \\
    --no-italic \\
    --json

Finding Shape Index:
  Use ppt_get_slide_info.py to list all shapes and their indices:
  
  uv run tools/ppt_get_slide_info.py --file presentation.pptx --slide 0 --json
  
  Look for "index" field in the shapes array. Only shapes with
  "has_text": true can be formatted with this tool.

Common Fonts (Cross-Platform Compatible):
  - Calibri (default Microsoft Office)
  - Arial (universal)
  - Times New Roman (classic serif)
  - Helvetica (Mac/design)
  - Georgia (readable serif)
  - Verdana (screen-optimized)
  - Tahoma (compact sans-serif)

Accessible Color Palette:
  High Contrast (WCAG AAA - 7:1):
  - Black: #000000
  - Dark Charcoal: #333333
  - Navy Blue: #003366
  
  Good Contrast (WCAG AA - 4.5:1):
  - Dark Gray: #595959
  - Corporate Blue: #0070C0
  - Forest Green: #006400
  - Dark Red: #8B0000
  
  Large Text Only (WCAG AA - 3:1):
  - Medium Gray: #767676
  - Light Blue: #4A90E2
  - Orange: #ED7D31

Validation Features:
  - Font size warnings (<12pt)
  - Color contrast checking (WCAG AA/AAA)
  - Before/after comparison
  - Shape type validation
  - Accessibility recommendations

Accessibility Guidelines:
  - Minimum font size: 12pt (14pt for presentations)
  - Color contrast: 4.5:1 for normal text, 3:1 for large text (≥18pt)
  - Avoid light colors on white background
  - Test on actual display (contrast varies by screen)
  - Consider colorblind users (avoid red/green alone)

Output Format:
  {
    "status": "warning",
    "slide_index": 0,
    "shape_index": 2,
    "before": {
      "font_size": 24
    },
    "after": {
      "font_size": 11,
      "color": "#CCCCCC"
    },
    "changes_applied": ["font_size", "color"],
    "validation": {
      "font_size": 11,
      "font_size_ok": false,
      "color_contrast": {
        "ratio": 2.1,
        "wcag_aa": false,
        "required_ratio": 4.5
      }
    },
    "warnings": [
      "Font size 11pt is below minimum recommended 12pt",
      "Color #CCCCCC has contrast ratio 2.1:1 (requires 4.5:1)"
    ],
    "recommendations": [
      "Use 12pt minimum, 14pt+ for presentations",
      "Use high-contrast colors: #000000, #333333, #0070C0"
    ]
  }

Related Tools:
  - ppt_get_slide_info.py: Find shape indices and current formatting
  - ppt_add_text_box.py: Add new text with formatting
  - ppt_set_title.py: Format title/subtitle placeholders
  - ppt_add_bullet_list.py: Add formatted lists

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
        '--shape',
        required=True,
        type=int,
        help='Shape index (0-based, use ppt_get_slide_info.py to find)'
    )
    
    parser.add_argument(
        '--font-name',
        help='Font name (e.g., Arial, Calibri)'
    )
    
    parser.add_argument(
        '--font-size',
        type=int,
        help='Font size in points (minimum recommended: 12pt)'
    )
    
    parser.add_argument(
        '--color',
        help='Text color hex (e.g., #0070C0). Contrast will be validated.'
    )
    
    parser.add_argument(
        '--bold',
        action='store_true',
        dest='bold',
        help='Make text bold'
    )
    
    parser.add_argument(
        '--no-bold',
        action='store_false',
        dest='bold',
        help='Remove bold formatting'
    )
    
    parser.add_argument(
        '--italic',
        action='store_true',
        dest='italic',
        help='Make text italic'
    )
    
    parser.add_argument(
        '--no-italic',
        action='store_false',
        dest='italic',
        help='Remove italic formatting'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default: true)'
    )
    
    parser.set_defaults(bold=None, italic=None)
    
    args = parser.parse_args()
    
    try:
        result = format_text(
            filepath=args.file,
            slide_index=args.slide,
            shape_index=args.shape,
            font_name=args.font_name,
            font_size=args.font_size,
            color=args.color,
            bold=args.bold,
            italic=args.italic
        )
        
        print(json.dumps(result, indent=2))
        sys.exit(0)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "file": str(args.file) if args.file else None,
            "slide_index": args.slide if hasattr(args, 'slide') else None,
            "shape_index": args.shape if hasattr(args, 'shape') else None
        }
        
        print(json.dumps(error_result, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()
