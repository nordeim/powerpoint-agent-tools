#!/usr/bin/env python3
"""
PowerPoint Format Text Tool
Format existing text (font, size, color, bold, italic)

Usage:
    uv python ppt_format_text.py --file presentation.pptx --slide 0 --shape 0 --font-name Arial --font-size 24 --color "#FF0000" --bold --json

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
    """Format text in specified shape."""
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    # Check that at least one formatting option is provided
    if all(v is None for v in [font_name, font_size, color, bold, italic]):
        raise ValueError("At least one formatting option must be specified")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Validate slide index
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides-1})"
            )
        
        # Get slide info to validate shape index
        slide_info = agent.get_slide_info(slide_index)
        if shape_index >= slide_info["shape_count"]:
            raise ValueError(
                f"Shape index {shape_index} out of range (0-{slide_info['shape_count']-1})"
            )
        
        # Format text
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
    
    return {
        "status": "success",
        "file": str(filepath),
        "slide_index": slide_index,
        "shape_index": shape_index,
        "formatting_applied": {
            "font_name": font_name,
            "font_size": font_size,
            "color": color,
            "bold": bold,
            "italic": italic
        }
    }


def main():
    parser = argparse.ArgumentParser(
        description="Format text in PowerPoint shape",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Change font and size
  uv python ppt_format_text.py \\
    --file presentation.pptx \\
    --slide 0 \\
    --shape 0 \\
    --font-name "Arial" \\
    --font-size 24 \\
    --json
  
  # Make text bold and red
  uv python ppt_format_text.py \\
    --file presentation.pptx \\
    --slide 1 \\
    --shape 2 \\
    --bold \\
    --color "#FF0000" \\
    --json
  
  # Comprehensive formatting
  uv python ppt_format_text.py \\
    --file presentation.pptx \\
    --slide 0 \\
    --shape 1 \\
    --font-name "Calibri" \\
    --font-size 18 \\
    --bold \\
    --italic \\
    --color "#0070C0" \\
    --json

Common Fonts:
  - Calibri (default Office)
  - Arial
  - Times New Roman
  - Helvetica
  - Georgia
  - Verdana
  - Tahoma

Color Examples:
  - Black: #000000
  - White: #FFFFFF
  - Red: #FF0000
  - Blue: #0070C0
  - Green: #00B050
  - Orange: #FFC000

Finding Shape Index:
  # Use ppt_get_slide_info.py to list shapes
  uv python ppt_get_slide_info.py --file presentation.pptx --slide 0 --json

Best Practices:
  - Use standard fonts for compatibility
  - Keep font sizes 12pt or larger
  - Ensure sufficient color contrast
  - Test on actual presentation display
  - Use bold for emphasis sparingly
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
        help='Shape index (0-based)'
    )
    
    parser.add_argument(
        '--font-name',
        help='Font name (e.g., Arial, Calibri)'
    )
    
    parser.add_argument(
        '--font-size',
        type=int,
        help='Font size in points'
    )
    
    parser.add_argument(
        '--color',
        help='Text color (hex, e.g., #FF0000)'
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
        '--json',
        action='store_true',
        help='Output JSON response'
    )
    
    args = parser.parse_args()
    
    try:
        result = format_text(
            filepath=args.file,
            slide_index=args.slide,
            shape_index=args.shape,
            font_name=args.font_name,
            font_size=args.font_size,
            color=args.color,
            bold=args.bold if args.bold else None,
            italic=args.italic if args.italic else None
        )
        
        if args.json:
            print(json.dumps(result, indent=2))
        else:
            print(f"✅ Formatted text in slide {result['slide_index']}, shape {result['shape_index']}")
            formatting = result['formatting_applied']
            if formatting['font_name']:
                print(f"   Font: {formatting['font_name']}")
            if formatting['font_size']:
                print(f"   Size: {formatting['font_size']}pt")
            if formatting['color']:
                print(f"   Color: {formatting['color']}")
            if formatting['bold']:
                print(f"   Bold: Yes")
            if formatting['italic']:
                print(f"   Italic: Yes")
        
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
