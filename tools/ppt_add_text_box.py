#!/usr/bin/env python3
"""
PowerPoint Add Text Box Tool
Add text box to slide with flexible positioning

Usage:
    uv python ppt_add_text_box.py --file presentation.pptx --slide 0 --text "Revenue: $1.5M" --position '{"left":"20%","top":"30%"}' --size '{"width":"60%","height":"10%"}' --json

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
    InvalidPositionError
)


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
    alignment: str = "left"
) -> Dict[str, Any]:
    """Add text box to slide."""
    
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
    
    return {
        "status": "success",
        "file": str(filepath),
        "slide_index": slide_index,
        "text": text[:50] + "..." if len(text) > 50 else text,
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
        "slide_shape_count": slide_info["shape_count"]
    }


def main():
    parser = argparse.ArgumentParser(
        description="Add text box to PowerPoint slide",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Position Formats:
  1. Percentage: {"left": "20%", "top": "30%"}
  2. Absolute inches: {"left": 2.0, "top": 3.0}
  3. Anchor-based: {"anchor": "center", "offset_x": 0, "offset_y": -1.0}
  4. Grid system: {"grid_row": 2, "grid_col": 3, "grid_size": 12}
  5. Excel-like: {"grid": "C4"}

Size Formats:
  - {"width": "60%", "height": "10%"}
  - {"width": 5.0, "height": 2.0}  (inches)

Anchor Points:
  top_left, top_center, top_right,
  center_left, center, center_right,
  bottom_left, bottom_center, bottom_right

Examples:
  # Percentage positioning (easiest for AI)
  uv python ppt_add_text_box.py \\
    --file presentation.pptx \\
    --slide 0 \\
    --text "Revenue: $1.5M" \\
    --position '{"left":"20%","top":"30%"}' \\
    --size '{"width":"60%","height":"10%"}' \\
    --font-size 24 \\
    --bold \\
    --json
  
  # Grid positioning (Excel-like)
  uv python ppt_add_text_box.py \\
    --file presentation.pptx \\
    --slide 1 \\
    --text "Q4 Summary" \\
    --position '{"grid":"C4"}' \\
    --size '{"width":"25%","height":"8%"}' \\
    --json
  
  # Anchor-based (centered text)
  uv python ppt_add_text_box.py \\
    --file presentation.pptx \\
    --slide 2 \\
    --text "Thank You!" \\
    --position '{"anchor":"center","offset_x":0,"offset_y":0}' \\
    --size '{"width":"80%","height":"15%"}' \\
    --font-size 48 \\
    --bold \\
    --alignment center \\
    --json
  
  # Bottom right copyright
  uv python ppt_add_text_box.py \\
    --file presentation.pptx \\
    --slide 0 \\
    --text "© 2024 Company Inc." \\
    --position '{"anchor":"bottom_right","offset_x":-0.5,"offset_y":-0.3}' \\
    --size '{"width":"2.5","height":"0.3"}' \\
    --font-size 10 \\
    --color "#808080" \\
    --json
  
  # Colored headline
  uv python ppt_add_text_box.py \\
    --file presentation.pptx \\
    --slide 1 \\
    --text "Key Takeaways" \\
    --position '{"left":"5%","top":"15%"}' \\
    --size '{"width":"90%","height":"8%"}' \\
    --font-name "Arial" \\
    --font-size 36 \\
    --bold \\
    --color "#0070C0" \\
    --json

Tips:
  - Use percentages for responsive layouts
  - Grid system ("C4") is intuitive for structured content
  - Anchor points great for headers/footers
  - Keep font size 18pt+ for readability
  - Use hex colors: #FF0000 (red), #0070C0 (blue), #00B050 (green)
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
        help='Text content'
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
        '--font-name',
        default='Calibri',
        help='Font name (default: Calibri)'
    )
    
    parser.add_argument(
        '--font-size',
        type=int,
        default=18,
        help='Font size in points (default: 18)'
    )
    
    parser.add_argument(
        '--bold',
        nargs='?',
        const='true',
        default='false',
        help='Bold text (optional: true/false)'
    )
    
    parser.add_argument(
        '--italic',
        nargs='?',
        const='true',
        default='false',
        help='Italic text (optional: true/false)'
    )
    
    parser.add_argument(
        '--color',
        help='Text color (hex, e.g., #FF0000)'
    )
    
    parser.add_argument(
        '--alignment',
        choices=['left', 'center', 'right', 'justify'],
        default='left',
        help='Text alignment (default: left)'
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
            
        # Helper to parse boolean string/flag
        def parse_bool(val):
            if isinstance(val, bool): return val
            if val is None: return False
            return str(val).lower() in ('true', 'yes', '1', 'on')

        result = add_text_box(
            filepath=args.file,
            slide_index=args.slide,
            text=args.text,
            position=position,
            size=size,
            font_name=args.font_name,
            font_size=args.font_size,
            bold=parse_bool(args.bold),
            italic=parse_bool(args.italic),
            color=args.color,
            alignment=args.alignment
        )
        
        if args.json:
            print(json.dumps(result, indent=2))
        else:
            print(f"✅ Added text box to slide {result['slide_index']}")
            print(f"   Text: {result['text']}")
            print(f"   Font: {result['formatting']['font_name']} {result['formatting']['font_size']}pt")
            if args.bold or args.italic:
                style = []
                if args.bold:
                    style.append("bold")
                if args.italic:
                    style.append("italic")
                print(f"   Style: {', '.join(style)}")
        
        sys.exit(0)
        
    except json.JSONDecodeError as e:
        error_result = {
            "status": "error",
            "error": f"Invalid JSON in position or size argument: {e}",
            "error_type": "JSONDecodeError",
            "hint": "Use single quotes around JSON and double quotes inside: '{\"left\":\"20%\",\"top\":\"30%\"}'"
        }
        
        if args.json:
            print(json.dumps(error_result, indent=2))
        else:
            print(f"❌ Error: {error_result['error']}", file=sys.stderr)
            print(f"   Hint: {error_result['hint']}", file=sys.stderr)
        
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
