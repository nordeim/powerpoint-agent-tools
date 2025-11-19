#!/usr/bin/env python3
"""
PowerPoint Add Bullet List Tool
Add bullet or numbered list to slide

Usage:
    uv python ppt_add_bullet_list.py --file presentation.pptx --slide 0 --items "Item 1,Item 2,Item 3" --position '{"left":"10%","top":"25%"}' --size '{"width":"80%","height":"60%"}' --json

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
    PowerPointAgent, PowerPointAgentError, SlideNotFoundError
)


def add_bullet_list(
    filepath: Path,
    slide_index: int,
    items: List[str],
    position: Dict[str, Any],
    size: Dict[str, Any],
    bullet_style: str = "bullet",
    font_size: int = 18,
    font_name: str = "Calibri",
    color: str = None,
    line_spacing: float = 1.0
) -> Dict[str, Any]:
    """Add bullet or numbered list to slide."""
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if not items:
        raise ValueError("At least one item required")
    
    if len(items) > 20:
        raise ValueError("Maximum 20 items per list (readability limit)")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Validate slide index
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides-1})"
            )
        
        # Add bullet list
        agent.add_bullet_list(
            slide_index=slide_index,
            items=items,
            position=position,
            size=size,
            bullet_style=bullet_style,
            font_size=font_size
        )
        
        # Get the last added shape for additional formatting if needed
        slide_info = agent.get_slide_info(slide_index)
        last_shape_idx = slide_info["shape_count"] - 1
        
        # Apply color if specified
        if color:
            agent.format_text(
                slide_index=slide_index,
                shape_index=last_shape_idx,
                color=color
            )
        
        # Save
        agent.save()
    
    return {
        "status": "success",
        "file": str(filepath),
        "slide_index": slide_index,
        "items_added": len(items),
        "items": items,
        "bullet_style": bullet_style,
        "formatting": {
            "font_size": font_size,
            "font_name": font_name,
            "color": color,
            "line_spacing": line_spacing
        }
    }


def main():
    parser = argparse.ArgumentParser(
        description="Add bullet or numbered list to PowerPoint slide",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
List Formats:
  1. Comma-separated: "Item 1,Item 2,Item 3"
  2. Newline-separated: "Item 1\\nItem 2\\nItem 3"
  3. JSON array from file: --items-file items.json

Bullet Styles:
  - bullet: Traditional bullet points (default)
  - numbered: 1. 2. 3. numbering
  - none: Plain list without bullets

Examples:
  # Simple bullet list
  uv python ppt_add_bullet_list.py \\
    --file presentation.pptx \\
    --slide 1 \\
    --items "Revenue up 45%,Customer growth 60%,Market share 23%" \\
    --position '{"left":"10%","top":"25%"}' \\
    --size '{"width":"80%","height":"60%"}' \\
    --json
  
  # Numbered list with custom formatting
  uv python ppt_add_bullet_list.py \\
    --file presentation.pptx \\
    --slide 2 \\
    --items "Define objectives,Analyze market,Develop strategy,Execute plan" \\
    --bullet-style numbered \\
    --position '{"left":"15%","top":"30%"}' \\
    --size '{"width":"70%","height":"50%"}' \\
    --font-size 20 \\
    --color "#0070C0" \\
    --json
  
  # Multi-line items (agenda)
  uv python ppt_add_bullet_list.py \\
    --file presentation.pptx \\
    --slide 0 \\
    --items "Introduction and welcome,Q4 financial results,2024 strategic priorities,Q&A session" \\
    --position '{"grid":"B3"}' \\
    --size '{"width":"60%","height":"50%"}' \\
    --font-size 22 \\
    --json
  
  # Key takeaways (centered)
  uv python ppt_add_bullet_list.py \\
    --file presentation.pptx \\
    --slide 10 \\
    --items "Strong Q4 performance,Exceeded revenue targets,Positive market reception" \\
    --position '{"anchor":"center","offset_x":0,"offset_y":0}' \\
    --size '{"width":"70%","height":"40%"}' \\
    --font-size 24 \\
    --color "#00B050" \\
    --json
  
  # From JSON file
  echo '["First point", "Second point", "Third point"]' > items.json
  uv python ppt_add_bullet_list.py \\
    --file presentation.pptx \\
    --slide 3 \\
    --items-file items.json \\
    --position '{"left":"10%","top":"25%"}' \\
    --size '{"width":"80%","height":"60%"}' \\
    --json

Best Practices:
  - Keep items concise (max 2 lines per bullet)
  - Use 3-7 items per slide for readability
  - Start each item with action verb for impact
  - Use parallel structure (all items same format)
  - Font size 18-24pt for body text
  - Leave white space (don't fill entire slide)

Formatting Tips:
  - Use color for emphasis (sparingly)
  - Increase font size for key points
  - Use numbered lists for sequential steps
  - Use bullet lists for unordered points
  - Consistent spacing improves readability
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
        '--items',
        help='Comma-separated list items'
    )
    
    parser.add_argument(
        '--items-file',
        type=Path,
        help='JSON file with array of items'
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
        '--bullet-style',
        choices=['bullet', 'numbered', 'none'],
        default='bullet',
        help='Bullet style (default: bullet)'
    )
    
    parser.add_argument(
        '--font-size',
        type=int,
        default=18,
        help='Font size in points (default: 18)'
    )
    
    parser.add_argument(
        '--font-name',
        default='Calibri',
        help='Font name (default: Calibri)'
    )
    
    parser.add_argument(
        '--color',
        help='Text color (hex, e.g., #0070C0)'
    )
    
    parser.add_argument(
        '--line-spacing',
        type=float,
        default=1.0,
        help='Line spacing multiplier (default: 1.0)'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        help='Output JSON response'
    )
    
    args = parser.parse_args()
    
    try:
        # Parse items
        if args.items_file:
            if not args.items_file.exists():
                raise FileNotFoundError(f"Items file not found: {args.items_file}")
            with open(args.items_file, 'r') as f:
                items = json.load(f)
            if not isinstance(items, list):
                raise ValueError("Items file must contain JSON array")
        elif args.items:
            # Split by comma or newline
            if '\\n' in args.items:
                items = args.items.split('\\n')
            else:
                items = args.items.split(',')
            items = [item.strip() for item in items if item.strip()]
        else:
            raise ValueError("Either --items or --items-file required")
        
        result = add_bullet_list(
            filepath=args.file,
            slide_index=args.slide,
            items=items,
            position=args.position,
            size=args.size,
            bullet_style=args.bullet_style,
            font_size=args.font_size,
            font_name=args.font_name,
            color=args.color,
            line_spacing=args.line_spacing
        )
        
        if args.json:
            print(json.dumps(result, indent=2))
        else:
            print(f"✅ Added {result['items_added']}-item {result['bullet_style']} list to slide {result['slide_index']}")
            print(f"   Items:")
            for i, item in enumerate(result['items'][:5], 1):
                prefix = f"{i}." if args.bullet_style == 'numbered' else "•"
                print(f"     {prefix} {item}")
            if len(result['items']) > 5:
                print(f"     ... and {len(result['items']) - 5} more")
        
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
