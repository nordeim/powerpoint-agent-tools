#!/usr/bin/env python3
"""
PowerPoint Format Chart Tool
Format existing chart (title, legend position)

Usage:
    uv python ppt_format_chart.py --file presentation.pptx --slide 1 --chart 0 --title "Revenue Growth" --legend bottom --json

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


def format_chart(
    filepath: Path,
    slide_index: int,
    chart_index: int = 0,
    title: str = None,
    legend_position: str = None
) -> Dict[str, Any]:
    """Format chart on slide."""
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if title is None and legend_position is None:
        raise ValueError("At least one formatting option (title or legend) must be specified")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Validate slide index
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides-1})"
            )
        
        # Format chart
        agent.format_chart(
            slide_index=slide_index,
            chart_index=chart_index,
            title=title,
            legend_position=legend_position
        )
        
        # Save
        agent.save()
    
    return {
        "status": "success",
        "file": str(filepath),
        "slide_index": slide_index,
        "chart_index": chart_index,
        "formatting_applied": {
            "title": title,
            "legend_position": legend_position
        }
    }


def main():
    parser = argparse.ArgumentParser(
        description="Format PowerPoint chart",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Set chart title
  uv python ppt_format_chart.py \\
    --file presentation.pptx \\
    --slide 1 \\
    --chart 0 \\
    --title "Revenue Growth Trend" \\
    --json
  
  # Position legend at bottom
  uv python ppt_format_chart.py \\
    --file presentation.pptx \\
    --slide 2 \\
    --chart 0 \\
    --legend bottom \\
    --json
  
  # Set title and legend
  uv python ppt_format_chart.py \\
    --file presentation.pptx \\
    --slide 3 \\
    --chart 0 \\
    --title "Q4 Performance" \\
    --legend right \\
    --json

Legend Positions:
  - bottom: Below chart (common)
  - right: Right side of chart (default)
  - top: Above chart
  - left: Left side of chart

Finding Charts:
  Charts are indexed in order they appear on the slide.
  First chart = 0, second = 1, etc.
  
  To find charts:
  uv python ppt_get_slide_info.py --file presentation.pptx --slide 1 --json

Best Practices:
  - Keep titles concise and descriptive
  - Place legend where it doesn't obscure data
  - Bottom legend works well for wide charts
  - Right legend works well for tall charts
  - Consider removing legend if only 1 series

Chart Formatting Limitations:
  Note: python-pptx has limited chart formatting support.
  This tool handles:
  - Chart title text
  - Legend position
  
  Not supported (requires PowerPoint):
  - Individual series colors
  - Axis formatting
  - Data labels
  - Chart styles/templates
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
        '--chart',
        type=int,
        default=0,
        help='Chart index on slide (default: 0)'
    )
    
    parser.add_argument(
        '--title',
        help='Chart title text'
    )
    
    parser.add_argument(
        '--legend',
        choices=['bottom', 'left', 'right', 'top'],
        help='Legend position'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        help='Output JSON response'
    )
    
    args = parser.parse_args()
    
    try:
        result = format_chart(
            filepath=args.file,
            slide_index=args.slide,
            chart_index=args.chart,
            title=args.title,
            legend_position=args.legend
        )
        
        if args.json:
            print(json.dumps(result, indent=2))
        else:
            print(f"✅ Formatted chart on slide {result['slide_index']}")
            if args.title:
                print(f"   Title: {result['formatting_applied']['title']}")
            if args.legend:
                print(f"   Legend: {result['formatting_applied']['legend_position']}")
        
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
