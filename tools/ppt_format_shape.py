#!/usr/bin/env python3
"""
PowerPoint Format Shape Tool
Update fill color, line color, and line width of an existing shape

Usage:
    uv python ppt_format_shape.py --file presentation.pptx --slide 0 --shape 1 --fill-color "#FF0000" --json
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


def format_shape(
    filepath: Path,
    slide_index: int,
    shape_index: int,
    fill_color: str = None,
    line_color: str = None,
    line_width: float = None
) -> Dict[str, Any]:
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if all(v is None for v in [fill_color, line_color, line_width]):
        raise ValueError("At least one formatting option required")

    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        total = agent.get_slide_count()
        if not 0 <= slide_index < total:
            raise SlideNotFoundError(f"Slide index {slide_index} out of range")
            
        agent.format_shape(
            slide_index=slide_index,
            shape_index=shape_index,
            fill_color=fill_color,
            line_color=line_color,
            line_width=line_width
        )
        agent.save()
        
    return {
        "status": "success",
        "file": str(filepath),
        "slide_index": slide_index,
        "shape_index": shape_index,
        "formatting": {
            "fill_color": fill_color,
            "line_color": line_color,
            "line_width": line_width
        }
    }


def main():
    parser = argparse.ArgumentParser(description="Format PowerPoint shape")
    parser.add_argument('--file', required=True, type=Path, help='PowerPoint file path')
    parser.add_argument('--slide', required=True, type=int, help='Slide index (0-based)')
    parser.add_argument('--shape', required=True, type=int, help='Shape index (0-based)')
    parser.add_argument('--fill-color', help='Fill hex color')
    parser.add_argument('--line-color', help='Line hex color')
    parser.add_argument('--line-width', type=float, help='Line width in points')
    parser.add_argument('--json', action='store_true', default=True, help='Output JSON')
    
    args = parser.parse_args()
    
    try:
        result = format_shape(
            filepath=args.file, 
            slide_index=args.slide, 
            shape_index=args.shape,
            fill_color=args.fill_color,
            line_color=args.line_color,
            line_width=args.line_width
        )
        print(json.dumps(result, indent=2))
        sys.exit(0)
    except Exception as e:
        print(json.dumps({"status": "error", "error": str(e), "error_type": type(e).__name__}, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()
