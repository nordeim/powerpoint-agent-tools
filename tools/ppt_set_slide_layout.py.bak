#!/usr/bin/env python3
"""
PowerPoint Set Slide Layout Tool
Change the layout of an existing slide

Usage:
    uv python ppt_set_slide_layout.py --file presentation.pptx --slide 0 --layout "Title Only" --json
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError, SlideNotFoundError, LayoutNotFoundError
)


def set_slide_layout(filepath: Path, slide_index: int, layout_name: str) -> Dict[str, Any]:
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        total = agent.get_slide_count()
        if not 0 <= slide_index < total:
            raise SlideNotFoundError(f"Slide index {slide_index} out of range")
            
        # Get available layouts to validate/fuzzy match
        available = agent.get_available_layouts()
        if layout_name not in available:
            # Try fuzzy match
            layout_lower = layout_name.lower()
            for avail in available:
                if layout_lower in avail.lower():
                    layout_name = avail
                    break
            else:
                 raise LayoutNotFoundError(f"Layout '{layout_name}' not found. Available: {available}")

        agent.set_slide_layout(slide_index, layout_name)
        agent.save()
        
    return {
        "status": "success",
        "file": str(filepath),
        "slide_index": slide_index,
        "new_layout": layout_name
    }


def main():
    parser = argparse.ArgumentParser(description="Set PowerPoint slide layout")
    parser.add_argument('--file', required=True, type=Path, help='PowerPoint file path')
    parser.add_argument('--slide', required=True, type=int, help='Slide index (0-based)')
    parser.add_argument('--layout', required=True, help='New layout name')
    parser.add_argument('--json', action='store_true', default=True, help='Output JSON')
    
    args = parser.parse_args()
    
    try:
        result = set_slide_layout(filepath=args.file, slide_index=args.slide, layout_name=args.layout)
        print(json.dumps(result, indent=2))
        sys.exit(0)
    except Exception as e:
        print(json.dumps({"status": "error", "error": str(e), "error_type": type(e).__name__}, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()
