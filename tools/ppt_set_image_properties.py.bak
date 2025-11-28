#!/usr/bin/env python3
"""
PowerPoint Set Image Properties Tool
Set alt text and transparency for images

Usage:
    uv python ppt_set_image_properties.py --file deck.pptx --slide 0 --shape 1 --alt-text "Logo" --json
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


def set_image_properties(
    filepath: Path,
    slide_index: int,
    shape_index: int,
    alt_text: str = None,
    transparency: float = None
) -> Dict[str, Any]:
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if alt_text is None and transparency is None:
        raise ValueError("At least alt-text or transparency must be set")

    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        total = agent.get_slide_count()
        if not 0 <= slide_index < total:
            raise SlideNotFoundError(f"Slide index {slide_index} out of range")
            
        agent.set_image_properties(
            slide_index=slide_index,
            shape_index=shape_index,
            alt_text=alt_text,
            transparency=transparency
        )
        agent.save()
        
    return {
        "status": "success",
        "file": str(filepath),
        "slide_index": slide_index,
        "shape_index": shape_index,
        "properties": {
            "alt_text": alt_text,
            "transparency": transparency
        }
    }


def main():
    parser = argparse.ArgumentParser(description="Set PowerPoint image properties")
    parser.add_argument('--file', required=True, type=Path, help='PowerPoint file path')
    parser.add_argument('--slide', required=True, type=int, help='Slide index (0-based)')
    parser.add_argument('--shape', required=True, type=int, help='Shape index (0-based)')
    parser.add_argument('--alt-text', help='Alt text for accessibility')
    parser.add_argument('--transparency', type=float, help='Transparency (0.0-1.0)')
    parser.add_argument('--json', action='store_true', default=True, help='Output JSON')
    
    args = parser.parse_args()
    
    try:
        result = set_image_properties(
            filepath=args.file, 
            slide_index=args.slide, 
            shape_index=args.shape,
            alt_text=args.alt_text,
            transparency=args.transparency
        )
        print(json.dumps(result, indent=2))
        sys.exit(0)
    except Exception as e:
        print(json.dumps({"status": "error", "error": str(e), "error_type": type(e).__name__}, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()
