#!/usr/bin/env python3
"""
PowerPoint Set Background Tool
Set slide background to color or image

Usage:
    uv python ppt_set_background.py --file deck.pptx --color "#FFFFFF" --json
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError, ColorHelper
)

def set_background(
    filepath: Path,
    color: str = None,
    image: Path = None,
    slide_index: int = None
) -> Dict[str, Any]:
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
        
    if not color and not image:
        raise ValueError("Must specify either --color or --image")

    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Determine target slides
        if slide_index is not None:
             target_slides = [agent.prs.slides[slide_index]]
        else:
             target_slides = agent.prs.slides
             
        for slide in target_slides:
            bg = slide.background
            fill = bg.fill
            
            if color:
                fill.solid()
                fill.fore_color.rgb = ColorHelper.from_hex(color)
            elif image:
                if not image.exists():
                    raise FileNotFoundError(f"Image not found: {image}")
                # Note: python-pptx background image support is limited in some versions
                # but user_picture is the standard method
                fill.user_picture(str(image))
                
        agent.save()
        
    return {
        "status": "success",
        "file": str(filepath),
        "slides_affected": len(target_slides),
        "type": "color" if color else "image"
    }

def main():
    parser = argparse.ArgumentParser(description="Set PowerPoint background")
    parser.add_argument('--file', required=True, type=Path, help='PowerPoint file path')
    parser.add_argument('--slide', type=int, help='Slide index (optional, defaults to all)')
    parser.add_argument('--color', help='Hex color code')
    parser.add_argument('--image', type=Path, help='Background image path')
    parser.add_argument('--json', action='store_true', default=True, help='Output JSON')
    
    args = parser.parse_args()
    
    try:
        result = set_background(
            filepath=args.file, 
            slide_index=args.slide, 
            color=args.color,
            image=args.image
        )
        print(json.dumps(result, indent=2))
        sys.exit(0)
    except Exception as e:
        print(json.dumps({"status": "error", "error": str(e), "error_type": type(e).__name__}, indent=2))
        sys.exit(1)

if __name__ == "__main__":
    main()
