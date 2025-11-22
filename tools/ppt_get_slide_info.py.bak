#!/usr/bin/env python3
"""
PowerPoint Get Slide Info Tool
Get detailed information about slide content (shapes, images, text)

Usage:
    uv python ppt_get_slide_info.py --file presentation.pptx --slide 0 --json

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


def get_slide_info(
    filepath: Path,
    slide_index: int
) -> Dict[str, Any]:
    """Get detailed slide information."""
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath, acquire_lock=False)
        
        # Validate slide index
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides-1})"
            )
        
        # Get slide info
        slide_info = agent.get_slide_info(slide_index)
    
    return {
        "status": "success",
        "file": str(filepath),
        "slide_index": slide_index,
        "layout": slide_info["layout"],
        "shape_count": slide_info["shape_count"],
        "shapes": slide_info["shapes"],
        "has_notes": slide_info["has_notes"]
    }


def main():
    parser = argparse.ArgumentParser(
        description="Get PowerPoint slide information",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Get info for first slide
  uv python ppt_get_slide_info.py \\
    --file presentation.pptx \\
    --slide 0 \\
    --json
  
  # Get info for specific slide
  uv python ppt_get_slide_info.py \\
    --file presentation.pptx \\
    --slide 5 \\
    --json

Output Information:
  - Slide layout name
  - Total shape count
  - List of all shapes with:
    - Shape index (for targeting with other tools)
    - Shape type (PLACEHOLDER, PICTURE, TEXT_BOX, etc.)
    - Shape name
    - Whether it contains text
    - Text preview (first 100 chars)
    - Image size (for pictures)

Use Cases:
  - Find shape indices for ppt_format_text.py
  - Locate images for ppt_replace_image.py
  - Inspect slide layout
  - Audit slide content
  - Debug presentation structure

Finding Shape Indices:
  Use this tool before:
  - ppt_format_text.py (need shape index)
  - ppt_replace_image.py (need image name)
  - ppt_format_shape.py (need shape index)

Example Output:
{
  "slide_index": 0,
  "layout": "Title Slide",
  "shape_count": 3,
  "shapes": [
    {
      "index": 0,
      "type": "PLACEHOLDER",
      "name": "Title 1",
      "has_text": true,
      "text": "My Presentation"
    },
    {
      "index": 1,
      "type": "PICTURE",
      "name": "company_logo",
      "has_text": false,
      "image_size_bytes": 45678
    }
  ]
}
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
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default)'
    )
    
    args = parser.parse_args()
    
    try:
        result = get_slide_info(
            filepath=args.file,
            slide_index=args.slide
        )
        
        print(json.dumps(result, indent=2))
        sys.exit(0)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__
        }
        
        print(json.dumps(error_result, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()
