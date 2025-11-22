#!/usr/bin/env python3
"""
PowerPoint Get Slide Info Tool
Get detailed information about slide content (shapes, images, text, positions)

Version 2.0.0 - Enhanced with Full Text and Position Data

Changes from v1.x:
- Fixed: Removed 100-character text truncation (was causing data loss)
- Enhanced: Now returns full text content with separate preview field
- Enhanced: Added position information (inches and percentages)
- Enhanced: Added size information (inches and percentages)
- Enhanced: Human-readable placeholder type names (e.g., "PLACEHOLDER (TITLE)" not "PLACEHOLDER (14)")
- Enhanced: Better shape type identification

This tool is critical for:
- Finding shape indices for ppt_format_text.py
- Locating images for ppt_replace_image.py
- Debugging positioning issues
- Auditing slide content
- Verifying footer/header presence

Usage:
    # Get info for first slide
    uv run tools/ppt_get_slide_info.py --file presentation.pptx --slide 0 --json
    
    # Get info for specific slide
    uv run tools/ppt_get_slide_info.py --file presentation.pptx --slide 5 --json
    
    # Inspect footer elements
    uv run tools/ppt_get_slide_info.py --file presentation.pptx --slide 1 --json | grep -i footer

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
    """
    Get detailed slide information including full text and positioning.
    
    Returns comprehensive information about:
    - All shapes on the slide
    - Full text content (no truncation)
    - Position (in inches and percentages)
    - Size (in inches and percentages)
    - Placeholder types (human-readable names)
    - Image metadata
    
    Args:
        filepath: Path to .pptx file
        slide_index: Slide index (0-based)
        
    Returns:
        Dict containing:
        - slide_index: Index of slide
        - layout: Layout name
        - shape_count: Total shapes
        - shapes: List of shape information dicts
        - has_notes: Whether slide has speaker notes
        
    Each shape dict contains:
        - index: Shape index (for targeting with other tools)
        - type: Shape type (with human-readable placeholder names)
        - name: Shape name
        - has_text: Boolean
        - text: Full text content (no truncation!)
        - text_length: Character count
        - text_preview: First 100 chars (if text > 100 chars)
        - position: Dict with inches and percentages
        - size: Dict with inches and percentages
        - image_size_bytes: For pictures only
    """
    
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
        
        # Get enhanced slide info from core (now includes full text and positions)
        slide_info = agent.get_slide_info(slide_index)
    
    return {
        "status": "success",
        "file": str(filepath),
        "slide_index": slide_info["slide_index"],
        "layout": slide_info["layout"],
        "shape_count": slide_info["shape_count"],
        "shapes": slide_info["shapes"],
        "has_notes": slide_info["has_notes"]
    }


def main():
    parser = argparse.ArgumentParser(
        description="Get PowerPoint slide information with full text and positioning",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Get info for first slide
  uv run tools/ppt_get_slide_info.py \\
    --file presentation.pptx \\
    --slide 0 \\
    --json
  
  # Get info for specific slide
  uv run tools/ppt_get_slide_info.py \\
    --file presentation.pptx \\
    --slide 5 \\
    --json
  
  # Find footer elements
  uv run tools/ppt_get_slide_info.py --file deck.pptx --slide 1 --json | jq '.shapes[] | select(.type | contains("FOOTER"))'
  
  # Find text boxes at bottom of slide (footer candidates)
  uv run tools/ppt_get_slide_info.py --file deck.pptx --slide 1 --json | jq '.shapes[] | select(.position.top_percent > "90%")'

Output Information:
  - Slide layout name
  - Total shape count
  - List of all shapes with:
    - Shape index (for targeting with other tools)
    - Shape type with human-readable placeholder names
      (e.g., "PLACEHOLDER (TITLE)" instead of "PLACEHOLDER (14)")
    - Shape name
    - Whether it contains text
    - FULL text content (no truncation in v2.0!)
    - Text length in characters
    - Text preview (first 100 chars if longer)
    - Position in both inches and percentages
    - Size in both inches and percentages
    - Image size for pictures

Use Cases:
  - Find shape indices for ppt_format_text.py
  - Locate images for ppt_replace_image.py
  - Inspect slide layout and structure
  - Audit slide content
  - Debug positioning issues
  - Verify footer/header presence
  - Check if text is truncated or overflowing

Finding Shape Indices:
  Use this tool before:
  - ppt_format_text.py (needs shape index)
  - ppt_replace_image.py (needs image name)
  - ppt_format_shape.py (needs shape index)
  - ppt_set_image_properties.py (needs shape index)

Example Output:
{
  "status": "success",
  "slide_index": 0,
  "layout": "Title Slide",
  "shape_count": 5,
  "shapes": [
    {
      "index": 0,
      "type": "PLACEHOLDER (TITLE)",
      "name": "Title 1",
      "has_text": true,
      "text": "My Presentation Title",
      "text_length": 21,
      "position": {
        "left_inches": 0.5,
        "top_inches": 1.0,
        "left_percent": "5.0%",
        "top_percent": "13.3%"
      },
      "size": {
        "width_inches": 9.0,
        "height_inches": 1.5,
        "width_percent": "90.0%",
        "height_percent": "20.0%"
      }
    },
    {
      "index": 3,
      "type": "TEXT_BOX",
      "name": "TextBox 4",
      "has_text": true,
      "text": "Bitcoin Market Report • November 2024",
      "text_length": 38,
      "position": {
        "left_inches": 0.5,
        "top_inches": 6.9,
        "left_percent": "5.0%",
        "top_percent": "92.0%"
      },
      "size": {
        "width_inches": 6.0,
        "height_inches": 0.375,
        "width_percent": "60.0%",
        "height_percent": "5.0%"
      }
    }
  ],
  "has_notes": false
}

Changes in v2.0:
  - Full text returned (no 100-char truncation)
  - Position and size data included
  - Placeholder types human-readable
  - Text preview field for long text
  - Better documentation
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
        help='Output JSON response (default: true)'
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
            "error_type": type(e).__name__,
            "file": str(args.file) if args.file else None,
            "slide_index": args.slide if hasattr(args, 'slide') else None
        }
        
        print(json.dumps(error_result, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()
