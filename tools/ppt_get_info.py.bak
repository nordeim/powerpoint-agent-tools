#!/usr/bin/env python3
"""
PowerPoint Get Info Tool
Get presentation metadata (slide count, dimensions, file size)

Usage:
    uv python ppt_get_info.py --file presentation.pptx --json

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
    PowerPointAgent, PowerPointAgentError
)


def get_info(filepath: Path) -> Dict[str, Any]:
    """Get presentation information."""
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath, acquire_lock=False)
        
        # Get presentation info
        info = agent.get_presentation_info()
    
    return {
        "status": "success",
        "file": info["file"],
        "slide_count": info["slide_count"],
        "file_size_bytes": info.get("file_size_bytes", 0),
        "file_size_mb": info.get("file_size_mb", 0),
        "slide_dimensions": {
            "width_inches": info["slide_width_inches"],
            "height_inches": info["slide_height_inches"],
            "aspect_ratio": info["aspect_ratio"]
        },
        "layouts": info["layouts"],
        "layout_count": len(info["layouts"]),
        "modified": info.get("modified")
    }


def main():
    parser = argparse.ArgumentParser(
        description="Get PowerPoint presentation information",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Get presentation info
  uv python ppt_get_info.py \\
    --file presentation.pptx \\
    --json

Output Information:
  - File path and size
  - Total slide count
  - Slide dimensions (width x height)
  - Aspect ratio (16:9, 4:3, etc.)
  - Available layouts
  - Last modified date

Use Cases:
  - Verify presentation structure
  - Check aspect ratio before editing
  - List available layouts
  - File size checking
  - Metadata inspection

Example Output:
{
  "file": "presentation.pptx",
  "slide_count": 15,
  "file_size_mb": 2.45,
  "slide_dimensions": {
    "width_inches": 10.0,
    "height_inches": 7.5,
    "aspect_ratio": "16:9"
  },
  "layouts": [
    "Title Slide",
    "Title and Content",
    "Section Header"
  ],
  "layout_count": 11
}

Aspect Ratios:
  - 16:9 (Widescreen): Most common, modern standard
  - 4:3 (Standard): Traditional, older format
  - 16:10: Some displays, between 16:9 and 4:3

Layout Information:
  The layouts list shows all slide layouts available in the presentation.
  Use these names with:
  - ppt_create_new.py --layout "Title Slide"
  - ppt_add_slide.py --layout "Title and Content"
  - ppt_set_slide_layout.py --layout "Section Header"
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file path'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default)'
    )
    
    args = parser.parse_args()
    
    try:
        result = get_info(filepath=args.file)
        
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
