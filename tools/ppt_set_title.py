#!/usr/bin/env python3
"""
PowerPoint Set Title Tool
Set slide title and optional subtitle

Usage:
    uv python ppt_set_title.py --file presentation.pptx --slide 0 --title "Q4 Results" --subtitle "Financial Review" --json

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


def set_title(
    filepath: Path,
    slide_index: int,
    title: str,
    subtitle: str = None
) -> Dict[str, Any]:
    """Set slide title and subtitle."""
    
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
        
        # Set title
        agent.set_title(slide_index, title, subtitle)
        
        # Get slide info
        slide_info = agent.get_slide_info(slide_index)
        
        # Save
        agent.save()
    
    return {
        "status": "success",
        "file": str(filepath),
        "slide_index": slide_index,
        "title": title,
        "subtitle": subtitle,
        "layout": slide_info["layout"],
        "shape_count": slide_info["shape_count"]
    }


def main():
    parser = argparse.ArgumentParser(
        description="Set PowerPoint slide title and subtitle",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Set title only
  uv python ppt_set_title.py --file presentation.pptx --slide 0 --title "Q4 Financial Results" --json
  
  # Set title and subtitle
  uv python ppt_set_title.py --file deck.pptx --slide 0 --title "2024 Strategy" --subtitle "Driving Growth & Innovation" --json
  
  # Update existing title
  uv python ppt_set_title.py --file presentation.pptx --slide 5 --title "Updated Section Title" --json
  
  # Set title on last slide
  uv python ppt_set_title.py --file presentation.pptx --slide -1 --title "Thank You" --json

Best Practices:
  - Keep titles concise (max 60 characters)
  - Use title case: "This Is Title Case"
  - Subtitles provide context, not repetition
  - First slide (index 0) should use "Title Slide" layout
  - Section headers benefit from clear, bold titles

Slide Index:
  - 0 = first slide
  - 1 = second slide
  - -1 = last slide (not yet supported in this version)
  
  To find total slides: uv python ppt_get_info.py --file your.pptx --json
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
        '--title',
        required=True,
        help='Title text'
    )
    
    parser.add_argument(
        '--subtitle',
        help='Optional subtitle text'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        help='Output JSON response'
    )
    
    args = parser.parse_args()
    
    try:
        result = set_title(
            filepath=args.file,
            slide_index=args.slide,
            title=args.title,
            subtitle=args.subtitle
        )
        
        if args.json:
            print(json.dumps(result, indent=2))
        else:
            print(f"✅ Set title on slide {result['slide_index']}")
            print(f"   Title: {result['title']}")
            if args.subtitle:
                print(f"   Subtitle: {result['subtitle']}")
            print(f"   Layout: {result['layout']}")
        
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
