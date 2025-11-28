#!/usr/bin/env python3
"""
PowerPoint Set Title Tool
Set slide title and optional subtitle with validation

Version 2.0.0 - Enhanced Validation and JSON-First

Changes from v1.x:
- Enhanced: Now validates title length and provides readability warnings
- Enhanced: Checks for placeholder availability
- Enhanced: JSON-first output (always returns JSON)
- Enhanced: Better error messages and suggestions
- Enhanced: Layout compatibility warnings
- Enhanced: Comprehensive documentation and examples
- Fixed: Consistent response format across all scenarios
- Fixed: Proper integration with core library v1.1.0 (fixed placeholder types)

Best Practices:
- Keep titles under 60 characters for readability
- Keep subtitles under 100 characters
- Use "Title Slide" layout for first slide (index 0)
- Use title case: "This Is Title Case"
- Subtitles provide context, not repetition

Usage:
    # Set title only
    uv run tools/ppt_set_title.py --file presentation.pptx --slide 0 --title "Q4 Results" --json
    
    # Set title and subtitle
    uv run tools/ppt_set_title.py --file deck.pptx --slide 0 --title "2024 Strategy" --subtitle "Growth & Innovation" --json
    
    # Update existing title
    uv run tools/ppt_set_title.py --file presentation.pptx --slide 5 --title "New Section Title" --json

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


def set_title(
    filepath: Path,
    slide_index: int,
    title: str,
    subtitle: str = None
) -> Dict[str, Any]:
    """
    Set slide title and subtitle with validation.
    
    This tool integrates with core library v1.1.0 which correctly handles:
    - PP_PLACEHOLDER.TITLE (type 1)
    - PP_PLACEHOLDER.CENTER_TITLE (type 3)
    - PP_PLACEHOLDER.SUBTITLE (type 4) - FIXED in v1.1.0
    
    Args:
        filepath: Path to PowerPoint file
        slide_index: Slide index (0-based)
        title: Title text
        subtitle: Optional subtitle text
        
    Returns:
        Dict containing:
        - status: "success" or "warning"
        - file: File path
        - slide_index: Modified slide
        - title: Title set
        - subtitle: Subtitle set (if any)
        - layout: Current layout name
        - warnings: List of validation warnings
        - recommendations: Suggested improvements
    """
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    warnings = []
    recommendations = []
    
    # Validation: Title length
    if len(title) > 60:
        warnings.append(
            f"Title is {len(title)} characters (recommended: ≤60 for readability). "
            "Consider shortening for better visual impact."
        )
    
    if len(title) > 100:
        warnings.append(
            "Title exceeds 100 characters and may not fit on slide. "
            "Strong recommendation to shorten."
        )
    
    # Validation: Subtitle length
    if subtitle and len(subtitle) > 100:
        warnings.append(
            f"Subtitle is {len(subtitle)} characters (recommended: ≤100). "
            "Long subtitles reduce readability."
        )
    
    # Validation: Title case check (basic)
    if title == title.upper() and len(title) > 10:
        recommendations.append(
            "Title is all uppercase. Consider using title case for better readability: "
            "'This Is Title Case' instead of 'THIS IS TITLE CASE'"
        )
    
    if title == title.lower() and len(title) > 10:
        recommendations.append(
            "Title is all lowercase. Consider using title case for professionalism."
        )
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Validate slide index
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides-1}). "
                f"This presentation has {total_slides} slides."
            )
        
        # Get slide info before modification
        slide_info_before = agent.get_slide_info(slide_index)
        layout_name = slide_info_before["layout"]
        
        # Layout recommendations
        if slide_index == 0 and "Title Slide" not in layout_name:
            recommendations.append(
                f"First slide has layout '{layout_name}'. "
                "Consider using 'Title Slide' layout for cover slides."
            )
        
        # Check for title/subtitle placeholders
        has_title_placeholder = False
        has_subtitle_placeholder = False
        
        for shape in slide_info_before["shapes"]:
            shape_type = shape["type"]
            if "TITLE" in shape_type or "CENTER_TITLE" in shape_type:
                has_title_placeholder = True
            if "SUBTITLE" in shape_type:
                has_subtitle_placeholder = True
        
        if not has_title_placeholder:
            warnings.append(
                f"Layout '{layout_name}' may not have a title placeholder. "
                "Title may not display as expected. Consider changing layout first."
            )
        
        if subtitle and not has_subtitle_placeholder:
            warnings.append(
                f"Layout '{layout_name}' does not have a subtitle placeholder. "
                "Subtitle will not be displayed. Consider using 'Title Slide' layout."
            )
        
        # Set title (delegates to core library v1.1.0 with fixed placeholder types)
        agent.set_title(slide_index, title, subtitle)
        
        # Get slide info after modification for verification
        slide_info_after = agent.get_slide_info(slide_index)
        
        # Save
        agent.save()
    
    # Build response
    status = "success" if len(warnings) == 0 else "warning"
    
    result = {
        "status": status,
        "file": str(filepath),
        "slide_index": slide_index,
        "title": title,
        "subtitle": subtitle,
        "layout": layout_name,
        "shape_count": slide_info_after["shape_count"],
        "placeholders_found": {
            "title": has_title_placeholder,
            "subtitle": has_subtitle_placeholder
        },
        "validation": {
            "title_length": len(title),
            "title_length_ok": len(title) <= 60,
            "subtitle_length": len(subtitle) if subtitle else 0,
            "subtitle_length_ok": len(subtitle) <= 100 if subtitle else True
        }
    }
    
    if warnings:
        result["warnings"] = warnings
    
    if recommendations:
        result["recommendations"] = recommendations
    
    return result


def main():
    parser = argparse.ArgumentParser(
        description="Set PowerPoint slide title and subtitle with validation (v2.0.0)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Set title only
  uv run tools/ppt_set_title.py \\
    --file presentation.pptx \\
    --slide 0 \\
    --title "Q4 Financial Results" \\
    --json
  
  # Set title and subtitle (first slide)
  uv run tools/ppt_set_title.py \\
    --file deck.pptx \\
    --slide 0 \\
    --title "2024 Strategic Plan" \\
    --subtitle "Driving Growth and Innovation" \\
    --json
  
  # Update section title (middle slide)
  uv run tools/ppt_set_title.py \\
    --file presentation.pptx \\
    --slide 5 \\
    --title "Market Analysis" \\
    --json
  
  # Set title on last slide
  uv run tools/ppt_set_title.py \\
    --file presentation.pptx \\
    --slide 10 \\
    --title "Thank You" \\
    --subtitle "Questions?" \\
    --json

Best Practices:
  Title Guidelines:
  - Keep under 60 characters (optimal readability)
  - Use title case: "This Is Title Case"
  - Be specific and descriptive
  - Avoid jargon and abbreviations
  - One clear message per title
  
  Subtitle Guidelines:
  - Keep under 100 characters
  - Provide context, not repetition
  - Use for date, location, or clarification
  - Optional on content slides
  
  Layout Recommendations:
  - Slide 0 (first): Use "Title Slide" layout
  - Section headers: Use "Section Header" layout
  - Content slides: Use "Title and Content" layout
  - Blank slides: Use "Title Only" layout

Validation:
  This tool performs automatic validation:
  - Title length (warns if >60 chars, strong warning if >100)
  - Subtitle length (warns if >100 chars)
  - Title case recommendations
  - Placeholder availability checks
  - Layout compatibility warnings

Output Format:
  Always returns JSON with:
  - status: "success" or "warning"
  - validation: Length checks and recommendations
  - warnings: Issues that should be addressed
  - recommendations: Optional improvements
  
  Example success output:
  {
    "status": "success",
    "slide_index": 0,
    "title": "Q4 Results",
    "subtitle": "Financial Review",
    "layout": "Title Slide",
    "validation": {
      "title_length": 10,
      "title_length_ok": true
    }
  }
  
  Example with warnings:
  {
    "status": "warning",
    "title": "This Is A Very Long Title That Exceeds Recommended Length",
    "warnings": [
      "Title is 61 characters (recommended: ≤60)"
    ],
    "recommendations": [
      "Consider shortening for better visual impact"
    ]
  }

Related Tools:
  - ppt_get_slide_info.py: Inspect slide layout and placeholders
  - ppt_set_slide_layout.py: Change slide layout
  - ppt_get_info.py: Get presentation info (total slides, layouts)
  - ppt_add_text_box.py: Add custom text if placeholders unavailable

Version: 2.0.0
Requires: core/powerpoint_agent_core.py v1.1.0+
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
        help='Slide index (0-based, e.g., 0 for first slide)'
    )
    
    parser.add_argument(
        '--title',
        required=True,
        help='Title text (recommended: ≤60 characters)'
    )
    
    parser.add_argument(
        '--subtitle',
        help='Optional subtitle text (recommended: ≤100 characters)'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        result = set_title(
            filepath=args.file,
            slide_index=args.slide,
            title=args.title,
            subtitle=args.subtitle
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
