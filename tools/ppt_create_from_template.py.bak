#!/usr/bin/env python3
"""
PowerPoint Create From Template Tool
Create new presentation from existing .pptx template

Usage:
    uv python ppt_create_from_template.py --template corporate_template.pptx --output new_presentation.pptx --slides 10 --json

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


def create_from_template(
    template: Path,
    output: Path,
    slides: int = 1,
    layout: str = "Title and Content"
) -> Dict[str, Any]:
    """Create presentation from template."""
    
    if not template.exists():
        raise FileNotFoundError(f"Template file not found: {template}")
    
    if not template.suffix.lower() == '.pptx':
        raise ValueError(f"Template must be .pptx file, got: {template.suffix}")
    
    if slides < 1:
        raise ValueError("Must create at least 1 slide")
    
    if slides > 100:
        raise ValueError("Maximum 100 slides per creation (performance limit)")
    
    with PowerPointAgent() as agent:
        # Create from template
        agent.create_new(template=template)
        
        # Get available layouts from template
        available_layouts = agent.get_available_layouts()
        
        # Validate layout
        if layout not in available_layouts:
            # Try to find closest match
            layout_lower = layout.lower()
            for avail in available_layouts:
                if layout_lower in avail.lower():
                    layout = avail
                    break
            else:
                # Use first available layout as fallback
                layout = available_layouts[0] if available_layouts else "Title Slide"
        
        # Template comes with at least 1 slide usually, check current count
        current_slides = agent.get_slide_count()
        
        # Add additional slides if needed
        slides_to_add = max(0, slides - current_slides)
        slide_indices = list(range(current_slides))
        
        for i in range(slides_to_add):
            idx = agent.add_slide(layout_name=layout)
            slide_indices.append(idx)
        
        # Save
        agent.save(output)
        
        # Get presentation info
        info = agent.get_presentation_info()
    
    return {
        "status": "success",
        "file": str(output),
        "template_used": str(template),
        "total_slides": info["slide_count"],
        "slides_requested": slides,
        "template_slides": current_slides,
        "slides_added": slides_to_add,
        "layout_used": layout,
        "available_layouts": info["layouts"],
        "file_size_bytes": output.stat().st_size if output.exists() else 0,
        "slide_dimensions": {
            "width_inches": info["slide_width_inches"],
            "height_inches": info["slide_height_inches"],
            "aspect_ratio": info["aspect_ratio"]
        }
    }


def main():
    parser = argparse.ArgumentParser(
        description="Create PowerPoint presentation from template",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Create from corporate template
  uv python ppt_create_from_template.py \\
    --template templates/corporate.pptx \\
    --output q4_report.pptx \\
    --slides 15 \\
    --json
  
  # Create presentation matching template's default layout
  uv python ppt_create_from_template.py \\
    --template templates/minimal.pptx \\
    --output demo.pptx \\
    --slides 5 \\
    --layout "Section Header" \\
    --json
  
  # Quick presentation from template
  uv python ppt_create_from_template.py \\
    --template templates/branded.pptx \\
    --output quick_deck.pptx \\
    --json

Use Cases:
  - Corporate presentations with branding
  - Consistent theme across team presentations
  - Pre-formatted layouts (fonts, colors, logos)
  - Department-specific templates
  - Client-specific branded decks

Template Benefits:
  - Consistent branding across organization
  - Pre-configured master slides
  - Corporate colors and fonts
  - Logo placements
  - Standard layouts
  - Accessibility features

Creating Templates:
  1. Design in PowerPoint with desired theme
  2. Configure master slides
  3. Set up color scheme
  4. Define standard layouts
  5. Save as .pptx template
  6. Use with this tool

Best Practices:
  - Maintain template library for different purposes
  - Version control templates
  - Document template usage guidelines
  - Test templates before distribution
  - Include variety of layouts in template
        """
    )
    
    parser.add_argument(
        '--template',
        required=True,
        type=Path,
        help='Path to template .pptx file'
    )
    
    parser.add_argument(
        '--output',
        required=True,
        type=Path,
        help='Output presentation path'
    )
    
    parser.add_argument(
        '--slides',
        type=int,
        default=1,
        help='Total number of slides desired (default: 1)'
    )
    
    parser.add_argument(
        '--layout',
        default='Title and Content',
        help='Layout for additional slides (default: "Title and Content")'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        help='Output JSON response'
    )
    
    args = parser.parse_args()
    
    try:
        # Validate output path
        if not args.output.suffix.lower() == '.pptx':
            args.output = args.output.with_suffix('.pptx')
        
        result = create_from_template(
            template=args.template,
            output=args.output,
            slides=args.slides,
            layout=args.layout
        )
        
        if args.json:
            print(json.dumps(result, indent=2))
        else:
            print(f"✅ Created presentation from template: {result['file']}")
            print(f"   Template: {result['template_used']}")
            print(f"   Total slides: {result['total_slides']}")
            print(f"   Template had: {result['template_slides']} slides")
            print(f"   Added: {result['slides_added']} slides")
            print(f"   Layout: {result['layout_used']}")
        
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
