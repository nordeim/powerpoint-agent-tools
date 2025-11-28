#!/usr/bin/env python3
"""
PowerPoint Insert Image Tool
Insert image into slide with automatic aspect ratio handling

Usage:
    uv python ppt_insert_image.py --file presentation.pptx --slide 0 --image logo.png --position '{"left":"10%","top":"10%"}' --size '{"width":"20%","height":"auto"}' --json

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
    PowerPointAgent, PowerPointAgentError, SlideNotFoundError,
    ImageNotFoundError, InvalidPositionError
)


def insert_image(
    filepath: Path,
    slide_index: int,
    image_path: Path,
    position: Dict[str, Any],
    size: Dict[str, Any] = None,
    compress: bool = False,
    alt_text: str = None
) -> Dict[str, Any]:
    """Insert image into slide."""
    
    if not filepath.exists():
        raise FileNotFoundError(f"Presentation file not found: {filepath}")
    
    if not image_path.exists():
        raise ImageNotFoundError(f"Image file not found: {image_path}")
    
    # Validate image format
    valid_extensions = {'.png', '.jpg', '.jpeg', '.gif', '.bmp'}
    if image_path.suffix.lower() not in valid_extensions:
        raise ValueError(
            f"Unsupported image format: {image_path.suffix}. "
            f"Supported: {valid_extensions}"
        )
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Validate slide index
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides-1})"
            )
        
        # Insert image
        agent.insert_image(
            slide_index=slide_index,
            image_path=image_path,
            position=position,
            size=size,
            compress=compress
        )
        
        # Set alt text if provided
        if alt_text:
            slide_info = agent.get_slide_info(slide_index)
            # Find the last shape (just added image)
            last_shape_idx = slide_info["shape_count"] - 1
            agent.set_image_properties(
                slide_index=slide_index,
                shape_index=last_shape_idx,
                alt_text=alt_text
            )
        
        # Get updated slide info
        slide_info = agent.get_slide_info(slide_index)
        
        # Save
        agent.save()
    
    # Get image file info
    image_size = image_path.stat().st_size
    
    return {
        "status": "success",
        "file": str(filepath),
        "slide_index": slide_index,
        "image_file": str(image_path),
        "image_size_bytes": image_size,
        "image_size_mb": round(image_size / (1024 * 1024), 2),
        "position": position,
        "size": size,
        "compressed": compress,
        "alt_text": alt_text,
        "slide_shape_count": slide_info["shape_count"]
    }


def main():
    parser = argparse.ArgumentParser(
        description="Insert image into PowerPoint slide",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Size Options:
  - {"width": "50%", "height": "auto"} - Auto-calculate height (recommended)
  - {"width": "auto", "height": "40%"} - Auto-calculate width
  - {"width": "30%", "height": "20%"} - Fixed dimensions
  - {"width": 3.0, "height": 2.0} - Absolute inches

Position Options:
  Same as ppt_add_text_box.py (percentage, anchor, grid, Excel-like)

Examples:
  # Insert logo (top-left, auto height)
  uv python ppt_insert_image.py \\
    --file presentation.pptx \\
    --slide 0 \\
    --image company_logo.png \\
    --position '{"left":"5%","top":"5%"}' \\
    --size '{"width":"15%","height":"auto"}' \\
    --alt-text "Company Logo" \\
    --json
  
  # Insert centered hero image
  uv python ppt_insert_image.py \\
    --file presentation.pptx \\
    --slide 1 \\
    --image product_photo.jpg \\
    --position '{"anchor":"center","offset_x":0,"offset_y":0}' \\
    --size '{"width":"80%","height":"auto"}' \\
    --compress \\
    --json
  
  # Insert screenshot (grid positioning)
  uv python ppt_insert_image.py \\
    --file presentation.pptx \\
    --slide 2 \\
    --image screenshot.png \\
    --position '{"grid":"B3"}' \\
    --size '{"width":"60%","height":"auto"}' \\
    --alt-text "Dashboard Screenshot - Q4 Metrics" \\
    --json
  
  # Insert chart export
  uv python ppt_insert_image.py \\
    --file presentation.pptx \\
    --slide 3 \\
    --image revenue_chart.png \\
    --position '{"left":"10%","top":"25%"}' \\
    --size '{"width":"80%","height":"auto"}' \\
    --compress \\
    --alt-text "Revenue Growth Chart 2020-2024" \\
    --json

Supported Formats:
  - PNG (recommended for logos, diagrams)
  - JPG/JPEG (recommended for photos)
  - GIF (animated not supported, will show first frame)
  - BMP (not recommended, large file size)

Best Practices:
  - Always use --alt-text for accessibility
  - Use "auto" for height/width to maintain aspect ratio
  - Use --compress for large images (>1MB)
  - Keep images under 2MB for best performance
  - Use PNG for transparency, JPG for photos
  - Recommended resolution: 1920x1080 max

Compression:
  --compress flag reduces image size by:
  - Resizing to max 1920px width
  - Converting RGBA to RGB
  - Optimizing JPEG quality to 85%
  - Typically reduces file size by 50-70%
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
        '--image',
        required=True,
        type=Path,
        help='Image file path'
    )
    
    parser.add_argument(
        '--position',
        required=True,
        type=json.loads,
        help='Position dict (JSON string)'
    )
    
    parser.add_argument(
        '--size',
        type=json.loads,
        help='Size dict (JSON string, optional - defaults to 50%% width with auto height)'
    )
    
    parser.add_argument(
        '--compress',
        action='store_true',
        help='Compress image before inserting (recommended for large images)'
    )
    
    parser.add_argument(
        '--alt-text',
        help='Alternative text for accessibility (highly recommended)'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        help='Output JSON response'
    )
    
    args = parser.parse_args()
    
    try:
        # Default size if not provided
        if not args.size:
            args.size = {"width": "50%", "height": "auto"}
        
        result = insert_image(
            filepath=args.file,
            slide_index=args.slide,
            image_path=args.image,
            position=args.position,
            size=args.size,
            compress=args.compress,
            alt_text=args.alt_text
        )
        
        if args.json:
            print(json.dumps(result, indent=2))
        else:
            print(f"✅ Inserted image into slide {result['slide_index']}")
            print(f"   Image: {result['image_file']}")
            print(f"   Size: {result['image_size_mb']} MB")
            if args.compress:
                print(f"   Compressed: Yes")
            if args.alt_text:
                print(f"   Alt text: {result['alt_text']}")
        
        sys.exit(0)
        
    except json.JSONDecodeError as e:
        error_result = {
            "status": "error",
            "error": f"Invalid JSON in position or size argument: {e}",
            "error_type": "JSONDecodeError",
            "hint": "Use single quotes around JSON: '{\"left\":\"20%\",\"top\":\"30%\"}'"
        }
        
        if args.json:
            print(json.dumps(error_result, indent=2))
        else:
            print(f"❌ Error: {error_result['error']}", file=sys.stderr)
            print(f"   Hint: {error_result['hint']}", file=sys.stderr)
        
        sys.exit(1)
        
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
