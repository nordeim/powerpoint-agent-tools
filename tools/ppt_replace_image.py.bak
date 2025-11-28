#!/usr/bin/env python3
"""
PowerPoint Replace Image Tool
Replace existing image (useful for logo/photo updates)

Usage:
    uv python ppt_replace_image.py --file presentation.pptx --slide 0 --old-image "logo" --new-image new_logo.png --json

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
    ImageNotFoundError
)


def replace_image(
    filepath: Path,
    slide_index: int,
    old_image: str,
    new_image: Path,
    compress: bool = False
) -> Dict[str, Any]:
    """Replace image in presentation."""
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if not new_image.exists():
        raise ImageNotFoundError(f"New image not found: {new_image}")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Validate slide index
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides-1})"
            )
        
        # Try to replace by name
        replaced = agent.replace_image(
            slide_index=slide_index,
            old_image_name=old_image,
            new_image_path=new_image,
            compress=compress
        )
        
        if not replaced:
            raise ImageNotFoundError(
                f"Image '{old_image}' not found on slide {slide_index}. "
                "Use ppt_get_slide_info.py to list images."
            )
        
        # Save
        agent.save()
    
    # Get new image size
    new_size = new_image.stat().st_size
    
    return {
        "status": "success",
        "file": str(filepath),
        "slide_index": slide_index,
        "old_image": old_image,
        "new_image": str(new_image),
        "new_image_size_bytes": new_size,
        "new_image_size_mb": round(new_size / (1024 * 1024), 2),
        "compressed": compress,
        "replaced": True
    }


def main():
    parser = argparse.ArgumentParser(
        description="Replace image in PowerPoint presentation",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Replace logo by name
  uv python ppt_replace_image.py \\
    --file presentation.pptx \\
    --slide 0 \\
    --old-image "company_logo" \\
    --new-image new_logo.png \\
    --json
  
  # Replace and compress
  uv python ppt_replace_image.py \\
    --file presentation.pptx \\
    --slide 5 \\
    --old-image "product_photo" \\
    --new-image updated_photo.jpg \\
    --compress \\
    --json
  
  # Replace image with partial name match
  uv python ppt_replace_image.py \\
    --file presentation.pptx \\
    --slide 0 \\
    --old-image "logo" \\
    --new-image rebrand_logo.png \\
    --json

Finding Images:
  # List images on slide
  uv python ppt_get_slide_info.py \\
    --file presentation.pptx \\
    --slide 0 \\
    --json

Use Cases:
  - Logo updates (rebranding)
  - Product photo updates
  - Team photo updates
  - Chart/diagram updates
  - Screenshot updates

Search Strategy:
  The tool searches for images by:
  1. Exact name match
  2. Partial name match (contains)
  3. First match wins

Image Compression:
  --compress flag reduces size by:
  - Resizing to max 1920px width
  - Converting to JPEG at 85% quality
  - Typically reduces size 50-70%

Best Practices:
  - Use descriptive image names in PowerPoint
  - Keep new images similar dimensions to old
  - Use --compress for large images (>1MB)
  - Test on a copy first
  - Verify aspect ratios match
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
        '--old-image',
        required=True,
        help='Name or pattern of image to replace'
    )
    
    parser.add_argument(
        '--new-image',
        required=True,
        type=Path,
        help='Path to new image file'
    )
    
    parser.add_argument(
        '--compress',
        action='store_true',
        help='Compress new image before inserting'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        help='Output JSON response'
    )
    
    args = parser.parse_args()
    
    try:
        result = replace_image(
            filepath=args.file,
            slide_index=args.slide,
            old_image=args.old_image,
            new_image=args.new_image,
            compress=args.compress
        )
        
        if args.json:
            print(json.dumps(result, indent=2))
        else:
            print(f"✅ Replaced image on slide {result['slide_index']}")
            print(f"   Old: {result['old_image']}")
            print(f"   New: {result['new_image']}")
            print(f"   Size: {result['new_image_size_mb']} MB")
            if args.compress:
                print(f"   Compressed: Yes")
        
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
