#!/usr/bin/env python3
"""
PowerPoint Export Images Tool
Export each slide as PNG or JPG image

Usage:
    uv python ppt_export_images.py --file presentation.pptx --output-dir output/ --format png --json

Exit Codes:
    0: Success
    1: Error occurred

Requirements:
    LibreOffice must be installed for image export
    - Linux: sudo apt install libreoffice-impress
    - macOS: brew install --cask libreoffice
    - Windows: Download from https://www.libreoffice.org/
"""

import sys
import json
import argparse
import subprocess
import shutil
from pathlib import Path
from typing import Dict, Any, List

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError
)


def check_libreoffice() -> bool:
    """Check if LibreOffice is installed."""
    return shutil.which('soffice') is not None or shutil.which('libreoffice') is not None


def export_images(
    filepath: Path,
    output_dir: Path,
    format: str = "png",
    prefix: str = "slide_"
) -> Dict[str, Any]:
    """Export slides as images."""
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if not filepath.suffix.lower() == '.pptx':
        raise ValueError(f"Input must be .pptx file, got: {filepath.suffix}")
    
    if format.lower() not in ['png', 'jpg', 'jpeg']:
        raise ValueError(f"Format must be png or jpg, got: {format}")
    
    # Normalize format
    format_ext = 'png' if format.lower() == 'png' else 'jpg'
    
    # Check LibreOffice
    if not check_libreoffice():
        raise RuntimeError(
            "LibreOffice not found. Image export requires LibreOffice.\n"
            "Install:\n"
            "  Linux: sudo apt install libreoffice-impress\n"
            "  macOS: brew install --cask libreoffice\n"
            "  Windows: https://www.libreoffice.org/download/"
        )
    
    # Create output directory
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # Get slide count first
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath, acquire_lock=False)
        slide_count = agent.get_slide_count()
    
    # Use LibreOffice to export
    # LibreOffice exports all slides when converting to image format
    cmd = [
        'soffice' if shutil.which('soffice') else 'libreoffice',
        '--headless',
        '--convert-to', format_ext,
        '--outdir', str(output_dir),
        str(filepath)
    ]
    
    result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
    
    if result.returncode != 0:
        raise PowerPointAgentError(
            f"Image export failed: {result.stderr}\n"
            f"Command: {' '.join(cmd)}"
        )
    
    # LibreOffice creates files named: presentation.png for single page
    # or presentation_01.png, presentation_02.png for multiple pages
    base_name = filepath.stem
    
    # Find exported images using a robust pattern match
    # LibreOffice naming can vary (e.g. file.png, file_01.png, file_1.png)
    candidates = sorted(output_dir.glob(f"{base_name}*.{format_ext}"))
    
    exported_files = []
    
    if not candidates:
        # Try without base_name prefix just in case (unlikely but possible if LO behaves weirdly)
        pass
    
    # Rename found files sequentially
    for i, old_file in enumerate(candidates):
        new_file = output_dir / f"{prefix}{i+1:03d}.{format_ext}"
        
        # Handle case where source and dest are same (e.g. prefix is same as original name start)
        if old_file != new_file:
            # If target exists (from previous run), remove it to avoid collision/confusion
            if new_file.exists():
                new_file.unlink()
            old_file.rename(new_file)
            exported_files.append(new_file)
        else:
            exported_files.append(old_file)
    
    if len(exported_files) == 0:
        raise PowerPointAgentError(
            "Export completed but no image files found. "
            f"Expected files in: {output_dir}"
        )
    
    # Get file sizes
    total_size = sum(f.stat().st_size for f in exported_files)
    
    return {
        "status": "success",
        "input_file": str(filepath),
        "output_dir": str(output_dir),
        "format": format_ext,
        "slides_exported": len(exported_files),
        "files": [str(f) for f in exported_files],
        "total_size_bytes": total_size,
        "total_size_mb": round(total_size / (1024 * 1024), 2),
        "average_size_mb": round(total_size / (1024 * 1024) / len(exported_files), 2) if exported_files else 0
    }


def main():
    parser = argparse.ArgumentParser(
        description="Export PowerPoint slides as images",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Export as PNG
  uv python ppt_export_images.py \\
    --file presentation.pptx \\
    --output-dir slides/ \\
    --format png \\
    --json
  
  # Export as JPG with custom prefix
  uv python ppt_export_images.py \\
    --file presentation.pptx \\
    --output-dir images/ \\
    --format jpg \\
    --prefix deck_ \\
    --json

Output Files:
  Files are named: <prefix><number>.<format>
  Examples:
    slide_001.png
    slide_002.png
    deck_001.jpg
    deck_002.jpg

Requirements:
  LibreOffice must be installed:
  
  Linux:
    sudo apt install libreoffice-impress
  
  macOS:
    brew install --cask libreoffice
  
  Windows:
    Download from https://www.libreoffice.org/download/

Use Cases:
  - Website screenshots
  - Social media sharing
  - Email attachments
  - Documentation
  - Thumbnails for preview
  - Archive for reference

Format Comparison:
  PNG:
    - Lossless compression
    - Better for text/diagrams
    - Larger file size
    - Supports transparency
  
  JPG:
    - Lossy compression
    - Better for photos
    - Smaller file size
    - No transparency

Performance:
  - Export time: ~2-3 seconds per slide
  - File size: 200-500 KB per slide (PNG)
  - File size: 100-300 KB per slide (JPG)

Troubleshooting:
  If export fails:
  1. Verify LibreOffice: soffice --version
  2. Check disk space
  3. Ensure file not corrupted
  4. Try shorter timeout for small files
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file to export'
    )
    
    parser.add_argument(
        '--output-dir',
        required=True,
        type=Path,
        help='Output directory for images'
    )
    
    parser.add_argument(
        '--format',
        choices=['png', 'jpg', 'jpeg'],
        default='png',
        help='Image format (default: png)'
    )
    
    parser.add_argument(
        '--prefix',
        default='slide_',
        help='Filename prefix (default: slide_)'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        help='Output JSON response'
    )
    
    args = parser.parse_args()
    
    try:
        result = export_images(
            filepath=args.file,
            output_dir=args.output_dir,
            format=args.format,
            prefix=args.prefix
        )
        
        if args.json:
            print(json.dumps(result, indent=2))
        else:
            print(f"✅ Exported {result['slides_exported']} slides to {result['output_dir']}")
            print(f"   Format: {result['format'].upper()}")
            print(f"   Total size: {result['total_size_mb']} MB")
            print(f"   Average: {result['average_size_mb']} MB per slide")
        
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
