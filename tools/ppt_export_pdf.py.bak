#!/usr/bin/env python3
"""
PowerPoint Export PDF Tool
Export presentation to PDF format

Usage:
    uv python ppt_export_pdf.py --file presentation.pptx --output presentation.pdf --json

Exit Codes:
    0: Success
    1: Error occurred

Requirements:
    LibreOffice must be installed for PDF export
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
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError
)


def check_libreoffice() -> bool:
    """Check if LibreOffice is installed."""
    return shutil.which('soffice') is not None or shutil.which('libreoffice') is not None


def export_pdf(
    filepath: Path,
    output: Path
) -> Dict[str, Any]:
    """Export presentation to PDF."""
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if not filepath.suffix.lower() == '.pptx':
        raise ValueError(f"Input must be .pptx file, got: {filepath.suffix}")
    
    # Check LibreOffice
    if not check_libreoffice():
        raise RuntimeError(
            "LibreOffice not found. PDF export requires LibreOffice.\n"
            "Install:\n"
            "  Linux: sudo apt install libreoffice-impress\n"
            "  macOS: brew install --cask libreoffice\n"
            "  Windows: https://www.libreoffice.org/download/"
        )
    
    # Ensure output directory exists
    output.parent.mkdir(parents=True, exist_ok=True)
    
    # Use LibreOffice to convert
    # Note: --headless runs without GUI, --convert-to pdf exports to PDF
    cmd = [
        'soffice' if shutil.which('soffice') else 'libreoffice',
        '--headless',
        '--convert-to', 'pdf',
        '--outdir', str(output.parent),
        str(filepath)
    ]
    
    result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
    
    if result.returncode != 0:
        raise PowerPointAgentError(
            f"PDF export failed: {result.stderr}\n"
            f"Command: {' '.join(cmd)}"
        )
    
    # LibreOffice names output file based on input, rename if needed
    expected_pdf = filepath.parent / f"{filepath.stem}.pdf"
    if output.parent != filepath.parent:
        expected_pdf = output.parent / f"{filepath.stem}.pdf"
    
    if expected_pdf != output and expected_pdf.exists():
        if output.exists():
            output.unlink()
        expected_pdf.rename(output)
    elif not output.exists() and expected_pdf.exists():
        expected_pdf.rename(output)
    
    if not output.exists():
        raise PowerPointAgentError("PDF export completed but output file not found")
    
    # Get file sizes
    input_size = filepath.stat().st_size
    output_size = output.stat().st_size
    
    return {
        "status": "success",
        "input_file": str(filepath),
        "output_file": str(output),
        "input_size_bytes": input_size,
        "input_size_mb": round(input_size / (1024 * 1024), 2),
        "output_size_bytes": output_size,
        "output_size_mb": round(output_size / (1024 * 1024), 2),
        "size_ratio": round(output_size / input_size, 2) if input_size > 0 else 0
    }


def main():
    parser = argparse.ArgumentParser(
        description="Export PowerPoint presentation to PDF",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Basic export
  uv python ppt_export_pdf.py \\
    --file presentation.pptx \\
    --output presentation.pdf \\
    --json
  
  # Export with automatic naming
  uv python ppt_export_pdf.py \\
    --file quarterly_report.pptx \\
    --output reports/q4_report.pdf \\
    --json

Requirements:
  LibreOffice must be installed:
  
  Linux:
    sudo apt update
    sudo apt install libreoffice-impress
  
  macOS:
    brew install --cask libreoffice
  
  Windows:
    Download from https://www.libreoffice.org/download/

Verification:
  # Check LibreOffice installation
  soffice --version
  # or
  libreoffice --version

Use Cases:
  - Share presentations as PDFs
  - Archive presentations
  - Print-ready versions
  - Email distribution (smaller, universal format)
  - Document repositories

PDF Benefits:
  - Universal compatibility
  - Prevents editing
  - Smaller file size typically
  - Better for printing
  - Preserves layout exactly

Limitations:
  - Animations not preserved
  - Embedded videos become static
  - No speaker notes in output
  - Transitions removed
  - Interactive elements static

Performance:
  - Export time: ~2-5 seconds per slide
  - Large presentations (100+ slides): ~3-5 minutes
  - File size typically 30-50% of .pptx

Troubleshooting:
  If export fails:
  1. Verify LibreOffice installed: soffice --version
  2. Check file not corrupted: open in PowerPoint
  3. Ensure disk space available
  4. Try shorter timeout for small files
  5. Check LibreOffice logs
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file to export'
    )
    
    parser.add_argument(
        '--output',
        required=True,
        type=Path,
        help='Output PDF file path'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        help='Output JSON response'
    )
    
    args = parser.parse_args()
    
    try:
        # Ensure output has .pdf extension
        if not args.output.suffix.lower() == '.pdf':
            args.output = args.output.with_suffix('.pdf')
        
        result = export_pdf(
            filepath=args.file,
            output=args.output
        )
        
        if args.json:
            print(json.dumps(result, indent=2))
        else:
            print(f"✅ Exported to PDF: {result['output_file']}")
            print(f"   Input: {result['input_size_mb']} MB")
            print(f"   Output: {result['output_size_mb']} MB")
            print(f"   Ratio: {int(result['size_ratio'] * 100)}%")
        
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
