# tools/ppt_export_images.py
```py
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
    
    # Use PDF-intermediate workflow for better reliability
    # 1. Export to PDF
    base_name = filepath.stem
    pdf_path = output_dir / f"{base_name}.pdf"
    
    # Use LibreOffice to export to PDF
    cmd_pdf = [
        'soffice' if shutil.which('soffice') else 'libreoffice',
        '--headless',
        '--convert-to', 'pdf',
        '--outdir', str(output_dir),
        str(filepath)
    ]
    
    result_pdf = subprocess.run(cmd_pdf, capture_output=True, text=True, timeout=120)
    
    if result_pdf.returncode != 0:
        raise PowerPointAgentError(
            f"PDF export failed: {result_pdf.stderr}\n"
            f"Command: {' '.join(cmd_pdf)}"
        )
        
    # 2. Convert PDF to Images using pdftoppm (if available)
    if shutil.which('pdftoppm'):
        # pdftoppm -png -r 150 input.pdf output_prefix
        cmd_img = [
            'pdftoppm',
            f"-{format_ext}",
            '-r', '150',  # 150 DPI
            str(pdf_path),
            str(output_dir / base_name)
        ]
        
        result_img = subprocess.run(cmd_img, capture_output=True, text=True, timeout=120)
        
        if result_img.returncode != 0:
             # Fallback to direct export if pdftoppm fails
             print(f"Warning: pdftoppm failed, falling back to LibreOffice direct export: {result_img.stderr}", file=sys.stderr)
             _export_direct(filepath, output_dir, format_ext)
        
        # Clean up PDF
        # Clean up PDF
        if pdf_path.exists():
            pdf_path.unlink()
            
    else:
        # Fallback to direct export if pdftoppm not installed
        print("Warning: pdftoppm not found, using LibreOffice direct export (may be incomplete)", file=sys.stderr)
        _export_direct(filepath, output_dir, format_ext)

    # 3. Process and rename files
    return _scan_and_process_results(filepath, output_dir, format_ext, prefix)


def _export_direct(filepath: Path, output_dir: Path, format_ext: str):
    """Direct export using LibreOffice (legacy method)."""
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


def _scan_and_process_results(
    filepath: Path,
    output_dir: Path,
    format_ext: str,
    prefix: str
) -> Dict[str, Any]:
    """Find, rename, and report exported images."""
    # LibreOffice/pdftoppm creates files named: presentation.png or presentation-1.png
    base_name = filepath.stem
    
    # Find exported images using a robust pattern match
    # Match base_name*.ext
    candidates = sorted(output_dir.glob(f"{base_name}*.{format_ext}"))
    
    exported_files = []
    
    # Rename found files sequentially
    for i, old_file in enumerate(candidates):
        new_file = output_dir / f"{prefix}{i+1:03d}.{format_ext}"
        
        # Handle case where source and dest are same
        if old_file != new_file:
            # If target exists (from previous run), remove it
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

```

# tools/ppt_export_pdf.py
```py
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

```

# tools/ppt_json_adapter.py
```py
#!/usr/bin/env python3
"""
ppt_json_adapter.py

Validates and normalizes JSON outputs from presentation CLI tools.
Usage:
  python ppt_json_adapter.py --schema ppt_get_info.schema.json --input raw.json

Behavior:
- Validates input JSON against provided schema.
- Maps common alias keys to canonical keys.
- Emits normalized JSON to stdout.
- On validation failure, emits structured error JSON and exits non-zero.
"""

import argparse
import json
import sys
import hashlib
from jsonschema import validate, ValidationError

# Alias mapping table for common drifted keys
ALIAS_MAP = {
    "slides_count": "slide_count",
    "slidesTotal": "slide_count",
    "slides_list": "slides",
    "probe_time": "probe_timestamp",
    "canWrite": "can_write",
    "canRead": "can_read",
    "maxImageSizeMB": "max_image_size_mb"
}

ERROR_TEMPLATE = {
    "error": {
        "error_code": None,
        "message": None,
        "details": None,
        "retryable": False
    }
}

def load_json(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

def map_aliases(obj):
    if isinstance(obj, dict):
        new = {}
        for k, v in obj.items():
            canonical = ALIAS_MAP.get(k, k)
            # recursively map nested dicts and lists
            if isinstance(v, dict):
                new[canonical] = map_aliases(v)
            elif isinstance(v, list):
                new[canonical] = [map_aliases(i) for i in v]
            else:
                new[canonical] = v
        return new
    elif isinstance(obj, list):
        return [map_aliases(i) for i in obj]
    else:
        return obj

def compute_presentation_version(info_obj):
    """
    Compute a stable presentation_version if missing.
    Uses slide ids and counts to produce a deterministic hash.
    """
    try:
        slides = info_obj.get("slides", [])
        slide_ids = ",".join([str(s.get("id", s.get("index", ""))) for s in slides])
        base = f"{info_obj.get('file','')}-{info_obj.get('slide_count',len(slides))}-{slide_ids}"
        return hashlib.sha256(base.encode("utf-8")).hexdigest()
    except Exception:
        return None

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--schema", required=True, help="Path to JSON Schema")
    parser.add_argument("--input", required=True, help="Path to raw JSON input")
    args = parser.parse_args()

    try:
        schema = load_json(args.schema)
    except Exception as e:
        print(json.dumps({**ERROR_TEMPLATE, "error": {"error_code": "SCHEMA_LOAD_ERROR", "message": str(e), "details": None, "retryable": False}}))
        sys.exit(5)

    try:
        raw = load_json(args.input)
    except Exception as e:
        print(json.dumps({**ERROR_TEMPLATE, "error": {"error_code": "INPUT_LOAD_ERROR", "message": str(e), "details": None, "retryable": True}}))
        sys.exit(3)

    # Map aliases
    normalized = map_aliases(raw)

    # If presentation_version missing for get_info, compute a best-effort version
    if "presentation_version" not in normalized and schema.get("title","").lower().find("ppt_get_info") != -1:
        pv = compute_presentation_version(normalized)
        if pv:
            normalized["presentation_version"] = pv

    # Validate
    try:
        validate(instance=normalized, schema=schema)
    except ValidationError as ve:
        err = {
            "error": {
                "error_code": "SCHEMA_VALIDATION_ERROR",
                "message": str(ve.message),
                "details": ve.schema_path,
                "retryable": False
            }
        }
        print(json.dumps(err))
        sys.exit(2)

    # Emit normalized JSON
    print(json.dumps(normalized, indent=2))
    sys.exit(0)

if __name__ == "__main__":
    main()

```

# tools/ppt_validate_presentation.py
```py
#!/usr/bin/env python3
"""
PowerPoint Validate Presentation Tool v3.1.0
Comprehensive validation for structure, accessibility, assets, and design quality.

Fully aligned with PowerPoint Agent Core v3.1.0 and System Prompt v3.0 validation gates.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_validate_presentation.py --file presentation.pptx --json
    uv run tools/ppt_validate_presentation.py --file presentation.pptx --policy strict --json

Exit Codes:
    0: Success (valid or only warnings within policy thresholds)
    1: Error occurred or critical issues exceed policy thresholds

Changelog v3.1.0:
- FIXED: Strict output hygiene (stderr redirected to devnull) to prevent JSON corruption
- ENHANCED: Robust error handling ensures valid JSON output even on crash
"""

import sys
import os

# CRITICAL: Redirect stderr to /dev/null immediately.
# This prevents libraries (pptx, warnings) from printing non-JSON text
# which corrupts pipelines that capture 2>&1.
sys.stderr = open(os.devnull, 'w')

import json
import argparse
import logging
from pathlib import Path
from typing import Dict, Any, List, Optional
from dataclasses import dataclass, field, asdict
from datetime import datetime

# Configure logging to null handler just in case
logging.basicConfig(level=logging.CRITICAL)

# Add parent directory to path for core import
sys.path.insert(0, str(Path(__file__).parent.parent))

try:
    from core.powerpoint_agent_core import (
        PowerPointAgent,
        PowerPointAgentError,
        __version__ as CORE_VERSION
    )
except ImportError:
    CORE_VERSION = "0.0.0"

# ============================================================================
# CONSTANTS & POLICIES
# ============================================================================

__version__ = "3.1.0"

VALIDATION_POLICIES = {
    "lenient": {
        "name": "Lenient",
        "thresholds": {
            "max_critical_issues": 10,
            "max_accessibility_issues": 20,
            "max_design_warnings": 50,
            "max_empty_slides": 5,
            "max_slides_without_titles": 10,
            "max_missing_alt_text": 20,
            "max_low_contrast": 10,
            "max_large_images": 10,
            "require_all_alt_text": False,
            "enforce_6x6_rule": False,
            "max_fonts": 10,
            "max_colors": 20,
        }
    },
    "standard": {
        "name": "Standard",
        "thresholds": {
            "max_critical_issues": 0,
            "max_accessibility_issues": 5,
            "max_design_warnings": 10,
            "max_empty_slides": 0,
            "max_slides_without_titles": 3,
            "max_missing_alt_text": 5,
            "max_low_contrast": 3,
            "max_large_images": 5,
            "require_all_alt_text": False,
            "enforce_6x6_rule": False,
            "max_fonts": 5,
            "max_colors": 10,
        }
    },
    "strict": {
        "name": "Strict",
        "thresholds": {
            "max_critical_issues": 0,
            "max_accessibility_issues": 0,
            "max_design_warnings": 3,
            "max_empty_slides": 0,
            "max_slides_without_titles": 0,
            "max_missing_alt_text": 0,
            "max_low_contrast": 0,
            "max_large_images": 3,
            "require_all_alt_text": True,
            "enforce_6x6_rule": True,
            "max_fonts": 3,
            "max_colors": 5,
        }
    }
}

# ============================================================================
# DATA CLASSES
# ============================================================================

@dataclass
class ValidationIssue:
    category: str
    severity: str
    message: str
    slide_index: Optional[int] = None
    shape_index: Optional[int] = None
    fix_command: Optional[str] = None
    details: Dict[str, Any] = field(default_factory=dict)
    
    def to_dict(self) -> Dict[str, Any]:
        return {k: v for k, v in asdict(self).items() if v is not None}

@dataclass
class ValidationSummary:
    total_issues: int = 0
    critical_count: int = 0
    warning_count: int = 0
    info_count: int = 0
    empty_slides: int = 0
    slides_without_titles: int = 0
    missing_alt_text: int = 0
    low_contrast: int = 0
    large_images: int = 0
    fonts_used: int = 0
    colors_detected: int = 0
    
    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

@dataclass
class ValidationPolicy:
    name: str
    thresholds: Dict[str, Any]
    description: str = ""
    
    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

# ============================================================================
# LOGIC
# ============================================================================

def get_policy(policy_name: str, custom_thresholds: Optional[Dict[str, Any]] = None) -> ValidationPolicy:
    if policy_name == "custom" and custom_thresholds:
        base = VALIDATION_POLICIES["standard"]["thresholds"].copy()
        base.update(custom_thresholds)
        return ValidationPolicy(name="Custom", thresholds=base)
    
    config = VALIDATION_POLICIES.get(policy_name, VALIDATION_POLICIES["standard"])
    return ValidationPolicy(name=config["name"], thresholds=config["thresholds"])

def validate_presentation(filepath: Path, policy: ValidationPolicy) -> Dict[str, Any]:
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    issues: List[ValidationIssue] = []
    summary = ValidationSummary()
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath, acquire_lock=False)
        presentation_info = agent.get_presentation_info()
        slide_count = presentation_info.get("slide_count", 0)
        
        # Run Validations
        core_val = agent.validate_presentation()
        acc_val = agent.check_accessibility()
        asset_val = agent.validate_assets()
        
        # Process Results
        _process_core_validation(core_val, issues, summary, filepath)
        _process_accessibility(acc_val, issues, summary, filepath)
        _process_assets(asset_val, issues, summary, filepath)
        _validate_design_rules(agent, issues, summary, policy, filepath)
        
        # Summarize
        summary.total_issues = len(issues)
        summary.critical_count = sum(1 for i in issues if i.severity == "critical")
        summary.warning_count = sum(1 for i in issues if i.severity == "warning")
        summary.info_count = sum(1 for i in issues if i.severity == "info")
        
        passed, violations = _check_policy_compliance(summary, policy)
        
        status = "valid"
        if summary.critical_count > 0: status = "critical"
        elif not passed: status = "failed"
        elif summary.warning_count > 0: status = "warnings"

    return {
        "status": status,
        "passed": passed,
        "file": str(filepath),
        "validated_at": datetime.utcnow().isoformat() + "Z",
        "policy": policy.to_dict(),
        "summary": summary.to_dict(),
        "policy_violations": violations,
        "issues": [i.to_dict() for i in issues],
        "recommendations": _generate_recommendations(issues, policy),
        "presentation_info": {
            "slide_count": slide_count,
            "file_size_mb": presentation_info.get("file_size_mb"),
            "aspect_ratio": presentation_info.get("aspect_ratio")
        }
    }

def _process_core_validation(val, issues, summary, filepath):
    i = val.get("issues", {})
    summary.empty_slides = len(i.get("empty_slides", []))
    summary.slides_without_titles = len(i.get("slides_without_titles", []))
    for idx in i.get("empty_slides", []):
        issues.append(ValidationIssue("structure", "critical", "Empty slide", slide_index=idx))
    for idx in i.get("slides_without_titles", []):
        issues.append(ValidationIssue("structure", "warning", "Slide missing title", slide_index=idx))
    if isinstance(i.get("fonts_used"), list):
        summary.fonts_used = len(i.get("fonts_used"))

def _process_accessibility(acc, issues, summary, filepath):
    i = acc.get("issues", {})
    summary.missing_alt_text = len(i.get("missing_alt_text", []))
    summary.low_contrast = len(i.get("low_contrast", []))
    for item in i.get("missing_alt_text", []):
        issues.append(ValidationIssue("accessibility", "critical", "Missing alt text", slide_index=item.get("slide")))
    for item in i.get("low_contrast", []):
        issues.append(ValidationIssue("accessibility", "warning", "Low contrast", slide_index=item.get("slide")))

def _process_assets(assets, issues, summary, filepath):
    i = assets.get("issues", {})
    summary.large_images = len(i.get("large_images", []))
    for item in i.get("large_images", []):
        issues.append(ValidationIssue("assets", "info", f"Large image ({item.get('size_mb')}MB)", slide_index=item.get("slide")))

def _validate_design_rules(agent, issues, summary, policy, filepath):
    if summary.fonts_used > policy.thresholds.get("max_fonts", 5):
        issues.append(ValidationIssue("design", "warning", f"Too many fonts ({summary.fonts_used})"))

def _check_policy_compliance(summary, policy):
    v = []
    t = policy.thresholds
    if summary.critical_count > t.get("max_critical_issues", 0): v.append("Critical issues exceed limit")
    if summary.empty_slides > t.get("max_empty_slides", 0): v.append("Empty slides exceed limit")
    if summary.slides_without_titles > t.get("max_slides_without_titles", 3): v.append("Untitled slides exceed limit")
    if summary.missing_alt_text > t.get("max_missing_alt_text", 5): v.append("Missing alt text exceeds limit")
    return len(v) == 0, v

def _generate_recommendations(issues, policy):
    recs = []
    if any(i.details.get("issue_type") == "empty_slide" for i in issues): recs.append({"priority":"high", "action":"Remove empty slides"})
    if any(i.category == "accessibility" for i in issues): recs.append({"priority":"medium", "action":"Fix accessibility issues"})
    return recs

# ============================================================================
# MAIN
# ============================================================================

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--file', required=True, type=Path)
    parser.add_argument('--policy', default='standard')
    parser.add_argument('--json', action='store_true', default=True)
    parser.add_argument('--max-missing-alt-text', type=int)
    parser.add_argument('--max-slides-without-titles', type=int)
    parser.add_argument('--max-empty-slides', type=int)
    parser.add_argument('--require-all-alt-text', action='store_true')
    parser.add_argument('--enforce-6x6', action='store_true')
    parser.add_argument('--summary-only', action='store_true')
    
    args = parser.parse_args()
    
    try:
        custom = {}
        if args.max_missing_alt_text is not None: custom["max_missing_alt_text"] = args.max_missing_alt_text
        if args.max_slides_without_titles is not None: custom["max_slides_without_titles"] = args.max_slides_without_titles
        
        policy = get_policy("custom", custom) if custom else get_policy(args.policy)
        
        result = validate_presentation(args.file, policy)
        
        # Explicitly write to stdout using sys.stdout.write to avoid print's behavior
        # and ensure it is the only thing on stdout
        sys.stdout.write(json.dumps(result, indent=2) + "\n")
        sys.stdout.flush()
        
        sys.exit(1 if result["status"] in ("critical", "failed") else 0)
        
    except Exception as e:
        # Crash handler: Print JSON error to stdout
        error_res = {
            "status": "error", 
            "error": str(e), 
            "error_type": type(e).__name__
        }
        sys.stdout.write(json.dumps(error_res, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)

if __name__ == "__main__":
    main()

```

