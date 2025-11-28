#!/usr/bin/env python3
"""
PowerPoint Check Accessibility Tool v3.1.0
Run WCAG 2.1 accessibility checks on presentation.

This tool performs comprehensive accessibility analysis including:
- Missing alt text detection for images
- Color contrast ratio checking
- Reading order analysis
- Font size compliance
- Text-only content verification

Usage:
    uv run tools/ppt_check_accessibility.py --file presentation.pptx --json

Exit Codes:
    0: Success (check 'passed' field for compliance status)
    1: Error occurred (file not found, crash)

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0
"""

import sys
import os

# --- HYGIENE BLOCK START ---
# CRITICAL: Redirect stderr to /dev/null immediately.
# This prevents libraries (pptx, warnings) from printing non-JSON text
# which corrupts pipelines that capture 2>&1.
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

import json
import argparse
import logging
import uuid
import time
from pathlib import Path
from typing import Dict, Any, Optional
from datetime import datetime

# Configure logging to null handler
logging.basicConfig(level=logging.CRITICAL)

# Add parent directory to path for core import
sys.path.insert(0, str(Path(__file__).parent.parent))

# ============================================================================
# CONSTANTS
# ============================================================================

__version__ = "3.1.0"

# ============================================================================
# SAFE IMPORTS
# ============================================================================

CORE_AVAILABLE = False
PowerPointAgent = None
PowerPointAgentError = None

try:
    from core.powerpoint_agent_core import (
        PowerPointAgent as _PowerPointAgent,
        PowerPointAgentError as _PowerPointAgentError
    )
    PowerPointAgent = _PowerPointAgent
    PowerPointAgentError = _PowerPointAgentError
    CORE_AVAILABLE = True
except ImportError as e:
    # Will be handled in main()
    IMPORT_ERROR = str(e)

# ============================================================================
# MAIN LOGIC
# ============================================================================

def check_accessibility(filepath: Path) -> Dict[str, Any]:
    """
    Run comprehensive accessibility checks on a PowerPoint presentation.
    
    Args:
        filepath: Path to PowerPoint file
        
    Returns:
        Dict containing:
            - status: "success" or "error"
            - file: Path to checked file
            - tool_version: Version of this tool
            - presentation_version: Hash of presentation state
            - validated_at: ISO timestamp
            - operation_id: Unique operation identifier
            - duration_ms: Time taken for check
            - passed: Boolean indicating if all checks passed
            - summary: Issue counts by category
            - issues: Detailed issue list from core
            
    Raises:
        FileNotFoundError: If file doesn't exist
        ImportError: If core module not available
    """
    if not CORE_AVAILABLE:
        raise ImportError(f"Core module not available: {IMPORT_ERROR}")
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    start_time = time.perf_counter()
    operation_id = str(uuid.uuid4())
    
    with PowerPointAgent(filepath) as agent:
        # acquire_lock=False because validation is read-only safe
        agent.open(filepath, acquire_lock=False)
        
        # Capture presentation version for audit trail
        presentation_version = agent.get_presentation_version()
        
        # Run core accessibility check
        core_result = agent.check_accessibility()
    
    duration_ms = int((time.perf_counter() - start_time) * 1000)
    
    # Extract issues from core result
    issues = core_result.get("issues", {})
    
    # Calculate summary statistics
    missing_alt_text = issues.get("missing_alt_text", [])
    low_contrast = issues.get("low_contrast", [])
    reading_order_issues = issues.get("reading_order_issues", [])
    small_text = issues.get("small_text", [])
    
    total_issues = (
        len(missing_alt_text) +
        len(low_contrast) +
        len(reading_order_issues) +
        len(small_text)
    )
    
    # Determine if accessibility check passed (no critical issues)
    # Critical: missing_alt_text, low_contrast
    critical_count = len(missing_alt_text) + len(low_contrast)
    passed = critical_count == 0
    
    return {
        "status": "success",
        "file": str(filepath.resolve()),
        "tool_version": __version__,
        "presentation_version": presentation_version,
        "validated_at": datetime.utcnow().isoformat() + "Z",
        "operation_id": operation_id,
        "duration_ms": duration_ms,
        "passed": passed,
        "summary": {
            "total_issues": total_issues,
            "critical_issues": critical_count,
            "missing_alt_text_count": len(missing_alt_text),
            "low_contrast_count": len(low_contrast),
            "reading_order_issues_count": len(reading_order_issues),
            "small_text_count": len(small_text)
        },
        "issues": issues,
        "wcag_level": "AA",
        "checks_performed": [
            "alt_text",
            "color_contrast",
            "reading_order",
            "font_size"
        ]
    }

# ============================================================================
# CLI INTERFACE
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description=f"Check PowerPoint accessibility (WCAG 2.1) v{__version__}",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
    # Check accessibility and output JSON
    uv run tools/ppt_check_accessibility.py --file presentation.pptx --json
    
    # Check a specific file
    uv run tools/ppt_check_accessibility.py --file /path/to/deck.pptx --json

Output includes:
    - passed: Boolean indicating compliance
    - summary: Issue counts by category
    - issues: Detailed list with slide/shape references
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
    
    # Check if core module is available before proceeding
    if not CORE_AVAILABLE:
        error_result = {
            "status": "error",
            "error": f"Core module import failed: {IMPORT_ERROR}",
            "error_type": "ImportError",
            "tool_version": __version__,
            "suggestion": "Ensure core/powerpoint_agent_core.py is available and dependencies are installed"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
    
    try:
        result = check_accessibility(filepath=args.file)
        sys.stdout.write(json.dumps(result, indent=2) + "\n")
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "tool_version": __version__,
            "suggestion": "Verify the file path exists and is accessible"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)


if __name__ == "__main__":
    main()
