#!/usr/bin/env python3
"""
PowerPoint Check Accessibility Tool v3.1.0
Run WCAG 2.1 accessibility checks on presentation

Usage:
    uv run tools/ppt_check_accessibility.py --file presentation.pptx --json

Exit Codes:
    0: Success (valid JSON returned, check 'status' field for findings)
    1: Error occurred (file not found, crash)

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
from typing import Dict, Any

# Configure logging to null handler just in case
logging.basicConfig(level=logging.CRITICAL)

# Add parent directory to path for core import
sys.path.insert(0, str(Path(__file__).parent.parent))

try:
    from core.powerpoint_agent_core import (
        PowerPointAgent,
        PowerPointAgentError
    )
except ImportError:
    # Fallback for standalone testing
    pass


def check_accessibility(filepath: Path) -> Dict[str, Any]:
    """Run accessibility checks."""
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    with PowerPointAgent(filepath) as agent:
        # acquire_lock=False because validation is read-only safe
        agent.open(filepath, acquire_lock=False)
        result = agent.check_accessibility()
        result["file"] = str(filepath)
    
    return result


def main():
    parser = argparse.ArgumentParser(
        description="Check PowerPoint accessibility (WCAG 2.1)",
        formatter_class=argparse.RawDescriptionHelpFormatter
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
        result = check_accessibility(filepath=args.file)
        
        # Strict JSON output to stdout
        sys.stdout.write(json.dumps(result, indent=2) + "\n")
        sys.exit(0)
        
    except Exception as e:
        # Crash handler: Ensure valid JSON is always returned
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)


if __name__ == "__main__":
    main()
