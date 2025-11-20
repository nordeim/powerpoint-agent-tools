#!/usr/bin/env python3
"""
PowerPoint Check Accessibility Tool
Run WCAG 2.1 accessibility checks on presentation

Usage:
    uv python ppt_check_accessibility.py --file presentation.pptx --json

Exit Codes:
    0: Success (even if issues found - check 'status')
    1: Error occurred (file not found, etc)
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


def check_accessibility(filepath: Path) -> Dict[str, Any]:
    """Run accessibility checks."""
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath, acquire_lock=False)  # Read-only check
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
        print(json.dumps(result, indent=2))
        sys.exit(0)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()
