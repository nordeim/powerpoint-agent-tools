#!/usr/bin/env python3
"""
PowerPoint Reorder Slides Tool
Move a slide to a new position

Usage:
    uv python ppt_reorder_slides.py --file presentation.pptx --from-index 3 --to-index 1 --json
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError, SlideNotFoundError
)


def reorder_slides(filepath: Path, from_index: int, to_index: int) -> Dict[str, Any]:
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        total = agent.get_slide_count()
        if not 0 <= from_index < total:
            raise SlideNotFoundError(f"Source index {from_index} out of range")
        if not 0 <= to_index < total:
            raise SlideNotFoundError(f"Target index {to_index} out of range")
            
        agent.reorder_slides(from_index, to_index)
        agent.save()
        
    return {
        "status": "success",
        "file": str(filepath),
        "moved_from": from_index,
        "moved_to": to_index,
        "total_slides": total
    }


def main():
    parser = argparse.ArgumentParser(description="Reorder PowerPoint slides")
    parser.add_argument('--file', required=True, type=Path, help='PowerPoint file path')
    parser.add_argument('--from-index', required=True, type=int, help='Current slide index')
    parser.add_argument('--to-index', required=True, type=int, help='New slide index')
    parser.add_argument('--json', action='store_true', default=True, help='Output JSON')
    
    args = parser.parse_args()
    
    try:
        result = reorder_slides(filepath=args.file, from_index=args.from_index, to_index=args.to_index)
        print(json.dumps(result, indent=2))
        sys.exit(0)
    except Exception as e:
        print(json.dumps({"status": "error", "error": str(e), "error_type": type(e).__name__}, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()
