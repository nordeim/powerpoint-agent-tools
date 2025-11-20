#!/usr/bin/env python3
"""
PowerPoint Duplicate Slide Tool
Clone an existing slide

Usage:
    uv python ppt_duplicate_slide.py --file presentation.pptx --index 0 --json
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


def duplicate_slide(filepath: Path, index: int) -> Dict[str, Any]:
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        total = agent.get_slide_count()
        if not 0 <= index < total:
            raise SlideNotFoundError(f"Index {index} out of range (0-{total-1})")
            
        new_index = agent.duplicate_slide(index)
        agent.save()
        
        # Get info about new slide
        info = agent.get_slide_info(new_index)
        
    return {
        "status": "success",
        "file": str(filepath),
        "source_index": index,
        "new_slide_index": new_index,
        "total_slides": new_index + 1,
        "layout": info["layout"]
    }


def main():
    parser = argparse.ArgumentParser(description="Duplicate PowerPoint slide")
    parser.add_argument('--file', required=True, type=Path, help='PowerPoint file path')
    parser.add_argument('--index', required=True, type=int, help='Source slide index (0-based)')
    parser.add_argument('--json', action='store_true', default=True, help='Output JSON')
    
    args = parser.parse_args()
    
    try:
        result = duplicate_slide(filepath=args.file, index=args.index)
        print(json.dumps(result, indent=2))
        sys.exit(0)
    except Exception as e:
        print(json.dumps({"status": "error", "error": str(e), "error_type": type(e).__name__}, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()
