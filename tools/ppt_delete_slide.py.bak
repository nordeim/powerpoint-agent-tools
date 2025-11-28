#!/usr/bin/env python3
"""
PowerPoint Delete Slide Tool
Remove a slide from the presentation

Usage:
    uv python ppt_delete_slide.py --file presentation.pptx --index 1 --json
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


def delete_slide(filepath: Path, index: int) -> Dict[str, Any]:
    """Delete slide at index."""
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        total_slides = agent.get_slide_count()
        if not 0 <= index < total_slides:
            raise SlideNotFoundError(f"Index {index} out of range (0-{total_slides-1})")
            
        agent.delete_slide(index)
        agent.save()
        
        new_count = agent.get_slide_count()
    
    return {
        "status": "success",
        "file": str(filepath),
        "deleted_index": index,
        "remaining_slides": new_count
    }


def main():
    parser = argparse.ArgumentParser(description="Delete PowerPoint slide")
    parser.add_argument('--file', required=True, type=Path, help='PowerPoint file path')
    parser.add_argument('--index', required=True, type=int, help='Slide index to delete (0-based)')
    parser.add_argument('--json', action='store_true', default=True, help='Output JSON')
    
    args = parser.parse_args()
    
    try:
        result = delete_slide(filepath=args.file, index=args.index)
        print(json.dumps(result, indent=2))
        sys.exit(0)
    except Exception as e:
        print(json.dumps({"status": "error", "error": str(e), "error_type": type(e).__name__}, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()
