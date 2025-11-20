#!/usr/bin/env python3
"""
PowerPoint Add Connector Tool
Draw a line between two shapes

Usage:
    uv python ppt_add_connector.py --file deck.pptx --slide 0 --from-shape 0 --to-shape 1 --json
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


def add_connector(
    filepath: Path,
    slide_index: int,
    from_shape: int,
    to_shape: int,
    connector_type: str = "straight"
) -> Dict[str, Any]:
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        total = agent.get_slide_count()
        if not 0 <= slide_index < total:
            raise SlideNotFoundError(f"Slide index {slide_index} out of range")
            
        agent.add_connector(
            slide_index=slide_index,
            from_shape=from_shape,
            to_shape=to_shape,
            connector_type=connector_type
        )
        agent.save()
        
    return {
        "status": "success",
        "file": str(filepath),
        "slide_index": slide_index,
        "connection": {
            "from": from_shape,
            "to": to_shape,
            "type": connector_type
        }
    }


def main():
    parser = argparse.ArgumentParser(description="Add connector between shapes")
    parser.add_argument('--file', required=True, type=Path, help='PowerPoint file path')
    parser.add_argument('--slide', required=True, type=int, help='Slide index (0-based)')
    parser.add_argument('--from-shape', required=True, type=int, help='Start shape index')
    parser.add_argument('--to-shape', required=True, type=int, help='End shape index')
    parser.add_argument('--type', default='straight', help='Connector type (straight)')
    parser.add_argument('--json', action='store_true', default=True, help='Output JSON')
    
    args = parser.parse_args()
    
    try:
        result = add_connector(
            filepath=args.file,
            slide_index=args.slide,
            from_shape=args.from_shape,
            to_shape=args.to_shape,
            connector_type=args.type
        )
        print(json.dumps(result, indent=2))
        sys.exit(0)
    except Exception as e:
        print(json.dumps({"status": "error", "error": str(e), "error_type": type(e).__name__}, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()
