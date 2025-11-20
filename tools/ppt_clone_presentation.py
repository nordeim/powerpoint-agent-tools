#!/usr/bin/env python3
"""
PowerPoint Clone Presentation Tool
Create an exact copy of a presentation

Usage:
    uv python ppt_clone_presentation.py --source base.pptx --output new_deck.pptx --json
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import PowerPointAgent


def clone_presentation(source: Path, output: Path) -> Dict[str, Any]:
    
    if not source.exists():
        raise FileNotFoundError(f"Source file not found: {source}")
    
    # Validate output path
    if not output.suffix.lower() == '.pptx':
        output = output.with_suffix('.pptx')

    # Uses agent to open and save-as, effectively cloning
    with PowerPointAgent(source) as agent:
        agent.open(source, acquire_lock=False) # Read-only open
        agent.save(output)
        
    return {
        "status": "success",
        "source": str(source),
        "output": str(output),
        "size_bytes": output.stat().st_size
    }


def main():
    parser = argparse.ArgumentParser(description="Clone PowerPoint presentation")
    parser.add_argument('--source', required=True, type=Path, help='Source file')
    parser.add_argument('--output', required=True, type=Path, help='Destination file')
    parser.add_argument('--json', action='store_true', default=True, help='Output JSON')
    
    args = parser.parse_args()
    
    try:
        result = clone_presentation(source=args.source, output=args.output)
        print(json.dumps(result, indent=2))
        sys.exit(0)
    except Exception as e:
        print(json.dumps({"status": "error", "error": str(e), "error_type": type(e).__name__}, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()
