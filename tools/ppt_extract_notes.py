#!/usr/bin/env python3
"""
PowerPoint Extract Notes Tool
Get speaker notes from all slides

Usage:
    uv python ppt_extract_notes.py --file presentation.pptx --json
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import PowerPointAgent


def extract_notes(filepath: Path) -> Dict[str, Any]:
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath, acquire_lock=False)
        notes = agent.extract_notes()
        
    return {
        "status": "success",
        "file": str(filepath),
        "notes_found": len(notes),
        "notes": notes # Dict {slide_index: text}
    }


def main():
    parser = argparse.ArgumentParser(description="Extract speaker notes")
    parser.add_argument('--file', required=True, type=Path, help='PowerPoint file path')
    parser.add_argument('--json', action='store_true', default=True, help='Output JSON')
    
    args = parser.parse_args()
    
    try:
        result = extract_notes(filepath=args.file)
        print(json.dumps(result, indent=2))
        sys.exit(0)
    except Exception as e:
        print(json.dumps({"status": "error", "error": str(e), "error_type": type(e).__name__}, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()
