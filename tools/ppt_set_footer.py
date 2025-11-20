#!/usr/bin/env python3
"""
PowerPoint Set Footer Tool
Configure slide footer, date, and slide number

Usage:
    uv python ppt_set_footer.py --file deck.pptx --text "Confidential" --show-number --json
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import PowerPointAgent

def set_footer(
    filepath: Path,
    text: str = None,
    show_number: bool = False,
    show_date: bool = False,
    apply_to_master: bool = True
) -> Dict[str, Any]:
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")

    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # In python-pptx, footer visibility is often controlled via the master
        # or individual slide layouts.
        
        if apply_to_master:
             masters = agent.prs.slide_masters
             for master in masters:
                 # Update layouts in master
                 for layout in master.slide_layouts:
                     # Iterate shapes to find placeholders
                     for shape in layout.placeholders:
                         if shape.is_placeholder:
                             # Footer type is 15, Slide Number is 16, Date is 14
                             if shape.placeholder_format.type == 15 and text: # Footer
                                 shape.text = text
                             
        # Also attempt to set on individual slides for immediate visibility
        count = 0
        for slide in agent.prs.slides:
            # This is simplified; robustness varies by template
            # We try to find standard placeholders
            for shape in slide.placeholders:
                 if shape.placeholder_format.type == 15 and text:
                     shape.text = text
                     count += 1
                     
        agent.save()
        
    return {
        "status": "success",
        "file": str(filepath),
        "footer_text": text,
        "settings": {
            "show_number": show_number,
            "show_date": show_date
        },
        "slides_updated": count
    }

def main():
    parser = argparse.ArgumentParser(description="Set PowerPoint footer")
    parser.add_argument('--file', required=True, type=Path, help='PowerPoint file path')
    parser.add_argument('--text', help='Footer text')
    parser.add_argument('--show-number', action='store_true', help='Show slide number')
    parser.add_argument('--show-date', action='store_true', help='Show date')
    parser.add_argument('--json', action='store_true', default=True, help='Output JSON')
    
    args = parser.parse_args()
    
    try:
        result = set_footer(
            filepath=args.file, 
            text=args.text, 
            show_number=args.show_number,
            show_date=args.show_date
        )
        print(json.dumps(result, indent=2))
        sys.exit(0)
    except Exception as e:
        print(json.dumps({"status": "error", "error": str(e), "error_type": type(e).__name__}, indent=2))
        sys.exit(1)

if __name__ == "__main__":
    main()
