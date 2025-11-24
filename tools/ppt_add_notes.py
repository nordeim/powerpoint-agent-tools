#!/usr/bin/env python3
"""
PowerPoint Add Speaker Notes Tool
Add or update speaker notes for a specific slide.

Usage:
    uv python ppt_add_notes.py --file deck.pptx --slide 0 --text "Talk about Q4 growth." --json
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

def add_notes(
    filepath: Path,
    slide_index: int,
    text: str,
    mode: str = "append"
) -> Dict[str, Any]:
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        total = agent.get_slide_count()
        if not 0 <= slide_index < total:
            raise SlideNotFoundError(f"Slide index {slide_index} out of range")
            
        slide = agent.prs.slides[slide_index]
        
        # Access or create notes slide
        if not slide.has_notes_slide:
            notes_slide = slide.notes_slide
        else:
            notes_slide = slide.notes_slide
            
        text_frame = notes_slide.notes_text_frame
        
        original_text = text_frame.text
        
        if mode == "overwrite":
            text_frame.text = text
        elif mode == "append":
            if text_frame.text:
                text_frame.text += "\n" + text
            else:
                text_frame.text = text
        elif mode == "prepend":
             if text_frame.text:
                text_frame.text = text + "\n" + text_frame.text
             else:
                text_frame.text = text
                
        agent.save()
        
    return {
        "status": "success",
        "file": str(filepath),
        "slide_index": slide_index,
        "mode": mode,
        "notes_length": len(text_frame.text),
        "preview": text_frame.text[:100] + "..." if len(text_frame.text) > 100 else text_frame.text
    }

def main():
    parser = argparse.ArgumentParser(description="Add speaker notes to slide")
    parser.add_argument('--file', required=True, type=Path, help='PowerPoint file path')
    parser.add_argument('--slide', required=True, type=int, help='Slide index (0-based)')
    parser.add_argument('--text', required=True, help='Notes content')
    parser.add_argument('--mode', choices=['append', 'overwrite', 'prepend'], default='append', help='Insertion mode')
    parser.add_argument('--json', action='store_true', default=True, help='Output JSON')
    
    args = parser.parse_args()
    
    try:
        result = add_notes(
            filepath=args.file,
            slide_index=args.slide,
            text=args.text,
            mode=args.mode
        )
        print(json.dumps(result, indent=2))
        sys.exit(0)
    except Exception as e:
        print(json.dumps({"status": "error", "error": str(e), "error_type": type(e).__name__}, indent=2))
        sys.exit(1)

if __name__ == "__main__":
    main()
