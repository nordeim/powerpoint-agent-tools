#!/usr/bin/env python3
"""
PowerPoint Add Speaker Notes Tool
Add, append, or overwrite speaker notes for a specific slide.

Usage:
    # Append note (default)
    uv run tools/ppt_add_notes.py --file deck.pptx --slide 0 --text "Talk about Q4 growth." --json
    
    # Overwrite existing notes
    uv run tools/ppt_add_notes.py --file deck.pptx --slide 0 --text "New script." --mode overwrite --json

Exit Codes:
    0: Success
    1: Error occurred
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any

# Add parent directory to path for core import
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
    """
    Add speaker notes to a slide.
    
    Args:
        filepath: Path to PowerPoint file
        slide_index: Index of slide to modify
        text: Text to add
        mode: 'append' (default), 'prepend', or 'overwrite'
    """
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if not text:
        raise ValueError("Notes text cannot be empty")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Validate slide index
        total = agent.get_slide_count()
        if not 0 <= slide_index < total:
            raise SlideNotFoundError(f"Slide index {slide_index} out of range (0-{total-1})")
            
        slide = agent.prs.slides[slide_index]
        
        # Access or create notes slide
        # python-pptx creates the notes slide automatically when accessed if it doesn't exist
        try:
            notes_slide = slide.notes_slide
            text_frame = notes_slide.notes_text_frame
        except Exception as e:
            raise PowerPointAgentError(f"Failed to access notes slide: {str(e)}")
        
        original_text = text_frame.text
        final_text = text
        
        if mode == "overwrite":
            text_frame.text = text
        elif mode == "append":
            if original_text and original_text.strip():
                text_frame.text = original_text + "\n" + text
                final_text = text_frame.text
            else:
                text_frame.text = text
        elif mode == "prepend":
             if original_text and original_text.strip():
                text_frame.text = text + "\n" + original_text
                final_text = text_frame.text
             else:
                text_frame.text = text
                
        agent.save()
        
    return {
        "status": "success",
        "file": str(filepath),
        "slide_index": slide_index,
        "mode": mode,
        "original_length": len(original_text) if original_text else 0,
        "new_length": len(final_text),
        "preview": final_text[:100] + "..." if len(final_text) > 100 else final_text
    }

def main():
    parser = argparse.ArgumentParser(
        description="Add speaker notes to PowerPoint slide",
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file path'
    )
    
    parser.add_argument(
        '--slide',
        required=True,
        type=int,
        help='Slide index (0-based)'
    )
    
    parser.add_argument(
        '--text',
        required=True,
        help='Notes content'
    )
    
    parser.add_argument(
        '--mode',
        choices=['append', 'overwrite', 'prepend'],
        default='append',
        help='Insertion mode (default: append)'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default: true)'
    )
    
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
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)

if __name__ == "__main__":
    main()
