#!/usr/bin/env python3
"""
PowerPoint Set Z-Order Tool
Manage shape layering (Bring to Front, Send to Back).

Usage:
    uv python ppt_set_z_order.py --file deck.pptx --slide 0 --shape 1 --action bring_to_front --json
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

def set_z_order(
    filepath: Path,
    slide_index: int,
    shape_index: int,
    action: str
) -> Dict[str, Any]:
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        total = agent.get_slide_count()
        if not 0 <= slide_index < total:
            raise SlideNotFoundError(f"Slide index {slide_index} out of range")
            
        slide = agent.prs.slides[slide_index]
        
        if not 0 <= shape_index < len(slide.shapes):
            raise ValueError(f"Shape index {shape_index} out of range")
            
        shape = slide.shapes[shape_index]
        
        # XML Manipulation for Z-Order
        sp_tree = slide.shapes._spTree
        element = shape.element
        
        # Current position
        current_index = -1
        for i, child in enumerate(sp_tree):
            if child == element:
                current_index = i
                break
                
        new_index = current_index
        max_index = len(sp_tree) - 1
        
        if action == 'bring_to_front':
            sp_tree.remove(element)
            sp_tree.append(element)
            new_index = max_index
        elif action == 'send_to_back':
            # Index 0 is often the background or master placeholder stuff, 
            # but in _spTree, strict order 0 is the very back.
            sp_tree.remove(element)
            sp_tree.insert(0, element)
            new_index = 0
        elif action == 'bring_forward':
            if current_index < max_index:
                sp_tree.remove(element)
                sp_tree.insert(current_index + 1, element)
                new_index = current_index + 1
        elif action == 'send_backward':
            if current_index > 0:
                sp_tree.remove(element)
                sp_tree.insert(current_index - 1, element)
                new_index = current_index - 1
                
        agent.save()
        
    return {
        "status": "success",
        "file": str(filepath),
        "slide_index": slide_index,
        "shape_index": shape_index, # Note: Index might change after reorder!
        "action": action,
        "z_order_change": {
            "from": current_index,
            "to": new_index
        }
    }

def main():
    parser = argparse.ArgumentParser(description="Set shape Z-Order (layering)")
    parser.add_argument('--file', required=True, type=Path, help='PowerPoint file path')
    parser.add_argument('--slide', required=True, type=int, help='Slide index (0-based)')
    parser.add_argument('--shape', required=True, type=int, help='Shape index (0-based)')
    parser.add_argument('--action', required=True, 
                        choices=['bring_to_front', 'send_to_back', 'bring_forward', 'send_backward'],
                        help='Layering action')
    parser.add_argument('--json', action='store_true', default=True, help='Output JSON')
    
    args = parser.parse_args()
    
    try:
        result = set_z_order(
            filepath=args.file,
            slide_index=args.slide,
            shape_index=args.shape,
            action=args.action
        )
        print(json.dumps(result, indent=2))
        sys.exit(0)
    except Exception as e:
        print(json.dumps({"status": "error", "error": str(e), "error_type": type(e).__name__}, indent=2))
        sys.exit(1)

if __name__ == "__main__":
    main()
