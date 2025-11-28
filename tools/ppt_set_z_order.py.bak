#!/usr/bin/env python3
"""
PowerPoint Set Z-Order Tool
Manage shape layering (Bring to Front, Send to Back).

Usage:
    uv python ppt_set_z_order.py --file deck.pptx --slide 0 --shape 1 --action bring_to_front --json

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

def set_z_order(
    filepath: Path,
    slide_index: int,
    shape_index: int,
    action: str
) -> Dict[str, Any]:
    """
    Change the Z-order (stacking order) of a shape.
    """
    
    if not filepath.suffix.lower() in ['.pptx', '.ppt']:
        raise ValueError("Invalid PowerPoint file format (must be .pptx or .ppt)")

    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Performance warning for large presentations
        slide_count = agent.get_slide_count()
        if slide_count > 50:
            print(f"⚠️  WARNING: Large presentation ({slide_count} slides) - operation may take longer", file=sys.stderr)
        
        # Validate slide index
        if not 0 <= slide_index < slide_count:
            raise SlideNotFoundError(f"Slide index {slide_index} out of range (0-{slide_count-1})")
            
        slide = agent.prs.slides[slide_index]
        
        # Validate shape index
        if not 0 <= shape_index < len(slide.shapes):
            raise ValueError(f"Shape index {shape_index} out of range (0-{len(slide.shapes)-1})")
            
        shape = slide.shapes[shape_index]
        
        # XML Manipulation for Z-Order
        # The ._spTree attribute is the lxml element containing shapes
        sp_tree = slide.shapes._spTree
        element = shape.element
        
        # Find current position
        current_index = -1
        for i, child in enumerate(sp_tree):
            if child == element:
                current_index = i
                break
        
        if current_index == -1:
            raise PowerPointAgentError("Could not locate shape in XML tree")
            
        new_index = current_index
        max_index = len(sp_tree) - 1
        
        # Execute Action
        if action == 'bring_to_front':
            sp_tree.remove(element)
            sp_tree.append(element)
            new_index = max_index
            
        elif action == 'send_to_back':
            sp_tree.remove(element)
            # Insert at 0 (behind everything else on this slide)
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
        
        # Post-manipulation validation
        if not _validate_xml_structure(sp_tree):
            raise PowerPointAgentError("XML structure corrupted during Z-order operation")
                
        agent.save()

def _validate_xml_structure(sp_tree) -> bool:
    """Validate XML tree integrity after manipulation"""
    return all(child is not None for child in sp_tree)
        
    return {
        "status": "success",
        "file": str(filepath),
        "slide_index": slide_index,
        "shape_index_target": shape_index,
        "action": action,
        "z_order_change": {
            "from": current_index,
            "to": new_index
        },
        "note": "Shape indices may shift after reordering."
    }

def main():
    parser = argparse.ArgumentParser(
        description="Set shape Z-Order (layering)",
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
        '--shape',
        required=True,
        type=int,
        help='Shape index (0-based)'
    )
    
    parser.add_argument(
        '--action',
        required=True, 
        choices=['bring_to_front', 'send_to_back', 'bring_forward', 'send_backward'],
        help='Layering action to perform'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default: true)'
    )
    
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
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)

if __name__ == "__main__":
    main()
