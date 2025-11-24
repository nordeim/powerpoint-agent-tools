#!/usr/bin/env python3
"""
PowerPoint Replace Text Tool
Find and replace text across presentation or in specific targets

Version 2.0.0 - Enhanced with Surgical Targeting

Changes from v1.0:
- Added: --slide and --shape arguments for targeted replacement
- Enhanced: Logic to handle specific scope vs global scope
- Enhanced: Detailed reporting on location of replacements

Usage:
    # Global replacement
    uv run tools/ppt_replace_text.py --file deck.pptx --find "Old" --replace "New" --json
    
    # Targeted replacement (Specific Slide)
    uv run tools/ppt_replace_text.py --file deck.pptx --slide 2 --find "Old" --replace "New" --json
    
    # Surgical replacement (Specific Shape)
    uv run tools/ppt_replace_text.py --file deck.pptx --slide 2 --shape 0 --find "Old" --replace "New" --json
"""

import sys
import re
import json
import argparse
from pathlib import Path
from typing import Dict, Any, List, Optional

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError, SlideNotFoundError
)

def perform_replacement_on_shape(shape, find: str, replace: str, match_case: bool) -> int:
    """
    Helper to replace text in a single shape.
    Returns number of replacements made.
    """
    if not hasattr(shape, 'text_frame'):
        return 0
        
    count = 0
    text_frame = shape.text_frame
    
    # Strategy 1: Replace in runs (preserves formatting)
    for paragraph in text_frame.paragraphs:
        for run in paragraph.runs:
            if match_case:
                if find in run.text:
                    run.text = run.text.replace(find, replace)
                    count += 1
            else:
                if find.lower() in run.text.lower():
                    pattern = re.compile(re.escape(find), re.IGNORECASE)
                    if pattern.search(run.text):
                        run.text = pattern.sub(replace, run.text)
                        count += 1
    
    if count > 0:
        return count
        
    # Strategy 2: Shape-level replacement (if runs didn't catch it due to splitting)
    # Only try this if Strategy 1 failed but we know the text exists
    try:
        full_text = shape.text
        should_replace = False
        if match_case:
            if find in full_text:
                should_replace = True
        else:
            if find.lower() in full_text.lower():
                should_replace = True
        
        if should_replace:
            if match_case:
                new_text = full_text.replace(find, replace)
                shape.text = new_text
                count += 1
            else:
                pattern = re.compile(re.escape(find), re.IGNORECASE)
                new_text = pattern.sub(replace, full_text)
                shape.text = new_text
                count += 1
    except:
        pass
        
    return count

def replace_text(
    filepath: Path,
    find: str,
    replace: str,
    slide_index: Optional[int] = None,
    shape_index: Optional[int] = None,
    match_case: bool = False,
    dry_run: bool = False
) -> Dict[str, Any]:
    """Find and replace text with optional targeting."""
    
    if not filepath.suffix.lower() in ['.pptx', '.ppt']:
        raise ValueError("Invalid PowerPoint file format (must be .pptx or .ppt)")

    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if not find:
        raise ValueError("Find text cannot be empty")
    
    # If shape is specified, slide must be specified
    if shape_index is not None and slide_index is None:
        raise ValueError("If --shape is specified, --slide must also be specified")
    
    action = "dry_run" if dry_run else "replace"
    total_replacements = 0
    locations = []
    
    with PowerPointAgent(filepath) as agent:
        # Open appropriately based on dry_run
        agent.open(filepath, acquire_lock=not dry_run)
        
        # Performance warning for large presentations
        slide_count = agent.get_slide_count()
        if slide_count > 50:
            print(f"⚠️  WARNING: Large presentation ({slide_count} slides) - operation may take longer", file=sys.stderr)
        
        # Determine scope
        target_slides = []
        if slide_index is not None:
            # Single slide scope
            if not 0 <= slide_index < agent.get_slide_count():
                raise SlideNotFoundError(f"Slide index {slide_index} out of range")
            target_slides = [(slide_index, agent.prs.slides[slide_index])]
        else:
            # Global scope
            target_slides = list(enumerate(agent.prs.slides))
            
        # Iterate scope
        for s_idx, slide in target_slides:
            
            # Determine shapes on this slide
            target_shapes = []
            if shape_index is not None:
                # Single shape scope
                if 0 <= shape_index < len(slide.shapes):
                    target_shapes = [(shape_index, slide.shapes[shape_index])]
                else:
                    # Warning or Error? Error seems appropriate for explicit target
                    raise ValueError(f"Shape index {shape_index} out of range on slide {s_idx}")
            else:
                # All shapes on slide
                target_shapes = list(enumerate(slide.shapes))
            
            # Execute on shapes
            for sh_idx, shape in target_shapes:
                if not hasattr(shape, 'text_frame'):
                    continue
                    
                # Logic for Dry Run (Count only)
                if dry_run:
                    text = shape.text_frame.text
                    occurrences = 0
                    if match_case:
                        occurrences = text.count(find)
                    else:
                        occurrences = text.lower().count(find.lower())
                    
                    if occurrences > 0:
                        total_replacements += occurrences
                        locations.append({
                            "slide": s_idx,
                            "shape": sh_idx,
                            "occurrences": occurrences,
                            "preview": text[:50] + "..." if len(text) > 50 else text
                        })
                
                # Logic for Actual Replacement
                else:
                    replacements = perform_replacement_on_shape(shape, find, replace, match_case)
                    if replacements > 0:
                        total_replacements += replacements
                        locations.append({
                            "slide": s_idx,
                            "shape": sh_idx,
                            "replacements": replacements
                        })
        
        if not dry_run:
            agent.save()
            
    return {
        "status": "success",
        "file": str(filepath),
        "action": action,
        "find": find,
        "replace": replace,
        "scope": {
            "slide": slide_index if slide_index is not None else "all",
            "shape": shape_index if shape_index is not None else "all"
        },
        "total_matches" if dry_run else "replacements_made": total_replacements,
        "locations": locations
    }

def main():
    parser = argparse.ArgumentParser(
        description="Find and replace text in PowerPoint (v2.0.0 - Targeted)",
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    
    parser.add_argument('--file', required=True, type=Path, help='PowerPoint file path')
    parser.add_argument('--find', required=True, help='Text to find')
    parser.add_argument('--replace', required=True, help='Replacement text')
    parser.add_argument('--slide', type=int, help='Target specific slide index')
    parser.add_argument('--shape', type=int, help='Target specific shape index (requires --slide)')
    parser.add_argument('--match-case', action='store_true', help='Case-sensitive matching')
    parser.add_argument('--dry-run', action='store_true', help='Preview changes without modifying')
    parser.add_argument('--json', action='store_true', default=True, help='Output JSON response')
    
    args = parser.parse_args()
    
    try:
        result = replace_text(
            filepath=args.file,
            find=args.find,
            replace=args.replace,
            slide_index=args.slide,
            shape_index=args.shape,
            match_case=args.match_case,
            dry_run=args.dry_run
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
