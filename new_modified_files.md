### 1. Creating `tools/ppt_add_notes.py`

This tool enables the "AI Architect" to script the presentation, adding a critical layer of "voice" to the visual slides.

```python
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
```

**Review & Validation (Tools Checklist):**
*   [x] Imports `core.powerpoint_agent_core`? Yes.
*   [x] JSON Output? Yes.
*   [x] Error Handling? Yes (`try...except`).
*   [x] Functionality? Adds notes via `slide.notes_slide`. Logic handles append/prepend/overwrite.

---

### 2. Creating `tools/ppt_set_z_order.py`

This tool gives the Architect control over the "Z-axis", essential for managing overlays and backgrounds.

```python
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
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Validate slide index
        total = agent.get_slide_count()
        if not 0 <= slide_index < total:
            raise SlideNotFoundError(f"Slide index {slide_index} out of range (0-{total-1})")
            
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
                
        agent.save()
        
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
```

**Review & Validation (Tools Checklist):**
*   [x] Imports `core`? Yes.
*   [x] XML manipulation? Uses `_spTree` correctly.
*   [x] Action logic? Handles all 4 cases (front, back, forward, backward).
*   [x] Error Handling? Checks bounds and XML existence.

---

### 3. Updating `tools/ppt_replace_text.py`

I will enhance this tool to support "Surgical Targeting" (Specific slide/shape) while maintaining global capability.

```python
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
```

**Review & Validation (Tools Checklist):**
*   [x] Surgical Targeting? Implemented `--slide` and `--shape` logic.
*   [x] Helper Function? `perform_replacement_on_shape` duplicates core logic but ensures isolation.
*   [x] Dry Run? Supported for scoped searches.
*   [x] Input Validation? Checks that `--shape` implies `--slide`.

---

### 4. Creating `tests/test_remedial_tools.py`

```python
import pytest
import json
import os
import sys
from pathlib import Path
from pptx import Presentation

# Allow importing tools directly
sys.path.insert(0, str(Path(__file__).parent.parent))

from tools.ppt_add_notes import add_notes
from tools.ppt_set_z_order import set_z_order
from tools.ppt_replace_text import replace_text

def test_add_notes(tmp_path):
    """Test adding speaker notes functionality."""
    pptx_path = tmp_path / "test_notes.pptx"
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6]) # Blank layout
    prs.save(pptx_path)
    
    # 1. Add initial note
    res1 = add_notes(pptx_path, 0, "First note", mode="append")
    assert res1["status"] == "success"
    
    # Verify
    prs = Presentation(pptx_path)
    assert prs.slides[0].notes_slide.notes_text_frame.text == "First note"
    
    # 2. Append note
    res2 = add_notes(pptx_path, 0, "Second note", mode="append")
    
    # Verify Append
    prs = Presentation(pptx_path)
    assert "First note" in prs.slides[0].notes_slide.notes_text_frame.text
    assert "Second note" in prs.slides[0].notes_slide.notes_text_frame.text
    
    # 3. Overwrite note
    res3 = add_notes(pptx_path, 0, "Overwritten", mode="overwrite")
    
    # Verify Overwrite
    prs = Presentation(pptx_path)
    assert prs.slides[0].notes_slide.notes_text_frame.text == "Overwritten"

def test_z_order(tmp_path):
    """Test shape layering (Z-Order)."""
    pptx_path = tmp_path / "test_z.pptx"
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # Create shapes (0=Back by default creation order)
    s1 = slide.shapes.add_shape(1, 0, 0, 100, 100)
    s1.name = "Shape 1"
    s2 = slide.shapes.add_shape(1, 10, 10, 100, 100)
    s2.name = "Shape 2" # On top
    
    prs.save(pptx_path)
    
    # Verify initial state (Shape 2 is last in list, so on top)
    prs = Presentation(pptx_path)
    assert prs.slides[0].shapes[-1].name == "Shape 2"
    
    # Action: Send Shape 2 (index 1) to back
    set_z_order(pptx_path, 0, 1, "send_to_back")
    
    # Verify swap
    prs = Presentation(pptx_path)
    # Shape 2 should now be at index 0 (back)
    assert prs.slides[0].shapes[0].name == "Shape 2"
    # Shape 1 should now be at index 1 (front)
    assert prs.slides[0].shapes[1].name == "Shape 1"

def test_targeted_text_replace(tmp_path):
    """Test surgical text replacement."""
    pptx_path = tmp_path / "test_replace.pptx"
    prs = Presentation()
    
    # Slide 0
    slide0 = prs.slides.add_slide(prs.slide_layouts[6])
    tb0 = slide0.shapes.add_textbox(0,0,100,100)
    tb0.text_frame.text = "Target"
    
    # Slide 1
    slide1 = prs.slides.add_slide(prs.slide_layouts[6])
    tb1 = slide1.shapes.add_textbox(0,0,100,100)
    tb1.text_frame.text = "Target"
    
    prs.save(pptx_path)
    
    # Action: Replace "Target" -> "Hit" ONLY on Slide 1
    replace_text(pptx_path, "Target", "Hit", slide_index=1)
    
    # Verify
    prs = Presentation(pptx_path)
    # Slide 0 should be UNCHANGED
    assert prs.slides[0].shapes[0].text == "Target"
    # Slide 1 should be CHANGED
    assert prs.slides[1].shapes[0].text == "Hit"

```

**Review & Validation (Test Checklist):**
*   [x] Imports correct tools? Yes.
*   [x] Tests Notes? Yes (Add, Append, Overwrite).
*   [x] Tests Z-Order? Yes (Create, Send to Back, Verify Index Change).
*   [x] Tests Targeting? Yes (Multi-slide setup, single slide target).

**Conclusion**: The remedial execution plan is fully implemented. The new tools provide essential capabilities for the AI Architect, and the enhancements add necessary precision. All code has been reviewed against the design specs.

I am ready to generate the files.

### File 1: tools/ppt_add_notes.py
```python
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
```

### File 2: tools/ppt_set_z_order.py
```python
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
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Validate slide index
        total = agent.get_slide_count()
        if not 0 <= slide_index < total:
            raise SlideNotFoundError(f"Slide index {slide_index} out of range (0-{total-1})")
            
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
                
        agent.save()
        
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
```

### File 3: tools/ppt_replace_text.py
```python
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
```

### File 4: tests/test_remedial_tools.py
```python
import pytest
import json
import os
import sys
from pathlib import Path
from pptx import Presentation

# Allow importing tools directly
sys.path.insert(0, str(Path(__file__).parent.parent))

from tools.ppt_add_notes import add_notes
from tools.ppt_set_z_order import set_z_order
from tools.ppt_replace_text import replace_text

def test_add_notes(tmp_path):
    """Test adding speaker notes functionality."""
    pptx_path = tmp_path / "test_notes.pptx"
    prs = Presentation()
    # Use a layout that exists (index 0 usually Title Slide)
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    prs.save(pptx_path)
    
    # 1. Add initial note
    res1 = add_notes(pptx_path, 0, "First note", mode="append")
    assert res1["status"] == "success"
    
    # Verify
    prs = Presentation(pptx_path)
    assert prs.slides[0].notes_slide.notes_text_frame.text == "First note"
    
    # 2. Append note
    res2 = add_notes(pptx_path, 0, "Second note", mode="append")
    
    # Verify Append
    prs = Presentation(pptx_path)
    assert "First note" in prs.slides[0].notes_slide.notes_text_frame.text
    assert "Second note" in prs.slides[0].notes_slide.notes_text_frame.text
    
    # 3. Overwrite note
    res3 = add_notes(pptx_path, 0, "Overwritten", mode="overwrite")
    
    # Verify Overwrite
    prs = Presentation(pptx_path)
    assert prs.slides[0].notes_slide.notes_text_frame.text == "Overwritten"

def test_z_order(tmp_path):
    """Test shape layering (Z-Order)."""
    pptx_path = tmp_path / "test_z.pptx"
    prs = Presentation()
    # Use blank layout (usually index 6)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # Create shapes (0=Back by default creation order)
    s1 = slide.shapes.add_shape(1, 0, 0, 100, 100)
    s1.name = "Shape 1"
    s2 = slide.shapes.add_shape(1, 10, 10, 100, 100)
    s2.name = "Shape 2" # On top
    
    prs.save(pptx_path)
    
    # Verify initial state (Shape 2 is last in list, so on top)
    prs = Presentation(pptx_path)
    assert prs.slides[0].shapes[-1].name == "Shape 2"
    
    # Action: Send Shape 2 (index 1) to back
    set_z_order(pptx_path, 0, 1, "send_to_back")
    
    # Verify swap
    prs = Presentation(pptx_path)
    # Shape 2 should now be at index 0 (back)
    assert prs.slides[0].shapes[0].name == "Shape 2"
    # Shape 1 should now be at index 1 (front)
    assert prs.slides[0].shapes[1].name == "Shape 1"

def test_targeted_text_replace(tmp_path):
    """Test surgical text replacement."""
    pptx_path = tmp_path / "test_replace.pptx"
    prs = Presentation()
    
    # Slide 0
    slide0 = prs.slides.add_slide(prs.slide_layouts[6])
    tb0 = slide0.shapes.add_textbox(0,0,100,100)
    tb0.text_frame.text = "Target"
    
    # Slide 1
    slide1 = prs.slides.add_slide(prs.slide_layouts[6])
    tb1 = slide1.shapes.add_textbox(0,0,100,100)
    tb1.text_frame.text = "Target"
    
    prs.save(pptx_path)
    
    # Action: Replace "Target" -> "Hit" ONLY on Slide 1
    replace_text(pptx_path, "Target", "Hit", slide_index=1)
    
    # Verify
    prs = Presentation(pptx_path)
    # Slide 0 should be UNCHANGED
    assert prs.slides[0].shapes[0].text_frame.text == "Target"
    # Slide 1 should be CHANGED
    assert prs.slides[1].shapes[0].text_frame.text == "Hit"
```
