# tools/ppt_add_notes.py
```py
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

# tools/ppt_add_notes.py (alternate version)
```python
#!/usr/bin/env python3
"""
PowerPoint Add Speaker Notes Tool
Add or update speaker notes for a specific slide.

Usage:
    uv python ppt_add_notes.py --file deck.pptx --slide 0 --text "Talk about Q4 growth." --json
    uv python ppt_add_notes.py --file deck.pptx --slide 0 --text "New point." --mode append --json

Exit Codes:
    0: Success
    1: Error occurred
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any

# Ensure core module can be imported
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
        filepath: Path to .pptx file
        slide_index: Index of slide (0-based)
        text: Notes text to add
        mode: 'append', 'prepend', or 'overwrite'
    """
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        total = agent.get_slide_count()
        if not 0 <= slide_index < total:
            raise SlideNotFoundError(f"Slide index {slide_index} out of range (0-{total-1})")
            
        slide = agent.prs.slides[slide_index]
        
        # Accessing .notes_slide property automatically creates it if missing
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
        
        # Save changes
        agent.save()
        
        # Capture final state for response
        final_text = text_frame.text
        
    return {
        "status": "success",
        "file": str(filepath),
        "slide_index": slide_index,
        "mode": mode,
        "notes_length": len(final_text),
        "preview": final_text[:100] + "..." if len(final_text) > 100 else final_text
    }

def main():
    parser = argparse.ArgumentParser(
        description="Add speaker notes to PowerPoint slide",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Add notes (appends by default)
  uv python ppt_add_notes.py --file presentation.pptx --slide 0 --text "Welcome everyone." --json
  
  # Overwrite existing notes
  uv python ppt_add_notes.py --file presentation.pptx --slide 1 --text "Specific metrics only." --mode overwrite --json
  
  # Prepend to beginning
  uv python ppt_add_notes.py --file presentation.pptx --slide 2 --text "Critical Update:" --mode prepend --json
        """
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
        help='Notes content text'
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

# tools/ppt_set_z_order.py
```py
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

# tools/ppt_set_z_order.py (alternate version)
```python
#!/usr/bin/env python3
"""
PowerPoint Set Z-Order Tool
Manage shape layering (Bring to Front, Send to Back, etc.)

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
    Change the z-order (stacking order) of a shape.
    
    Args:
        filepath: Path to .pptx file
        slide_index: Index of slide
        shape_index: Index of shape to move
        action: Movement action
    """
    
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
        
        # Access the XML element tree for the shapes
        # This allows us to reorder the elements in the list, which determines Z-order
        sp_tree = slide.shapes._spTree
        element = shape.element
        
        # Determine current position
        current_index = -1
        for i, child in enumerate(sp_tree):
            if child == element:
                current_index = i
                break
        
        if current_index == -1:
            raise RuntimeError("Could not locate shape in XML tree")
                
        new_index = current_index
        max_index = len(sp_tree) - 1
        
        # Perform Z-Order operation
        if action == 'bring_to_front':
            if current_index < max_index:
                sp_tree.remove(element)
                sp_tree.append(element)
                new_index = max_index
        
        elif action == 'send_to_back':
            # Note: Index 0 is the absolute back
            if current_index > 0:
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
        
        # Save changes
        agent.save()
        
    return {
        "status": "success",
        "file": str(filepath),
        "slide_index": slide_index,
        "original_shape_index": shape_index,
        "action": action,
        "z_order_change": {
            "from_xml_index": current_index,
            "to_xml_index": new_index
        },
        "note": "Shape index (0-based access) may have changed due to reordering."
    }

def main():
    parser = argparse.ArgumentParser(
        description="Set shape Z-Order (layering)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Z-Order Actions:
  - bring_to_front: Moves shape to the very top (covers everything else)
  - send_to_back: Moves shape to the very bottom (behind everything else)
  - bring_forward: Moves shape one step up in the stack
  - send_backward: Moves shape one step down in the stack

Note on Indices:
  Changing Z-order changes the index of the shape in the `shapes` collection.
  If you move shape 0 to front, it becomes shape N. Future commands should
  re-check indices using ppt_get_slide_info.py.
        """
    )
    
    parser.add_argument(
        '--file', required=True, type=Path, help='PowerPoint file path'
    )
    parser.add_argument(
        '--slide', required=True, type=int, help='Slide index (0-based)'
    )
    parser.add_argument(
        '--shape', required=True, type=int, help='Shape index (0-based)'
    )
    parser.add_argument(
        '--action', required=True, 
        choices=['bring_to_front', 'send_to_back', 'bring_forward', 'send_backward'],
        help='Layering action'
    )
    parser.add_argument(
        '--json', action='store_true', default=True, help='Output JSON response'
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
        print(json.dumps({"status": "error", "error": str(e), "error_type": type(e).__name__}, indent=2))
        sys.exit(1)

if __name__ == "__main__":
    main()

```

# tools/ppt_replace_text.py
```py
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

# tools/ppt_replace_text.py (alternate version)
```python
#!/usr/bin/env python3
"""
PowerPoint Replace Text Tool
Find and replace text in presentation (Global or Targeted)

Version 2.0.0 - Enhanced with Targeted Replacement

Changes:
- Added: Targeted replacement by --slide and --shape indices
- Added: Scoping logic to limit search area
- Fixed: Re-implementation using core logic with added filtering

Usage:
    # Global replacement (Original behavior)
    uv python ppt_replace_text.py --file deck.pptx --find "Old" --replace "New" --json
    
    # Targeted replacement (Single slide)
    uv python ppt_replace_text.py --file deck.pptx --find "Old" --replace "New" --slide 2 --json
    
    # Surgical replacement (Specific shape)
    uv python ppt_replace_text.py --file deck.pptx --find "Old" --replace "New" --slide 2 --shape 0 --json
"""

import sys
import json
import argparse
import re
from pathlib import Path
from typing import Dict, Any, Optional

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError, SlideNotFoundError
)

def perform_replacement_on_shape(shape, find: str, replace: str, match_case: bool) -> int:
    """Helper to perform replacement on a single shape object."""
    if not shape.has_text_frame:
        return 0
        
    count = 0
    text_frame = shape.text_frame
    
    # Strategy 1: Replace within runs (preserves formatting)
    for paragraph in text_frame.paragraphs:
        for run in paragraph.runs:
            if match_case:
                if find in run.text:
                    run.text = run.text.replace(find, replace)
                    count += 1 # Count runs modified
            else:
                if find.lower() in run.text.lower():
                    pattern = re.compile(re.escape(find), re.IGNORECASE)
                    if pattern.search(run.text):
                        new_text = pattern.sub(replace, run.text)
                        run.text = new_text
                        count += 1

    # Strategy 2: If no runs modified but text exists, try shape-level replacement
    # This handles text split across runs
    if count == 0 and shape.text:
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
                # Only replace if actually different
                if new_text != full_text:
                    shape.text = new_text
                    count += 1
            else:
                pattern = re.compile(re.escape(find), re.IGNORECASE)
                new_text = pattern.sub(replace, full_text)
                if new_text != full_text:
                    shape.text = new_text
                    count += 1
                    
    return count

def replace_text(
    filepath: Path,
    find: str,
    replace: str,
    match_case: bool = False,
    dry_run: bool = False,
    slide_index: Optional[int] = None,
    shape_index: Optional[int] = None
) -> Dict[str, Any]:
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if not find:
        raise ValueError("Find text cannot be empty")
        
    replacements = 0
    locations = []
    
    with PowerPointAgent(filepath) as agent:
        # Lock unless dry run
        agent.open(filepath, acquire_lock=not dry_run)
        
        # Determine scope
        slides_to_process = []
        
        if slide_index is not None:
            # Target specific slide
            total_slides = agent.get_slide_count()
            if not 0 <= slide_index < total_slides:
                raise SlideNotFoundError(f"Slide index {slide_index} out of range")
            slides_to_process = [(slide_index, agent.prs.slides[slide_index])]
        else:
            # Global scope
            slides_to_process = list(enumerate(agent.prs.slides))
            
        for s_idx, slide in slides_to_process:
            shapes_to_process = []
            
            if slide_index is not None and shape_index is not None:
                # Target specific shape
                if not 0 <= shape_index < len(slide.shapes):
                    raise ValueError(f"Shape index {shape_index} out of range on slide {s_idx}")
                shapes_to_process = [(shape_index, slide.shapes[shape_index])]
            else:
                # All shapes on this slide
                shapes_to_process = list(enumerate(slide.shapes))
                
            for sh_idx, shape in shapes_to_process:
                if not hasattr(shape, 'text_frame'):
                    continue
                    
                # For dry run, just peek
                if dry_run:
                    text = shape.text_frame.text
                    match_count = 0
                    if match_case:
                        match_count = text.count(find)
                    else:
                        match_count = text.lower().count(find.lower())
                    
                    if match_count > 0:
                        replacements += match_count
                        locations.append({
                            "slide": s_idx,
                            "shape": sh_idx,
                            "matches": match_count,
                            "context": text[:50] + "..." if len(text) > 50 else text
                        })
                else:
                    # Perform replacement
                    count = perform_replacement_on_shape(shape, find, replace, match_case)
                    if count > 0:
                        replacements += count
                        locations.append({
                            "slide": s_idx,
                            "shape": sh_idx,
                            "changes": count
                        })
        
        if not dry_run:
            agent.save()
            
    return {
        "status": "success" if not dry_run else "dry_run",
        "file": str(filepath),
        "find": find,
        "replace": replace,
        "scope": {
            "slide": slide_index if slide_index is not None else "all",
            "shape": shape_index if shape_index is not None else "all"
        },
        "matches_found" if dry_run else "replacements_made": replacements,
        "locations_preview": locations[:10],
        "total_locations": len(locations)
    }

def main():
    parser = argparse.ArgumentParser(
        description="Find and replace text (Global or Targeted)",
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    
    parser.add_argument('--file', required=True, type=Path, help='PowerPoint file path')
    parser.add_argument('--find', required=True, help='Text to find')
    parser.add_argument('--replace', required=True, help='Replacement text')
    parser.add_argument('--match-case', action='store_true', help='Case-sensitive matching')
    
    # New arguments for targeting
    parser.add_argument('--slide', type=int, help='Target specific slide index (optional)')
    parser.add_argument('--shape', type=int, help='Target specific shape index (optional, requires --slide)')
    
    parser.add_argument('--dry-run', action='store_true', help='Preview changes without modifying')
    parser.add_argument('--json', action='store_true', default=True, help='Output JSON response')
    
    args = parser.parse_args()
    
    try:
        if args.shape is not None and args.slide is None:
            raise ValueError("--shape requires --slide to be specified")
            
        result = replace_text(
            filepath=args.file,
            find=args.find,
            replace=args.replace,
            match_case=args.match_case,
            dry_run=args.dry_run,
            slide_index=args.slide,
            shape_index=args.shape
        )
        print(json.dumps(result, indent=2))
        sys.exit(0)
    except Exception as e:
        print(json.dumps({"status": "error", "error": str(e), "error_type": type(e).__name__}, indent=2))
        sys.exit(1)

if __name__ == "__main__":
    main()

```

# tests/test_remedial_tools.py
```py
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

# tests/test_remedial_tools.py (alternate version)
```python
import pytest
import json
import sys
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches

# Add project root to path
sys.path.insert(0, str(Path(__file__).parent.parent))

# Import the tools directly to test logic
# (Alternatively, could use subprocess to test CLI, but direct import is faster for unit tests)
from tools.ppt_add_notes import add_notes
from tools.ppt_set_z_order import set_z_order
from tools.ppt_replace_text import replace_text

@pytest.fixture
def temp_deck(tmp_path):
    """Create a temporary deck for testing."""
    path = tmp_path / "test.pptx"
    prs = Presentation()
    
    # Slide 0: Title Slide
    slide_0 = prs.slides.add_slide(prs.slide_layouts[0])
    slide_0.shapes.title.text = "Slide Zero"
    
    # Slide 1: Blank with Shapes
    slide_1 = prs.slides.add_slide(prs.slide_layouts[6]) # Blank
    
    # Add shapes for Z-order and Text tests
    # Shape 0 (Back initially)
    s1 = slide_1.shapes.add_shape(1, Inches(1), Inches(1), Inches(2), Inches(2))
    s1.name = "BackShape"
    s1.text = "TargetText"
    
    # Shape 1 (Front initially)
    s2 = slide_1.shapes.add_shape(1, Inches(2), Inches(2), Inches(2), Inches(2))
    s2.name = "FrontShape"
    s2.text = "TargetText" # Same text for replacement test
    
    prs.save(path)
    return path

def test_add_notes_workflow(temp_deck):
    """Test adding, appending, and overwriting notes."""
    
    # 1. Add new notes
    res = add_notes(temp_deck, 0, "Initial note.")
    assert res["status"] == "success"
    
    # Verify
    prs = Presentation(temp_deck)
    assert "Initial note." in prs.slides[0].notes_slide.notes_text_frame.text
    
    # 2. Append notes
    add_notes(temp_deck, 0, "Second note.", mode="append")
    prs = Presentation(temp_deck)
    text = prs.slides[0].notes_slide.notes_text_frame.text
    assert "Initial note." in text
    assert "Second note." in text
    
    # 3. Overwrite notes
    add_notes(temp_deck, 0, "Overwritten.", mode="overwrite")
    prs = Presentation(temp_deck)
    assert prs.slides[0].notes_slide.notes_text_frame.text == "Overwritten."

def test_z_order_workflow(temp_deck):
    """Test manipulating Z-order."""
    
    # Initial check: Shape 0 is index 0, Shape 1 is index 1
    prs = Presentation(temp_deck)
    slide = prs.slides[1]
    assert slide.shapes[0].name == "BackShape"
    assert slide.shapes[1].name == "FrontShape"
    
    # Action: Send FrontShape (index 1) to back
    res = set_z_order(temp_deck, 1, 1, "send_to_back")
    assert res["status"] == "success"
    
    # Verify: FrontShape should now be at index 0
    prs = Presentation(temp_deck)
    slide = prs.slides[1]
    assert slide.shapes[0].name == "FrontShape"
    assert slide.shapes[1].name == "BackShape"

def test_targeted_replace_text(temp_deck):
    """Test replacing text in a specific shape only."""
    
    # Scenario: Both shapes on Slide 1 have "TargetText".
    # We want to replace ONLY the one in Shape 0 (now FrontShape due to prev test? No, fresh fixture)
    # Fixture is recreated per test function in pytest by default unless scoped module.
    # Default pytest fixture scope is 'function', so we get a fresh deck.
    
    # Target Slide 1, Shape 0 ("BackShape")
    res = replace_text(
        temp_deck, 
        find="TargetText", 
        replace="REPLACED", 
        slide_index=1, 
        shape_index=0
    )
    
    assert res["status"] == "success"
    assert res["replacements_made"] == 1
    
    # Verify
    prs = Presentation(temp_deck)
    slide = prs.slides[1]
    
    # Shape 0 should change
    assert slide.shapes[0].text == "REPLACED"
    
    # Shape 1 should remain unchanged
    assert slide.shapes[1].text == "TargetText"

def test_global_replace_text(temp_deck):
    """Test global replacement works as legacy."""
    res = replace_text(temp_deck, find="TargetText", replace="GLOBAL")
    
    assert res["status"] == "success"
    # Should replace 2 occurrences (Shape 0 and Shape 1 on Slide 1)
    assert res["replacements_made"] == 2
    
    prs = Presentation(temp_deck)
    assert prs.slides[1].shapes[0].text == "GLOBAL"
    assert prs.slides[1].shapes[1].text == "GLOBAL"

```

