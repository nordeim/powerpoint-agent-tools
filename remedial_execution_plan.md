# Remedial Execution Plan: Toolkit Enhancement & Expansion

## 📋 Executive Summary
Based on the deep analysis of the current `codebase_files_set.md`, I have identified specific deficits that hinder the "AI Presentation Architect" from achieving full autonomy. The current toolkit excels at *creation* but lacks precision in *refinement* (layering) and *annotation* (speaker notes).

This plan outlines the development of **2 New Tools** and **2 Critical Enhancements** to bridge these gaps.

### 🎯 Strategic Objectives
1.  **Enable "The Voice"**: Implement Speaker Notes management to allow the AI to script the presentation.
2.  **Master "The Z-Axis"**: Implement Z-Order control to manage visual layering (e.g., overlays, backgrounds).
3.  **Surgical Precision**: Enhance text replacement to target specific shapes, preventing global regex accidents.
4.  **Logistical Efficiency**: Enable batch operations for slide reordering.

---

## 🛠️ Phase 1: New Tool Development

### 1. New Tool: `ppt_add_notes.py`
**Purpose**: Adds or appends speaker notes to a specific slide. Essential for generating presentation scripts.

#### Design Specification
*   **Core Logic**: Check `slide.has_notes_slide`. If False, access `slide.notes_slide` (which creates it). Access `notes_slide.notes_text_frame`.
*   **Features**:
    *   `--text`: Content to add.
    *   `--mode`: `append` (default) or `overwrite`.
    *   `--json`: Standard output format.

#### Implementation Plan
```python
# tools/ppt_add_notes.py sketch
def add_notes(filepath, slide_index, text, mode="append"):
    with PowerPointAgent(filepath) as agent:
        slide = agent.get_slide(slide_index)
        notes_slide = slide.notes_slide # Auto-creates if missing
        text_frame = notes_slide.notes_text_frame
        
        if mode == "overwrite":
            text_frame.text = text
        else:
            # Append logic preserving existing text
            if text_frame.text:
                text_frame.text += "\n" + text
            else:
                text_frame.text = text
```

#### ✅ Integrated Checklist
- [ ] Import core libraries and `PowerPointAgent`.
- [ ] Implement `add_notes` function with mode switching.
- [ ] Add CLI argument parsing (`--file`, `--slide`, `--text`, `--mode`).
- [ ] Implement JSON output schema.
- [ ] **Test Case**: Add notes to a slide that has none.
- [ ] **Test Case**: Append notes to a slide that already has notes.
- [ ] **Test Case**: Verify notes persist after save.

---

### 2. New Tool: `ppt_set_z_order.py`
**Purpose**: Manage the visual stacking order of shapes (Bring to Front, Send to Back). Critical for placing text over images or backgrounds.

#### Design Specification
*   **Core Logic**: `python-pptx` does not have a direct `z_order` API. We must manipulate the XML elements in the `slide.shapes._spTree`.
*   **Operations**:
    *   `bring_to_front`: Move element to end of `_spTree`.
    *   `send_to_back`: Move element to index 0 of `_spTree`.
    *   `bring_forward`: Swap with next element.
    *   `send_backward`: Swap with previous element.

#### Implementation Plan
```python
# tools/ppt_set_z_order.py sketch
def set_z_order(filepath, slide_index, shape_index, action):
    with PowerPointAgent(filepath) as agent:
        slide = agent.get_slide(slide_index)
        shape = slide.shapes[shape_index]
        
        # XML Manipulation Magic
        sp_tree = slide.shapes._spTree
        element = shape.element
        
        if action == 'bring_to_front':
            sp_tree.remove(element)
            sp_tree.append(element)
        elif action == 'send_to_back':
            sp_tree.remove(element)
            sp_tree.insert(0, element)
        # ... handle forward/backward
```

#### ✅ Integrated Checklist
- [ ] Import `python-pptx` internals access.
- [ ] Implement XML tree manipulation logic safely.
- [ ] Add CLI arguments (`--action` choices: `bring_to_front`, `send_to_back`, etc.).
- [ ] Implement JSON output showing new Z-index.
- [ ] **Test Case**: Create two overlapping shapes, swap their order.
- [ ] **Test Case**: Verify visual result in output file.

---

## 🔧 Phase 2: Critical Enhancements

### 3. Enhancement: `ppt_replace_text.py` (Surgical Targeting)
**Critique**: Currently finds/replaces globally. This is dangerous for specific edits.
**Remedy**: Add optional `--slide` and `--shape` arguments.

#### Modified Design
*   **Input**: Add `--slide N` and `--shape N` optional args.
*   **Logic**:
    *   If `slide` and `shape` provided: Get specific shape, replace text ONLY in that shape.
    *   If only `slide` provided: Loop through shapes ONLY on that slide.
    *   Else: Default to global replacement.

#### ✅ Integrated Checklist
- [ ] Update `argparse` to accept `--slide` and `--shape`.
- [ ] Refactor `replace_text` function in `core` or tool to handle scoping.
- [ ] Ensure "Dry Run" reports specific location correctly.
- [ ] **Test Case**: Replace text in Slide 1, Shape 0, ignore identical text in Slide 2.

---

### 4. Enhancement: `ppt_reorder_slides.py` (Batch Processing)
**Critique**: Moving one slide at a time is inefficient and error-prone for major restructuring.
**Remedy**: Allow moving multiple slides or defining a new order array.

#### Modified Design
*   **Input**: Change from `--from-index`/`--to-index` to `--order`.
*   **Argument**: `--order "0,3,1,2"` (New sequence of original indices).
*   **Logic**:
    *   Validate that `--order` contains all existing indices exactly once (permutation check).
    *   Create a new slide list based on the order.
    *   Reassign `prs.slides._sldIdLst`.

#### ✅ Integrated Checklist
- [ ] Update `argparse` to accept `--order` (list of ints).
- [ ] Implement validation logic (ensure no slides are lost or duplicated).
- [ ] Implement atomic reordering logic.
- [ ] **Test Case**: Reversing slide order (e.g., 0,1,2 -> 2,1,0).

---

## 🧪 Phase 3: Validation & Testing Plan

I will generate a new test file `tests/test_remedial_tools.py` to validate these specific additions.

**Test Strategy:**
1.  **Setup**: Create a temporary deck with known content (3 slides, 2 shapes per slide).
2.  **Test Notes**: Add notes to Slide 0, verify extraction.
3.  **Test Z-Order**: Add Red Rect, then Blue Rect (Blue covers Red). Send Blue to back. Verify XML order.
4.  **Test Targeting**: Replace text "Placeholder" with "Targeted" only on Slide 1. Verify Slide 0 is unchanged.
5.  **Test Reorder**: Shuffle slides, verify title text order matches new sequence.

---

## 📜 Remedial Code Generation

I will now generate the code for the two **NEW** tools (`ppt_add_notes.py`, `ppt_set_z_order.py`) and the **TEST** suite.

### 1. `tools/ppt_add_notes.py`

```python
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
```

### 2. `tools/ppt_set_z_order.py`

```python
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
```

### 3. Remedial Tests: `tests/test_remedial.py`

```python
import pytest
import json
import os
from pathlib import Path
from pptx import Presentation

# Assuming tools are accessible via subprocess or direct import if refactored
# For this plan, we assume direct testing of the generated files functionality

def test_add_notes(tmp_path):
    """Test adding speaker notes."""
    pptx_path = tmp_path / "test_notes.pptx"
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    prs.save(pptx_path)
    
    # Run tool (simulated)
    from tools.ppt_add_notes import add_notes
    result = add_notes(pptx_path, 0, "Hello World", mode="overwrite")
    
    assert result["status"] == "success"
    
    # Verify
    prs_new = Presentation(pptx_path)
    assert prs_new.slides[0].notes_slide.notes_text_frame.text == "Hello World"

def test_z_order(tmp_path):
    """Test z-order manipulation."""
    pptx_path = tmp_path / "test_z.pptx"
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6]) # Blank
    
    # Add Shape 1 (Back)
    s1 = slide.shapes.add_shape(1, 0, 0, 100, 100)
    s1.name = "Shape 1"
    
    # Add Shape 2 (Front)
    s2 = slide.shapes.add_shape(1, 10, 10, 100, 100)
    s2.name = "Shape 2"
    
    prs.save(pptx_path)
    
    # Verify initial order (index)
    prs_check = Presentation(pptx_path)
    assert prs_check.slides[0].shapes[0].name == "Shape 1"
    assert prs_check.slides[0].shapes[1].name == "Shape 2"
    
    # Run tool: Send Shape 2 to back
    from tools.ppt_set_z_order import set_z_order
    # Shape 2 is at index 1
    set_z_order(pptx_path, 0, 1, "send_to_back")
    
    # Verify swap
    prs_final = Presentation(pptx_path)
    assert prs_final.slides[0].shapes[0].name == "Shape 2" # Now at back
    assert prs_final.slides[0].shapes[1].name == "Shape 1" # Now at front
```

### 4. Updated `ppt_replace_text.py` (Enhancement)

I will modify the `replace_text` function in `core/powerpoint_agent_core.py` (conceptually) and the tool wrapper.

**Plan for Tool Modification:**
1.  Update `ppt_replace_text.py` arguments.
2.  Pass `slide_index` and `shape_index` to a new method `replace_text_targeted` in the Core, OR logic-gate the existing `replace_text` tool wrapper to only target specific elements.

*Self-Correction:* For minimizing core changes, the Tool wrapper can handle the filtering.

```python
# In tools/ppt_replace_text.py

def replace_text(..., slide_index=None, shape_index=None):
    # ... inside Agent context ...
    count = 0
    
    # Determine target shapes
    targets = []
    if slide_index is not None:
        slide = agent.get_slide(slide_index)
        if shape_index is not None:
            targets = [slide.shapes[shape_index]]
        else:
            targets = slide.shapes
    else:
        # Global
        for s in agent.prs.slides:
            targets.extend(s.shapes)
            
    # Loop targets and replace
    for shape in targets:
        if not shape.has_text_frame: continue
        # ... replacement logic ...
```

## 🚀 Execution Summary
This remedial plan addresses the "Voice" (Notes) and "Visual Control" (Z-Order) gaps immediately. It upgrades safety on text replacement. This brings the toolkit from "Capable" to "Professional Grade".

Ready to execute creation of:
1. `tools/ppt_add_notes.py`
2. `tools/ppt_set_z_order.py`
3. Updated `tools/ppt_replace_text.py`
4. `tests/test_remedial_tools.py`

