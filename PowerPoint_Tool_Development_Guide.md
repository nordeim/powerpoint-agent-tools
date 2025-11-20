# CONTRIBUTING_TOOLS.md

This handbook is the definitive reference for creating new CLI tools for the PowerPoint Agent suite. It is **self-sufficient**, containing all necessary templates, API references, and patterns required to build a compliant tool without examining the source code of other files.

---

## 1. The Design Contract

All tools must strictly adhere to these 4 principles to ensure compatibility with the AI Agent:

1.  **Atomic & Stateless:** Tools must open a file, perform one specific action, save, and exit. Do not assume previous state.
2.  **CLI Interface:** Use `argparse`. Complex data (positions, lists) must be passed as **JSON strings**.
3.  **JSON Output:**
    *   **STDOUT:** Must contain *only* the final JSON response.
    *   **STDERR:** Use for logging/debugging.
    *   **Exit Codes:** `0` for Success, `1` for Error.
4.  **Path Safety:** All file paths must be validated using `pathlib.Path` before execution.

---

## 2. The Master Template

Copy this code to start a new tool (e.g., `tools/ppt_new_feature.py`).

```python
#!/usr/bin/env python3
"""
[Tool Name]
[Short Description of what the tool does]

Usage:
    uv python ppt_new_feature.py --file deck.pptx --param value --json
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any

# 1. PATH SETUP: Allow importing 'core' without installation
sys.path.insert(0, str(Path(__file__).parent.parent))

# 2. IMPORTS: Bring in the Core Agent and Exceptions
from core.powerpoint_agent_core import (
    PowerPointAgent, 
    PowerPointAgentError, 
    SlideNotFoundError
)

def logic_function(filepath: Path, param: str) -> Dict[str, Any]:
    """
    The main logic handler.
    1. Validate Inputs
    2. Open Agent
    3. Execute Core Method
    4. Save & Return Info
    """
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")

    # Context Manager handles Open/Lock/Close automatically
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # --- CORE LOGIC HERE ---
        # Example: Validating a slide index
        total_slides = agent.get_slide_count()
        if total_slides == 0:
             raise PowerPointAgentError("Presentation is empty")
             
        # Example: Calling a core method
        # result_data = agent.some_core_method(param)
        
        # Save changes
        agent.save()
        
        # Get fresh info for response
        info = agent.get_presentation_info()

    # Return standardized success dictionary
    return {
        "status": "success",
        "file": str(filepath),
        "action_performed": "new_feature",
        "details": {
            "param_used": param,
            "total_slides": info["slide_count"]
        }
    }

def main():
    # 3. ARGUMENT PARSING
    parser = argparse.ArgumentParser(description="Tool Description")
    
    # Standard Argument: File Path
    parser.add_argument(
        '--file', 
        required=True, 
        type=Path, 
        help='PowerPoint file path'
    )
    
    # Custom Arguments
    parser.add_argument(
        '--param', 
        required=True, 
        help='Description of parameter'
    )
    
    # Standard Argument: JSON Flag (Required convention)
    parser.add_argument(
        '--json', 
        action='store_true', 
        default=True, 
        help='Output JSON response'
    )
    
    args = parser.parse_args()
    
    # 4. ERROR HANDLING & OUTPUT
    try:
        result = logic_function(filepath=args.file, param=args.param)
        print(json.dumps(result, indent=2))
        sys.exit(0)
        
    except Exception as e:
        # Standard Error Response Format
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

---

## 3. Data Structures Reference

When passing complex arguments to `PowerPointAgent` methods, use these dictionary schemas.

### **Position Dictionary** (`Dict[str, Any]`)
Used in: `add_text_box`, `insert_image`, `add_chart`, `add_shape`

*   **Percentage (Recommended):** `{"left": "10%", "top": "20%"}`
*   **Absolute (Inches):** `{"left": 1.5, "top": 2.0}`
*   **Anchor:** `{"anchor": "center", "offset_x": 0, "offset_y": -0.5}`
    *   *Anchors:* `top_left`, `top_center`, `top_right`, `center_left`, `center`, `center_right`, `bottom_left`, `bottom_center`, `bottom_right`
*   **Grid:** `{"grid_row": 2, "grid_col": 2, "grid_size": 12}`

### **Size Dictionary** (`Dict[str, Any]`)
Used in: `add_text_box`, `insert_image`, `add_chart`, `add_shape`

*   **Percentage:** `{"width": "50%", "height": "50%"}`
*   **Absolute:** `{"width": 5.0, "height": 3.0}`
*   **Auto (Aspect Ratio):** `{"width": "50%", "height": "auto"}`

### **Colors**
*   **Format:** Hex String `"#FF0000"` or `"#0070C0"`.

---

## 4. Core API Cheatsheet

You do not need to check `powerpoint_agent_core.py`. Use this reference for available methods on the `agent` instance.

### **File & Info**
| Method | Args | Returns |
| :--- | :--- | :--- |
| `create_new()` | `template: Path=None` | `None` |
| `open()` | `filepath: Path` | `None` |
| `save()` | `filepath: Path=None` | `None` |
| `get_slide_count()` | *None* | `int` |
| `get_presentation_info()` | *None* | `Dict` (metadata) |
| `get_slide_info()` | `slide_index: int` | `Dict` (shapes/text) |

### **Slide Manipulation**
| Method | Args | Returns |
| :--- | :--- | :--- |
| `add_slide()` | `layout_name: str, index: int=None` | `int` (new index) |
| `delete_slide()` | `index: int` | `None` |
| `duplicate_slide()` | `index: int` | `int` (new index) |
| `reorder_slides()` | `from_index: int, to_index: int` | `None` |
| `set_slide_layout()` | `slide_index: int, layout_name: str` | `None` |

### **Content Creation**
| Method | Args | Notes |
| :--- | :--- | :--- |
| `add_text_box()` | `slide_index, text, position, size, font_name, font_size, bold, italic, color, alignment` | See Data Structures |
| `add_bullet_list()` | `slide_index, items: List[str], position, size, bullet_style` | Styles: `bullet`, `numbered` |
| `set_title()` | `slide_index, title: str, subtitle: str` | Uses layout placeholders |
| `insert_image()` | `slide_index, image_path, position, size, compress: bool` | Handles `auto` size |
| `add_shape()` | `slide_index, shape_type, position, size, fill_color, line_color` | Types: `rectangle`, `arrow`, etc. |
| `add_chart()` | `slide_index, chart_type, data: Dict, position, size` | Data: `{"categories":[], "series":[]}` |
| `add_table()` | `slide_index, rows, cols, position, size, data: List[List]` | Data is 2D array |

### **Formatting & Editing**
| Method | Args |
| :--- | :--- |
| `format_text()` | `slide_index, shape_index, font_name, font_size, bold, color` |
| `format_shape()` | `slide_index, shape_index, fill_color, line_color, line_width` |
| `replace_text()` | `find: str, replace: str, match_case: bool` |
| `replace_image()` | `slide_index, old_image_name, new_image_path` |
| `set_image_properties()` | `slide_index, shape_index, alt_text, transparency` |

### **Validation**
| Method | Returns |
| :--- | :--- |
| `check_accessibility()` | `Dict` (WCAG issues) |
| `validate_presentation()` | `Dict` (Empty slides, missing assets) |

---

## 5. Implementation Checklist

Before committing a new tool, verify:

1.  [ ] **JSON Argument Parsing:** If your tool takes a `Position` or `Size`, are you parsing the input string via `json.loads` in `argparse`?
2.  [ ] **Exit Codes:** Does it return `0` on success and `1` on exception?
3.  [ ] **File Existence:** Do you check `if not filepath.exists()` before opening?
4.  [ ] **Self-Contained:** Does the tool run without needing to modify `core.py`?
5.  [ ] **Slide Bounds:** Do you check `if not 0 <= index < total_slides`?
