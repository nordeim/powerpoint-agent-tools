# Core Analysis & Sub-Plan Validation

## 1. Meticulous Core Analysis
I have performed a line-by-line analysis of the provided `core/powerpoint_agent_core.py` to verify if it supports the implementation of the requested 10 tools.

**Findings:**
The `PowerPointAgent` class in `core/powerpoint_agent_core.py` is **already fully equipped** with the underlying logic for all 10 proposed tools. The original author anticipated these requirements well.

*   **Accessibility:** `check_accessibility()` is implemented (Lines 1067-1071).
*   **Slide Management:** `delete_slide`, `duplicate_slide`, `reorder_slides`, and `set_slide_layout` are all implemented (Lines 684-736, 1029-1035).
*   **Shape Formatting:** `format_shape` is implemented (Lines 858-880).
*   **Image Properties:** `set_image_properties` is implemented (Lines 966-996).
*   **Cloning:** Can be achieved via `open()` and `save(new_path)` (Lines 642-660).
*   **Notes:** `extract_notes` is implemented (Lines 1122-1136).
*   **Connectors:** `add_connector` is implemented (Lines 915-935).

**Conclusion:**
No modifications to `core/powerpoint_agent_core.py` or `core/__init__.py` are necessary. The core library is robust and ready. I can proceed directly to implementing the CLI tool wrappers.

## 2. Updated Execution Plan
I will generate the following 10 files in the `tools/` directory. Each file will import the core library, handle command-line arguments, invoke the specific core method, and output the result as JSON.

### **Tool List:**
1.  `tools/ppt_check_accessibility.py`
2.  `tools/ppt_delete_slide.py`
3.  `tools/ppt_duplicate_slide.py`
4.  `tools/ppt_reorder_slides.py`
5.  `tools/ppt_set_slide_layout.py`
6.  `tools/ppt_format_shape.py`
7.  `tools/ppt_set_image_properties.py`
8.  `tools/ppt_clone_presentation.py`
9.  `tools/ppt_extract_notes.py`
10. `tools/ppt_add_connector.py`

I will now generate these files.

### 1. `tools/ppt_check_accessibility.py`

```python
#!/usr/bin/env python3
"""
PowerPoint Check Accessibility Tool
Run WCAG 2.1 accessibility checks on presentation

Usage:
    uv python ppt_check_accessibility.py --file presentation.pptx --json

Exit Codes:
    0: Success (even if issues found - check 'status')
    1: Error occurred (file not found, etc)
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError
)


def check_accessibility(filepath: Path) -> Dict[str, Any]:
    """Run accessibility checks."""
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath, acquire_lock=False)  # Read-only check
        result = agent.check_accessibility()
        result["file"] = str(filepath)
    
    return result


def main():
    parser = argparse.ArgumentParser(
        description="Check PowerPoint accessibility (WCAG 2.1)",
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file path'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default)'
    )
    
    args = parser.parse_args()
    
    try:
        result = check_accessibility(filepath=args.file)
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

### 2. `tools/ppt_delete_slide.py`

```python
#!/usr/bin/env python3
"""
PowerPoint Delete Slide Tool
Remove a slide from the presentation

Usage:
    uv python ppt_delete_slide.py --file presentation.pptx --index 1 --json
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


def delete_slide(filepath: Path, index: int) -> Dict[str, Any]:
    """Delete slide at index."""
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        total_slides = agent.get_slide_count()
        if not 0 <= index < total_slides:
            raise SlideNotFoundError(f"Index {index} out of range (0-{total_slides-1})")
            
        agent.delete_slide(index)
        agent.save()
        
        new_count = agent.get_slide_count()
    
    return {
        "status": "success",
        "file": str(filepath),
        "deleted_index": index,
        "remaining_slides": new_count
    }


def main():
    parser = argparse.ArgumentParser(description="Delete PowerPoint slide")
    parser.add_argument('--file', required=True, type=Path, help='PowerPoint file path')
    parser.add_argument('--index', required=True, type=int, help='Slide index to delete (0-based)')
    parser.add_argument('--json', action='store_true', default=True, help='Output JSON')
    
    args = parser.parse_args()
    
    try:
        result = delete_slide(filepath=args.file, index=args.index)
        print(json.dumps(result, indent=2))
        sys.exit(0)
    except Exception as e:
        print(json.dumps({"status": "error", "error": str(e), "error_type": type(e).__name__}, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()
```

### 3. `tools/ppt_duplicate_slide.py`

```python
#!/usr/bin/env python3
"""
PowerPoint Duplicate Slide Tool
Clone an existing slide

Usage:
    uv python ppt_duplicate_slide.py --file presentation.pptx --index 0 --json
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


def duplicate_slide(filepath: Path, index: int) -> Dict[str, Any]:
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        total = agent.get_slide_count()
        if not 0 <= index < total:
            raise SlideNotFoundError(f"Index {index} out of range (0-{total-1})")
            
        new_index = agent.duplicate_slide(index)
        agent.save()
        
        # Get info about new slide
        info = agent.get_slide_info(new_index)
        
    return {
        "status": "success",
        "file": str(filepath),
        "source_index": index,
        "new_slide_index": new_index,
        "total_slides": new_index + 1,
        "layout": info["layout"]
    }


def main():
    parser = argparse.ArgumentParser(description="Duplicate PowerPoint slide")
    parser.add_argument('--file', required=True, type=Path, help='PowerPoint file path')
    parser.add_argument('--index', required=True, type=int, help='Source slide index (0-based)')
    parser.add_argument('--json', action='store_true', default=True, help='Output JSON')
    
    args = parser.parse_args()
    
    try:
        result = duplicate_slide(filepath=args.file, index=args.index)
        print(json.dumps(result, indent=2))
        sys.exit(0)
    except Exception as e:
        print(json.dumps({"status": "error", "error": str(e), "error_type": type(e).__name__}, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()
```

### 4. `tools/ppt_reorder_slides.py`

```python
#!/usr/bin/env python3
"""
PowerPoint Reorder Slides Tool
Move a slide to a new position

Usage:
    uv python ppt_reorder_slides.py --file presentation.pptx --from-index 3 --to-index 1 --json
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


def reorder_slides(filepath: Path, from_index: int, to_index: int) -> Dict[str, Any]:
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        total = agent.get_slide_count()
        if not 0 <= from_index < total:
            raise SlideNotFoundError(f"Source index {from_index} out of range")
        if not 0 <= to_index < total:
            raise SlideNotFoundError(f"Target index {to_index} out of range")
            
        agent.reorder_slides(from_index, to_index)
        agent.save()
        
    return {
        "status": "success",
        "file": str(filepath),
        "moved_from": from_index,
        "moved_to": to_index,
        "total_slides": total
    }


def main():
    parser = argparse.ArgumentParser(description="Reorder PowerPoint slides")
    parser.add_argument('--file', required=True, type=Path, help='PowerPoint file path')
    parser.add_argument('--from-index', required=True, type=int, help='Current slide index')
    parser.add_argument('--to-index', required=True, type=int, help='New slide index')
    parser.add_argument('--json', action='store_true', default=True, help='Output JSON')
    
    args = parser.parse_args()
    
    try:
        result = reorder_slides(filepath=args.file, from_index=args.from_index, to_index=args.to_index)
        print(json.dumps(result, indent=2))
        sys.exit(0)
    except Exception as e:
        print(json.dumps({"status": "error", "error": str(e), "error_type": type(e).__name__}, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()
```

### 5. `tools/ppt_set_slide_layout.py`

```python
#!/usr/bin/env python3
"""
PowerPoint Set Slide Layout Tool
Change the layout of an existing slide

Usage:
    uv python ppt_set_slide_layout.py --file presentation.pptx --slide 0 --layout "Title Only" --json
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError, SlideNotFoundError, LayoutNotFoundError
)


def set_slide_layout(filepath: Path, slide_index: int, layout_name: str) -> Dict[str, Any]:
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        total = agent.get_slide_count()
        if not 0 <= slide_index < total:
            raise SlideNotFoundError(f"Slide index {slide_index} out of range")
            
        # Get available layouts to validate/fuzzy match
        available = agent.get_available_layouts()
        if layout_name not in available:
            # Try fuzzy match
            layout_lower = layout_name.lower()
            for avail in available:
                if layout_lower in avail.lower():
                    layout_name = avail
                    break
            else:
                 raise LayoutNotFoundError(f"Layout '{layout_name}' not found. Available: {available}")

        agent.set_slide_layout(slide_index, layout_name)
        agent.save()
        
    return {
        "status": "success",
        "file": str(filepath),
        "slide_index": slide_index,
        "new_layout": layout_name
    }


def main():
    parser = argparse.ArgumentParser(description="Set PowerPoint slide layout")
    parser.add_argument('--file', required=True, type=Path, help='PowerPoint file path')
    parser.add_argument('--slide', required=True, type=int, help='Slide index (0-based)')
    parser.add_argument('--layout', required=True, help='New layout name')
    parser.add_argument('--json', action='store_true', default=True, help='Output JSON')
    
    args = parser.parse_args()
    
    try:
        result = set_slide_layout(filepath=args.file, slide_index=args.slide, layout_name=args.layout)
        print(json.dumps(result, indent=2))
        sys.exit(0)
    except Exception as e:
        print(json.dumps({"status": "error", "error": str(e), "error_type": type(e).__name__}, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()
```

### 6. `tools/ppt_format_shape.py`

```python
#!/usr/bin/env python3
"""
PowerPoint Format Shape Tool
Update fill color, line color, and line width of an existing shape

Usage:
    uv python ppt_format_shape.py --file presentation.pptx --slide 0 --shape 1 --fill-color "#FF0000" --json
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


def format_shape(
    filepath: Path,
    slide_index: int,
    shape_index: int,
    fill_color: str = None,
    line_color: str = None,
    line_width: float = None
) -> Dict[str, Any]:
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if all(v is None for v in [fill_color, line_color, line_width]):
        raise ValueError("At least one formatting option required")

    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        total = agent.get_slide_count()
        if not 0 <= slide_index < total:
            raise SlideNotFoundError(f"Slide index {slide_index} out of range")
            
        agent.format_shape(
            slide_index=slide_index,
            shape_index=shape_index,
            fill_color=fill_color,
            line_color=line_color,
            line_width=line_width
        )
        agent.save()
        
    return {
        "status": "success",
        "file": str(filepath),
        "slide_index": slide_index,
        "shape_index": shape_index,
        "formatting": {
            "fill_color": fill_color,
            "line_color": line_color,
            "line_width": line_width
        }
    }


def main():
    parser = argparse.ArgumentParser(description="Format PowerPoint shape")
    parser.add_argument('--file', required=True, type=Path, help='PowerPoint file path')
    parser.add_argument('--slide', required=True, type=int, help='Slide index (0-based)')
    parser.add_argument('--shape', required=True, type=int, help='Shape index (0-based)')
    parser.add_argument('--fill-color', help='Fill hex color')
    parser.add_argument('--line-color', help='Line hex color')
    parser.add_argument('--line-width', type=float, help='Line width in points')
    parser.add_argument('--json', action='store_true', default=True, help='Output JSON')
    
    args = parser.parse_args()
    
    try:
        result = format_shape(
            filepath=args.file, 
            slide_index=args.slide, 
            shape_index=args.shape,
            fill_color=args.fill_color,
            line_color=args.line_color,
            line_width=args.line_width
        )
        print(json.dumps(result, indent=2))
        sys.exit(0)
    except Exception as e:
        print(json.dumps({"status": "error", "error": str(e), "error_type": type(e).__name__}, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()
```

### 7. `tools/ppt_set_image_properties.py`

```python
#!/usr/bin/env python3
"""
PowerPoint Set Image Properties Tool
Set alt text and transparency for images

Usage:
    uv python ppt_set_image_properties.py --file deck.pptx --slide 0 --shape 1 --alt-text "Logo" --json
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


def set_image_properties(
    filepath: Path,
    slide_index: int,
    shape_index: int,
    alt_text: str = None,
    transparency: float = None
) -> Dict[str, Any]:
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if alt_text is None and transparency is None:
        raise ValueError("At least alt-text or transparency must be set")

    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        total = agent.get_slide_count()
        if not 0 <= slide_index < total:
            raise SlideNotFoundError(f"Slide index {slide_index} out of range")
            
        agent.set_image_properties(
            slide_index=slide_index,
            shape_index=shape_index,
            alt_text=alt_text,
            transparency=transparency
        )
        agent.save()
        
    return {
        "status": "success",
        "file": str(filepath),
        "slide_index": slide_index,
        "shape_index": shape_index,
        "properties": {
            "alt_text": alt_text,
            "transparency": transparency
        }
    }


def main():
    parser = argparse.ArgumentParser(description="Set PowerPoint image properties")
    parser.add_argument('--file', required=True, type=Path, help='PowerPoint file path')
    parser.add_argument('--slide', required=True, type=int, help='Slide index (0-based)')
    parser.add_argument('--shape', required=True, type=int, help='Shape index (0-based)')
    parser.add_argument('--alt-text', help='Alt text for accessibility')
    parser.add_argument('--transparency', type=float, help='Transparency (0.0-1.0)')
    parser.add_argument('--json', action='store_true', default=True, help='Output JSON')
    
    args = parser.parse_args()
    
    try:
        result = set_image_properties(
            filepath=args.file, 
            slide_index=args.slide, 
            shape_index=args.shape,
            alt_text=args.alt_text,
            transparency=args.transparency
        )
        print(json.dumps(result, indent=2))
        sys.exit(0)
    except Exception as e:
        print(json.dumps({"status": "error", "error": str(e), "error_type": type(e).__name__}, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()
```

### 8. `tools/ppt_clone_presentation.py`

```python
#!/usr/bin/env python3
"""
PowerPoint Clone Presentation Tool
Create an exact copy of a presentation

Usage:
    uv python ppt_clone_presentation.py --source base.pptx --output new_deck.pptx --json
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import PowerPointAgent


def clone_presentation(source: Path, output: Path) -> Dict[str, Any]:
    
    if not source.exists():
        raise FileNotFoundError(f"Source file not found: {source}")
    
    # Validate output path
    if not output.suffix.lower() == '.pptx':
        output = output.with_suffix('.pptx')

    # Uses agent to open and save-as, effectively cloning
    with PowerPointAgent(source) as agent:
        agent.open(source, acquire_lock=False) # Read-only open
        agent.save(output)
        
    return {
        "status": "success",
        "source": str(source),
        "output": str(output),
        "size_bytes": output.stat().st_size
    }


def main():
    parser = argparse.ArgumentParser(description="Clone PowerPoint presentation")
    parser.add_argument('--source', required=True, type=Path, help='Source file')
    parser.add_argument('--output', required=True, type=Path, help='Destination file')
    parser.add_argument('--json', action='store_true', default=True, help='Output JSON')
    
    args = parser.parse_args()
    
    try:
        result = clone_presentation(source=args.source, output=args.output)
        print(json.dumps(result, indent=2))
        sys.exit(0)
    except Exception as e:
        print(json.dumps({"status": "error", "error": str(e), "error_type": type(e).__name__}, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()
```

### 9. `tools/ppt_extract_notes.py`

```python
#!/usr/bin/env python3
"""
PowerPoint Extract Notes Tool
Get speaker notes from all slides

Usage:
    uv python ppt_extract_notes.py --file presentation.pptx --json
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import PowerPointAgent


def extract_notes(filepath: Path) -> Dict[str, Any]:
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath, acquire_lock=False)
        notes = agent.extract_notes()
        
    return {
        "status": "success",
        "file": str(filepath),
        "notes_found": len(notes),
        "notes": notes # Dict {slide_index: text}
    }


def main():
    parser = argparse.ArgumentParser(description="Extract speaker notes")
    parser.add_argument('--file', required=True, type=Path, help='PowerPoint file path')
    parser.add_argument('--json', action='store_true', default=True, help='Output JSON')
    
    args = parser.parse_args()
    
    try:
        result = extract_notes(filepath=args.file)
        print(json.dumps(result, indent=2))
        sys.exit(0)
    except Exception as e:
        print(json.dumps({"status": "error", "error": str(e), "error_type": type(e).__name__}, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()
```

### 10. `tools/ppt_add_connector.py`

```python
#!/usr/bin/env python3
"""
PowerPoint Add Connector Tool
Draw a line between two shapes

Usage:
    uv python ppt_add_connector.py --file deck.pptx --slide 0 --from-shape 0 --to-shape 1 --json
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


def add_connector(
    filepath: Path,
    slide_index: int,
    from_shape: int,
    to_shape: int,
    connector_type: str = "straight"
) -> Dict[str, Any]:
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        total = agent.get_slide_count()
        if not 0 <= slide_index < total:
            raise SlideNotFoundError(f"Slide index {slide_index} out of range")
            
        agent.add_connector(
            slide_index=slide_index,
            from_shape=from_shape,
            to_shape=to_shape,
            connector_type=connector_type
        )
        agent.save()
        
    return {
        "status": "success",
        "file": str(filepath),
        "slide_index": slide_index,
        "connection": {
            "from": from_shape,
            "to": to_shape,
            "type": connector_type
        }
    }


def main():
    parser = argparse.ArgumentParser(description="Add connector between shapes")
    parser.add_argument('--file', required=True, type=Path, help='PowerPoint file path')
    parser.add_argument('--slide', required=True, type=int, help='Slide index (0-based)')
    parser.add_argument('--from-shape', required=True, type=int, help='Start shape index')
    parser.add_argument('--to-shape', required=True, type=int, help='End shape index')
    parser.add_argument('--type', default='straight', help='Connector type (straight)')
    parser.add_argument('--json', action='store_true', default=True, help='Output JSON')
    
    args = parser.parse_args()
    
    try:
        result = add_connector(
            filepath=args.file,
            slide_index=args.slide,
            from_shape=args.from_shape,
            to_shape=args.to_shape,
            connector_type=args.type
        )
        print(json.dumps(result, indent=2))
        sys.exit(0)
    except Exception as e:
        print(json.dumps({"status": "error", "error": str(e), "error_type": type(e).__name__}, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()
```
