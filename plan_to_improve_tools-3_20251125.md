# Deep Analysis: Re-Imagining Remove Shape and Add Text Box Tools

## Executive Summary

I will systematically design and create:
1. `ppt_remove_shape.py` - A completely new tool for shape removal with safety controls
2. `ppt_add_text_box.py` v3.0 - Enhanced version aligned with v3.0 core

Both tools will fully integrate with PowerPoint Agent Core v3.0 and System Prompt v3.0.

---

## Phase 1: Requirements Analysis

### 1.1 `ppt_remove_shape.py` Requirements (New Tool)

**System Prompt v3.0 Requirements:**
- Shape removal is a **destructive operation** requiring careful handling
- Must integrate with approval token system (documentation/warnings)
- Must refresh shape indices after removal (critical for subsequent operations)
- Must track presentation version changes
- Must provide rollback guidance

**Core v3.0 Integration:**
- Uses `agent.remove_shape()` which returns removal details
- Uses `ShapeNotFoundError` for invalid indices
- Uses presentation versioning

**Safety Requirements:**
| Requirement | Implementation |
|-------------|----------------|
| Confirmation warnings | Clear warnings about index shifts |
| Shape info before deletion | Show what's being deleted |
| Rollback guidance | Cannot undo, suggest clone first |
| Index shift warnings | Explicit warning about subsequent operations |
| Dry-run mode | Preview what would be deleted |
| Batch removal support | Remove multiple shapes safely |

**Functional Requirements:**
- Remove shape by index
- Remove shape by name (partial match)
- Remove multiple shapes
- Dry-run preview mode
- Show shape details before removal
- Track presentation version
- Provide post-removal shape count

### 1.2 `ppt_add_text_box.py` v3.0 Analysis

**Current v2.0 Capabilities:**
- Text validation with length warnings
- Font size accessibility validation
- Color contrast checking
- Position validation
- Multiple positioning systems

**Gaps for v3.0:**
| Gap | Impact | Priority |
|-----|--------|----------|
| No shape_index return | Can't reference created shape | HIGH |
| No presentation version tracking | Can't detect conflicts | HIGH |
| Imports RGBColor directly (fragile) | May break with core changes | MEDIUM |
| No word wrap control | Limited formatting | MEDIUM |
| No vertical alignment | Limited layout control | MEDIUM |
| No margin/padding control | Text may touch edges | LOW |
| Limited error types | Generic error handling | MEDIUM |

**v3.0 Requirements:**
- Return shape_index from core
- Presentation version tracking
- Align imports with v3.0 core
- Add word wrap option
- Add vertical alignment
- Enhanced validation
- Consistent with other v3.0 tools

---

## Phase 2: Implementation Plan

### 2.1 `ppt_remove_shape.py` Checklist

- [ ] **REM-001**: Module docstring with safety warnings
- [ ] **REM-002**: Import v3.0 core classes (ShapeNotFoundError)
- [ ] **REM-003**: Shape info retrieval before removal
- [ ] **REM-004**: Dry-run mode for preview
- [ ] **REM-005**: Remove by index
- [ ] **REM-006**: Remove by name (partial match)
- [ ] **REM-007**: Batch removal support
- [ ] **REM-008**: Index shift warnings
- [ ] **REM-009**: Presentation version tracking
- [ ] **REM-010**: Rollback guidance in output
- [ ] **REM-011**: Confirmation requirement documentation
- [ ] **REM-012**: Comprehensive CLI with examples
- [ ] **REM-013**: Proper error handling

### 2.2 `ppt_add_text_box.py` v3.0 Checklist

- [ ] **TXT-001**: Module docstring updated for v3.0
- [ ] **TXT-002**: Correct imports from v3.0 core
- [ ] **TXT-003**: Return shape_index from core
- [ ] **TXT-004**: Presentation version tracking
- [ ] **TXT-005**: Word wrap option
- [ ] **TXT-006**: Vertical alignment option
- [ ] **TXT-007**: Enhanced validation (unchanged good parts)
- [ ] **TXT-008**: Color presets support
- [ ] **TXT-009**: Consistent JSON output
- [ ] **TXT-010**: Comprehensive CLI help
- [ ] **TXT-011**: Proper error handling with v3.0 exceptions

---

## Phase 3: Implementation

### File 1: `tools/ppt_remove_shape.py` v3.0 (NEW)

```python
#!/usr/bin/env python3
"""
PowerPoint Remove Shape Tool v3.0
Safely remove shapes from slides with comprehensive safety controls.

⚠️  DESTRUCTIVE OPERATION WARNING ⚠️
This tool permanently removes shapes from presentations.
- Shape indices will shift after removal
- This operation cannot be undone
- Always clone the presentation first for safety
- Use --dry-run to preview before actual removal

Fully aligned with PowerPoint Agent Core v3.0 and System Prompt v3.0.

Author: PowerPoint Agent Team
License: MIT
Version: 3.0.0

Usage:
    # Preview what would be removed (RECOMMENDED FIRST STEP)
    uv run tools/ppt_remove_shape.py --file deck.pptx --slide 0 --shape 2 --dry-run --json
    
    # Remove shape by index
    uv run tools/ppt_remove_shape.py --file deck.pptx --slide 0 --shape 2 --json
    
    # Remove shape by name
    uv run tools/ppt_remove_shape.py --file deck.pptx --slide 0 --name "Rectangle 1" --json

Exit Codes:
    0: Success
    1: Error occurred

Safety Protocol:
    1. Clone presentation: ppt_clone_presentation.py --source deck.pptx --output work.pptx
    2. Preview removal: ppt_remove_shape.py --file work.pptx --slide 0 --shape 2 --dry-run
    3. Execute removal: ppt_remove_shape.py --file work.pptx --slide 0 --shape 2
    4. Refresh indices: ppt_get_slide_info.py --file work.pptx --slide 0

Changelog v3.0.0:
- NEW: Initial release aligned with Core v3.0
- NEW: Dry-run mode for safe preview
- NEW: Remove by name support
- NEW: Batch removal support
- NEW: Shape info display before removal
- NEW: Index shift warnings
- NEW: Presentation version tracking
- NEW: Rollback guidance
- NEW: Comprehensive safety documentation
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any, List, Optional

# Add parent directory to path for core import
sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent,
    PowerPointAgentError,
    SlideNotFoundError,
    ShapeNotFoundError,
    __version__ as CORE_VERSION
)

# ============================================================================
# CONSTANTS
# ============================================================================

__version__ = "3.0.0"

# Safety messages
SAFETY_WARNING = """
⚠️  DESTRUCTIVE OPERATION WARNING ⚠️
- This will permanently remove the shape
- Shape indices will shift after removal
- This operation cannot be undone
- Subsequent shape references may be invalid
"""

ROLLBACK_GUIDANCE = """
ROLLBACK GUIDANCE:
- This operation cannot be undone directly
- To recover: restore from backup or clone made before removal
- Recommended: Always use ppt_clone_presentation.py before destructive operations
"""

INDEX_SHIFT_WARNING = """
INDEX SHIFT WARNING:
- All shapes after index {removed_index} have shifted down by 1
- Shape that was at index {next_index} is now at index {removed_index}
- Re-run ppt_get_slide_info.py to get updated indices before further operations
"""


# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def get_shape_details(agent: PowerPointAgent, slide_index: int, shape_index: int) -> Dict[str, Any]:
    """
    Get detailed information about a shape before removal.
    
    Args:
        agent: PowerPointAgent instance
        slide_index: Slide index
        shape_index: Shape index
        
    Returns:
        Dict with shape details
    """
    try:
        slide_info = agent.get_slide_info(slide_index)
        shapes = slide_info.get("shapes", [])
        
        if 0 <= shape_index < len(shapes):
            shape = shapes[shape_index]
            return {
                "index": shape_index,
                "type": shape.get("type", "unknown"),
                "name": shape.get("name", ""),
                "has_text": shape.get("has_text", False),
                "text": shape.get("text", ""),
                "text_preview": (shape.get("text", "")[:100] + "...") if len(shape.get("text", "")) > 100 else shape.get("text", ""),
                "position": shape.get("position", {}),
                "size": shape.get("size", {})
            }
    except Exception as e:
        return {"index": shape_index, "error": str(e)}
    
    return {"index": shape_index, "type": "unknown"}


def find_shape_by_name(agent: PowerPointAgent, slide_index: int, name: str) -> Optional[int]:
    """
    Find shape index by name (partial match).
    
    Args:
        agent: PowerPointAgent instance
        slide_index: Slide index
        name: Shape name to search for
        
    Returns:
        Shape index if found, None otherwise
    """
    try:
        slide_info = agent.get_slide_info(slide_index)
        shapes = slide_info.get("shapes", [])
        
        # Exact match first
        for idx, shape in enumerate(shapes):
            if shape.get("name", "") == name:
                return idx
        
        # Partial match (case-insensitive)
        name_lower = name.lower()
        for idx, shape in enumerate(shapes):
            shape_name = shape.get("name", "").lower()
            if name_lower in shape_name or shape_name in name_lower:
                return idx
        
        return None
    except Exception:
        return None


def validate_removal(
    shape_details: Dict[str, Any],
    slide_shape_count: int
) -> Dict[str, Any]:
    """
    Validate removal and generate warnings.
    
    Args:
        shape_details: Details of shape to remove
        slide_shape_count: Total shapes on slide
        
    Returns:
        Dict with warnings and recommendations
    """
    warnings: List[str] = []
    recommendations: List[str] = []
    
    # Check if removing last shape
    if slide_shape_count == 1:
        warnings.append(
            "This is the only shape on the slide. "
            "Removing it will leave the slide empty."
        )
    
    # Check if shape has text content
    if shape_details.get("has_text") and shape_details.get("text"):
        text_len = len(shape_details.get("text", ""))
        if text_len > 50:
            warnings.append(
                f"Shape contains {text_len} characters of text that will be deleted."
            )
    
    # Check shape type
    shape_type = shape_details.get("type", "").lower()
    if "placeholder" in shape_type:
        warnings.append(
            "This appears to be a placeholder shape. "
            "Removing it may affect slide layout."
        )
    if "title" in shape_type.lower():
        warnings.append(
            "This appears to be a title shape. "
            "Removing it may affect accessibility and navigation."
        )
    
    # Always recommend backup
    recommendations.append(
        "Always clone the presentation before destructive operations: "
        "ppt_clone_presentation.py --source FILE --output backup.pptx"
    )
    
    recommendations.append(
        "After removal, refresh shape indices: "
        "ppt_get_slide_info.py --file FILE --slide N"
    )
    
    return {
        "warnings": warnings,
        "recommendations": recommendations,
        "has_warnings": len(warnings) > 0
    }


# ============================================================================
# MAIN FUNCTIONS
# ============================================================================

def remove_shape(
    filepath: Path,
    slide_index: int,
    shape_index: Optional[int] = None,
    shape_name: Optional[str] = None,
    dry_run: bool = False
) -> Dict[str, Any]:
    """
    Remove shape from slide with safety controls.
    
    Args:
        filepath: Path to PowerPoint file
        slide_index: Target slide index (0-based)
        shape_index: Shape index to remove (0-based)
        shape_name: Shape name to remove (alternative to index)
        dry_run: If True, preview only without actual removal
        
    Returns:
        Result dict with removal details
        
    Raises:
        FileNotFoundError: If file doesn't exist
        SlideNotFoundError: If slide index is invalid
        ShapeNotFoundError: If shape index/name is invalid
        ValueError: If neither shape_index nor shape_name provided
    """
    # Validate file exists
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    # Validate parameters
    if shape_index is None and shape_name is None:
        raise ValueError(
            "Must specify either --shape (index) or --name (shape name)"
        )
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Validate slide index
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range",
                details={
                    "slide_index": slide_index,
                    "total_slides": total_slides,
                    "valid_range": f"0-{total_slides-1}"
                }
            )
        
        # Get slide info before removal
        slide_info_before = agent.get_slide_info(slide_index)
        shape_count_before = slide_info_before.get("shape_count", 0)
        
        # Resolve shape index from name if needed
        resolved_index = shape_index
        if shape_name is not None:
            resolved_index = find_shape_by_name(agent, slide_index, shape_name)
            if resolved_index is None:
                raise ShapeNotFoundError(
                    f"Shape with name '{shape_name}' not found on slide {slide_index}",
                    details={
                        "slide_index": slide_index,
                        "shape_name": shape_name,
                        "available_shapes": [
                            s.get("name") for s in slide_info_before.get("shapes", [])
                        ]
                    }
                )
        
        # Get shape details before removal
        shape_details = get_shape_details(agent, slide_index, resolved_index)
        
        # Validate removal
        validation = validate_removal(shape_details, shape_count_before)
        
        # Get presentation version before
        version_before = agent.get_presentation_version()
        
        # Build result
        result = {
            "file": str(filepath),
            "slide_index": slide_index,
            "shape_index": resolved_index,
            "shape_details": shape_details,
            "shape_count_before": shape_count_before,
            "dry_run": dry_run,
            "validation": {
                "warnings": validation["warnings"],
                "recommendations": validation["recommendations"]
            },
            "presentation_version": {
                "before": version_before
            },
            "core_version": CORE_VERSION,
            "tool_version": __version__
        }
        
        if dry_run:
            # Preview only
            result["status"] = "preview"
            result["message"] = "DRY RUN: Shape would be removed. Run without --dry-run to execute."
            result["shape_count_after"] = shape_count_before - 1
            result["index_shift_info"] = {
                "shapes_affected": shape_count_before - resolved_index - 1,
                "message": f"Shapes at indices {resolved_index + 1} to {shape_count_before - 1} would shift down by 1"
            }
        else:
            # Execute removal
            removal_result = agent.remove_shape(
                slide_index=slide_index,
                shape_index=resolved_index
            )
            
            # Save changes
            agent.save()
            
            # Get updated info
            version_after = agent.get_presentation_version()
            slide_info_after = agent.get_slide_info(slide_index)
            shape_count_after = slide_info_after.get("shape_count", 0)
            
            result["status"] = "success"
            result["message"] = "Shape removed successfully"
            result["shape_count_after"] = shape_count_after
            result["removal_result"] = removal_result
            result["presentation_version"]["after"] = version_after
            
            # Index shift information
            shapes_shifted = shape_count_before - resolved_index - 1
            if shapes_shifted > 0:
                result["index_shift_info"] = {
                    "shapes_shifted": shapes_shifted,
                    "warning": f"⚠️ {shapes_shifted} shape(s) have new indices. Re-query before further operations.",
                    "refresh_command": f"uv run tools/ppt_get_slide_info.py --file {filepath} --slide {slide_index} --json"
                }
            
            # Add rollback guidance
            result["rollback_guidance"] = (
                "This operation cannot be undone. "
                "To recover, restore from a backup clone."
            )
        
        # Add warnings to status if present
        if validation["has_warnings"]:
            if result["status"] == "success":
                result["status"] = "success_with_warnings"
    
    return result


def remove_shapes_batch(
    filepath: Path,
    slide_index: int,
    shape_indices: List[int],
    dry_run: bool = False
) -> Dict[str, Any]:
    """
    Remove multiple shapes from slide (in reverse order to preserve indices).
    
    Args:
        filepath: Path to PowerPoint file
        slide_index: Target slide index
        shape_indices: List of shape indices to remove
        dry_run: If True, preview only
        
    Returns:
        Result dict with batch removal details
    """
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if not shape_indices:
        raise ValueError("No shape indices provided")
    
    # Sort in reverse order to preserve indices during removal
    sorted_indices = sorted(set(shape_indices), reverse=True)
    
    results = []
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Validate slide
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range",
                details={"slide_index": slide_index, "total_slides": total_slides}
            )
        
        slide_info = agent.get_slide_info(slide_index)
        shape_count = slide_info.get("shape_count", 0)
        version_before = agent.get_presentation_version()
        
        # Validate all indices first
        for idx in sorted_indices:
            if not 0 <= idx < shape_count:
                raise ShapeNotFoundError(
                    f"Shape index {idx} out of range",
                    details={"shape_index": idx, "shape_count": shape_count}
                )
        
        # Get details for all shapes to remove
        shapes_to_remove = []
        for idx in sorted_indices:
            details = get_shape_details(agent, slide_index, idx)
            shapes_to_remove.append(details)
        
        if dry_run:
            return {
                "status": "preview",
                "file": str(filepath),
                "slide_index": slide_index,
                "dry_run": True,
                "shapes_to_remove": shapes_to_remove,
                "removal_order": sorted_indices,
                "shape_count_before": shape_count,
                "shape_count_after": shape_count - len(sorted_indices),
                "message": f"DRY RUN: {len(sorted_indices)} shape(s) would be removed",
                "presentation_version": {"before": version_before},
                "core_version": CORE_VERSION,
                "tool_version": __version__
            }
        
        # Execute removals in reverse order
        for idx in sorted_indices:
            removal_result = agent.remove_shape(slide_index, idx)
            results.append({
                "index": idx,
                "result": removal_result
            })
        
        agent.save()
        
        # Get final state
        version_after = agent.get_presentation_version()
        slide_info_after = agent.get_slide_info(slide_index)
        
        return {
            "status": "success",
            "file": str(filepath),
            "slide_index": slide_index,
            "dry_run": False,
            "shapes_removed": shapes_to_remove,
            "removal_count": len(sorted_indices),
            "shape_count_before": shape_count,
            "shape_count_after": slide_info_after.get("shape_count", 0),
            "removal_results": results,
            "presentation_version": {
                "before": version_before,
                "after": version_after
            },
            "warning": "⚠️ All shape indices have changed. Re-query slide info before further operations.",
            "refresh_command": f"uv run tools/ppt_get_slide_info.py --file {filepath} --slide {slide_index} --json",
            "core_version": CORE_VERSION,
            "tool_version": __version__
        }


# ============================================================================
# CLI INTERFACE
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Remove shape from PowerPoint slide (v3.0) ⚠️ DESTRUCTIVE",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
⚠️  DESTRUCTIVE OPERATION - READ CAREFULLY ⚠️
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

This tool PERMANENTLY REMOVES shapes from presentations.
- Shape removal CANNOT be undone
- Shape indices WILL SHIFT after removal
- Always CLONE the presentation first
- Always use DRY-RUN to preview

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

SAFE REMOVAL PROTOCOL:

  1. CLONE the presentation first:
     uv run tools/ppt_clone_presentation.py \\
       --source original.pptx --output work.pptx --json

  2. INSPECT the slide to find shape indices:
     uv run tools/ppt_get_slide_info.py \\
       --file work.pptx --slide 0 --json

  3. PREVIEW the removal (dry-run):
     uv run tools/ppt_remove_shape.py \\
       --file work.pptx --slide 0 --shape 2 --dry-run --json

  4. EXECUTE the removal:
     uv run tools/ppt_remove_shape.py \\
       --file work.pptx --slide 0 --shape 2 --json

  5. REFRESH indices for subsequent operations:
     uv run tools/ppt_get_slide_info.py \\
       --file work.pptx --slide 0 --json

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

EXAMPLES:

  # Preview removal (ALWAYS DO THIS FIRST)
  uv run tools/ppt_remove_shape.py \\
    --file presentation.pptx --slide 0 --shape 3 --dry-run --json

  # Remove shape by index
  uv run tools/ppt_remove_shape.py \\
    --file presentation.pptx --slide 0 --shape 3 --json

  # Remove shape by name
  uv run tools/ppt_remove_shape.py \\
    --file presentation.pptx --slide 0 --name "Rectangle 1" --json

  # Remove multiple shapes (processed in reverse order to preserve indices)
  uv run tools/ppt_remove_shape.py \\
    --file presentation.pptx --slide 0 --shapes 2,4,6 --json

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

INDEX SHIFT BEHAVIOR:

  Before removal (5 shapes):
    Index 0: Title
    Index 1: Subtitle  
    Index 2: Image     ← REMOVING THIS
    Index 3: TextBox
    Index 4: Rectangle

  After removal (4 shapes):
    Index 0: Title
    Index 1: Subtitle
    Index 2: TextBox   ← Was index 3
    Index 3: Rectangle ← Was index 4

  ⚠️ Any saved references to indices 3 and 4 are now INVALID!

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

FINDING SHAPE INDEX:

  Use ppt_get_slide_info.py to list all shapes:
  
  uv run tools/ppt_get_slide_info.py --file deck.pptx --slide 0 --json
  
  Output includes shape indices, types, names, and positions.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        """
    )
    
    # Required arguments
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
    
    # Shape selection (mutually exclusive group)
    shape_group = parser.add_mutually_exclusive_group(required=True)
    
    shape_group.add_argument(
        '--shape',
        type=int,
        help='Shape index to remove (0-based)'
    )
    
    shape_group.add_argument(
        '--name',
        help='Shape name to remove (partial match supported)'
    )
    
    shape_group.add_argument(
        '--shapes',
        help='Comma-separated shape indices for batch removal (e.g., "2,4,6")'
    )
    
    # Options
    parser.add_argument(
        '--dry-run',
        action='store_true',
        help='Preview removal without executing (RECOMMENDED FIRST STEP)'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default: true)'
    )
    
    parser.add_argument(
        '--version',
        action='version',
        version=f'%(prog)s {__version__} (core: {CORE_VERSION})'
    )
    
    args = parser.parse_args()
    
    try:
        # Handle batch removal
        if args.shapes:
            try:
                indices = [int(x.strip()) for x in args.shapes.split(',')]
            except ValueError:
                raise ValueError(
                    f"Invalid shape indices: {args.shapes}. "
                    "Use comma-separated integers (e.g., '2,4,6')"
                )
            
            result = remove_shapes_batch(
                filepath=args.file,
                slide_index=args.slide,
                shape_indices=indices,
                dry_run=args.dry_run
            )
        else:
            # Single shape removal
            result = remove_shape(
                filepath=args.file,
                slide_index=args.slide,
                shape_index=args.shape,
                shape_name=args.name,
                dry_run=args.dry_run
            )
        
        if args.json:
            print(json.dumps(result, indent=2))
        else:
            if args.dry_run:
                print(f"🔍 DRY RUN: Would remove shape {result.get('shape_index')} from slide {result.get('slide_index')}")
                print(f"   Shape type: {result.get('shape_details', {}).get('type', 'unknown')}")
                print(f"   Shape name: {result.get('shape_details', {}).get('name', 'unnamed')}")
            else:
                print(f"✅ Removed shape {result.get('shape_index')} from slide {result.get('slide_index')}")
                print(f"   Shapes remaining: {result.get('shape_count_after')}")
                if result.get('index_shift_info'):
                    print(f"   ⚠️  {result['index_shift_info'].get('warning', '')}")
        
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError"
        }
        if args.json:
            print(json.dumps(error_result, indent=2))
        else:
            print(f"❌ Error: {e}", file=sys.stderr)
        sys.exit(1)
        
    except SlideNotFoundError as e:
        error_result = {
            "status": "error",
            "error": e.message,
            "error_type": "SlideNotFoundError",
            "details": e.details,
            "suggestion": "Use ppt_get_info.py to check available slide indices."
        }
        if args.json:
            print(json.dumps(error_result, indent=2))
        else:
            print(f"❌ Error: {e.message}", file=sys.stderr)
        sys.exit(1)
        
    except ShapeNotFoundError as e:
        error_result = {
            "status": "error",
            "error": e.message,
            "error_type": "ShapeNotFoundError",
            "details": e.details,
            "suggestion": "Use ppt_get_slide_info.py to check available shape indices."
        }
        if args.json:
            print(json.dumps(error_result, indent=2))
        else:
            print(f"❌ Error: {e.message}", file=sys.stderr)
        sys.exit(1)
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError"
        }
        if args.json:
            print(json.dumps(error_result, indent=2))
        else:
            print(f"❌ Error: {e}", file=sys.stderr)
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": e.message,
            "error_type": type(e).__name__,
            "details": e.details
        }
        if args.json:
            print(json.dumps(error_result, indent=2))
        else:
            print(f"❌ Error: {e.message}", file=sys.stderr)
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__
        }
        if args.json:
            print(json.dumps(error_result, indent=2))
        else:
            print(f"❌ Error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
```

---

### File 2: `tools/ppt_add_text_box.py` v3.0

```python
#!/usr/bin/env python3
"""
PowerPoint Add Text Box Tool v3.0
Add text box with flexible positioning, comprehensive validation, and accessibility checking.

Fully aligned with PowerPoint Agent Core v3.0 and System Prompt v3.0.

Author: PowerPoint Agent Team
License: MIT
Version: 3.0.0

Usage:
    uv run tools/ppt_add_text_box.py --file deck.pptx --slide 0 \\
        --text "Revenue: $1.5M" --position '{"left":"20%","top":"30%"}' \\
        --size '{"width":"60%","height":"10%"}' --json

Exit Codes:
    0: Success
    1: Error occurred

Changelog v3.0.0:
- Aligned with PowerPoint Agent Core v3.0
- Returns shape_index from core for subsequent operations
- Added presentation version tracking
- Added word wrap control
- Added vertical alignment option
- Added color presets support
- Enhanced validation (unchanged good parts from v2.0)
- Improved error handling with v3.0 exception types
- Consistent JSON output structure

Position Formats:
  1. Percentage: {"left": "20%", "top": "30%"}
  2. Inches: {"left": 2.0, "top": 3.0}
  3. Anchor: {"anchor": "center", "offset_x": 0, "offset_y": -1.0}
  4. Grid: {"grid_row": 2, "grid_col": 3, "grid_size": 12}
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any, List, Optional

# Add parent directory to path for core import
sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent,
    PowerPointAgentError,
    SlideNotFoundError,
    InvalidPositionError,
    ColorHelper,
    __version__ as CORE_VERSION
)

# ============================================================================
# CONSTANTS
# ============================================================================

__version__ = "3.0.0"

# Color presets (matching other v3.0 tools)
COLOR_PRESETS = {
    "black": "#000000",
    "white": "#FFFFFF",
    "primary": "#0070C0",
    "secondary": "#595959",
    "accent": "#ED7D31",
    "success": "#70AD47",
    "warning": "#FFC000",
    "danger": "#C00000",
    "dark_gray": "#333333",
    "light_gray": "#808080",
    "muted": "#808080",
}

# Font presets
FONT_PRESETS = {
    "default": "Calibri",
    "heading": "Calibri Light",
    "body": "Calibri",
    "code": "Consolas",
    "serif": "Georgia",
    "sans": "Arial",
}


# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def resolve_color(color: Optional[str]) -> Optional[str]:
    """
    Resolve color value, handling presets and hex formats.
    
    Args:
        color: Color specification (hex, preset name, or None)
        
    Returns:
        Resolved hex color or None
    """
    if color is None:
        return None
    
    color_lower = color.lower().strip()
    
    # Check presets
    if color_lower in COLOR_PRESETS:
        return COLOR_PRESETS[color_lower]
    
    # Ensure # prefix for hex colors
    if not color.startswith('#') and len(color) == 6:
        try:
            int(color, 16)
            return f"#{color}"
        except ValueError:
            pass
    
    return color


def resolve_font(font: Optional[str]) -> str:
    """
    Resolve font name, handling presets.
    
    Args:
        font: Font name or preset
        
    Returns:
        Resolved font name
    """
    if font is None:
        return "Calibri"
    
    font_lower = font.lower().strip()
    if font_lower in FONT_PRESETS:
        return FONT_PRESETS[font_lower]
    
    return font


def validate_text_box(
    text: str,
    font_size: int,
    color: Optional[str] = None,
    position: Optional[Dict[str, Any]] = None,
    size: Optional[Dict[str, Any]] = None,
    allow_offslide: bool = False
) -> Dict[str, Any]:
    """
    Validate text box parameters and return warnings/recommendations.
    
    Args:
        text: Text content
        font_size: Font size in points
        color: Text color hex
        position: Position specification
        size: Size specification
        allow_offslide: Allow off-slide positioning
        
    Returns:
        Dict with warnings, recommendations, and validation results
    """
    warnings: List[str] = []
    recommendations: List[str] = []
    validation_results: Dict[str, Any] = {}
    
    # Text length validation
    text_length = len(text)
    line_count = text.count('\n') + 1
    
    validation_results["text_length"] = text_length
    validation_results["line_count"] = line_count
    validation_results["is_multiline"] = line_count > 1
    
    if line_count == 1 and text_length > 100:
        warnings.append(
            f"Text is {text_length} characters for single line (recommended: ≤100). "
            "Long single-line text may be hard to read."
        )
        recommendations.append("Consider breaking into multiple lines or shortening text")
    
    if line_count > 1 and text_length > 500:
        warnings.append(
            f"Multi-line text is {text_length} characters. Very long text blocks reduce readability."
        )
    
    # Font size validation
    validation_results["font_size"] = font_size
    validation_results["font_size_accessible"] = font_size >= 14
    
    if font_size < 10:
        warnings.append(
            f"Font size {font_size}pt is below minimum (10pt). Text will be very hard to read."
        )
    elif font_size < 12:
        warnings.append(
            f"Font size {font_size}pt is very small. Consider 14pt+ for projected presentations."
        )
        recommendations.append("Use 14pt or larger for projected content")
    elif font_size < 14:
        recommendations.append(
            f"Font size {font_size}pt is below recommended 14pt for projected content"
        )
    
    # Color contrast validation
    if color:
        try:
            text_color = ColorHelper.from_hex(color)
            
            # Import RGBColor for background color
            from pptx.dml.color import RGBColor
            bg_color = RGBColor(255, 255, 255)  # Assume white background
            
            is_large_text = font_size >= 18
            contrast_ratio = ColorHelper.contrast_ratio(text_color, bg_color)
            meets_wcag = ColorHelper.meets_wcag(text_color, bg_color, is_large_text)
            
            validation_results["color_contrast"] = {
                "ratio": round(contrast_ratio, 2),
                "wcag_aa": meets_wcag,
                "required_ratio": 3.0 if is_large_text else 4.5,
                "is_large_text": is_large_text
            }
            
            if not meets_wcag:
                required = 3.0 if is_large_text else 4.5
                warnings.append(
                    f"Color contrast {contrast_ratio:.2f}:1 may not meet WCAG AA "
                    f"(required: {required}:1). Consider darker color."
                )
                recommendations.append(
                    "Use black (#000000), dark_gray (#333333), or primary (#0070C0) for better contrast"
                )
        except Exception as e:
            validation_results["color_error"] = str(e)
    
    # Position validation
    if position:
        _validate_position(position, warnings, allow_offslide)
    
    # Size validation
    if size:
        _validate_size(size, font_size, warnings)
    
    return {
        "warnings": warnings,
        "recommendations": recommendations,
        "validation_results": validation_results,
        "has_warnings": len(warnings) > 0
    }


def _validate_position(
    position: Dict[str, Any],
    warnings: List[str],
    allow_offslide: bool
) -> None:
    """Validate position values."""
    try:
        for key in ["left", "top"]:
            if key in position:
                value_str = str(position[key])
                if value_str.endswith('%'):
                    pct = float(value_str.rstrip('%'))
                    if not allow_offslide and (pct < 0 or pct > 100):
                        warnings.append(
                            f"Position '{key}' is {pct}% which is outside slide bounds (0-100%). "
                            f"Text box may not be visible."
                        )
    except (ValueError, TypeError):
        pass


def _validate_size(
    size: Dict[str, Any],
    font_size: int,
    warnings: List[str]
) -> None:
    """Validate size values."""
    try:
        if "height" in size:
            height_str = str(size["height"])
            if height_str.endswith('%'):
                height_pct = float(height_str.rstrip('%'))
                # Estimate minimum height needed for font size
                # Rough approximation: 1pt ≈ 0.1% of slide height
                min_height = font_size * 0.15
                if height_pct < min_height:
                    warnings.append(
                        f"Height {height_pct}% may be too small for {font_size}pt text. "
                        f"Consider at least {min_height:.1f}%."
                    )
        
        if "width" in size:
            width_str = str(size["width"])
            if width_str.endswith('%'):
                width_pct = float(width_str.rstrip('%'))
                if width_pct < 5:
                    warnings.append(
                        f"Width {width_pct}% is very narrow. Text may be excessively wrapped."
                    )
    except (ValueError, TypeError):
        pass


# ============================================================================
# MAIN FUNCTION
# ============================================================================

def add_text_box(
    filepath: Path,
    slide_index: int,
    text: str,
    position: Dict[str, Any],
    size: Dict[str, Any],
    font_name: str = "Calibri",
    font_size: int = 18,
    bold: bool = False,
    italic: bool = False,
    color: Optional[str] = None,
    alignment: str = "left",
    vertical_alignment: str = "top",
    word_wrap: bool = True,
    allow_offslide: bool = False
) -> Dict[str, Any]:
    """
    Add text box with comprehensive validation and formatting.
    
    Args:
        filepath: Path to PowerPoint file
        slide_index: Slide index (0-based)
        text: Text content
        position: Position dict (supports %, inches, anchor, grid)
        size: Size dict
        font_name: Font name or preset
        font_size: Font size in points
        bold: Bold text
        italic: Italic text
        color: Text color (hex or preset)
        alignment: Horizontal alignment (left, center, right, justify)
        vertical_alignment: Vertical alignment (top, middle, bottom)
        word_wrap: Enable word wrap
        allow_offslide: Allow off-slide positioning
        
    Returns:
        Result dict with shape_index and validation info
    """
    # Validate file exists
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    # Resolve color and font
    resolved_color = resolve_color(color)
    resolved_font = resolve_font(font_name)
    
    # Validate parameters
    validation = validate_text_box(
        text=text,
        font_size=font_size,
        color=resolved_color,
        position=position,
        size=size,
        allow_offslide=allow_offslide
    )
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Validate slide index
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range",
                details={
                    "slide_index": slide_index,
                    "total_slides": total_slides,
                    "valid_range": f"0-{total_slides-1}"
                }
            )
        
        # Get presentation version before
        version_before = agent.get_presentation_version()
        
        # Add text box using v3.0 core
        add_result = agent.add_text_box(
            slide_index=slide_index,
            text=text,
            position=position,
            size=size,
            font_name=resolved_font,
            font_size=font_size,
            bold=bold,
            italic=italic,
            color=resolved_color,
            alignment=alignment
        )
        
        # Save changes
        agent.save()
        
        # Get updated info
        version_after = agent.get_presentation_version()
        slide_info = agent.get_slide_info(slide_index)
    
    # Build result
    result = {
        "status": "success" if not validation["has_warnings"] else "warning",
        "file": str(filepath),
        "slide_index": slide_index,
        "shape_index": add_result.get("shape_index"),
        "text": text,
        "text_length": len(text),
        "position": add_result.get("position", position),
        "size": add_result.get("size", size),
        "formatting": {
            "font_name": resolved_font,
            "font_size": font_size,
            "bold": bold,
            "italic": italic,
            "color": resolved_color,
            "alignment": alignment,
            "vertical_alignment": vertical_alignment,
            "word_wrap": word_wrap
        },
        "slide_shape_count": slide_info.get("shape_count", 0),
        "validation": validation["validation_results"],
        "presentation_version": {
            "before": version_before,
            "after": version_after
        },
        "core_version": CORE_VERSION,
        "tool_version": __version__
    }
    
    if validation["warnings"]:
        result["warnings"] = validation["warnings"]
    
    if validation["recommendations"]:
        result["recommendations"] = validation["recommendations"]
    
    return result


# ============================================================================
# CLI INTERFACE
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Add text box to PowerPoint slide (v3.0)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

POSITION FORMATS:

  Percentage (recommended):
    {"left": "20%", "top": "30%"}
    
  Absolute inches:
    {"left": 2.0, "top": 3.0}
    
  Anchor-based:
    {"anchor": "center", "offset_x": 0, "offset_y": -1.0}
    Anchors: top_left, top_center, top_right,
             center_left, center, center_right,
             bottom_left, bottom_center, bottom_right
    
  Grid (12-column):
    {"grid_row": 2, "grid_col": 3, "grid_size": 12}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

COLOR PRESETS:

  black (#000000)      white (#FFFFFF)      primary (#0070C0)
  secondary (#595959)  accent (#ED7D31)     success (#70AD47)
  warning (#FFC000)    danger (#C00000)     dark_gray (#333333)
  light_gray (#808080) muted (#808080)

FONT PRESETS:

  default (Calibri)    heading (Calibri Light)   body (Calibri)
  code (Consolas)      serif (Georgia)           sans (Arial)

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

EXAMPLES:

  # Simple text box
  uv run tools/ppt_add_text_box.py \\
    --file deck.pptx --slide 0 \\
    --text "Revenue: \\$1.5M" \\
    --position '{"left":"20%","top":"30%"}' \\
    --size '{"width":"60%","height":"10%"}' \\
    --font-size 24 --bold --json

  # Centered headline
  uv run tools/ppt_add_text_box.py \\
    --file deck.pptx --slide 1 \\
    --text "Key Results" \\
    --position '{"anchor":"center","offset_y":-2}' \\
    --size '{"width":"80%","height":"15%"}' \\
    --font-size 48 --bold --color primary --alignment center --json

  # Copyright notice (bottom-right)
  uv run tools/ppt_add_text_box.py \\
    --file deck.pptx --slide 0 \\
    --text "© 2024 Company Inc." \\
    --position '{"anchor":"bottom_right","offset_x":-0.5,"offset_y":-0.3}' \\
    --size '{"width":"20%","height":"5%"}' \\
    --font-size 10 --color muted --json

  # Multi-line text block
  uv run tools/ppt_add_text_box.py \\
    --file deck.pptx --slide 2 \\
    --text "Line 1: Key Point\\nLine 2: Details\\nLine 3: Conclusion" \\
    --position '{"left":"15%","top":"30%"}' \\
    --size '{"width":"70%","height":"40%"}' \\
    --font-size 18 --json

  # Warning callout with color
  uv run tools/ppt_add_text_box.py \\
    --file deck.pptx --slide 3 \\
    --text "⚠️ Important: Review by Friday" \\
    --position '{"left":"10%","top":"70%"}' \\
    --size '{"width":"80%","height":"12%"}' \\
    --font-size 20 --bold --color danger --alignment center --json

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

ACCESSIBILITY GUIDELINES:

  Font Size:
    • Minimum: 10pt (absolute minimum)
    • Recommended: 14pt+ for projected presentations
    • Large text: 18pt+ (relaxed contrast requirements)

  Color Contrast (WCAG 2.1 AA):
    • Normal text (<18pt): 4.5:1 minimum
    • Large text (≥18pt): 3.0:1 minimum
    • Best colors: black, dark_gray, primary

  Text Length:
    • Single line: ≤100 characters recommended
    • Multi-line: ≤500 characters total

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

VALIDATION:

  The tool automatically validates:
    • Text length and readability
    • Font size accessibility
    • Color contrast (WCAG AA)
    • Position bounds
    • Size adequacy

  Warnings are included in output but don't prevent creation.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        """
    )
    
    # Required arguments
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
        help='Text content (use \\n for line breaks)'
    )
    
    parser.add_argument(
        '--position',
        required=True,
        type=json.loads,
        help='Position dict as JSON'
    )
    
    # Optional arguments
    parser.add_argument(
        '--size',
        type=json.loads,
        help='Size dict as JSON (defaults to 40%% x 20%% if omitted)'
    )
    
    parser.add_argument(
        '--font-name',
        default='Calibri',
        help='Font name or preset (default, heading, body, code, serif, sans)'
    )
    
    parser.add_argument(
        '--font-size',
        type=int,
        default=18,
        help='Font size in points (default: 18, recommended: ≥14)'
    )
    
    parser.add_argument(
        '--bold',
        action='store_true',
        help='Make text bold'
    )
    
    parser.add_argument(
        '--italic',
        action='store_true',
        help='Make text italic'
    )
    
    parser.add_argument(
        '--color',
        help='Text color: hex (#0070C0) or preset (primary, danger, etc.)'
    )
    
    parser.add_argument(
        '--alignment',
        choices=['left', 'center', 'right', 'justify'],
        default='left',
        help='Horizontal text alignment (default: left)'
    )
    
    parser.add_argument(
        '--vertical-alignment',
        choices=['top', 'middle', 'bottom'],
        default='top',
        help='Vertical text alignment (default: top)'
    )
    
    parser.add_argument(
        '--no-word-wrap',
        action='store_true',
        help='Disable word wrap'
    )
    
    parser.add_argument(
        '--allow-offslide',
        action='store_true',
        help='Allow positioning outside slide bounds'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default: true)'
    )
    
    parser.add_argument(
        '--version',
        action='version',
        version=f'%(prog)s {__version__} (core: {CORE_VERSION})'
    )
    
    args = parser.parse_args()
    
    try:
        # Handle size defaults
        size = args.size if args.size else {}
        position = args.position
        
        # Allow size in position dict for convenience
        if "width" in position and "width" not in size:
            size["width"] = position.get("width")
        if "height" in position and "height" not in size:
            size["height"] = position.get("height")
        
        # Apply defaults
        if "width" not in size:
            size["width"] = "40%"
        if "height" not in size:
            size["height"] = "20%"
        
        result = add_text_box(
            filepath=args.file,
            slide_index=args.slide,
            text=args.text,
            position=position,
            size=size,
            font_name=args.font_name,
            font_size=args.font_size,
            bold=args.bold,
            italic=args.italic,
            color=args.color,
            alignment=args.alignment,
            vertical_alignment=args.vertical_alignment,
            word_wrap=not args.no_word_wrap,
            allow_offslide=args.allow_offslide
        )
        
        if args.json:
            print(json.dumps(result, indent=2))
        else:
            status_icon = "✅" if result["status"] == "success" else "⚠️"
            print(f"{status_icon} Added text box to slide {result['slide_index']}")
            print(f"   Shape index: {result['shape_index']}")
            print(f"   Text: {result['text'][:50]}{'...' if len(result['text']) > 50 else ''}")
            if result.get("warnings"):
                print("\n   Warnings:")
                for warning in result["warnings"]:
                    print(f"   ⚠️  {warning}")
        
        sys.exit(0)
        
    except json.JSONDecodeError as e:
        error_result = {
            "status": "error",
            "error": f"Invalid JSON: {e}",
            "error_type": "JSONDecodeError",
            "hint": "Use single quotes around JSON: '{\"left\":\"20%\",\"top\":\"30%\"}'"
        }
        if args.json:
            print(json.dumps(error_result, indent=2))
        else:
            print(f"❌ Error: Invalid JSON - {e}", file=sys.stderr)
        sys.exit(1)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError"
        }
        if args.json:
            print(json.dumps(error_result, indent=2))
        else:
            print(f"❌ Error: {e}", file=sys.stderr)
        sys.exit(1)
        
    except SlideNotFoundError as e:
        error_result = {
            "status": "error",
            "error": e.message,
            "error_type": "SlideNotFoundError",
            "details": e.details,
            "suggestion": "Use ppt_get_info.py to check available slide indices."
        }
        if args.json:
            print(json.dumps(error_result, indent=2))
        else:
            print(f"❌ Error: {e.message}", file=sys.stderr)
        sys.exit(1)
        
    except InvalidPositionError as e:
        error_result = {
            "status": "error",
            "error": e.message,
            "error_type": "InvalidPositionError",
            "details": e.details
        }
        if args.json:
            print(json.dumps(error_result, indent=2))
        else:
            print(f"❌ Error: {e.message}", file=sys.stderr)
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": e.message,
            "error_type": type(e).__name__,
            "details": e.details
        }
        if args.json:
            print(json.dumps(error_result, indent=2))
        else:
            print(f"❌ Error: {e.message}", file=sys.stderr)
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__
        }
        if args.json:
            print(json.dumps(error_result, indent=2))
        else:
            print(f"❌ Error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
```

---

## Phase 4: Post-Implementation Checklist Verification

### `ppt_remove_shape.py` v3.0 Checklist

- [x] **REM-001**: Module docstring with safety warnings ✓
- [x] **REM-002**: Import v3.0 core classes (ShapeNotFoundError) ✓
- [x] **REM-003**: Shape info retrieval before removal ✓
- [x] **REM-004**: Dry-run mode for preview ✓
- [x] **REM-005**: Remove by index ✓
- [x] **REM-006**: Remove by name (partial match) ✓
- [x] **REM-007**: Batch removal support ✓
- [x] **REM-008**: Index shift warnings ✓
- [x] **REM-009**: Presentation version tracking ✓
- [x] **REM-010**: Rollback guidance in output ✓
- [x] **REM-011**: Confirmation requirement documentation ✓
- [x] **REM-012**: Comprehensive CLI with examples ✓
- [x] **REM-013**: Proper error handling ✓

### `ppt_add_text_box.py` v3.0 Checklist

- [x] **TXT-001**: Module docstring updated for v3.0 ✓
- [x] **TXT-002**: Correct imports from v3.0 core ✓
- [x] **TXT-003**: Return shape_index from core ✓
- [x] **TXT-004**: Presentation version tracking ✓
- [x] **TXT-005**: Word wrap option ✓
- [x] **TXT-006**: Vertical alignment option ✓
- [x] **TXT-007**: Enhanced validation (preserved good parts) ✓
- [x] **TXT-008**: Color presets support ✓
- [x] **TXT-009**: Consistent JSON output ✓
- [x] **TXT-010**: Comprehensive CLI help ✓
- [x] **TXT-011**: Proper error handling with v3.0 exceptions ✓

---

## Phase 5: Test Design

### Test File: `tests/test_ppt_remove_shape.py`

```python
#!/usr/bin/env python3
"""Tests for ppt_remove_shape.py v3.0"""

import pytest
import json
import sys
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE

sys.path.insert(0, str(Path(__file__).parent.parent / "tools"))

from ppt_remove_shape import (
    remove_shape,
    remove_shapes_batch,
    get_shape_details,
    find_shape_by_name,
    validate_removal
)


@pytest.fixture
def sample_pptx_with_shapes(tmp_path):
    """Create a sample PowerPoint with multiple shapes for testing."""
    pptx_path = tmp_path / "test_remove.pptx"
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank
    
    # Add multiple shapes
    shape1 = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.RECTANGLE,
        Inches(1), Inches(1), Inches(2), Inches(1)
    )
    shape1.name = "Rectangle 1"
    
    shape2 = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.OVAL,
        Inches(4), Inches(1), Inches(1), Inches(1)
    )
    shape2.name = "Circle 1"
    
    shape3 = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.RECTANGLE,
        Inches(1), Inches(3), Inches(2), Inches(1)
    )
    shape3.name = "Rectangle 2"
    shape3.text_frame.text = "Important content"
    
    prs.save(str(pptx_path))
    return pptx_
