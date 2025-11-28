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
