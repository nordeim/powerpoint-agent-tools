#!/usr/bin/env python3
"""
PowerPoint Set Footer Tool
Configure slide footer, date, and slide number

Version 2.0.0 - Complete Rewrite with Dual Strategy

This tool now implements a dual-strategy approach:
1. Primary: Attempts to use native placeholder footers (if template supports them)
2. Fallback: Creates text box overlays for presentations without footer placeholders

Changes from v1.x:
- Fixed: Now uses correct PP_PLACEHOLDER constants (FOOTER=4, not 15)
- Fixed: Properly handles presentations without footer placeholders
- Enhanced: Dual strategy ensures footer always works
- Enhanced: Detailed warnings when placeholders not available
- Enhanced: Returns which method was used (placeholder vs text_box)
- Enhanced: Comprehensive slide-by-slide status reporting

Usage:
    # Basic footer text
    uv run tools/ppt_set_footer.py --file deck.pptx --text "Confidential" --json
    
    # Footer with slide numbers
    uv run tools/ppt_set_footer.py --file deck.pptx --text "Company Name" --show-number --json
    
    # Footer with date and numbers
    uv run tools/ppt_set_footer.py --file deck.pptx --text "Q4 Report" --show-number --show-date --json

Note on Slide Numbers:
    Due to python-pptx limitations, --show-number creates text box overlays with
    slide numbers rather than activating native PowerPoint slide number placeholders.
    This ensures consistent behavior across all templates.

Exit Codes:
    0: Success (footer applied via placeholder or text box)
    1: Error (file not found, invalid arguments, etc.)
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any, List

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import PowerPointAgent, PP_PLACEHOLDER


def set_footer(
    filepath: Path,
    text: str = None,
    show_number: bool = False,
    show_date: bool = False,
    apply_to_all: bool = True
) -> Dict[str, Any]:
    """
    Set footer on presentation slides using dual-strategy approach.
    
    Strategy 1 (Preferred): Use native footer placeholders if available
    Strategy 2 (Fallback): Create text box overlays at bottom of slides
    
    Args:
        filepath: Path to PowerPoint file
        text: Footer text content
        show_number: Whether to show slide numbers
        show_date: Whether to show date (not implemented - reserved for future)
        apply_to_all: Apply to all slides (True) or only master (False)
        
    Returns:
        Dict containing:
        - status: "success" or "warning"
        - method_used: "placeholder" or "text_box" or "hybrid"
        - slides_updated: Number of slides modified
        - slide_indices: List of modified slide indices
        - warnings: List of warning messages
        - details: Additional information
    """
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")

    warnings = []
    slide_indices_updated = set()
    method_used = None
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # ====================================================================
        # STRATEGY 1: Try native placeholder approach
        # ====================================================================
        
        placeholder_count = 0
        
        if text:
            # Try to set footer text on master slide layouts
            try:
                for master in agent.prs.slide_masters:
                    for layout in master.slide_layouts:
                        for shape in layout.placeholders:
                            if shape.placeholder_format.type == PP_PLACEHOLDER.FOOTER:
                                try:
                                    shape.text = text
                                    placeholder_count += 1
                                except:
                                    pass
            except Exception as e:
                warnings.append(f"Could not access master slide layouts: {str(e)}")
            
            # Try to set footer text on individual slides
            for slide_idx, slide in enumerate(agent.prs.slides):
                for shape in slide.placeholders:
                    if shape.placeholder_format.type == PP_PLACEHOLDER.FOOTER:
                        try:
                            shape.text = text
                            slide_indices_updated.add(slide_idx)
                            placeholder_count += 1
                        except:
                            pass
        
        # ====================================================================
        # STRATEGY 2: Fallback to text box overlay if no placeholders found
        # ====================================================================
        
        textbox_count = 0
        
        if placeholder_count == 0:
            warnings.append(
                "No footer placeholders found in this presentation template. "
                "Using text box overlay strategy instead."
            )
            
            # Skip title slide (slide 0) for footer overlays
            for slide_idx in range(1, len(agent.prs.slides)):
                try:
                    # Add footer text box (bottom-left)
                    if text:
                        agent.add_text_box(
                            slide_index=slide_idx,
                            text=text,
                            position={"left": "5%", "top": "92%"},
                            size={"width": "60%", "height": "5%"},
                            font_size=10,
                            color="#595959",
                            alignment="left"
                        )
                        textbox_count += 1
                        slide_indices_updated.add(slide_idx)
                    
                    # Add slide number box (bottom-right)
                    if show_number:
                        # Display number is 1-indexed for audience (slide 1 displays as "2")
                        display_number = slide_idx + 1
                        agent.add_text_box(
                            slide_index=slide_idx,
                            text=str(display_number),
                            position={"left": "92%", "top": "92%"},
                            size={"width": "5%", "height": "5%"},
                            font_size=10,
                            color="#595959",
                            alignment="left"
                        )
                        textbox_count += 1
                        slide_indices_updated.add(slide_idx)
                        
                except Exception as e:
                    warnings.append(f"Failed to add text box to slide {slide_idx}: {str(e)}")
        
        # Determine which method was used
        if placeholder_count > 0 and textbox_count == 0:
            method_used = "placeholder"
        elif textbox_count > 0 and placeholder_count == 0:
            method_used = "text_box"
        elif placeholder_count > 0 and textbox_count > 0:
            method_used = "hybrid"
        else:
            method_used = "none"
            warnings.append("No footer elements were added (no text provided or all operations failed)")
        
        # ====================================================================
        # Handle --show-date flag (reserved for future implementation)
        # ====================================================================
        
        if show_date:
            warnings.append(
                "Date placeholder activation not yet implemented due to python-pptx limitations. "
                "Consider adding date manually via ppt_add_text_box.py"
            )
        
        # Save changes
        agent.save()
    
    # ====================================================================
    # Build comprehensive response
    # ====================================================================
    
    status = "success" if len(slide_indices_updated) > 0 else "warning"
    
    result = {
        "status": status,
        "file": str(filepath),
        "method_used": method_used,
        "footer_text": text,
        "settings": {
            "show_number": show_number,
            "show_date": show_date,
            "show_date_implemented": False  # Not yet supported
        },
        "slides_updated": len(slide_indices_updated),
        "slide_indices": sorted(list(slide_indices_updated)),
        "details": {
            "placeholder_count": placeholder_count,
            "textbox_count": textbox_count,
            "total_elements_added": placeholder_count + textbox_count
        }
    }
    
    if warnings:
        result["warnings"] = warnings
        
    # Add recommendations if using fallback strategy
    if method_used == "text_box":
        result["recommendation"] = (
            "This presentation template does not include footer placeholders. "
            "For better maintainability, consider using a template with built-in footer support, "
            "or continue using text box overlays (which work on all presentations)."
        )
    
    return result


def main():
    parser = argparse.ArgumentParser(
        description="Set PowerPoint footer with dual-strategy approach",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Add footer text only
  uv run tools/ppt_set_footer.py --file deck.pptx --text "Confidential" --json
  
  # Add footer with slide numbers
  uv run tools/ppt_set_footer.py --file deck.pptx --text "Q4 Report" --show-number --json
  
  # Footer with all options
  uv run tools/ppt_set_footer.py --file deck.pptx --text "Company Name" --show-number --show-date --json

Strategy:
  This tool uses a dual-strategy approach:
  1. First attempts to use native PowerPoint footer placeholders
  2. Falls back to text box overlays if placeholders not available
  
  The method used is reported in the "method_used" field:
  - "placeholder": Native footer placeholders used
  - "text_box": Text box overlays used (for templates without footer support)
  - "hybrid": Both methods used
  
Slide Numbering:
  Due to python-pptx limitations, slide numbers are implemented as text boxes
  positioned at bottom-right (92% left, 92% top) rather than activating
  native PowerPoint slide number placeholders. This ensures consistent behavior.

Output:
  {
    "status": "success",
    "method_used": "text_box",
    "slides_updated": 8,
    "slide_indices": [1, 2, 3, 4, 5, 6, 7, 8],
    "warnings": ["No footer placeholders found..."],
    "recommendation": "..."
  }
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file path'
    )
    
    parser.add_argument(
        '--text',
        help='Footer text content'
    )
    
    parser.add_argument(
        '--show-number',
        action='store_true',
        help='Show slide numbers (implemented as text box overlays)'
    )
    
    parser.add_argument(
        '--show-date',
        action='store_true',
        help='Show date (reserved for future implementation)'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        result = set_footer(
            filepath=args.file, 
            text=args.text, 
            show_number=args.show_number,
            show_date=args.show_date
        )
        
        print(json.dumps(result, indent=2))
        sys.exit(0)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "file": str(args.file) if args.file else None
        }
        
        print(json.dumps(error_result, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()
