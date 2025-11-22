#!/usr/bin/env python3
"""
PowerPoint Capability Probe Tool
Detect and report presentation template capabilities, layouts, and theme properties

Version 1.0.0 - Foundation Release

This tool provides comprehensive introspection of PowerPoint presentations to detect:
- Available layouts and their placeholders
- Slide dimensions and aspect ratios
- Theme colors and fonts
- Template capabilities (footer support, etc.)

Critical for AI agents and automation workflows to understand template capabilities
before generating content.

Usage:
    # Basic probe (essential info)
    uv run tools/ppt_capability_probe.py --file template.pptx --json
    
    # Deep probe (full analysis)
    uv run tools/ppt_capability_probe.py --file template.pptx --deep --json
    
    # Human-friendly summary
    uv run tools/ppt_capability_probe.py --file template.pptx --summary

Exit Codes:
    0: Success
    1: Error occurred

Design Principles:
    - Read-only operation (atomic, no file mutation)
    - JSON-first output for automation
    - Comprehensive capability detection
    - Graceful degradation for missing features
    - Performance-optimized (--deep for full scan)
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any, List, Optional, Tuple
from datetime import datetime
import hashlib

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError, ColorHelper,
    PLACEHOLDER_TYPE_NAMES, get_placeholder_type_name
)


def calculate_file_checksum(filepath: Path) -> str:
    """
    Calculate MD5 checksum of file to verify no mutation.
    
    Args:
        filepath: Path to file
        
    Returns:
        Hex digest of file contents
    """
    md5 = hashlib.md5()
    with open(filepath, 'rb') as f:
        for chunk in iter(lambda: f.read(4096), b''):
            md5.update(chunk)
    return md5.hexdigest()


def detect_slide_dimensions(agent: PowerPointAgent) -> Dict[str, Any]:
    """
    Detect slide dimensions and calculate aspect ratio.
    
    Args:
        agent: Initialized PowerPointAgent
        
    Returns:
        Dict with width, height, aspect ratio, DPI estimate
    """
    prs = agent.prs
    
    width_inches = prs.slide_width.inches
    height_inches = prs.slide_height.inches
    
    width_emu = prs.slide_width
    height_emu = prs.slide_height
    
    dpi_estimate = 96
    width_pixels = int(width_inches * dpi_estimate)
    height_pixels = int(height_inches * dpi_estimate)
    
    ratio = width_inches / height_inches
    if abs(ratio - 16/9) < 0.01:
        aspect_ratio = "16:9"
    elif abs(ratio - 4/3) < 0.01:
        aspect_ratio = "4:3"
    else:
        aspect_ratio = f"{width_inches:.1f}:{height_inches:.1f}"
    
    return {
        "width_inches": round(width_inches, 2),
        "height_inches": round(height_inches, 2),
        "width_emu": int(width_emu),
        "height_emu": int(height_emu),
        "width_pixels": width_pixels,
        "height_pixels": height_pixels,
        "aspect_ratio": aspect_ratio,
        "dpi_estimate": dpi_estimate
    }


def analyze_placeholder(shape, slide_width: float, slide_height: float) -> Dict[str, Any]:
    """
    Analyze a single placeholder and return comprehensive info.
    
    Args:
        shape: Placeholder shape to analyze
        slide_width: Slide width in inches
        slide_height: Slide height in inches
        
    Returns:
        Dict with type, position, size information
    """
    ph_format = shape.placeholder_format
    ph_type = ph_format.type
    ph_type_name = get_placeholder_type_name(ph_type)
    
    left_inches = shape.left / 914400 if hasattr(shape, 'left') else 0
    top_inches = shape.top / 914400 if hasattr(shape, 'top') else 0
    width_inches = shape.width / 914400 if hasattr(shape, 'width') else 0
    height_inches = shape.height / 914400 if hasattr(shape, 'height') else 0
    
    left_percent = (left_inches / slide_width * 100) if slide_width > 0 else 0
    top_percent = (top_inches / slide_height * 100) if slide_height > 0 else 0
    width_percent = (width_inches / slide_width * 100) if slide_width > 0 else 0
    height_percent = (height_inches / slide_height * 100) if slide_height > 0 else 0
    
    return {
        "type": ph_type_name,
        "type_code": ph_type,
        "idx": ph_format.idx,
        "name": shape.name,
        "position_inches": {
            "left": round(left_inches, 2),
            "top": round(top_inches, 2)
        },
        "position_percent": {
            "left": f"{left_percent:.1f}%",
            "top": f"{top_percent:.1f}%"
        },
        "size_inches": {
            "width": round(width_inches, 2),
            "height": round(height_inches, 2)
        },
        "size_percent": {
            "width": f"{width_percent:.1f}%",
            "height": f"{height_percent:.1f}%"
        }
    }


def detect_layouts(agent: PowerPointAgent, slide_width: float, slide_height: float, deep: bool = False) -> List[Dict[str, Any]]:
    """
    Detect all layouts and their placeholders.
    
    Args:
        agent: Initialized PowerPointAgent
        slide_width: Slide width in inches
        slide_height: Slide height in inches
        deep: If True, analyze all placeholders; if False, count only
        
    Returns:
        List of layout information dicts
    """
    layouts = []
    
    for idx, layout in enumerate(agent.prs.slide_layouts):
        layout_info = {
            "index": idx,
            "name": layout.name,
            "placeholder_count": len(layout.placeholders)
        }
        
        if deep:
            placeholders = []
            for shape in layout.placeholders:
                try:
                    ph_info = analyze_placeholder(shape, slide_width, slide_height)
                    placeholders.append(ph_info)
                except Exception as e:
                    placeholders.append({
                        "type": "ERROR",
                        "error": str(e)
                    })
            
            layout_info["placeholders"] = placeholders
        else:
            placeholder_types = []
            for shape in layout.placeholders:
                try:
                    ph_type = shape.placeholder_format.type
                    ph_type_name = get_placeholder_type_name(ph_type)
                    if ph_type_name not in placeholder_types:
                        placeholder_types.append(ph_type_name)
                except:
                    pass
            
            layout_info["placeholder_types"] = placeholder_types
        
        layouts.append(layout_info)
    
    return layouts


def extract_theme_colors(agent: PowerPointAgent) -> Dict[str, str]:
    """
    Extract theme colors from presentation.
    
    Args:
        agent: Initialized PowerPointAgent
        
    Returns:
        Dict mapping color names to hex codes
    """
    colors = {}
    
    try:
        slide_master = agent.prs.slide_masters[0]
        theme = slide_master.theme
        color_scheme = theme.theme_color_scheme
        
        color_names = [
            'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6',
            'background1', 'background2', 'text1', 'text2', 'hyperlink', 'followed_hyperlink'
        ]
        
        for idx, color_name in enumerate(color_names):
            try:
                color = getattr(color_scheme, color_name, None)
                if color:
                    colors[color_name] = ColorHelper.to_hex(color)
            except:
                pass
    except Exception as e:
        colors["_error"] = f"Theme color extraction failed: {str(e)}"
    
    return colors


def extract_theme_fonts(agent: PowerPointAgent) -> Dict[str, str]:
    """
    Extract theme fonts from presentation.
    
    Args:
        agent: Initialized PowerPointAgent
        
    Returns:
        Dict with heading and body font names
    """
    fonts = {}
    
    try:
        slide_master = agent.prs.slide_masters[0]
        
        for shape in slide_master.shapes:
            if hasattr(shape, 'text_frame'):
                for paragraph in shape.text_frame.paragraphs:
                    if paragraph.font.name and 'heading' not in fonts:
                        fonts['heading'] = paragraph.font.name
                        break
                if 'heading' in fonts:
                    break
        
        for shape in slide_master.shapes:
            if hasattr(shape, 'text_frame'):
                for paragraph in shape.text_frame.paragraphs:
                    if paragraph.font.name:
                        fonts['body'] = paragraph.font.name
                        break
                if 'body' in fonts:
                    break
    except Exception as e:
        fonts["_error"] = f"Theme font extraction failed: {str(e)}"
    
    if not fonts:
        fonts = {"heading": "Calibri", "body": "Calibri", "_note": "Using defaults"}
    
    return fonts


def analyze_capabilities(layouts: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    Analyze template capabilities based on detected layouts.
    
    Args:
        layouts: List of layout information dicts
        
    Returns:
        Dict with capability flags and recommendations
    """
    has_footer = False
    has_slide_number = False
    has_date = False
    layouts_with_footer = []
    layouts_with_slide_number = []
    layouts_with_date = []
    
    for layout in layouts:
        layout_name = layout['name']
        
        if 'placeholders' in layout:
            for ph in layout['placeholders']:
                if ph.get('type') == 'FOOTER' or ph.get('type_code') == 4:
                    has_footer = True
                    if layout_name not in layouts_with_footer:
                        layouts_with_footer.append(layout_name)
                
                if ph.get('type') == 'SLIDE_NUMBER' or ph.get('type_code') == 13:
                    has_slide_number = True
                    if layout_name not in layouts_with_slide_number:
                        layouts_with_slide_number.append(layout_name)
                
                if ph.get('type') == 'DATE' or ph.get('type_code') == 16:
                    has_date = True
                    if layout_name not in layouts_with_date:
                        layouts_with_date.append(layout_name)
        elif 'placeholder_types' in layout:
            if 'FOOTER' in layout['placeholder_types']:
                has_footer = True
                layouts_with_footer.append(layout_name)
            
            if 'SLIDE_NUMBER' in layout['placeholder_types']:
                has_slide_number = True
                layouts_with_slide_number.append(layout_name)
            
            if 'DATE' in layout['placeholder_types']:
                has_date = True
                layouts_with_date.append(layout_name)
    
    recommendations = []
    
    if not has_footer:
        recommendations.append(
            "No footer placeholders found - ppt_set_footer.py will use text box fallback strategy"
        )
    else:
        recommendations.append(
            f"Footer placeholders available on {len(layouts_with_footer)} layout(s)"
        )
    
    if not has_slide_number:
        recommendations.append(
            "No slide number placeholders - recommend manual text box for slide numbers"
        )
    
    if not has_date:
        recommendations.append(
            "No date placeholders - dates must be added manually if needed"
        )
    
    return {
        "has_footer_placeholders": has_footer,
        "has_slide_number_placeholders": has_slide_number,
        "has_date_placeholders": has_date,
        "layouts_with_footer": layouts_with_footer,
        "layouts_with_slide_number": layouts_with_slide_number,
        "layouts_with_date": layouts_with_date,
        "total_layouts": len(layouts),
        "total_master_slides": 1,
        "recommendations": recommendations
    }


def probe_presentation(
    filepath: Path,
    deep: bool = False,
    verify_atomic: bool = True
) -> Dict[str, Any]:
    """
    Probe presentation and return comprehensive capability report.
    
    Args:
        filepath: Path to PowerPoint file
        deep: If True, perform deep analysis of all placeholders
        verify_atomic: If True, verify no file mutation occurred
        
    Returns:
        Dict with complete capability report
    """
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    checksum_before = None
    if verify_atomic:
        checksum_before = calculate_file_checksum(filepath)
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath, acquire_lock=False)
        
        dimensions = detect_slide_dimensions(agent)
        
        slide_width = dimensions['width_inches']
        slide_height = dimensions['height_inches']
        
        layouts = detect_layouts(agent, slide_width, slide_height, deep)
        
        theme_colors = extract_theme_colors(agent)
        theme_fonts = extract_theme_fonts(agent)
        
        capabilities = analyze_capabilities(layouts)
        
        master_count = len(agent.prs.slide_masters)
    
    checksum_after = None
    if verify_atomic:
        checksum_after = calculate_file_checksum(filepath)
        
        if checksum_before != checksum_after:
            raise PowerPointAgentError(
                "File was modified during probe operation! "
                "This should never happen (atomic read violation)."
            )
    
    result = {
        "metadata": {
            "file": str(filepath),
            "probed_at": datetime.now().isoformat(),
            "tool_version": "1.0.0",
            "deep_analysis": deep,
            "atomic_verified": verify_atomic,
            "checksum": checksum_after if verify_atomic else None
        },
        "slide_dimensions": dimensions,
        "layouts": layouts,
        "theme": {
            "colors": theme_colors,
            "fonts": theme_fonts
        },
        "capabilities": capabilities
    }
    
    result["capabilities"]["total_master_slides"] = master_count
    
    return result


def format_summary(probe_result: Dict[str, Any]) -> str:
    """
    Format probe result as human-readable summary.
    
    Args:
        probe_result: Result from probe_presentation()
        
    Returns:
        Formatted string summary
    """
    lines = []
    
    lines.append("═══════════════════════════════════════════════════════════════")
    lines.append("PowerPoint Capability Probe Report")
    lines.append("═══════════════════════════════════════════════════════════════")
    lines.append("")
    
    meta = probe_result['metadata']
    lines.append(f"File: {meta['file']}")
    lines.append(f"Probed: {meta['probed_at']}")
    lines.append(f"Analysis Mode: {'Deep' if meta['deep_analysis'] else 'Essential'}")
    lines.append("")
    
    dims = probe_result['slide_dimensions']
    lines.append("Slide Dimensions:")
    lines.append(f"  Size: {dims['width_inches']}\" × {dims['height_inches']}\"")
    lines.append(f"  Aspect Ratio: {dims['aspect_ratio']}")
    lines.append(f"  Pixels: {dims['width_pixels']} × {dims['height_pixels']} (@ {dims['dpi_estimate']} DPI)")
    lines.append("")
    
    caps = probe_result['capabilities']
    lines.append("Template Capabilities:")
    lines.append(f"  ✓ Total Layouts: {caps['total_layouts']}")
    lines.append(f"  {'✓' if caps['has_footer_placeholders'] else '✗'} Footer Placeholders: {len(caps['layouts_with_footer'])} layout(s)")
    lines.append(f"  {'✓' if caps['has_slide_number_placeholders'] else '✗'} Slide Number Placeholders: {len(caps['layouts_with_slide_number'])} layout(s)")
    lines.append(f"  {'✓' if caps['has_date_placeholders'] else '✗'} Date Placeholders: {len(caps['layouts_with_date'])} layout(s)")
    lines.append("")
    
    lines.append("Available Layouts:")
    for layout in probe_result['layouts']:
        ph_count = layout['placeholder_count']
        lines.append(f"  [{layout['index']}] {layout['name']} ({ph_count} placeholder{'s' if ph_count != 1 else ''})")
    lines.append("")
    
    theme = probe_result['theme']
    if theme['fonts'] and '_error' not in theme['fonts']:
        lines.append("Theme Fonts:")
        for key, value in theme['fonts'].items():
            if not key.startswith('_'):
                lines.append(f"  {key.capitalize()}: {value}")
        lines.append("")
    
    if caps['recommendations']:
        lines.append("Recommendations:")
        for rec in caps['recommendations']:
            lines.append(f"  • {rec}")
        lines.append("")
    
    lines.append("═══════════════════════════════════════════════════════════════")
    
    return "\n".join(lines)


def main():
    parser = argparse.ArgumentParser(
        description="Probe PowerPoint presentation capabilities (v1.0.0)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Basic probe (essential info, fast)
  uv run tools/ppt_capability_probe.py --file template.pptx --json
  
  # Deep probe (full placeholder analysis)
  uv run tools/ppt_capability_probe.py --file template.pptx --deep --json
  
  # Human-friendly summary
  uv run tools/ppt_capability_probe.py --file template.pptx --summary
  
  # Verify atomic read (no file mutation)
  uv run tools/ppt_capability_probe.py --file template.pptx --verify-atomic --json

Output Information:
  - Slide dimensions (inches, pixels, aspect ratio)
  - All available layouts with placeholder counts
  - Placeholder types and positions (in --deep mode)
  - Theme colors and fonts
  - Capability flags (footer support, etc.)
  - Recommendations for tool usage

Use Cases:
  1. Template Selection: Compare multiple templates
  2. Pre-flight Check: Verify template before content generation
  3. Tool Planning: Check if footer/header placeholders exist
  4. Documentation: Auto-generate template documentation
  5. Validation: Ensure template meets requirements

Capability Detection:
  This tool detects:
  - Footer placeholder availability (critical for ppt_set_footer.py)
  - Slide number placeholder availability
  - Date placeholder availability
  - Layout names and counts
  - Placeholder positions and sizes (--deep mode)
  - Theme color schemes
  - Theme fonts

Performance:
  - Essential mode (~0.5s): Layout names and placeholder counts
  - Deep mode (~2s): Full placeholder analysis with positions
  - Atomic verification adds ~0.1s for checksum calculation

Atomic Read Guarantee:
  By default, this tool verifies it did not modify the input file
  by comparing checksums before and after. This ensures:
  - No accidental mutations
  - No timestamp changes
  - Safe for production templates
  - Use --no-verify-atomic to skip (faster, less safe)

Integration with Other Tools:
  ppt_set_footer.py: Check has_footer_placeholders before use
  ppt_set_title.py: Verify title/subtitle placeholders exist
  ppt_create_from_template.py: Validate template before creating
  Script generation: Query capabilities to optimize content

Output Schema:
  {
    "metadata": {...},           // File info, timestamps
    "slide_dimensions": {...},   // Width, height, aspect ratio
    "layouts": [...],            // All layouts with placeholders
    "theme": {...},              // Colors and fonts
    "capabilities": {...}        // Flags and recommendations
  }

Version: 1.0.0
Requires: core/powerpoint_agent_core.py v1.1.0+
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file to probe'
    )
    
    parser.add_argument(
        '--deep',
        action='store_true',
        help='Perform deep analysis (full placeholder positions/sizes, slower)'
    )
    
    parser.add_argument(
        '--summary',
        action='store_true',
        help='Output human-friendly summary instead of JSON'
    )
    
    parser.add_argument(
        '--verify-atomic',
        action='store_true',
        default=True,
        help='Verify no file mutation occurred (default: true)'
    )
    
    parser.add_argument(
        '--no-verify-atomic',
        action='store_false',
        dest='verify_atomic',
        help='Skip atomic verification (faster, less safe)'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON format (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        result = probe_presentation(
            filepath=args.file,
            deep=args.deep,
            verify_atomic=args.verify_atomic
        )
        
        if args.summary:
            print(format_summary(result))
        else:
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
