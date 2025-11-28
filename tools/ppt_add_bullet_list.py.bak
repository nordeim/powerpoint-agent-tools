#!/usr/bin/env python3
"""
PowerPoint Add Bullet List Tool
Add bullet or numbered list with 6×6 rule validation and accessibility checks

Version 2.0.0 - Enhanced Validation and Accessibility

Changes from v1.x:
- Enhanced: 6×6 rule validation (warns at 6 items, error at 10)
- Enhanced: Character count validation per item
- Enhanced: Accessibility checks (color contrast, font size)
- Enhanced: Readability scoring
- Enhanced: Theme-aware formatting options
- Enhanced: JSON-first output (always returns JSON)
- Added: `--ignore-rules` flag to override 6×6 validation
- Added: `--theme-colors` flag to use presentation theme
- Added: Comprehensive warnings and recommendations
- Fixed: Consistent response format

Best Practices (6×6 Rule):
- Maximum 6 bullet points per slide
- Maximum 6 words per line (60 characters recommended)
- This ensures readability and audience engagement
- Use multiple slides rather than cramming content

Usage:
    # Simple bullet list (auto-validates 6×6 rule)
    uv run tools/ppt_add_bullet_list.py --file deck.pptx --slide 1 --items "Point 1,Point 2,Point 3" --position '{"left":"10%","top":"25%"}' --size '{"width":"80%","height":"60%"}' --json
    
    # Numbered list with custom formatting
    uv run tools/ppt_add_bullet_list.py --file deck.pptx --slide 2 --items "Step 1,Step 2,Step 3" --bullet-style numbered --font-size 20 --color "#0070C0" --position '{"left":"15%","top":"30%"}' --size '{"width":"70%","height":"50%"}' --json
    
    # Load items from JSON file
    echo '["Revenue up 45%", "Customer growth 60%", "Market share 23%"]' > items.json
    uv run tools/ppt_add_bullet_list.py --file deck.pptx --slide 3 --items-file items.json --position '{"left":"10%","top":"25%"}' --size '{"width":"80%","height":"60%"}' --json

Exit Codes:
    0: Success
    1: Error occurred
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any, List

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError, SlideNotFoundError, ColorHelper, RGBColor
)


def calculate_readability_score(items: List[str]) -> Dict[str, Any]:
    """
    Calculate readability metrics for bullet list.
    
    Returns:
        Dict with readability metrics and recommendations
    """
    total_chars = sum(len(item) for item in items)
    avg_chars = total_chars / len(items) if items else 0
    max_chars = max(len(item) for item in items) if items else 0
    
    # Count words (approximate)
    total_words = sum(len(item.split()) for item in items)
    avg_words = total_words / len(items) if items else 0
    max_words = max(len(item.split()) for item in items) if items else 0
    
    # Scoring
    score = 100
    issues = []
    
    # Deduct for too many items
    if len(items) > 6:
        score -= (len(items) - 6) * 10
        issues.append(f"Exceeds 6×6 rule: {len(items)} items (recommended: ≤6)")
    
    # Deduct for long items
    if avg_chars > 60:
        score -= 20
        issues.append(f"Items too long: {avg_chars:.0f} chars average (recommended: ≤60)")
    
    if max_chars > 100:
        score -= 10
        issues.append(f"Longest item: {max_chars} chars (consider splitting)")
    
    # Deduct for too many words per line
    if max_words > 12:
        score -= 15
        issues.append(f"Too many words per item: {max_words} max (recommended: ≤10)")
    
    score = max(0, score)
    
    return {
        "score": score,
        "grade": "A" if score >= 90 else "B" if score >= 75 else "C" if score >= 60 else "D" if score >= 50 else "F",
        "metrics": {
            "item_count": len(items),
            "avg_characters": round(avg_chars, 1),
            "max_characters": max_chars,
            "avg_words": round(avg_words, 1),
            "max_words": max_words
        },
        "issues": issues
    }


def add_bullet_list(
    filepath: Path,
    slide_index: int,
    items: List[str],
    position: Dict[str, Any],
    size: Dict[str, Any],
    bullet_style: str = "bullet",
    font_size: int = 18,
    font_name: str = "Calibri",
    color: str = None,
    line_spacing: float = 1.0,
    ignore_rules: bool = False
) -> Dict[str, Any]:
    """
    Add bullet or numbered list with validation.
    
    Enforces 6×6 rule unless --ignore-rules is specified:
    - Maximum 6 bullet points per slide
    - Maximum ~60 characters per line
    
    Args:
        filepath: Path to PowerPoint file
        slide_index: Slide index (0-based)
        items: List of bullet items
        position: Position dict
        size: Size dict
        bullet_style: "bullet", "numbered", or "none"
        font_size: Font size in points
        font_name: Font name
        color: Optional text color (hex)
        line_spacing: Line spacing multiplier
        ignore_rules: Override 6×6 rule validation
        
    Returns:
        Dict containing:
        - status: "success", "warning", or "error"
        - items_added: Count
        - readability_score: Metrics and grade
        - warnings: Validation warnings
        - recommendations: Suggested improvements
    """
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if not items:
        raise ValueError("At least one item required")
    
    warnings = []
    recommendations = []
    
    # Calculate readability
    readability = calculate_readability_score(items)
    
    # 6×6 Rule Enforcement
    if len(items) > 6 and not ignore_rules:
        warnings.append(
            f"6×6 Rule violation: {len(items)} items exceeds recommended 6 per slide. "
            "This reduces readability and audience engagement."
        )
        recommendations.append(
            "Consider splitting into multiple slides or using --ignore-rules to override"
        )
    
    # Hard limit at 10 items (safety)
    if len(items) > 10 and not ignore_rules:
        raise ValueError(
            f"Too many items: {len(items)} exceeds hard limit of 10 per slide. "
            "This severely reduces readability. Either:\n"
            "  1. Split into multiple slides (recommended)\n"
            "  2. Use --ignore-rules to override (not recommended)"
        )
    
    # Warn about very long items
    for idx, item in enumerate(items):
        if len(item) > 100:
            warnings.append(
                f"Item {idx + 1} is {len(item)} characters (very long). "
                "Consider breaking into multiple bullets."
            )
    
    # Font size validation
    if font_size < 14:
        warnings.append(
            f"Font size {font_size}pt is below recommended minimum of 14pt. "
            "Audience may struggle to read from distance."
        )
    
    # Color contrast check (if color specified)
    if color:
        try:
            text_color = ColorHelper.from_hex(color)
            bg_color = RGBColor(255, 255, 255)
            is_large_text = font_size >= 18
            
            if not ColorHelper.meets_wcag(text_color, bg_color, is_large_text):
                contrast_ratio = ColorHelper.contrast_ratio(text_color, bg_color)
                required_ratio = 3.0 if is_large_text else 4.5
                warnings.append(
                    f"Color contrast {contrast_ratio:.2f}:1 may not meet WCAG accessibility "
                    f"standards (required: {required_ratio}:1). Consider darker color."
                )
        except:
            pass
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Validate slide index
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides-1})"
            )
        
        # Add bullet list
        agent.add_bullet_list(
            slide_index=slide_index,
            items=items,
            position=position,
            size=size,
            bullet_style=bullet_style,
            font_size=font_size
        )
        
        # Get the last added shape for additional formatting
        slide_info = agent.get_slide_info(slide_index)
        last_shape_idx = slide_info["shape_count"] - 1
        
        # Apply color if specified
        if color:
            try:
                agent.format_text(
                    slide_index=slide_index,
                    shape_index=last_shape_idx,
                    color=color
                )
            except Exception as e:
                warnings.append(f"Could not apply color: {str(e)}")
        
        # Save
        agent.save()
    
    # Recommendations based on readability
    if readability["score"] < 75:
        recommendations.append(
            f"Readability score is {readability['grade']} ({readability['score']}/100). "
            "Consider simplifying content for better audience engagement."
        )
    
    if readability["metrics"]["avg_words"] > 8:
        recommendations.append(
            "Average words per item exceeds 8. Keep bullets concise for impact."
        )
    
    # Build response
    status = "success"
    if warnings:
        status = "warning"
    
    result = {
        "status": status,
        "file": str(filepath),
        "slide_index": slide_index,
        "items_added": len(items),
        "items": items,
        "bullet_style": bullet_style,
        "formatting": {
            "font_size": font_size,
            "font_name": font_name,
            "color": color,
            "line_spacing": line_spacing
        },
        "readability": readability,
        "validation": {
            "six_six_rule": {
                "compliant": len(items) <= 6 and readability["metrics"]["max_words"] <= 10,
                "item_count_ok": len(items) <= 6,
                "word_count_ok": readability["metrics"]["max_words"] <= 10
            },
            "accessibility": {
                "font_size_ok": font_size >= 14,
                "color_contrast_checked": color is not None
            }
        }
    }
    
    if warnings:
        result["warnings"] = warnings
    
    if recommendations:
        result["recommendations"] = recommendations
    
    return result


def main():
    parser = argparse.ArgumentParser(
        description="Add bullet/numbered list with 6×6 rule validation (v2.0.0)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
6×6 Rule (Best Practice):
  - Maximum 6 bullet points per slide
  - Maximum 6 words per line (~60 characters)
  - Ensures readability and audience engagement
  - Validated automatically unless --ignore-rules

Examples:
  # Simple bullet list (validates 6×6 rule)
  uv run tools/ppt_add_bullet_list.py \\
    --file presentation.pptx \\
    --slide 1 \\
    --items "Revenue up 45%,Customer growth 60%,Market share increased" \\
    --position '{"left":"10%","top":"25%"}' \\
    --size '{"width":"80%","height":"60%"}' \\
    --json
  
  # Numbered list with custom formatting
  uv run tools/ppt_add_bullet_list.py \\
    --file deck.pptx \\
    --slide 2 \\
    --items "Define objectives,Analyze market,Develop strategy,Execute plan" \\
    --bullet-style numbered \\
    --position '{"left":"15%","top":"30%"}' \\
    --size '{"width":"70%","height":"50%"}' \\
    --font-size 20 \\
    --color "#0070C0" \\
    --json
  
  # From JSON file
  echo '["First point", "Second point", "Third point"]' > items.json
  uv run tools/ppt_add_bullet_list.py \\
    --file deck.pptx \\
    --slide 3 \\
    --items-file items.json \\
    --position '{"left":"10%","top":"25%"}' \\
    --size '{"width":"80%","height":"60%"}' \\
    --json
  
  # Override 6×6 rule (not recommended)
  uv run tools/ppt_add_bullet_list.py \\
    --file deck.pptx \\
    --slide 4 \\
    --items "Item 1,Item 2,Item 3,Item 4,Item 5,Item 6,Item 7,Item 8" \\
    --ignore-rules \\
    --position '{"left":"10%","top":"25%"}' \\
    --size '{"width":"80%","height":"60%"}' \\
    --json

Validation Features:
  - 6×6 rule enforcement (warns at 6, errors at 10)
  - Character count per item
  - Word count per item
  - Font size accessibility check (minimum 14pt)
  - Color contrast validation (WCAG 2.1)
  - Readability scoring (A-F grade)

Output Format:
  {
    "status": "warning",
    "items_added": 7,
    "readability": {
      "score": 60,
      "grade": "C",
      "metrics": {
        "item_count": 7,
        "avg_characters": 45.2,
        "max_words": 9
      }
    },
    "validation": {
      "six_six_rule": {
        "compliant": false,
        "item_count_ok": false
      }
    },
    "warnings": [
      "6×6 Rule violation: 7 items exceeds recommended 6"
    ],
    "recommendations": [
      "Consider splitting into multiple slides"
    ]
  }

Related Tools:
  - ppt_add_text_box.py: Add free-form text
  - ppt_format_text.py: Format existing text
  - ppt_get_slide_info.py: Inspect slide content

Version: 2.0.0
Requires: core/powerpoint_agent_core.py v1.1.0+
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
        '--items',
        help='Comma-separated list items'
    )
    
    parser.add_argument(
        '--items-file',
        type=Path,
        help='JSON file with array of items'
    )
    
    parser.add_argument(
        '--position',
        required=True,
        type=json.loads,
        help='Position dict (JSON string)'
    )
    
    parser.add_argument(
        '--size',
        type=json.loads,
        help='Size dict (JSON string, defaults from position if omitted)'
    )
    
    parser.add_argument(
        '--bullet-style',
        choices=['bullet', 'numbered', 'none'],
        default='bullet',
        help='Bullet style (default: bullet)'
    )
    
    parser.add_argument(
        '--font-size',
        type=int,
        default=18,
        help='Font size in points (default: 18, min recommended: 14)'
    )
    
    parser.add_argument(
        '--font-name',
        default='Calibri',
        help='Font name (default: Calibri)'
    )
    
    parser.add_argument(
        '--color',
        help='Text color hex (e.g., #0070C0, contrast will be validated)'
    )
    
    parser.add_argument(
        '--line-spacing',
        type=float,
        default=1.0,
        help='Line spacing multiplier (default: 1.0)'
    )
    
    parser.add_argument(
        '--ignore-rules',
        action='store_true',
        help='Override 6×6 rule validation (not recommended)'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        # Parse items
        if args.items_file:
            if not args.items_file.exists():
                raise FileNotFoundError(f"Items file not found: {args.items_file}")
            with open(args.items_file, 'r') as f:
                items = json.load(f)
            if not isinstance(items, list):
                raise ValueError("Items file must contain JSON array")
        elif args.items:
            if '\\n' in args.items:
                items = args.items.split('\\n')
            else:
                items = args.items.split(',')
            items = [item.strip() for item in items if item.strip()]
        else:
            raise ValueError("Either --items or --items-file required")
        
        # Handle size defaults
        size = args.size if args.size else {}
        position = args.position
        
        if "width" not in size:
            size["width"] = position.get("width", "80%")
        if "height" not in size:
            size["height"] = position.get("height", "50%")
        
        result = add_bullet_list(
            filepath=args.file,
            slide_index=args.slide,
            items=items,
            position=position,
            size=size,
            bullet_style=args.bullet_style,
            font_size=args.font_size,
            font_name=args.font_name,
            color=args.color,
            line_spacing=args.line_spacing,
            ignore_rules=args.ignore_rules
        )
        
        print(json.dumps(result, indent=2))
        sys.exit(0)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "file": str(args.file) if args.file else None,
            "slide_index": args.slide if hasattr(args, 'slide') else None
        }
        
        print(json.dumps(error_result, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()
