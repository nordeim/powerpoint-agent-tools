#!/usr/bin/env python3
"""
PowerPoint Validate Presentation Tool
Comprehensive validation for missing assets, errors, quality issues

Usage:
    uv python ppt_validate_presentation.py --file presentation.pptx --json

Exit Codes:
    0: Success (no issues or only warnings)
    1: Error occurred or critical issues found
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


def validate_presentation(filepath: Path) -> Dict[str, Any]:
    """Validate presentation for issues."""
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath, acquire_lock=False)
        
        # Run all validation checks
        validation = agent.validate_presentation()
        accessibility = agent.check_accessibility()
        assets = agent.validate_assets()
    
    # Combine results
    all_issues = {
        "validation": validation,
        "accessibility": accessibility,
        "assets": assets
    }
    
    # Calculate totals
    total_issues = (
        validation.get("total_issues", 0) +
        accessibility.get("total_issues", 0) +
        assets.get("total_issues", 0)
    )
    
    # Determine severity
    has_critical = (
        len(validation.get("issues", {}).get("empty_slides", [])) > 0 or
        len(accessibility.get("issues", {}).get("missing_alt_text", [])) > 5
    )
    
    return {
        "status": "critical" if has_critical else ("warnings" if total_issues > 0 else "valid"),
        "file": str(filepath),
        "total_issues": total_issues,
        "summary": {
            "empty_slides": len(validation.get("issues", {}).get("empty_slides", [])),
            "slides_without_titles": len(validation.get("issues", {}).get("slides_without_titles", [])),
            "missing_alt_text": len(accessibility.get("issues", {}).get("missing_alt_text", [])),
            "low_contrast": len(accessibility.get("issues", {}).get("low_contrast", [])),
            "low_resolution_images": len(assets.get("issues", {}).get("low_resolution_images", [])),
            "large_images": len(assets.get("issues", {}).get("large_images", []))
        },
        "details": all_issues,
        "recommendations": generate_recommendations(all_issues)
    }


def generate_recommendations(issues: Dict[str, Any]) -> list:
    """Generate actionable recommendations based on issues."""
    recommendations = []
    
    validation = issues.get("validation", {}).get("issues", {})
    accessibility = issues.get("accessibility", {}).get("issues", {})
    assets = issues.get("assets", {}).get("issues", {})
    
    # Empty slides
    if validation.get("empty_slides"):
        recommendations.append({
            "priority": "high",
            "issue": "Empty slides found",
            "action": "Remove empty slides or add content",
            "affected_slides": validation["empty_slides"]
        })
    
    # Missing titles
    if validation.get("slides_without_titles"):
        recommendations.append({
            "priority": "medium",
            "issue": "Slides without titles",
            "action": "Add titles using ppt_set_title.py",
            "affected_slides": validation["slides_without_titles"][:5]
        })
    
    # Missing alt text
    if len(accessibility.get("missing_alt_text", [])) > 0:
        recommendations.append({
            "priority": "high",
            "issue": "Images without alt text",
            "action": "Add alt text using ppt_set_image_properties.py",
            "count": len(accessibility["missing_alt_text"])
        })
    
    # Low contrast
    if len(accessibility.get("low_contrast", [])) > 0:
        recommendations.append({
            "priority": "medium",
            "issue": "Low color contrast text",
            "action": "Improve text/background contrast for readability",
            "count": len(accessibility["low_contrast"])
        })
    
    # Low resolution images
    if assets.get("low_resolution_images"):
        recommendations.append({
            "priority": "medium",
            "issue": "Low resolution images",
            "action": "Replace with higher resolution images (96 DPI minimum)",
            "count": len(assets["low_resolution_images"])
        })
    
    # Large images
    if assets.get("large_images"):
        recommendations.append({
            "priority": "low",
            "issue": "Large images (>2MB)",
            "action": "Compress images using ppt_replace_image.py --compress",
            "count": len(assets["large_images"])
        })
    
    return recommendations


def main():
    parser = argparse.ArgumentParser(
        description="Validate PowerPoint presentation for quality issues",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Validate presentation
  uv python ppt_validate_presentation.py \\
    --file presentation.pptx \\
    --json

Validation Checks:
  1. Structure Validation:
     - Empty slides
     - Slides without titles
     - Text overflow
     - Inconsistent formatting
  
  2. Accessibility (WCAG 2.1):
     - Missing alt text on images
     - Low color contrast
     - Missing slide titles
     - Reading order issues
  
  3. Asset Quality:
     - Low resolution images (<96 DPI)
     - Large images (>2MB)
     - Total embedded file size
     - Missing external links

Output Status:
  - "valid": No issues found
  - "warnings": Minor issues that should be fixed
  - "critical": Major issues that must be fixed

Example Output:
{
  "status": "warnings",
  "total_issues": 5,
  "summary": {
    "empty_slides": 0,
    "slides_without_titles": 2,
    "missing_alt_text": 3,
    "low_contrast": 0
  },
  "recommendations": [
    {
      "priority": "high",
      "issue": "Images without alt text",
      "action": "Add alt text using ppt_set_image_properties.py",
      "count": 3
    }
  ]
}

Use Cases:
  - Pre-publication checks
  - Accessibility compliance
  - Quality assurance
  - Best practices enforcement
  - File size optimization

Fixing Issues:
  Use these tools to fix found issues:
  - Empty slides: ppt_delete_slide.py
  - Missing titles: ppt_set_title.py
  - Missing alt text: ppt_set_image_properties.py
  - Large images: ppt_replace_image.py --compress
  - Low contrast: ppt_format_text.py --color

Best Practices:
  - Run validation before sharing
  - Fix critical issues first
  - Address accessibility for compliance
  - Optimize large presentations
  - Regular validation during development
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file to validate'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default)'
    )
    
    args = parser.parse_args()
    
    try:
        result = validate_presentation(filepath=args.file)
        
        print(json.dumps(result, indent=2))
        
        # Exit with error code if critical issues
        if result["status"] == "critical":
            sys.exit(1)
        else:
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
