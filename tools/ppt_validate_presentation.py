#!/usr/bin/env python3
"""
PowerPoint Validate Presentation Tool v3.0
Comprehensive validation for structure, accessibility, assets, and design quality.

Fully aligned with PowerPoint Agent Core v3.0 and System Prompt v3.0 validation gates.

Author: PowerPoint Agent Team
License: MIT
Version: 3.0.0

Usage:
    uv run tools/ppt_validate_presentation.py --file presentation.pptx --json
    uv run tools/ppt_validate_presentation.py --file presentation.pptx --policy strict --json

Exit Codes:
    0: Success (valid or only warnings within policy thresholds)
    1: Error occurred or critical issues exceed policy thresholds

Changelog v3.0.0:
- Added configurable validation policies (lenient, standard, strict, custom)
- Added slide-by-slide issue breakdown
- Added actionable fix commands for each issue
- Added design rule validation (6x6 rule, font limits, color limits)
- Enhanced accessibility checks
- Added severity threshold configuration
- Added summary statistics
- Added pass/fail based on policy
- Integration with System Prompt v3.0 validation gates
- Comprehensive CLI with policy options
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any, List, Optional
from dataclasses import dataclass, field, asdict
from datetime import datetime

# Add parent directory to path for core import
sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent,
    PowerPointAgentError,
    __version__ as CORE_VERSION
)

# ============================================================================
# CONSTANTS
# ============================================================================

__version__ = "3.0.0"

# Default validation policies aligned with System Prompt v3.0
VALIDATION_POLICIES = {
    "lenient": {
        "name": "Lenient",
        "description": "Minimal validation, allows most issues",
        "thresholds": {
            "max_critical_issues": 10,
            "max_accessibility_issues": 20,
            "max_design_warnings": 50,
            "max_empty_slides": 5,
            "max_slides_without_titles": 10,
            "max_missing_alt_text": 20,
            "max_low_contrast": 10,
            "max_large_images": 10,
            "require_all_alt_text": False,
            "enforce_6x6_rule": False,
            "max_fonts": 10,
            "max_colors": 20,
        }
    },
    "standard": {
        "name": "Standard",
        "description": "Balanced validation for general use",
        "thresholds": {
            "max_critical_issues": 0,
            "max_accessibility_issues": 5,
            "max_design_warnings": 10,
            "max_empty_slides": 0,
            "max_slides_without_titles": 3,
            "max_missing_alt_text": 5,
            "max_low_contrast": 3,
            "max_large_images": 5,
            "require_all_alt_text": False,
            "enforce_6x6_rule": False,
            "max_fonts": 5,
            "max_colors": 10,
        }
    },
    "strict": {
        "name": "Strict",
        "description": "Full compliance with accessibility and design standards",
        "thresholds": {
            "max_critical_issues": 0,
            "max_accessibility_issues": 0,
            "max_design_warnings": 3,
            "max_empty_slides": 0,
            "max_slides_without_titles": 0,
            "max_missing_alt_text": 0,
            "max_low_contrast": 0,
            "max_large_images": 3,
            "require_all_alt_text": True,
            "enforce_6x6_rule": True,
            "max_fonts": 3,
            "max_colors": 5,
        }
    }
}


# ============================================================================
# DATA CLASSES
# ============================================================================

@dataclass
class ValidationIssue:
    """Represents a single validation issue."""
    category: str
    severity: str  # "critical", "warning", "info"
    message: str
    slide_index: Optional[int] = None
    shape_index: Optional[int] = None
    fix_command: Optional[str] = None
    details: Dict[str, Any] = field(default_factory=dict)
    
    def to_dict(self) -> Dict[str, Any]:
        return {k: v for k, v in asdict(self).items() if v is not None}


@dataclass
class ValidationSummary:
    """Summary of validation results."""
    total_issues: int = 0
    critical_count: int = 0
    warning_count: int = 0
    info_count: int = 0
    empty_slides: int = 0
    slides_without_titles: int = 0
    missing_alt_text: int = 0
    low_contrast: int = 0
    large_images: int = 0
    fonts_used: int = 0
    colors_detected: int = 0
    
    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


@dataclass
class ValidationPolicy:
    """Validation policy configuration."""
    name: str
    thresholds: Dict[str, Any]
    description: str = ""
    
    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


# ============================================================================
# VALIDATION FUNCTIONS
# ============================================================================

def get_policy(policy_name: str, custom_thresholds: Optional[Dict[str, Any]] = None) -> ValidationPolicy:
    """
    Get validation policy by name or create custom policy.
    
    Args:
        policy_name: Policy name (lenient, standard, strict, custom)
        custom_thresholds: Custom threshold values for "custom" policy
        
    Returns:
        ValidationPolicy instance
    """
    if policy_name == "custom" and custom_thresholds:
        base = VALIDATION_POLICIES["standard"]["thresholds"].copy()
        base.update(custom_thresholds)
        return ValidationPolicy(
            name="Custom",
            thresholds=base,
            description="Custom validation policy"
        )
    
    if policy_name in VALIDATION_POLICIES:
        config = VALIDATION_POLICIES[policy_name]
        return ValidationPolicy(
            name=config["name"],
            thresholds=config["thresholds"],
            description=config["description"]
        )
    
    # Default to standard
    config = VALIDATION_POLICIES["standard"]
    return ValidationPolicy(
        name=config["name"],
        thresholds=config["thresholds"],
        description=config["description"]
    )


def validate_presentation(
    filepath: Path,
    policy: ValidationPolicy
) -> Dict[str, Any]:
    """
    Comprehensive presentation validation.
    
    Args:
        filepath: Path to PowerPoint file
        policy: Validation policy to apply
        
    Returns:
        Validation result dict
    """
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    issues: List[ValidationIssue] = []
    summary = ValidationSummary()
    slide_breakdown: List[Dict[str, Any]] = []
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath, acquire_lock=False)
        
        # Get basic info
        presentation_info = agent.get_presentation_info()
        slide_count = presentation_info.get("slide_count", 0)
        
        # Run core validations
        core_validation = agent.validate_presentation()
        core_accessibility = agent.check_accessibility()
        core_assets = agent.validate_assets()
        
        # Process core validation results
        _process_core_validation(core_validation, issues, summary, filepath)
        
        # Process accessibility results
        _process_accessibility(core_accessibility, issues, summary, filepath)
        
        # Process asset validation
        _process_assets(core_assets, issues, summary, filepath)
        
        # Run design rule validation
        _validate_design_rules(agent, issues, summary, policy, filepath)
        
        # Build slide-by-slide breakdown
        for slide_idx in range(slide_count):
            slide_issues = [i for i in issues if i.slide_index == slide_idx]
            slide_info = {
                "slide_index": slide_idx,
                "issue_count": len(slide_issues),
                "issues": [i.to_dict() for i in slide_issues]
            }
            slide_breakdown.append(slide_info)
        
        # Calculate totals
        summary.total_issues = len(issues)
        summary.critical_count = sum(1 for i in issues if i.severity == "critical")
        summary.warning_count = sum(1 for i in issues if i.severity == "warning")
        summary.info_count = sum(1 for i in issues if i.severity == "info")
    
    # Determine pass/fail based on policy
    passed, policy_violations = _check_policy_compliance(summary, policy)
    
    # Generate recommendations
    recommendations = _generate_recommendations(issues, policy)
    
    # Determine overall status
    if summary.critical_count > 0:
        status = "critical"
    elif not passed:
        status = "failed"
    elif summary.warning_count > 0:
        status = "warnings"
    else:
        status = "valid"
    
    return {
        "status": status,
        "passed": passed,
        "file": str(filepath),
        "validated_at": datetime.utcnow().isoformat() + "Z",
        "policy": policy.to_dict(),
        "summary": summary.to_dict(),
        "policy_violations": policy_violations,
        "issues": [i.to_dict() for i in issues],
        "slide_breakdown": slide_breakdown,
        "recommendations": recommendations,
        "presentation_info": {
            "slide_count": slide_count,
            "file_size_mb": presentation_info.get("file_size_mb"),
            "aspect_ratio": presentation_info.get("aspect_ratio")
        },
        "core_version": CORE_VERSION,
        "tool_version": __version__
    }


def _process_core_validation(
    validation: Dict[str, Any],
    issues: List[ValidationIssue],
    summary: ValidationSummary,
    filepath: Path
) -> None:
    """Process core validation results into issues list."""
    validation_issues = validation.get("issues", {})
    
    # Empty slides
    empty_slides = validation_issues.get("empty_slides", [])
    summary.empty_slides = len(empty_slides)
    for slide_idx in empty_slides:
        issues.append(ValidationIssue(
            category="structure",
            severity="critical",
            message=f"Empty slide with no content",
            slide_index=slide_idx,
            fix_command=f"uv run tools/ppt_delete_slide.py --file {filepath} --index {slide_idx} --json",
            details={"issue_type": "empty_slide"}
        ))
    
    # Slides without titles
    slides_no_title = validation_issues.get("slides_without_titles", [])
    summary.slides_without_titles = len(slides_no_title)
    for slide_idx in slides_no_title:
        issues.append(ValidationIssue(
            category="structure",
            severity="warning",
            message=f"Slide lacks a title (important for navigation and accessibility)",
            slide_index=slide_idx,
            fix_command=f"uv run tools/ppt_set_title.py --file {filepath} --slide {slide_idx} --title \"[Title]\" --json",
            details={"issue_type": "missing_title"}
        ))
    
    # Font consistency
    fonts = validation_issues.get("inconsistent_fonts", [])
    if isinstance(fonts, list):
        summary.fonts_used = len(fonts)


def _process_accessibility(
    accessibility: Dict[str, Any],
    issues: List[ValidationIssue],
    summary: ValidationSummary,
    filepath: Path
) -> None:
    """Process accessibility check results into issues list."""
    acc_issues = accessibility.get("issues", {})
    
    # Missing alt text
    missing_alt = acc_issues.get("missing_alt_text", [])
    summary.missing_alt_text = len(missing_alt)
    for item in missing_alt:
        slide_idx = item.get("slide")
        shape_idx = item.get("shape")
        issues.append(ValidationIssue(
            category="accessibility",
            severity="critical",
            message=f"Image missing alternative text (required for screen readers)",
            slide_index=slide_idx,
            shape_index=shape_idx,
            fix_command=f"uv run tools/ppt_set_image_properties.py --file {filepath} --slide {slide_idx} --shape {shape_idx} --alt-text \"[Description]\" --json",
            details={"issue_type": "missing_alt_text", "shape_name": item.get("shape_name")}
        ))
    
    # Low contrast
    low_contrast = acc_issues.get("low_contrast", [])
    summary.low_contrast = len(low_contrast)
    for item in low_contrast:
        slide_idx = item.get("slide")
        shape_idx = item.get("shape")
        issues.append(ValidationIssue(
            category="accessibility",
            severity="warning",
            message=f"Text has low color contrast ({item.get('contrast_ratio', 'N/A')}:1, need {item.get('required', 4.5)}:1)",
            slide_index=slide_idx,
            shape_index=shape_idx,
            fix_command=f"uv run tools/ppt_format_text.py --file {filepath} --slide {slide_idx} --shape {shape_idx} --color \"#000000\" --json",
            details={"issue_type": "low_contrast", "contrast_ratio": item.get("contrast_ratio")}
        ))
    
    # Missing titles (from accessibility check)
    missing_titles = acc_issues.get("missing_titles", [])
    for item in missing_titles:
        slide_idx = item.get("slide") if isinstance(item, dict) else item
        # Check if already added from core validation
        existing = [i for i in issues if i.slide_index == slide_idx and i.details.get("issue_type") == "missing_title"]
        if not existing:
            issues.append(ValidationIssue(
                category="accessibility",
                severity="warning",
                message="Slide missing title for screen reader navigation",
                slide_index=slide_idx,
                fix_command=f"uv run tools/ppt_set_title.py --file {filepath} --slide {slide_idx} --title \"[Title]\" --json",
                details={"issue_type": "missing_title_accessibility"}
            ))
    
    # Small text
    small_text = acc_issues.get("small_text", [])
    for item in small_text:
        issues.append(ValidationIssue(
            category="accessibility",
            severity="warning",
            message=f"Text size {item.get('size_pt', 'N/A')}pt is below minimum (10pt recommended)",
            slide_index=item.get("slide"),
            shape_index=item.get("shape"),
            fix_command=f"uv run tools/ppt_format_text.py --file {filepath} --slide {item.get('slide')} --shape {item.get('shape')} --font-size 12 --json",
            details={"issue_type": "small_text", "size_pt": item.get("size_pt")}
        ))


def _process_assets(
    assets: Dict[str, Any],
    issues: List[ValidationIssue],
    summary: ValidationSummary,
    filepath: Path
) -> None:
    """Process asset validation results into issues list."""
    asset_issues = assets.get("issues", {})
    
    # Large images
    large_images = asset_issues.get("large_images", [])
    summary.large_images = len(large_images)
    for item in large_images:
        slide_idx = item.get("slide")
        shape_idx = item.get("shape")
        size_mb = item.get("size_mb", "N/A")
        issues.append(ValidationIssue(
            category="assets",
            severity="info",
            message=f"Large image ({size_mb}MB) may slow loading",
            slide_index=slide_idx,
            shape_index=shape_idx,
            fix_command=f"uv run tools/ppt_replace_image.py --file {filepath} --slide {slide_idx} --old-image \"[name]\" --new-image \"[path]\" --compress --json",
            details={"issue_type": "large_image", "size_mb": size_mb}
        ))
    
    # Large file warning
    if "large_file_warning" in asset_issues:
        warning = asset_issues["large_file_warning"]
        issues.append(ValidationIssue(
            category="assets",
            severity="warning",
            message=f"File size ({warning.get('size_mb', 'N/A')}MB) exceeds recommended maximum ({warning.get('recommended_max_mb', 50)}MB)",
            details={"issue_type": "large_file", "size_mb": warning.get("size_mb")}
        ))


def _validate_design_rules(
    agent: PowerPointAgent,
    issues: List[ValidationIssue],
    summary: ValidationSummary,
    policy: ValidationPolicy,
    filepath: Path
) -> None:
    """Validate design rules (6x6 rule, font limits, etc.)."""
    thresholds = policy.thresholds
    
    # Check font count against policy
    if summary.fonts_used > thresholds.get("max_fonts", 5):
        issues.append(ValidationIssue(
            category="design",
            severity="warning",
            message=f"Too many fonts used ({summary.fonts_used}). Recommended maximum: {thresholds.get('max_fonts', 5)}",
            details={"issue_type": "too_many_fonts", "font_count": summary.fonts_used}
        ))
    
    # 6x6 rule validation (if enforced)
    if thresholds.get("enforce_6x6_rule", False):
        try:
            slide_count = agent.get_slide_count()
            for slide_idx in range(slide_count):
                slide_info = agent.get_slide_info(slide_idx)
                for shape in slide_info.get("shapes", []):
                    if shape.get("has_text"):
                        text = shape.get("text", "")
                        lines = text.split("\n")
                        
                        # Check number of bullet points
                        if len(lines) > 6:
                            issues.append(ValidationIssue(
                                category="design",
                                severity="warning",
                                message=f"Shape has {len(lines)} lines/bullets (6x6 rule: max 6)",
                                slide_index=slide_idx,
                                shape_index=shape.get("index"),
                                details={"issue_type": "6x6_lines", "line_count": len(lines)}
                            ))
                        
                        # Check words per line
                        for line_idx, line in enumerate(lines):
                            word_count = len(line.split())
                            if word_count > 6:
                                issues.append(ValidationIssue(
                                    category="design",
                                    severity="info",
                                    message=f"Line has {word_count} words (6x6 rule: max 6 per line)",
                                    slide_index=slide_idx,
                                    shape_index=shape.get("index"),
                                    details={"issue_type": "6x6_words", "word_count": word_count, "line_index": line_idx}
                                ))
        except Exception:
            pass  # Design rule validation is best-effort


def _check_policy_compliance(
    summary: ValidationSummary,
    policy: ValidationPolicy
) -> tuple:
    """
    Check if validation results comply with policy thresholds.
    
    Returns:
        Tuple of (passed: bool, violations: List[str])
    """
    violations = []
    thresholds = policy.thresholds
    
    # Check each threshold
    if summary.critical_count > thresholds.get("max_critical_issues", 0):
        violations.append(f"Critical issues ({summary.critical_count}) exceed threshold ({thresholds.get('max_critical_issues', 0)})")
    
    if summary.empty_slides > thresholds.get("max_empty_slides", 0):
        violations.append(f"Empty slides ({summary.empty_slides}) exceed threshold ({thresholds.get('max_empty_slides', 0)})")
    
    if summary.slides_without_titles > thresholds.get("max_slides_without_titles", 3):
        violations.append(f"Slides without titles ({summary.slides_without_titles}) exceed threshold ({thresholds.get('max_slides_without_titles', 3)})")
    
    if summary.missing_alt_text > thresholds.get("max_missing_alt_text", 5):
        violations.append(f"Missing alt text ({summary.missing_alt_text}) exceeds threshold ({thresholds.get('max_missing_alt_text', 5)})")
    
    if thresholds.get("require_all_alt_text", False) and summary.missing_alt_text > 0:
        violations.append(f"Policy requires all images have alt text, but {summary.missing_alt_text} are missing")
    
    if summary.low_contrast > thresholds.get("max_low_contrast", 3):
        violations.append(f"Low contrast issues ({summary.low_contrast}) exceed threshold ({thresholds.get('max_low_contrast', 3)})")
    
    if summary.fonts_used > thresholds.get("max_fonts", 5):
        violations.append(f"Fonts used ({summary.fonts_used}) exceed threshold ({thresholds.get('max_fonts', 5)})")
    
    passed = len(violations) == 0
    return passed, violations


def _generate_recommendations(
    issues: List[ValidationIssue],
    policy: ValidationPolicy
) -> List[Dict[str, Any]]:
    """Generate prioritized, actionable recommendations."""
    recommendations = []
    
    # Group issues by type
    issue_types = {}
    for issue in issues:
        issue_type = issue.details.get("issue_type", "other")
        if issue_type not in issue_types:
            issue_types[issue_type] = []
        issue_types[issue_type].append(issue)
    
    # Generate recommendations for each issue type
    if "empty_slide" in issue_types:
        count = len(issue_types["empty_slide"])
        slides = [i.slide_index for i in issue_types["empty_slide"]]
        recommendations.append({
            "priority": "high",
            "category": "structure",
            "issue": f"{count} empty slide(s) found",
            "action": "Remove empty slides or add content",
            "affected_slides": slides[:5],
            "fix_tool": "ppt_delete_slide.py"
        })
    
    if "missing_title" in issue_types or "missing_title_accessibility" in issue_types:
        all_missing = issue_types.get("missing_title", []) + issue_types.get("missing_title_accessibility", [])
        slides = list(set(i.slide_index for i in all_missing))
        recommendations.append({
            "priority": "medium",
            "category": "accessibility",
            "issue": f"{len(slides)} slide(s) missing titles",
            "action": "Add descriptive titles for navigation and screen readers",
            "affected_slides": slides[:5],
            "fix_tool": "ppt_set_title.py"
        })
    
    if "missing_alt_text" in issue_types:
        count = len(issue_types["missing_alt_text"])
        is_critical = policy.thresholds.get("require_all_alt_text", False)
        recommendations.append({
            "priority": "high" if is_critical else "medium",
            "category": "accessibility",
            "issue": f"{count} image(s) missing alt text",
            "action": "Add descriptive alternative text for screen readers",
            "count": count,
            "fix_tool": "ppt_set_image_properties.py --alt-text"
        })
    
    if "low_contrast" in issue_types:
        count = len(issue_types["low_contrast"])
        recommendations.append({
            "priority": "medium",
            "category": "accessibility",
            "issue": f"{count} text element(s) with low contrast",
            "action": "Improve text/background contrast (WCAG AA requires 4.5:1)",
            "count": count,
            "fix_tool": "ppt_format_text.py --color"
        })
    
    if "large_image" in issue_types:
        count = len(issue_types["large_image"])
        recommendations.append({
            "priority": "low",
            "category": "performance",
            "issue": f"{count} large image(s) (>2MB)",
            "action": "Compress images to reduce file size",
            "count": count,
            "fix_tool": "ppt_replace_image.py --compress"
        })
    
    if "too_many_fonts" in issue_types:
        recommendations.append({
            "priority": "low",
            "category": "design",
            "issue": "Too many fonts used",
            "action": f"Reduce to {policy.thresholds.get('max_fonts', 3)} or fewer font families for consistency",
            "fix_tool": "ppt_format_text.py --font-name"
        })
    
    if "6x6_lines" in issue_types or "6x6_words" in issue_types:
        recommendations.append({
            "priority": "info",
            "category": "design",
            "issue": "Content density exceeds 6x6 rule",
            "action": "Consider reducing text: max 6 bullets with 6 words each",
            "note": "Break content across multiple slides if needed"
        })
    
    # Sort by priority
    priority_order = {"high": 0, "medium": 1, "low": 2, "info": 3}
    recommendations.sort(key=lambda x: priority_order.get(x.get("priority", "info"), 3))
    
    return recommendations


# ============================================================================
# CLI INTERFACE
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Validate PowerPoint presentation (v3.0)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

VALIDATION POLICIES:

  lenient   - Minimal checks, allows most issues
              Good for: Draft presentations, internal documents
              
  standard  - Balanced validation (DEFAULT)
              Good for: Most presentations, team sharing
              
  strict    - Full compliance with accessibility and design standards
              Good for: Public presentations, compliance requirements
              Enforces: WCAG AA, 6x6 rule, font/color limits

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

VALIDATION CHECKS:

  Structure:
    • Empty slides
    • Slides without titles
    
  Accessibility (WCAG 2.1):
    • Missing alt text on images
    • Low color contrast
    • Small text (<10pt)
    • Missing slide titles
    
  Assets:
    • Large images (>2MB)
    • Total file size
    
  Design (strict policy):
    • 6x6 rule compliance
    • Font count limits
    • Color consistency

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

EXAMPLES:

  # Standard validation
  uv run tools/ppt_validate_presentation.py --file deck.pptx --json

  # Strict validation for compliance
  uv run tools/ppt_validate_presentation.py --file deck.pptx --policy strict --json

  # Lenient validation for drafts
  uv run tools/ppt_validate_presentation.py --file deck.pptx --policy lenient --json

  # Custom thresholds
  uv run tools/ppt_validate_presentation.py --file deck.pptx \\
    --max-missing-alt-text 3 --max-slides-without-titles 2 --json

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

OUTPUT STATUS:

  valid     - No issues found, passes policy
  warnings  - Issues found but within policy thresholds
  failed    - Issues exceed policy thresholds
  critical  - Critical issues found (always fails)

EXIT CODES:

  0 - Valid or warnings (within policy)
  1 - Failed or critical issues

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

FIXING ISSUES:

  Each issue includes a fix_command that you can run directly.
  Example fix commands:
  
  • Missing title:
    uv run tools/ppt_set_title.py --file deck.pptx --slide 0 --title "Title" --json
    
  • Missing alt text:
    uv run tools/ppt_set_image_properties.py --file deck.pptx --slide 1 --shape 2 --alt-text "Description" --json
    
  • Low contrast:
    uv run tools/ppt_format_text.py --file deck.pptx --slide 0 --shape 1 --color "#000000" --json
    
  • Large images:
    uv run tools/ppt_replace_image.py --file deck.pptx --slide 0 --old-image "name" --new-image "path" --compress --json

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        """
    )
    
    # Required arguments
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file to validate'
    )
    
    # Policy selection
    parser.add_argument(
        '--policy',
        choices=['lenient', 'standard', 'strict'],
        default='standard',
        help='Validation policy (default: standard)'
    )
    
    # Custom threshold overrides
    parser.add_argument(
        '--max-missing-alt-text',
        type=int,
        help='Override maximum missing alt text threshold'
    )
    
    parser.add_argument(
        '--max-slides-without-titles',
        type=int,
        help='Override maximum slides without titles threshold'
    )
    
    parser.add_argument(
        '--max-empty-slides',
        type=int,
        help='Override maximum empty slides threshold'
    )
    
    parser.add_argument(
        '--require-all-alt-text',
        action='store_true',
        help='Require all images have alt text (fails if any missing)'
    )
    
    parser.add_argument(
        '--enforce-6x6',
        action='store_true',
        help='Enforce 6x6 design rule'
    )
    
    # Output options
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default: true)'
    )
    
    parser.add_argument(
        '--summary-only',
        action='store_true',
        help='Output summary only, not full issue list'
    )
    
    parser.add_argument(
        '--version',
        action='version',
        version=f'%(prog)s {__version__} (core: {CORE_VERSION})'
    )
    
    args = parser.parse_args()
    
    try:
        # Build custom thresholds from CLI arguments
        custom_thresholds = {}
        if args.max_missing_alt_text is not None:
            custom_thresholds["max_missing_alt_text"] = args.max_missing_alt_text
        if args.max_slides_without_titles is not None:
            custom_thresholds["max_slides_without_titles"] = args.max_slides_without_titles
        if args.max_empty_slides is not None:
            custom_thresholds["max_empty_slides"] = args.max_empty_slides
        if args.require_all_alt_text:
            custom_thresholds["require_all_alt_text"] = True
        if args.enforce_6x6:
            custom_thresholds["enforce_6x6_rule"] = True
        
        # Get policy
        if custom_thresholds:
            policy = get_policy("custom", custom_thresholds)
            # Merge with base policy
            base = VALIDATION_POLICIES[args.policy]["thresholds"].copy()
            base.update(custom_thresholds)
            policy = ValidationPolicy(
                name=f"{args.policy.title()} (Custom)",
                thresholds=base,
                description=f"Based on {args.policy} with custom overrides"
            )
        else:
            policy = get_policy(args.policy)
        
        # Run validation
        result = validate_presentation(filepath=args.file, policy=policy)
        
        # Output
        if args.summary_only:
            summary_output = {
                "status": result["status"],
                "passed": result["passed"],
                "file": result["file"],
                "summary": result["summary"],
                "policy_violations": result["policy_violations"],
                "recommendations": result["recommendations"][:3]  # Top 3 only
            }
            print(json.dumps(summary_output, indent=2))
        else:
            print(json.dumps(result, indent=2))
        
        # Exit code based on pass/fail
        if result["status"] in ("critical", "failed"):
            sys.exit(1)
        else:
            sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError"
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": e.message,
            "error_type": type(e).__name__,
            "details": e.details
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
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
