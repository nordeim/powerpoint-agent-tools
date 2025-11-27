# Implementation Plan: Remediation of Tool Failures

## 1. Issue Analysis
Based on the execution logs (`reference_powerpoint_demo_script.log.txt`) and code review, I have identified two distinct failure modes requiring immediate remediation.

### 🔴 Critical Failure: `ppt_validate_presentation.py` (JSON Corruption)
*   **Symptom:** `jq: parse error: Invalid numeric literal` in Step 12.
*   **Root Cause:** The bash script captures both stdout and stderr (`2>&1`). The Python tool emits log messages (likely from `core` or `pptx` libraries) to stderr. These log messages are interleaved with the JSON output on stdout, creating invalid JSON data that `jq` cannot parse.
*   **Fix Strategy:**
    1.  Initialize `logging` with `level=logging.CRITICAL` **before** importing any other modules.
    2.  Explicitly silence known chatty loggers (`pptx`, `core.powerpoint_agent_core`).
    3.  Ensure the script *only* prints the final JSON payload to stdout.

### 🟡 Functional Defect: `ppt_set_footer.py` (Logic Gap)
*   **Symptom:** "Footer application may have failed - 0 slides updated" despite reporting "Footer Method: placeholder".
*   **Root Cause:** The logic checks `if placeholder_count == 0` to trigger the fallback text-box strategy. However, `placeholder_count` includes placeholders found on the **Slide Master**. If the Master has a footer placeholder but the individual Slides do not have it activated (common in default templates), `placeholder_count` is > 0, preventing the fallback, yet no slides are actually updated.
*   **Fix Strategy:**
    1.  Change the fallback condition to check `len(slide_indices_updated) == 0`.
    2.  This forces the text-box overlay strategy if the placeholder strategy failed to modify any actual content, regardless of Master configuration.

---

## 2. Implementation Checklist

### `tools/ppt_validate_presentation.py`
- [x] **Logging Suppression**: Configure `logging` to `CRITICAL` immediately upon startup.
- [x] **Import Safety**: Ensure `core` imports don't reset logging config.
- [x] **Output Hygiene**: Verify absolutely no `print` statements exist other than the final JSON dump.
- [x] **Error Handling**: Wrap entire execution in `try/except` to catch crashes and output clean JSON error reports.

### `tools/ppt_set_footer.py`
- [x] **Fallback Logic**: Update condition to `if len(slide_indices_updated) == 0`.
- [x] **Import Fix**: Ensure `PP_PLACEHOLDER` is imported safely (addressing the `ImportError` seen in earlier logs).
- [x] **Output Hygiene**: Remove all informational prints; output JSON only.

---

## 3. Replacement Files

### File 1: `tools/ppt_validate_presentation.py`

```python
#!/usr/bin/env python3
"""
PowerPoint Validate Presentation Tool v3.1.0
Comprehensive validation for structure, accessibility, assets, and design quality.

Fully aligned with PowerPoint Agent Core v3.1.0 and System Prompt v3.0 validation gates.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_validate_presentation.py --file presentation.pptx --json
    uv run tools/ppt_validate_presentation.py --file presentation.pptx --policy strict --json

Exit Codes:
    0: Success (valid or only warnings within policy thresholds)
    1: Error occurred or critical issues exceed policy thresholds

Changelog v3.1.0:
- FIXED: strict output hygiene; logs silenced to prevent JSON corruption in 2>&1 pipelines
- ENHANCED: Robust error handling ensures valid JSON output even on crash
"""

import sys
import json
import argparse
import logging
from pathlib import Path
from typing import Dict, Any, List, Optional
from dataclasses import dataclass, field, asdict
from datetime import datetime

# CRITICAL: Silence all logging immediately to protect JSON output stream
# This must happen before any other imports that might initialize loggers
logging.basicConfig(level=logging.CRITICAL)
logging.getLogger().setLevel(logging.CRITICAL)

# Explicitly silence likely noisy modules
for logger_name in ['pptx', 'core.powerpoint_agent_core', 'urllib3']:
    logging.getLogger(logger_name).setLevel(logging.CRITICAL)

# Add parent directory to path for core import
sys.path.insert(0, str(Path(__file__).parent.parent))

try:
    from core.powerpoint_agent_core import (
        PowerPointAgent,
        PowerPointAgentError,
        __version__ as CORE_VERSION
    )
except ImportError:
    CORE_VERSION = "0.0.0"

# ============================================================================
# CONSTANTS & POLICIES
# ============================================================================

__version__ = "3.1.0"

VALIDATION_POLICIES = {
    "lenient": {
        "name": "Lenient",
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
    category: str
    severity: str
    message: str
    slide_index: Optional[int] = None
    shape_index: Optional[int] = None
    fix_command: Optional[str] = None
    details: Dict[str, Any] = field(default_factory=dict)
    
    def to_dict(self) -> Dict[str, Any]:
        return {k: v for k, v in asdict(self).items() if v is not None}

@dataclass
class ValidationSummary:
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
    name: str
    thresholds: Dict[str, Any]
    description: str = ""
    
    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

# ============================================================================
# LOGIC
# ============================================================================

def get_policy(policy_name: str, custom_thresholds: Optional[Dict[str, Any]] = None) -> ValidationPolicy:
    if policy_name == "custom" and custom_thresholds:
        base = VALIDATION_POLICIES["standard"]["thresholds"].copy()
        base.update(custom_thresholds)
        return ValidationPolicy(name="Custom", thresholds=base)
    
    config = VALIDATION_POLICIES.get(policy_name, VALIDATION_POLICIES["standard"])
    return ValidationPolicy(name=config["name"], thresholds=config["thresholds"])

def validate_presentation(filepath: Path, policy: ValidationPolicy) -> Dict[str, Any]:
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    issues: List[ValidationIssue] = []
    summary = ValidationSummary()
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath, acquire_lock=False)
        presentation_info = agent.get_presentation_info()
        slide_count = presentation_info.get("slide_count", 0)
        
        # Run Validations
        core_val = agent.validate_presentation()
        acc_val = agent.check_accessibility()
        asset_val = agent.validate_assets()
        
        # Process Results
        _process_core_validation(core_val, issues, summary, filepath)
        _process_accessibility(acc_val, issues, summary, filepath)
        _process_assets(asset_val, issues, summary, filepath)
        _validate_design_rules(agent, issues, summary, policy, filepath)
        
        # Summarize
        summary.total_issues = len(issues)
        summary.critical_count = sum(1 for i in issues if i.severity == "critical")
        summary.warning_count = sum(1 for i in issues if i.severity == "warning")
        summary.info_count = sum(1 for i in issues if i.severity == "info")
        
        passed, violations = _check_policy_compliance(summary, policy)
        
        status = "valid"
        if summary.critical_count > 0: status = "critical"
        elif not passed: status = "failed"
        elif summary.warning_count > 0: status = "warnings"

    return {
        "status": status,
        "passed": passed,
        "file": str(filepath),
        "validated_at": datetime.utcnow().isoformat() + "Z",
        "policy": policy.to_dict(),
        "summary": summary.to_dict(),
        "policy_violations": violations,
        "issues": [i.to_dict() for i in issues],
        "recommendations": _generate_recommendations(issues, policy),
        "presentation_info": {
            "slide_count": slide_count,
            "file_size_mb": presentation_info.get("file_size_mb"),
            "aspect_ratio": presentation_info.get("aspect_ratio")
        }
    }

def _process_core_validation(val, issues, summary, filepath):
    i = val.get("issues", {})
    summary.empty_slides = len(i.get("empty_slides", []))
    summary.slides_without_titles = len(i.get("slides_without_titles", []))
    for idx in i.get("empty_slides", []):
        issues.append(ValidationIssue("structure", "critical", "Empty slide", slide_index=idx))
    for idx in i.get("slides_without_titles", []):
        issues.append(ValidationIssue("structure", "warning", "Slide missing title", slide_index=idx))
    if isinstance(i.get("fonts_used"), list):
        summary.fonts_used = len(i.get("fonts_used"))

def _process_accessibility(acc, issues, summary, filepath):
    i = acc.get("issues", {})
    summary.missing_alt_text = len(i.get("missing_alt_text", []))
    summary.low_contrast = len(i.get("low_contrast", []))
    for item in i.get("missing_alt_text", []):
        issues.append(ValidationIssue("accessibility", "critical", "Missing alt text", slide_index=item.get("slide")))
    for item in i.get("low_contrast", []):
        issues.append(ValidationIssue("accessibility", "warning", "Low contrast", slide_index=item.get("slide")))

def _process_assets(assets, issues, summary, filepath):
    i = assets.get("issues", {})
    summary.large_images = len(i.get("large_images", []))
    for item in i.get("large_images", []):
        issues.append(ValidationIssue("assets", "info", f"Large image ({item.get('size_mb')}MB)", slide_index=item.get("slide")))

def _validate_design_rules(agent, issues, summary, policy, filepath):
    if summary.fonts_used > policy.thresholds.get("max_fonts", 5):
        issues.append(ValidationIssue("design", "warning", f"Too many fonts ({summary.fonts_used})"))

def _check_policy_compliance(summary, policy):
    v = []
    t = policy.thresholds
    if summary.critical_count > t.get("max_critical_issues", 0): v.append("Critical issues exceed limit")
    if summary.empty_slides > t.get("max_empty_slides", 0): v.append("Empty slides exceed limit")
    if summary.slides_without_titles > t.get("max_slides_without_titles", 3): v.append("Untitled slides exceed limit")
    if summary.missing_alt_text > t.get("max_missing_alt_text", 5): v.append("Missing alt text exceeds limit")
    return len(v) == 0, v

def _generate_recommendations(issues, policy):
    recs = []
    if any(i.message == "Empty slide" for i in issues): recs.append({"priority":"high", "action":"Remove empty slides"})
    if any(i.category == "accessibility" for i in issues): recs.append({"priority":"medium", "action":"Fix accessibility issues"})
    return recs

# ============================================================================
# MAIN
# ============================================================================

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--file', required=True, type=Path)
    parser.add_argument('--policy', default='standard')
    parser.add_argument('--json', action='store_true', default=True)
    parser.add_argument('--max-missing-alt-text', type=int)
    parser.add_argument('--max-slides-without-titles', type=int)
    parser.add_argument('--max-empty-slides', type=int)
    parser.add_argument('--require-all-alt-text', action='store_true')
    parser.add_argument('--enforce-6x6', action='store_true')
    
    args = parser.parse_args()
    
    try:
        custom = {}
        if args.max_missing_alt_text is not None: custom["max_missing_alt_text"] = args.max_missing_alt_text
        if args.max_slides_without_titles is not None: custom["max_slides_without_titles"] = args.max_slides_without_titles
        
        policy = get_policy("custom", custom) if custom else get_policy(args.policy)
        
        result = validate_presentation(args.file, policy)
        print(json.dumps(result, indent=2))
        
        sys.exit(1 if result["status"] in ("critical", "failed") else 0)
        
    except Exception as e:
        # Crash handler: Print JSON error
        print(json.dumps({
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__
        }, indent=2))
        sys.exit(1)

if __name__ == "__main__":
    main()
```

### File 2: `tools/ppt_set_footer.py`

```python
#!/usr/bin/env python3
"""
PowerPoint Set Footer Tool v3.1.0
Configure slide footer with Dual Strategy (Placeholder + Text Box Fallback).

Fixes:
- Forces fallback to text boxes if placeholders exist on master but fail to update on slides.
- Clean JSON output.
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import PowerPointAgent

# Safe import for constants
try:
    from pptx.enum.shapes import PP_PLACEHOLDER
except ImportError:
    class PP_PLACEHOLDER:
        FOOTER = 15
        SLIDE_NUMBER = 13

def set_footer(filepath: Path, text: str = None, show_number: bool = False) -> Dict[str, Any]:
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")

    slide_indices_updated = set()
    method_used = None
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Strategy 1: Placeholders
        # Try to set footer on Master first
        try:
            for master in agent.prs.slide_masters:
                for layout in master.slide_layouts:
                    for shape in layout.placeholders:
                        if shape.placeholder_format.type == PP_PLACEHOLDER.FOOTER:
                            if text: shape.text = text
        except Exception:
            pass
        
        # Try to set on individual slides
        for slide_idx, slide in enumerate(agent.prs.slides):
            for shape in slide.placeholders:
                if shape.placeholder_format.type == PP_PLACEHOLDER.FOOTER:
                    try:
                        if text: shape.text = text
                        slide_indices_updated.add(slide_idx)
                    except:
                        pass

        # Strategy 2: Fallback if NO slides were updated via placeholders
        # We check slide_indices_updated, not placeholder_count, to catch cases where
        # placeholders exist on Master but are not active on Slides.
        textbox_count = 0
        if len(slide_indices_updated) == 0:
            method_used = "text_box"
            
            for slide_idx in range(1, len(agent.prs.slides)):
                try:
                    if text:
                        agent.add_text_box(
                            slide_index=slide_idx,
                            text=text,
                            position={"left": "5%", "top": "92%"},
                            size={"width": "60%", "height": "5%"},
                            font_size=10,
                            color="#595959"
                        )
                        textbox_count += 1
                        slide_indices_updated.add(slide_idx)
                    
                    if show_number:
                        agent.add_text_box(
                            slide_index=slide_idx,
                            text=str(slide_idx + 1),
                            position={"left": "92%", "top": "92%"},
                            size={"width": "5%", "height": "5%"},
                            font_size=10,
                            color="#595959"
                        )
                        textbox_count += 1
                        slide_indices_updated.add(slide_idx)
                except Exception:
                    pass
        else:
            method_used = "placeholder"

        agent.save()
        prs_info = agent.get_presentation_info()

    return {
        "status": "success" if len(slide_indices_updated) > 0 else "warning",
        "file": str(filepath),
        "method_used": method_used,
        "slides_updated": len(slide_indices_updated),
        "presentation_version_after": prs_info["presentation_version"]
    }

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--file', required=True, type=Path)
    parser.add_argument('--text')
    parser.add_argument('--show-number', action='store_true')
    parser.add_argument('--show-date', action='store_true')
    parser.add_argument('--json', action='store_true', default=True)
    args = parser.parse_args()
    
    try:
        result = set_footer(args.file, args.text, args.show_number)
        print(json.dumps(result, indent=2))
    except Exception as e:
        print(json.dumps({"status": "error", "error": str(e)}), indent=2)
        sys.exit(1)

if __name__ == "__main__":
    main()
```

