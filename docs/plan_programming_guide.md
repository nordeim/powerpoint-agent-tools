# Implementation Plan: PowerPoint Agent Tools Programming & Troubleshooting Guide

## 1. Objective
Create a comprehensive "Single Source of Truth" documentation (`PROGRAMMING_GUIDE.md`) for the PowerPoint Agent Tools ecosystem (v3.1.0). This document will codify the architectural principles, security standards, and operational lessons learned from recent debugging sessions to prevent regression by AI agents or human developers.

## 2. Document Structure & Content Strategy

### Section 1: Architectural Philosophy ("The Physics of the System")
*   **Hub-and-Spoke Model**: Define the relationship between `PowerPointAgent` (Core) and CLI Tools.
*   **Statelessness**: Explain why every operation is atomic (Open → Modify → Save → Close).
*   **The "Golden Rule" of Output**: Explain the strict separation of `STDOUT` (JSON only) and `STDERR` (Logs/Warnings), and the "Draconian" suppression required for pipeline compatibility.

### Section 2: Core Library Deep Dive (`powerpoint_agent_core.py`)
*   **Context Management**: How file locking and resource cleanup work.
*   **Versioning Logic**: Explain the Geometry-Aware Hashing (Position+Size+Text) used to detect changes.
*   **Security & Governance**:
    *   Path Traversal Protection (`allowed_base_dirs`).
    *   Approval Tokens for destructive operations.
*   **XML Manipulation**: Document the internal hacks used for Opacity and Z-Order (features missing in `python-pptx`).

### Section 3: The Perfect Tool Pattern (Template)
*   Provide a **Reference Implementation** of a tool script that includes:
    *   Standard boilerplate for imports.
    *   **Crucial**: `sys.stderr` redirection to `/dev/null` (The "Hygiene Fix").
    *   Argument parsing best practices.
    *   Handling Dictionary return values (The v3.1.0 API change).
    *   Global `try/except` block for JSON error reporting.

### Section 4: Common Pitfalls & Anti-Patterns ("The Graveyard of Bugs")
*   **The "Stdout Leak"**: Why `print()` is forbidden for anything but the final JSON.
*   **The "Master Slide" Fallacy**: Explaining why checking `placeholder_count` on a Master doesn't mean the Slide has a footer.
*   **The "Return Value" Trap**: Confusing v2.0 (int) vs v3.1.0 (dict) return types.
*   **The "Silent Clamp"**: Why we now raise `SlideNotFoundError` instead of auto-correcting indices.

### Section 5: Troubleshooting Playbook
*   **Symptom-Based Diagnosis**:
    *   `jq: parse error`: Diagnosis -> Stdout pollution.
    *   `ImportError`: Diagnosis -> Circular imports or path issues.
    *   `SlideNotFoundError`: Diagnosis -> 0-based vs 1-based confusion or "Ghost Slides".

## 3. Execution Steps
1.  **Drafting**: I will synthesize the knowledge from the recent `ppt_validate_presentation.py` and `ppt_set_footer.py` fixes.
2.  **Validation**: I will review the guide to ensure the code snippets match the exact v3.1.0 syntax used in the finalized core file.
3.  **Formatting**: Use clear Markdown with syntax highlighting, alert blocks, and table-based matrices.

---

I will now generate the `PROGRAMMING_GUIDE.md`.

```markdown
# PowerPoint Agent Tools: Programming & Troubleshooting Guide (v3.1.0)

**Version:** 1.0  
**Applicable Core Version:** 3.1.0  
**Target Audience:** AI Coding Agents, Human Developers, DevOps Engineers

---

## 1. Architectural Philosophy

The PowerPoint Agent Tools ecosystem is designed around a **Hub-and-Spoke** architecture optimized for **stateless, atomic, and machine-parseable** operations.

### 1.1 The Hub: `PowerPointAgent` (Core)
*   **Location**: `core/powerpoint_agent_core.py`
*   **Role**: The "Operating System." It handles all direct interaction with the `.pptx` binary, manages file locking, enforces security (path traversal, approval tokens), and calculates state hashes (versioning).
*   **Statefulness**: The *Core* instance is stateful while open (holds the `Presentation` object), but the *Tools* using it must treat it as ephemeral.

### 1.2 The Spokes: CLI Tools (`tools/*.py`)
*   **Location**: `tools/`
*   **Role**: Thin wrappers around Core methods.
*   **The Prime Directive**: **JSON IN, JSON OUT.**
*   **Statelessness**: A tool must fully initialize, execute its task, save, and exit. It assumes no memory of previous commands.

---

## 2. The "Golden Rules" of Development

Violating these rules will break the CI/CD pipeline or cause the AI orchestrator to fail.

### 🔒 Rule 1: Output Hygiene is Non-Negotiable
The most common failure mode is **Stdout Pollution**.
*   **The Problem**: Tools are often chained via pipes (e.g., `tool.py | jq .`). If a library (like `pptx`) prints a warning, or if you use `print("Doing work...")`, the output becomes invalid JSON.
*   **The Requirement**: 
    1.  **STDOUT** is exclusively for the final JSON payload.
    2.  **STDERR** is for logs/debugs (but be careful, some pipelines capture `2>&1`).
    3.  **The Fix**: In v3.1.0, we apply **Draconian Suppression** at the top of every tool.

### 🔒 Rule 2: Fail Safely with JSON
If a tool crashes, it **must** still print a valid JSON object to stdout so the orchestrator knows *why* it failed.
*   **Bad**: Python Traceback dumped to shell.
*   **Good**: `{"status": "error", "error": "IndexError...", "error_type": "IndexError"}`

### 🔒 Rule 3: Versioning is Mandatory
Every mutation (write) operation must capture the presentation state *before* and *after* the change.
*   **Why**: To detect race conditions and verify that the AI's intent was actually applied.
*   **Implementation**: Core v3.1.0 methods return `presentation_version_before` and `presentation_version_after`. Tools must pass these through.

---

## 3. Reference Tool Implementation (The "Perfect" Tool)

Use this template for creating new tools or refactoring existing ones.

```python
#!/usr/bin/env python3
"""
Standard Tool Template v3.1.0
"""
import sys
import os

# --- HYGIENE BLOCK START ---
# CRITICAL: Redirect stderr to /dev/null immediately to prevent library noise.
# This guarantees that `jq` or other parsers only see your JSON on stdout.
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

import json
import argparse
import logging
from pathlib import Path
from typing import Dict, Any

# Configure logging to null (redundant but safe)
logging.basicConfig(level=logging.CRITICAL)

# Add parent directory to path to import core
sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import PowerPointAgent, PowerPointAgentError

def my_tool_logic(filepath: Path, param: str) -> Dict[str, Any]:
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")

    # Context Manager handles Open/Save/Close/Locking
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # 1. Capture Version
        # Note: Core v3.1.0 mutation methods do this internally, 
        # but we capture the result from the method return.
        
        # 2. Execute Core Method
        # V3.1.0 CHANGE: Core methods return DICTIONARIES, not just ints.
        result = agent.some_mutation_method(param)
        
        # 3. Save
        agent.save()
        
        # 4. Get Final State (if not in result)
        final_info = agent.get_presentation_info()

    # 5. Construct Clean Response
    return {
        "status": "success",
        "file": str(filepath),
        "data": result,
        "presentation_version_after": final_info["presentation_version"]
    }

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--file', required=True, type=Path)
    parser.add_argument('--param', required=True)
    parser.add_argument('--json', action='store_true', default=True)
    args = parser.parse_args()

    try:
        result = my_tool_logic(args.file, args.param)
        # THE ONLY PRINT STATEMENT ALLOWED:
        sys.stdout.write(json.dumps(result, indent=2) + "\n")
        sys.exit(0)
    except Exception as e:
        # CATCH-ALL ERROR HANDLER
        error_res = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__
        }
        sys.stdout.write(json.dumps(error_res, indent=2) + "\n")
        sys.exit(1)

if __name__ == "__main__":
    main()
```

---

## 4. Core Library Internals & Gotchas

### 4.1 Geometry-Aware Versioning
*   **Concept**: A hash of the presentation state.
*   **Gotcha**: In v2.0, only text content was hashed. If you moved a box, the version didn't change.
*   **Fix**: In v3.1.0, `get_presentation_version` hashes `{left}:{top}:{width}:{height}` for every shape. Moving a shape by 1 pixel *will* change the version hash.

### 4.2 XML Manipulation (Opacity & Z-Order)
`python-pptx` does not support transparency or Z-order natively. We use `lxml` to hack the XML tree directly.
*   **Opacity**: We inject `<a:alpha val="50000"/>` into the color element. Note that Office uses a 0-100,000 scale (50000 = 50%). The Core converts 0.0-1.0 floats to this scale automatically.
*   **Z-Order**: We physically move the XML element within the `<p:spTree>`.
    *   **Warning**: Moving an element in the XML tree **changes its index** in the `shapes` collection.
    *   **Rule**: Always re-query `get_slide_info` after a Z-order change.

### 4.3 Footer Strategy (The "Master Trap")
*   **The Trap**: `agent.prs.slide_masters[0].placeholders` might contain a footer. This makes you think the footer works.
*   **The Reality**: Individual slides might have "Hide Background Graphics" on, or simply haven't instantiated that placeholder.
*   **The Fix**: `ppt_set_footer.py` uses a **Dual Strategy**:
    1.  Try setting the placeholder.
    2.  Check if any slides were *actually* updated (`slide_indices_updated`).
    3.  If 0 slides updated, fall back to creating a Text Box Overlay.

---

## 5. Troubleshooting Playbook

### Error: `jq: parse error: Invalid numeric literal`
*   **Meaning**: Your tool printed something that isn't JSON to stdout.
*   **Likely Culprits**:
    *   `print("Processing...")`
    *   A library warning (e.g., `DeprecationWarning`).
    *   `logging` configured to write to stdout (default behavior in some setups).
*   **Fix**: Apply the **Hygiene Block** (redirect stderr to devnull) and ensure you use `sys.stdout.write(json.dumps(...))` only once.

### Error: `TypeError: '<=' not supported between 'int' and 'dict'`
*   **Meaning**: You treated a Core v3.1.0 return value as an integer.
*   **Context**: `agent.add_slide()` used to return `slide_index` (int). Now it returns `{'slide_index': 1, 'version': ...}` (dict).
*   **Fix**: Change `idx = agent.add_slide(...)` to `res = agent.add_slide(...); idx = res['slide_index']`.

### Error: `SlideNotFoundError`
*   **Meaning**: You requested an index that doesn't exist.
*   **Context**: `python-pptx` is 0-indexed.
*   **The Ghost Slide**: If your script crashes mid-loop, you might have partially created slides.
*   **Fix**: Use `ppt_get_info.py` to verify the current slide count before assuming index N exists.

### Error: `ImportError: cannot import name 'PP_PLACEHOLDER'`
*   **Meaning**: You tried to import a constant from `core` that doesn't exist or isn't exported.
*   **Fix**: Import standard constants directly from `pptx.enum.shapes` or define fallback constants within your tool if strict dependency isolation is needed.

---

## 6. Tool Development Checklist

Before submitting a PR or running a new tool:

- [ ] **Imports**: Does it import `core.powerpoint_agent_core`?
- [ ] **Hygiene**: Is `sys.stderr` redirected to `/dev/null` at the very top?
- [ ] **Logging**: Is `logging` configured to `CRITICAL`?
- [ ] **Output**: is there exactly **one** `print` or `sys.stdout.write` call?
- [ ] **Error Handling**: Is there a `try/except Exception` block wrapping `main`?
- [ ] **Return Types**: Are you handling Core v3.1.0 dictionary returns correctly?
- [ ] **Versioning**: Are you passing `presentation_version_before/after` in the output?
- [ ] **Validation**: Did you run the tool manually with `--json` to verify the output format?
```

---

https://drive.google.com/file/d/1-lONMjzWFWo7D4WzPd9eaa91UWEDlYs2/view?usp=sharing, https://drive.google.com/file/d/1-rr9ukBHpWeP7vuR6aF2zwHC3X8SqD1_/view?usp=sharing, https://drive.google.com/file/d/102XGFegBYo08DCJX9lDGzIUtOLFtIp_T/view?usp=sharing, https://drive.google.com/file/d/12PNhT3WR6XXS3bD9dLY-2m-omtTvtV6K/view?usp=sharing, https://drive.google.com/file/d/12VOcUdRmkIzKse86Tg9rPOSJO6pCkNI7/view?usp=sharing, https://drive.google.com/file/d/16AOBSW-wOFGG4OGUQr_3JWn-jG4PPfdg/view?usp=sharing, https://drive.google.com/file/d/1Cx9KQl49KeLEfk4YartHZRVZ58Sw0D1d/view?usp=sharing, https://drive.google.com/file/d/1Ii2cMBBQz6uE4BzSAGsBAbCRPBFuHjka/view?usp=sharing, https://drive.google.com/file/d/1JshOaGOfH4B-8MnCmZXgg1a5qNQHoadM/view?usp=sharing, https://drive.google.com/file/d/1P0SD-9Kujkfv2aI5dJP0223StaZiw1Pk/view?usp=sharing, https://drive.google.com/file/d/1PvAd3BjO5zrCctMsdIUEEeeWUx5ZwEAd/view?usp=sharing, https://drive.google.com/file/d/1Q6eHXO5Z8o__ljMhqk06WhKa80g0zQbe/view?usp=sharing, https://drive.google.com/file/d/1QB9cpMPZJ9bSTd_eItL2wAdEFd7snsS1/view?usp=sharing, https://drive.google.com/file/d/1QnyQNIaJQQSKNA1HETuPm2CHibXZ5SX1/view?usp=sharing, https://drive.google.com/file/d/1Qpu7_xRKT4n3Nh-hv_0_3cRVRcLEezyz/view?usp=sharing, https://drive.google.com/file/d/1R78u0aY1P4CDJLQtjYQNVajFDbWmT72Z/view?usp=sharing, https://drive.google.com/file/d/1UaAS_xsYy1sAfk4O3VQKtFK3kHW3Snf-/view?usp=sharing, https://drive.google.com/file/d/1YQYcBWjBTyDFjV5uPzN7i99AuU6TYY6X/view?usp=sharing, https://drive.google.com/file/d/1Zu35u4UR26HXz9B4N72cGmQiOr3m2Qnx/view?usp=sharing, https://drive.google.com/file/d/1_HKEmJmqE7NBn2ChE0t4qAcs938HKIhD/view?usp=sharing, https://drive.google.com/file/d/1aJuyEThrTHgBCZDX2d7nETttMuEGDMaG/view?usp=sharing, https://aistudio.google.com/app/prompts?state=%7B%22ids%22:%5B%221cAEnnh5wHg4Prz3dkfl9fDWwLgyfwiFE%22%5D,%22action%22:%22open%22,%22userId%22:%22108686197475781557359%22,%22resourceKeys%22:%7B%7D%7D&usp=sharing, https://drive.google.com/file/d/1dy2e5Jg4V5SEuIz_Bj2qjlgyHt86MSei/view?usp=sharing, https://drive.google.com/file/d/1fEVzFAsU45yp18RqfZxDkNpRirZUtxZi/view?usp=sharing, https://drive.google.com/file/d/1j8mlvHVDYeeEyzXDcxEgoFS2bE10_L88/view?usp=sharing, https://drive.google.com/file/d/1jHco55_r7fW44hLc6K-ow7LgN_jazXyC/view?usp=sharing, https://drive.google.com/file/d/1loLT1KawLUdcF1WDwpxttpqaOeKmP6Qu/view?usp=sharing, https://drive.google.com/file/d/1m_hHKIjY2h_Pnk7LpvC8PlFCMdo9w4zc/view?usp=sharing, https://drive.google.com/file/d/1mj1C9BJ2H7soOF_rgrI82PU_kPB6izyd/view?usp=sharing, https://drive.google.com/file/d/1pEH1bNKBHejDcsct6qVsY-uqCDsoyIbL/view?usp=sharing, https://drive.google.com/file/d/1pJ_AAf1MO9sk7ksnPPF0d-IB_lAaZsRJ/view?usp=sharing, https://drive.google.com/file/d/1tyLEEbaVG1g5JROTxH4zWnDz-TkO1bfb/view?usp=sharing, https://drive.google.com/file/d/1u3D_EuBv_5944Y2OFeHlrQOquuj-APOz/view?usp=sharing, https://drive.google.com/file/d/1vkXILnVtm98xu_vwSHDPsNW0tLSNpkbV/view?usp=sharing, https://drive.google.com/file/d/1yUbIjDUSyB1pClSeZQ46slcBW548-VaF/view?usp=sharing

