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
