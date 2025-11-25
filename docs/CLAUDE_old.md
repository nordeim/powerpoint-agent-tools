# CLAUDE.md - PowerPoint Agent Tools Briefing

## 🌍 Context & Purpose
**PowerPoint Agent Tools** is a suite of 30+ stateless CLI utilities designed for AI agents to manipulate `.pptx` files. 
- **Core Library:** `python-pptx` wrapped in a semantic API (`core/powerpoint_agent_core.py`).
- **Interface:** Command Line Interface (CLI) with JSON I/O.
- **Philosophy:** Atomic, Stateless, Composable, Visual-Aware.

## 🛠️ Build & Runtime
- **Package Manager:** `uv` (recommended) or `pip`.
- **Python Version:** 3.8+.
- **Dependencies:** `python-pptx`, `Pillow`, `pandas`, `LibreOffice` (for export).

### Common Commands
```bash
# Install Dependencies
uv pip install -r requirements.txt

# Run a Tool (Example)
uv python tools/ppt_get_info.py --file deck.pptx --json

# Run Tests
pytest tests/

# Lint/Format (Standard)
ruff check .
black .
```

## 🏗️ Architecture
The system uses a **Hub-and-Spoke** architecture:

1.  **The Core (`core/`)**:
    *   `powerpoint_agent_core.py`: The Monolith. Handles all XML manipulation, file locking, and validation.
    *   **Pattern:** Context Manager (`with PowerPointAgent(path) as agent:`).
    *   **Key Classes:** `PowerPointAgent`, `Position`, `Size`, `AssetValidator`.

2.  **The Tools (`tools/`)**:
    *   Thin wrappers around Core methods.
    *   **Responsibility:** Argument parsing (`argparse`), Input validation, JSON formatting.
    *   **Naming:** `ppt_<verb>_<noun>.py` (e.g., `ppt_add_slide.py`).

## 📝 Coding Standards

### 1. New Tool Checklist
When creating a new tool in `tools/`:
- [ ] **Imports:** `sys.path.insert(0, ...)` to allow core import without package install.
- [ ] **Output:** MUST print a single JSON object to `stdout`.
- [ ] **Errors:** Catch all exceptions, print JSON `{ "status": "error", "error": "..." }`, exit code `1`.
- [ ] **Success:** Print JSON `{ "status": "success", ... }`, exit code `0`.
- [ ] **Arguments:** Use `argparse`. Complex inputs (Position/Size/Data) must be JSON strings.

### 2. Style
- **Type Hints:** Mandatory for all functions.
- **Docstrings:** Required for modules and functions.
- **Path Handling:** Use `pathlib.Path`. Always validate `.exists()` for inputs.

## 🧠 "Gotchas" & Critical Patterns

### 1. The Inspection Pattern
**Problem:** AI guesses Shape IDs (e.g., "Shape 1") and fails.
**Solution:**
1.  **NEVER** guess an index.
2.  **ALWAYS** run `ppt_get_slide_info.py` first to map the slide.
3.  Use the returned `shape_index` to target edits.

### 2. Positioning System
Do not use raw numbers unless specified. Use the **Position Dictionary**:
*   **Responsive (Preferred):** `{'left': '10%', 'top': '20%'}`
*   **Anchor:** `{'anchor': 'bottom_right', 'offset_x': -1.0}`
*   **Grid:** `{'grid': 'C4'}`

### 3. Statelessness
The agent does not hold the file open. Every tool call opens the file, modifies it, saves it, and closes it. Race conditions are handled by `core.FileLock`, but avoid parallel writes to the same file if possible.

### 4. Chart Limitations
`python-pptx` has limited support for *updating* existing charts.
*   **Preferred:** Delete the old chart and `ppt_add_chart.py` a new one.
*   **Update:** Only use `ppt_update_chart_data.py` if strictly necessary and schematic matches.

## 📂 Directory Structure
```text
.
├── core/
│   ├── powerpoint_agent_core.py  # logic
│   └── __init__.py               # exports
├── tools/                        # 30+ CLI scripts
├── tests/                        # pytest suite
├── assets/                       # test images/data
└── AGENT_SYSTEM_PROMPT.md        # User-facing prompt
```
