# Implementation Plan: Standardize Validation in Shape and Table Tools

## Goal Description
Enhance `tools/ppt_add_shape.py` and `tools/ppt_add_table.py` with robust validation logic similar to `tools/ppt_add_text_box.py`. This includes checks for off-slide positioning, minimum size, and color contrast (for shapes), ensuring higher quality and accessibility of generated presentations.

## User Review Required
- **None**: This is a direct request to standardize validation.

## Proposed Changes

### Tools
#### [MODIFY] [tools/ppt_add_shape.py](file:///home/project/powerpoint-agent-tools/tools/ppt_add_shape.py)
-   **Add Function**: `validate_shape_params(position, size, fill_color)`
    -   **Position**: Warn if off-slide (0-100%) unless `--allow-offslide` (need to add this flag).
    -   **Size**: Warn if too small (< 1% width/height).
    -   **Contrast**: Check fill color against white background (if provided).
-   **Update `main`**: Call validation, include warnings/recommendations in JSON output.
-   **Add Argument**: `--allow-offslide` to suppress position warnings.

#### [MODIFY] [tools/ppt_add_table.py](file:///home/project/powerpoint-agent-tools/tools/ppt_add_table.py)
-   **Add Function**: `validate_table_params(rows, cols, position, size)`
    -   **Position**: Warn if off-slide.
    -   **Size**: Warn if too small for the number of rows/cols (e.g., < 5% height for 10 rows).
-   **Update `main`**: Call validation, include warnings/recommendations in JSON output.
-   **Add Argument**: `--allow-offslide` to suppress position warnings.

## Verification Plan

### Automated Tests
1.  **Shape Validation Test (`tests/verify_shape_validation.py`)**:
    -   Test adding shape with valid params (no warnings).
    -   Test adding shape off-slide (expect warning).
    -   Test adding tiny shape (expect warning).
    -   Test adding low-contrast shape (expect warning).

2.  **Table Validation Test (`tests/verify_table_validation.py`)**:
    -   Test adding table with valid params.
    -   Test adding table off-slide.
    -   Test adding tiny table.

---

# Task: Standardize Validation in Shape and Table Tools

- [x] **Phase 1: Analysis**
    - [x] Analyze `ppt_add_text_box.py` validation logic <!-- id: 0 -->
    - [x] Analyze `ppt_add_shape.py` current state <!-- id: 1 -->
    - [x] Analyze `ppt_add_table.py` current state <!-- id: 2 -->
    - [x] Check `core/powerpoint_agent_core.py` for helpers <!-- id: 3 -->
    - [x] Design baseline and new validation tests <!-- id: 4 -->
- [x] **Phase 2: Planning**
    - [x] Create `implementation_plan.md` <!-- id: 5 -->
- [x] **Phase 3: Execution**
    - [x] Create baseline tests <!-- id: 6 -->
    - [x] Update `ppt_add_shape.py` <!-- id: 7 -->
    - [x] Update `ppt_add_table.py` <!-- id: 8 -->
- [x] **Phase 4: Verification**
    - [x] Run validation tests <!-- id: 9 -->
        - [x] Fix syntax errors in tools
        - [x] Verify shape contrast check (Found bug in Core)
        - [x] Check ColorHelper usage in tools
        - [x] Fix ColorHelper in Core

