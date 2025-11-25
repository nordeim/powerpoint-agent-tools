# Implementation Plan - Core File Improvements (Opacity Support)

## Goal Description
Enhance `core/powerpoint_agent_core.py` to support fill and line opacity for shapes. This involves direct OOXML manipulation as `python-pptx` does not natively support transparency.

## User Review Required
> [!IMPORTANT]
> **OOXML Manipulation**: This update introduces direct XML manipulation using `lxml`. This is necessary for opacity but increases complexity.
> **Deprecation**: The `transparency` argument in `format_shape` is deprecated in favor of `fill_opacity`.

## Proposed Changes

### [MODIFY] [powerpoint_agent_core.py](file:///home/project/powerpoint-agent-tools/core/powerpoint_agent_core.py)

#### 1. Imports
- [ ] Add `from lxml import etree`
- [ ] Add `from pptx.oxml.ns import qn`

#### 2. Helper Methods (New)
- [ ] Add `_set_fill_opacity`
- [ ] Add `_set_line_opacity`
- [ ] Add `_ensure_line_solid_fill`
- [ ] Add `_log_warning`

#### 3. Method Updates
- [ ] Update `add_shape`:
    - Add `fill_opacity` (default 1.0) and `line_opacity` (default 1.0) arguments.
    - Implement XML manipulation logic via helpers.
- [ ] Update `format_shape`:
    - Add `fill_opacity` and `line_opacity` arguments.
    - Handle deprecated `transparency` argument (convert to `fill_opacity`).
    - Implement XML manipulation logic via helpers.

## Verification Plan

### Automated Tests
1.  **Opacity Verification**:
    - Create `tests/test_core_opacity.py`.
    - Create a presentation, add shapes with various opacity levels.
    - Inspect the underlying XML to verify `<a:alpha>` tags are present and correct.
    - Verify `format_shape` correctly updates existing opacity.

### Manual Verification
- Run the test script and check the generated `test_opacity.pptx` manually if possible (or rely on XML inspection in the test).

