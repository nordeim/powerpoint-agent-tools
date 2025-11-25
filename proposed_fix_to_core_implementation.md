# Implementation Plan - Fix Core Format Shape

## Goal Description
Implement a complete, working `format_shape()` method in `core/powerpoint_agent_core.py` that supports opacity/transparency, fixes critical bugs (NameError, unused parameters), and adds comprehensive testing.

## User Review Required
> [!IMPORTANT]
> This fix addresses a critical `NameError` and unused parameters in the previous implementation. It also deprecates the `transparency` parameter in favor of `fill_opacity`.

## Proposed Changes

### Core Library
#### [MODIFY] [powerpoint_agent_core.py](file:///home/project/powerpoint-agent-tools/core/powerpoint_agent_core.py)
- **Method**: `format_shape`
- **Changes**:
    - Add `fill_opacity` and `line_opacity` parameters.
    - Deprecate `transparency` parameter (convert to `fill_opacity`).
    - Implement logic to apply fill and line opacity using helper methods.
    - Return detailed change status.
    - Update docstring.

### Tests
#### [NEW] [test_shape_opacity.py](file:///home/project/powerpoint-agent-tools/tests/test_shape_opacity.py)
- **Purpose**: Comprehensive test suite for shape opacity features.
- **Tests**:
    - `test_add_shape_default_opacity`
    - `test_add_shape_fill_opacity_overlay`
    - `test_add_shape_fill_opacity_half`
    - `test_add_shape_fill_opacity_zero`
    - `test_add_shape_line_opacity`
    - `test_add_shape_both_opacities`
    - `test_add_shape_opacity_validation_*`
    - `test_format_shape_fill_opacity`
    - `test_format_shape_line_opacity`
    - `test_format_shape_transparency_deprecated`
    - `test_format_shape_fill_opacity_overrides_transparency`
    - `test_format_shape_color_and_opacity`

## Verification Plan

### Automated Tests
- Run the new test suite:
  ```bash
  uv run pytest tests/test_shape_opacity.py -v
  ```
- Verify all tests pass.

---

# Task: Implement Core Fix

- [x] Analyze `proposed_fix_to_core.md` and `core/powerpoint_agent_core.py` <!-- id: 0 -->
- [x] Create Implementation Plan <!-- id: 1 -->
- [x] Apply fix to `core/powerpoint_agent_core.py` <!-- id: 2 -->
- [x] Verify the fix <!-- id: 3 -->

