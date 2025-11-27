# Implementation Plan: PowerPoint Agent Core Handbook (v3.1.0)

## 1. Objective
Create a definitive "Programming Model & API Handbook" for `core/powerpoint_agent_core.py` (v3.1.0). This document will serve as the technical manual for developers (AI or human) building tools on top of this library. It will distill the actual code logic, method signatures, return structures, and internal behaviors into a usable reference.

## 2. Source of Truth
The content is derived strictly from the `powerpoint_agent_core.py` v3.1.0 file generated in the previous steps, which includes:
*   Security hardening (Path traversal, `errno` locking).
*   Governance (Approval tokens).
*   Observability (Geo-aware version hashing).
*   Logic restoration (`_copy_shape`, etc.).

## 3. Document Structure

### Part 1: Core Philosophy & Architecture
*   **Scope**: What the core does vs. what tools do.
*   **State Management**: The Context Manager pattern (`with PowerPointAgent...`).
*   **Safety Mechanisms**: File locking, atomic writes, path validation.

### Part 2: Data Structures & Schemas
*   **Input Schemas**: Detailed JSON structures for `Position`, `Size`, `Color`, and `Crop`.
*   **Output Schemas**: The standard dictionary structure returned by mutation methods (including versioning fields).

### Part 3: The `PowerPointAgent` API Reference
A detailed breakdown of methods grouped by domain, detailing **Arguments**, **Return Values**, **Side Effects**, and **Specific Exceptions**.
*   File Operations (`open`, `save`, `clone`)
*   Slide Operations (`add`, `delete`, `reorder`)
*   Content Operations (`text`, `lists`, `titles`)
*   Visuals (`shapes`, `images`, `z-order`, `opacity`)
*   Data Viz (`charts`, `tables`)
*   Validation & Info (`accessibility`, `versioning`)

### Part 4: Security & Governance Model
*   **Approval Tokens**: How `_validate_token` works and which methods enforce it.
*   **Path Validation**: Directory restrictions and traversal protection.

### Part 5: Internal "Magic" (For Troubleshooting)
*   **XML Hacking**: How Opacity and Z-Order are implemented via `lxml` (since `python-pptx` lacks them).
*   **Version Hashing**: The exact logic of the SHA-256 calculation (Text + Geometry).
*   **Error Hierarchy**: The class inheritance tree of exceptions.

---

## 4. Execution

I will now generate the `PowerPoint_Agent_Core_Handbook.md`.

```markdown
# PowerPoint Agent Core: Programming Handbook (v3.1.0)

**Version:** 3.1.0  
**Library:** `core/powerpoint_agent_core.py`  
**License:** MIT  

---

## 1. Introduction

The `PowerPointAgent` core library is the foundational engine for the PowerPoint Agent Tools ecosystem. It provides a **stateless, atomic, and security-hardened** interface for manipulating `.pptx` files. Unlike the raw `python-pptx` library, this core handles file locking, complex positioning logic, accessibility compliance, and operation auditing (versioning).

### 1.1 Key Capabilities (v3.1.0)
*   **Context-Safe**: Handles file opening/closing/locking automatically.
*   **Observability**: Tracks presentation state via deterministic SHA-256 hashing (Geometry + Content).
*   **Governance**: Enforces "Approval Tokens" for destructive actions (`delete_slide`, `remove_shape`).
*   **Visual Fidelity**: Implements XML hacks for features missing in `python-pptx` (Opacity, Z-Order).
*   **Accessibility**: Built-in WCAG 2.1 AA checking and Color Contrast calculation.

---

## 2. Usage Pattern (The "Hub" Model)

Tools interacting with this core **must** use the Context Manager pattern to ensure file safety.

```python
from core.powerpoint_agent_core import PowerPointAgent

# Atomic Operation Pattern
with PowerPointAgent(filepath) as agent:
    # 1. Acquire Lock & Load
    agent.open(filepath)
    
    # 2. Mutate (Capture return dict)
    result = agent.add_shape(...)
    
    # 3. Save (Atomic Write)
    agent.save()
    
    # 4. Release Lock (Automatic on exit)
```

---

## 3. Data Structures & Input Schemas

The Core uses flexible dictionary-based inputs to abstract layout complexity.

### 3.1 Positioning (`Position.from_dict`)
Used in `add_shape`, `add_text_box`, `insert_image`, etc.

| Type | Schema | Description |
|------|--------|-------------|
| **Percentage** | `{"left": "10%", "top": "20%"}` | Relative to slide dimensions. **Preferred.** |
| **Absolute** | `{"left": 1.5, "top": 2.0}` | Inches from top-left. |
| **Anchor** | `{"anchor": "center", "offset_x": 0, "offset_y": 0}` | Relative to standard points (e.g., `bottom_right`). |
| **Grid** | `{"grid_row": 2, "grid_col": 3, "grid_size": 12}` | 12-column grid system layout. |

**Anchor Points**: `top_left`, `top_center`, `top_right`, `center_left`, `center`, `center_right`, `bottom_left`, `bottom_center`, `bottom_right`.

### 3.2 Sizing (`Size.from_dict`)
| Type | Schema | Description |
|------|--------|-------------|
| **Percentage** | `{"width": "50%", "height": "50%"}` | % of slide width/height. |
| **Absolute** | `{"width": 5.0, "height": 3.0}` | Inches. |
| **Auto** | `{"width": "50%", "height": "auto"}` | Preserves aspect ratio. |

### 3.3 Colors
*   **Format**: Hex strings (`"#FF0000"` or `"0070C0"`).
*   **Validation**: Validated by `ColorHelper`. Invalid hex raises `ValueError`.

---

## 4. API Reference

### 4.1 File Operations

#### `open(filepath, acquire_lock=True)`
*   **Purpose**: Loads presentation and optionally acquires cross-platform file lock.
*   **Safety**: Uses `errno.EEXIST` (POSIX) / `FileExistsError` (Windows) for atomic locking.
*   **Throws**: `PathValidationError`, `FileLockError`, `PowerPointAgentError`.

#### `save(filepath=None)`
*   **Purpose**: Saves changes. If `filepath` is None, overwrites source.
*   **Safety**: Ensures parent directories exist.

#### `clone_presentation(output_path)`
*   **Purpose**: Creates a working copy.
*   **Returns**: A *new* `PowerPointAgent` instance pointed at the cloned file.

### 4.2 Slide Operations

#### `add_slide(layout_name, index=None)`
*   **Returns**:
    ```json
    {
      "slide_index": 5,
      "layout_name": "Title and Content",
      "total_slides": 6,
      "presentation_version_before": "...",
      "presentation_version_after": "..."
    }
    ```
*   **Validation**: Raises `SlideNotFoundError` if `index` is out of bounds (removed silent clamping).

#### `delete_slide(index, approval_token=None)`
*   **Security**: **Requires** valid `approval_token` matching scope `delete:slide`.
*   **Throws**: `ApprovalTokenError` if token is invalid/missing.

#### `duplicate_slide(index)` / `reorder_slides(from_index, to_index)`
*   **Behavior**: Performs deep copy of shapes including text runs and styles.

### 4.3 Shape & Visual Operations

#### `add_shape(slide_index, shape_type, position, size, ...)`
*   **Arguments**:
    *   `fill_opacity` (float 0.0-1.0): **New in v3.1.0**. (1.0 = Opaque).
    *   `line_opacity` (float 0.0-1.0).
    *   `shape_type`: String key (e.g., `"rectangle"`, `"arrow_right"`).
*   **Returns**: Dictionary containing `shape_index` and applied styling.
*   **Internal**: Uses `_set_fill_opacity` to inject OOXML `<a:alpha>` tags.

#### `format_shape(slide_index, shape_index, ...)`
*   **Arguments**: `fill_color`, `fill_opacity`, `line_color`, etc.
*   **Deprecation**: `transparency` param is deprecated; auto-converts to `1.0 - fill_opacity` with a warning.

#### `remove_shape(slide_index, shape_index, approval_token=None)`
*   **Security**: **Requires** valid `approval_token` matching scope `remove:shape`.
*   **Warning**: Removing a shape shifts the indices of all subsequent shapes on that slide.

#### `set_z_order(slide_index, shape_index, action)`
*   **Actions**: `"bring_to_front"`, `"send_to_back"`, `"bring_forward"`, `"send_backward"`.
*   **Internal**: Physically moves the XML element in `<p:spTree>`.
*   **Critical Side Effect**: **Invalidates Shape Indices**. Tools must warn users to re-query `get_slide_info`.

### 4.4 Text & Content

#### `add_text_box` / `add_bullet_list`
*   **Features**: Auto-fit text, specific font styling, alignment mapping.
*   **Returns**: `shape_index` of the created text container.

#### `replace_text(find, replace, match_case=False)`
*   **Scope**: Global (entire presentation) or scoped (slide/shape).
*   **Intelligence**: Tries to preserve formatting by replacing inside text runs first.

#### `set_footer(text, show_number, show_date)`
*   **Mechanism**: Iterates through *all* slides to find placeholders (type 7, 6, 5).
*   **Returns**: `slides_processed` count. (Note: Does not create text boxes; relies on Tool layer for fallback).

### 4.5 Charts & Images

#### `add_chart(chart_type, data, ...)`
*   **Supported Types**: Column, Bar, Line, Pie, Area, Scatter, Doughnut.
*   **Data Format**: `{"categories": ["A", "B"], "series": [{"name": "S1", "values": [1, 2]}]}`.

#### `update_chart_data(slide_index, chart_index, data)`
*   **Robustness**: Tries `chart.replace_data()`. If that fails (version compat), it **recreates** the chart in-place to ensure data is applied.

#### `insert_image` / `replace_image`
*   **Features**: Auto-aspect ratio calculation, optional compression (if Pillow installed), Alt Text setting.

---

## 5. Security & Governance

### 5.1 Path Validation (`PathValidator`)
*   **Traversal Protection**: If `allowed_base_dirs` is set, checks `path.is_relative_to(base)`.
*   **Extension Check**: Enforces `.pptx`, `.pptm`, `.potx`.

### 5.2 Approval Tokens
Destructive methods call `_validate_token(token, scope)`.
*   **Scope Constants**:
    *   `APPROVAL_SCOPE_DELETE_SLIDE` ("delete:slide")
    *   `APPROVAL_SCOPE_REMOVE_SHAPE` ("remove:shape")
*   **Failure**: Raises `ApprovalTokenError` (Exit Code 4 in tools).

---

## 6. Observability & Versioning

### 6.1 Presentation Versioning (`get_presentation_version`)
Returns a SHA-256 hash (prefix 16 chars) representing the state.

**Input for Hash**:
1.  Slide Count.
2.  Layout Names per slide.
3.  **Shape Geometry**: `{left}:{top}:{width}:{height}` (Detects moves/resizes).
4.  **Text Content**: SHA-256 of text runs.

**Usage**:
*   Tools capture `version_before` immediately after opening.
*   Tools capture `version_after` immediately before returning.
*   This allows the Orchestrator to detect "Phantom Edits" (race conditions).

### 6.2 Logging
*   **Logger**: `core.powerpoint_agent_core`
*   **Stream**: `sys.stderr` (Strictly separated from stdout to protect JSON pipelines).

---

## 7. Internal Troubleshooting

### 7.1 "Ghost Slides"
*   **Symptom**: Validation reports issues on `slide_index: 9` when you only added 9 slides (0-8).
*   **Cause**: Calling `add_slide` inside a loop that executes one time too many, or a tool crashing *after* creation but *before* configuration.
*   **Detection**: `ppt_validate_presentation.py` will flag "Empty slide" or "Slide missing title".

### 7.2 Opacity Not Working
*   **Cause**: `python-pptx` doesn't support alpha channels.
*   **Fix**: The Core uses `_set_fill_opacity` to inject XML.
*   **Requirement**: The shape must have a **Solid Fill** first. `add_shape` handles this automatically.

### 7.3 "Invalid Numeric Literal" in JSON
*   **Cause**: `stdout` pollution. A library printed text mixed with the JSON.
*   **Fix in Tools**:
    ```python
    import sys, os
    sys.stderr = open(os.devnull, 'w')  # The "Nuclear Option"
    ```
    (Note: The Core itself logs to stderr, so this must be done in the *Tool* wrapper).

---

## 8. Exception Hierarchy

All exceptions inherit from `PowerPointAgentError` and provide `to_dict()`.

*   `PowerPointAgentError`
    *   `SlideNotFoundError`
    *   `ShapeNotFoundError`
    *   `ChartNotFoundError`
    *   `LayoutNotFoundError`
    *   `ImageNotFoundError`
    *   `InvalidPositionError`
    *   `TemplateError`
    *   `ThemeError`
    *   `AccessibilityError`
    *   `AssetValidationError`
    *   `FileLockError`
    *   `PathValidationError`
    *   `ApprovalTokenError` (New in v3.1.0)
```

