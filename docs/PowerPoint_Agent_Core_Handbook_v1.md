# PowerPoint Agent Core: Programming Handbook (v3.1.1)

**Version:** 3.1.1  
**Library:** `core/powerpoint_agent_core.py`  
**License:** MIT  

---

## 1. Introduction

The `PowerPointAgent` core library is the foundational engine for the PowerPoint Agent Tools ecosystem. It provides a **stateless, atomic, and security-hardened** interface for manipulating `.pptx` files. Unlike the raw `python-pptx` library, this core handles file locking, complex positioning logic, accessibility compliance, and operation auditing (versioning).

### 1.1 Key Capabilities (v3.1.1)
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

## 3. Critical Protocols (Non-Negotiable)

### 3.1 Version Tracking Protocol
Every mutation method **must** capture the presentation state before and after execution to maintain the audit trail.

**Standard Pattern**:
```python
# 1. Capture version BEFORE changes
version_before = agent.get_presentation_version()

# 2. Perform operations
result = agent.some_operation()

# 3. Capture version AFTER changes
version_after = agent.get_presentation_version()

# 4. Return both versions in response
return {
    "status": "success",
    "presentation_version_before": version_before,
    "presentation_version_after": version_after,
    "result": result
}
```

### 3.2 Shape Index Freshness Protocol
Structural operations invalidate shape indices. Tools **must** refresh their knowledge of the slide state after any of these operations before performing further actions on shapes.

**Invalidating Operations**:
| Operation | Effect | Required Action |
|-----------|--------|-----------------|
| `add_shape()` | Adds index at end | Refresh if targeting new shape |
| `remove_shape()` | Shifts subsequent indices down | **CRITICAL**: Refresh immediately |
| `set_z_order()` | Reorders indices | **CRITICAL**: Refresh immediately |
| `delete_slide()` | Invalidates all indices | Reload slide info |
| `add_slide()` | New slide context | Query new slide info |

**Correct Pattern**:
```python
# 1. Perform structural change
agent.remove_shape(slide_index=0, shape_index=5)

# 2. REFRESH INDICES (Do not assume index 6 is now 5)
slide_info = agent.get_slide_info(slide_index=0)

# 3. Find target shape by name/properties in refreshed list
target_shape = next(s for s in slide_info["shapes"] if s["name"] == "TargetBox")
agent.format_shape(slide_index=0, shape_index=target_shape["index"], ...)
```

---

## 4. Data Structures & Input Schemas

The Core uses flexible dictionary-based inputs to abstract layout complexity.

### 4.1 Positioning (`Position.from_dict`)
Used in `add_shape`, `add_text_box`, `insert_image`, etc.

| Type | Schema | Description |
|------|--------|-------------|
| **Percentage** | `{"left": "10%", "top": "20%"}` | Relative to slide dimensions. **Preferred.** |
| **Absolute** | `{"left": 1.5, "top": 2.0}` | Inches from top-left. |
| **Anchor** | `{"anchor": "center", "offset_x": 0, "offset_y": 0}` | Relative to standard points (e.g., `bottom_right`). |
| **Grid** | `{"grid_row": 2, "grid_col": 3, "grid_size": 12}` | 12-column grid system layout. |

**Anchor Points**: `top_left`, `top_center`, `top_right`, `center_left`, `center`, `center_right`, `bottom_left`, `bottom_center`, `bottom_right`.

### 4.2 Sizing (`Size.from_dict`)
| Type | Schema | Description |
|------|--------|-------------|
| **Percentage** | `{"width": "50%", "height": "50%"}` | % of slide width/height. |
| **Absolute** | `{"width": 5.0, "height": 3.0}` | Inches. |
| **Auto** | `{"width": "50%", "height": "auto"}` | Preserves aspect ratio. |

---

## 5. API Reference

### 5.1 File Operations

#### `open(filepath, acquire_lock=True)`
*   **Purpose**: Loads presentation and optionally acquires file lock.
*   **Safety**: Uses `errno.EEXIST` (via `os.open` with `O_CREAT|O_EXCL`) for atomic locking across POSIX and Windows.
*   **Throws**: `PathValidationError`, `FileLockError`, `PowerPointAgentError`.

#### `save(filepath=None)`
*   **Purpose**: Saves changes. If `filepath` is None, overwrites source.
*   **Safety**: Ensures parent directories exist.

#### `clone_presentation(output_path)`
*   **Purpose**: Creates a working copy.
*   **Returns**: A *new* `PowerPointAgent` instance pointed at the cloned file.

### 5.2 Slide Operations

#### `add_slide(layout_name, index=None)`
*   **Returns**: `Dict` (v3.1+) containing `slide_index`, `layout_name`, `total_slides`, `presentation_version_before/after`.
*   **Validation**: Raises `SlideNotFoundError` if `index` is out of bounds (removed silent clamping).

#### `delete_slide(index, approval_token=None)`
*   **Security**: **Requires** valid `approval_token` matching scope `delete:slide`.
*   **Throws**: `ApprovalTokenError` if token is invalid/missing.

#### `duplicate_slide(index)` / `reorder_slides(from_index, to_index)`
*   **Behavior**: Performs deep copy of shapes including text runs and styles.

### 5.3 Shape & Visual Operations

#### `add_shape(slide_index, shape_type, position, size, ...)`
*   **Arguments**:
    *   `fill_opacity` (float 0.0-1.0): **New in v3.1.0**. (1.0 = Opaque).
    *   `line_opacity` (float 0.0-1.0).
    *   `shape_type`: String key (e.g., `"rectangle"`, `"arrow_right"`).
*   **Returns**: Dictionary containing `shape_index` and applied styling.
*   **Internal**: Uses `_set_fill_opacity` to inject OOXML `<a:alpha>` tags.

#### `format_shape(slide_index, shape_index, ...)`
*   **Arguments**: `fill_color`, `fill_opacity`, `line_color`, etc.
*   **Deprecation**: `transparency` param is deprecated; explicit conversion to `1.0 - fill_opacity` occurs, logging a warning.

#### `remove_shape(slide_index, shape_index, approval_token=None)`
*   **Security**: **Requires** valid `approval_token` matching scope `remove:shape`.
*   **Warning**: Removing a shape shifts the indices of all subsequent shapes on that slide.

#### `set_z_order(slide_index, shape_index, action)`
*   **Actions**: `"bring_to_front"`, `"send_to_back"`, `"bring_forward"`, `"send_backward"`.
*   **Internal**: Physically moves the XML element in `<p:spTree>`.
*   **Critical Side Effect**: **Invalidates Shape Indices**. Tools must warn users to re-query `get_slide_info`.

### 5.4 Text & Content

#### `add_text_box` / `add_bullet_list`
*   **Features**: Auto-fit text, specific font styling, alignment mapping.
*   **Returns**: `shape_index` of the created text container.

#### `replace_text(find, replace, match_case=False)`
*   **Scope**: Global (entire presentation) or scoped (slide/shape).
*   **Intelligence**: Tries to preserve formatting by replacing inside text runs first.

#### `set_footer(text, show_number, show_date)`
*   **Mechanism**: Iterates through *all* slides to find placeholders (type 7, 6, 5).
*   **Returns**: `slides_processed` count. (Note: Does not create text boxes; relies on Tool layer for fallback).

### 5.5 Charts & Images

#### `add_chart(chart_type, data, ...)`
*   **Supported Types**: Column, Bar, Line, Pie, Area, Scatter, Doughnut.
*   **Data Format**: `{"categories": ["A", "B"], "series": [{"name": "S1", "values": [1, 2]}]}`.

#### `update_chart_data(slide_index, chart_index, data)`
*   **Strategy**:
    1.  Try `chart.replace_data()` (Best, preserves format).
    2.  Catch `AttributeError` (Older pptx versions).
    3.  Fallback: Recreate chart in-place (Preserves position/size/title, resets some style).

#### `insert_image` / `replace_image`
*   **Features**: Auto-aspect ratio calculation, optional compression (if Pillow installed), Alt Text setting.

---

## 6. Security & Governance

### 6.1 Path Validation (`PathValidator`)
*   **Traversal Protection**: If `allowed_base_dirs` is set, checks `path.is_relative_to(base)`.
*   **Extension Check**: Enforces `.pptx`, `.pptm`, `.potx`.

### 6.2 Approval Tokens
Destructive methods call `_validate_token(token, scope)`.
*   **Scope Constants**:
    *   `APPROVAL_SCOPE_DELETE_SLIDE` ("delete:slide")
    *   `APPROVAL_SCOPE_REMOVE_SHAPE` ("remove:shape")
*   **Failure**: Raises `ApprovalTokenError` (Exit Code 4 in tools).

---

## 7. Observability & Versioning

### 7.1 Presentation Versioning (`get_presentation_version`)
Returns a SHA-256 hash (prefix 16 chars) representing the state.

**Input for Hash**:
1.  Slide Count.
2.  Layout Names per slide.
3.  **Shape Geometry**: `{left}:{top}:{width}:{height}` (Detects moves/resizes).
4.  **Text Content**: SHA-256 of text runs.

### 7.2 Error Handling Matrix

| Code | Category | Meaning | Response Format |
|---|---|---|---|
| 0 | Success | Completed | `{"status": "success", ...}` |
| 1 | Usage | Invalid args | `{"status": "error", "error_type": "ValueError", ...}` |
| 2 | Validation | Schema invalid | `{"status": "error", "error_type": "ValidationError", ...}` |
| 3 | Transient | Lock/Network | `{"status": "error", "error_type": "FileLockError", ...}` |
| 4 | Permission | Token missing | `{"status": "error", "error_type": "ApprovalTokenError", ...}` |
| 5 | Internal | Crash | `{"status": "error", "error_type": "PowerPointAgentError", ...}` |

---

## 8. Performance Characteristics

Understanding the cost of operations is vital for building efficient agents.

| Operation | Complexity | 10-Slide Deck | 50-Slide Deck | Notes |
|-----------|------------|---------------|---------------|-------|
| `get_presentation_version()` | O(N) Shapes | ~15ms | ~75ms | Scales linearly with total shape count. Called twice per mutation. |
| `capability_probe(deep=True)` | O(M) Layouts | ~120ms | ~600ms+ | Creates/destroys slides. Has 15s timeout. |
| `add_shape()` | O(1) | ~8ms | ~8ms | Constant time (XML injection). |
| `replace_text(global)` | O(N) TextRuns | ~25ms | ~125ms | Regex matching across all text runs. |
| `save()` | I/O Bound | ~50ms | ~200ms+ | Dominated by disk write speed and file size (images). |

**Optimization Guidelines**:
*   **Batching**: Not supported natively (stateless tools), but context managers in custom scripts can batch mutations before a single `save()`.
*   **Shallow Probes**: Use `deep=False` in `capability_probe` unless layout geometry is strictly required.
*   **Limits**: Avoid decks >100 slides or >50MB for interactive agent sessions to prevent timeouts.

---

## 9. Internal "Magic" (Troubleshooting)

### 9.1 Opacity Injection
`python-pptx` lacks transparency support. We use `lxml` to inject:
```xml
<a:solidFill>
  <a:srgbClr val="FF0000">
    <a:alpha val="50000"/> <!-- 50% Opacity -->
  </a:srgbClr>
</a:solidFill>
```
**Note**: Office uses 0-100,000 scale. Core converts 0.0-1.0 floats automatically.

### 9.2 Z-Order Manipulation
We physically move the `<p:sp>` element within the `<p:spTree>` XML list.
*   `bring_to_front`: Move to end of list.
*   `send_to_back`: Move to index 2 (after background/master refs).

---

## 10. Backward Compatibility (Migration)

### v3.0 → v3.1 Migration Guide
*   **Return Values**: Methods like `add_slide` and `add_shape` now return **Dictionaries**, not Integers.
    *   *Old*: `idx = agent.add_slide(...)`
    *   *New*: `res = agent.add_slide(...); idx = res["slide_index"]`
*   **Transparency**: Use `fill_opacity` instead of `transparency`.
*   **Safety**: `add_slide(index=999)` now raises `SlideNotFoundError` instead of clamping to end. Catch this exception if loose behavior is needed.
