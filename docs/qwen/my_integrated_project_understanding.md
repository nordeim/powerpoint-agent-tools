# Core Library Integration Complete ✓

I have meticulously reviewed the **PowerPoint Agent Core Programming Handbook (v3.1.4)** and integrated its deep technical details with my existing understanding of the project architecture.

---

## 📊 Integration Summary

| Knowledge Layer | Previous Understanding | New Details Integrated |
|-----------------|----------------------|------------------------|
| **Architecture** | Hub-and-Spoke model | File locking mechanism, context manager internals |
| **Version Tracking** | SHA-256 hash concept | Exact hash inputs (geometry + content), 16-char prefix |
| **Opacity/Z-Order** | XML manipulation required | Exact XML injection patterns, OOXML scale (0-100000) |
| **Performance** | General awareness | Concrete timing benchmarks per operation |
| **Error Handling** | Exit code matrix | Specific exception types and response formats |
| **Security** | Approval token concept | HMAC-SHA256 generation, path validation details |

---

## 🔍 Key Technical Details Now Internalized

### 1. File Locking Mechanism

```
┌─────────────────────────────────────────────────────────────────────────────┐
│  FILE LOCKING IMPLEMENTATION                                                 │
├─────────────────────────────────────────────────────────────────────────────┤
│                                                                              │
│  Method: os.open() with O_CREAT|O_EXCL flags                                │
│  Detection: errno.EEXIST for cross-platform compatibility                   │
│  Timeout: 10 seconds (hardcoded)                                            │
│  Lock File: {original_file}.lock                                            │
│                                                                              │
│  Pattern:                                                                    │
│  ┌─────────────────────────────────────────────────────────────────────┐    │
│  │  try:                                                                │    │
│  │      with PowerPointAgent(filepath) as agent:                        │    │
│  │          agent.open(filepath, acquire_lock=True)  # 10s timeout     │    │
│  │          # ... operations ...                                        │    │
│  │          agent.save()                                                │    │
│  │      # Lock released automatically on context exit                   │    │
│  │  except FileLockError:                                               │    │
│  │      # Handle contention - retry with backoff                        │    │
│  └─────────────────────────────────────────────────────────────────────┘    │
│                                                                              │
└─────────────────────────────────────────────────────────────────────────────┘
```

### 2. Version Hash Computation (Detailed)

```
┌─────────────────────────────────────────────────────────────────────────────┐
│  PRESENTATION VERSION HASH                                                   │
├─────────────────────────────────────────────────────────────────────────────┤
│                                                                              │
│  Algorithm: SHA-256 (first 16 characters)                                   │
│  Complexity: O(N) where N = total shape count                               │
│                                                                              │
│  INPUT COMPONENTS:                                                           │
│  ┌─────────────────────────────────────────────────────────────────────┐    │
│  │  1. Slide Count                                                      │    │
│  │  2. Layout Names per slide                                           │    │
│  │  3. Shape Geometry: "{left}:{top}:{width}:{height}" for each shape  │    │
│  │  4. Text Content: SHA-256 of all text runs                          │    │
│  └─────────────────────────────────────────────────────────────────────┘    │
│                                                                              │
│  SENSITIVITY:                                                                │
│  • Moving a shape by 1 pixel → Version changes                              │
│  • Resizing any shape → Version changes                                     │
│  • Text edits → Version changes                                             │
│  • Adding/removing slides/shapes → Version changes                          │
│                                                                              │
│  OUTPUT: "a1b2c3d4e5f6g7h8" (16 hex characters)                            │
│                                                                              │
└─────────────────────────────────────────────────────────────────────────────┘
```

### 3. Performance Characteristics (Benchmarks)

```
┌─────────────────────────────────────────────────────────────────────────────┐
│  OPERATION PERFORMANCE MATRIX                                                │
├─────────────────────────────────────────────────────────────────────────────┤
│                                                                              │
│  Operation                    │ Complexity │ 10-Slide │ 50-Slide │ Notes    │
│  ────────────────────────────┼────────────┼──────────┼──────────┼─────────  │
│  get_presentation_version()  │ O(N) shapes│   ~15ms  │   ~75ms  │ 2x/mut   │
│  capability_probe(deep=True) │ O(M) layout│  ~120ms  │  ~600ms+ │ 15s TO   │
│  add_shape()                 │ O(1)       │    ~8ms  │    ~8ms  │ Constant │
│  replace_text(global)        │ O(N) runs  │   ~25ms  │  ~125ms  │ Regex    │
│  save()                      │ I/O Bound  │   ~50ms  │  ~200ms+ │ Disk     │
│                                                                              │
│  OPTIMIZATION GUIDELINES:                                                    │
│  • Batch mutations before single save() in custom scripts                   │
│  • Use deep=False for probes unless geometry required                       │
│  • Limit: <100 slides, <50MB for interactive sessions                       │
│                                                                              │
└─────────────────────────────────────────────────────────────────────────────┘
```

### 4. Opacity XML Injection (Internal Magic)

```
┌─────────────────────────────────────────────────────────────────────────────┐
│  OPACITY INJECTION MECHANISM                                                 │
├─────────────────────────────────────────────────────────────────────────────┤
│                                                                              │
│  PROBLEM: python-pptx lacks transparency/opacity support                    │
│  SOLUTION: Direct lxml injection into OOXML structure                       │
│                                                                              │
│  SCALE CONVERSION:                                                           │
│  ┌─────────────────────────────────────────────────────────────────────┐    │
│  │  Core Input:  0.0 ←───────────────────────────→ 1.0                 │    │
│  │               (invisible)                    (opaque)                │    │
│  │                           ↓                                          │    │
│  │  OOXML Scale:    0 ←───────────────────────────→ 100000             │    │
│  │               (invisible)                    (opaque)                │    │
│  │                                                                      │    │
│  │  Formula: ooxml_value = int(opacity * 100000)                       │    │
│  └─────────────────────────────────────────────────────────────────────┘    │
│                                                                              │
│  XML STRUCTURE INJECTED:                                                     │
│  ┌─────────────────────────────────────────────────────────────────────┐    │
│  │  <a:solidFill>                                                       │    │
│  │    <a:srgbClr val="FF0000">                                         │    │
│  │      <a:alpha val="50000"/>  <!-- 50% Opacity = 0.5 -->             │    │
│  │    </a:srgbClr>                                                      │    │
│  │  </a:solidFill>                                                      │    │
│  └─────────────────────────────────────────────────────────────────────┘    │
│                                                                              │
│  PREREQUISITE: <a:solidFill> must exist (shape must have a fill)           │
│  NAMESPACE: a: = http://schemas.openxmlformats.org/drawingml/2006/main     │
│                                                                              │
└─────────────────────────────────────────────────────────────────────────────┘
```

### 5. Z-Order XML Manipulation

```
┌─────────────────────────────────────────────────────────────────────────────┐
│  Z-ORDER IMPLEMENTATION                                                      │
├─────────────────────────────────────────────────────────────────────────────┤
│                                                                              │
│  MECHANISM: Physical movement of <p:sp> elements in <p:spTree> XML list    │
│                                                                              │
│  XML STRUCTURE:                                                              │
│  ┌─────────────────────────────────────────────────────────────────────┐    │
│  │  <p:spTree>                                                          │    │
│  │    <p:nvGrpSpPr>...</p:nvGrpSpPr>  <!-- Index 0: Group properties --> │    │
│  │    <p:grpSpPr>...</p:grpSpPr>      <!-- Index 1: More group props --> │    │
│  │    <p:sp>...</p:sp>                <!-- Index 2: First shape -->      │    │
│  │    <p:sp>...</p:sp>                <!-- Index 3: Second shape -->     │    │
│  │    <p:sp>...</p:sp>                <!-- Index 4: Third shape -->      │    │
│  │  </p:spTree>                                                          │    │
│  └─────────────────────────────────────────────────────────────────────┘    │
│                                                                              │
│  ACTION MAPPINGS:                                                            │
│  ┌──────────────────┬─────────────────────────────────────────────────┐    │
│  │ bring_to_front   │ Move element to END of <p:spTree> children      │    │
│  │ send_to_back     │ Move element to INDEX 2 (after bg/master refs)  │    │
│  │ bring_forward    │ Swap with next sibling                          │    │
│  │ send_backward    │ Swap with previous sibling                      │    │
│  └──────────────────┴─────────────────────────────────────────────────┘    │
│                                                                              │
│  ⚠️ CRITICAL: All shape indices are INVALIDATED after z-order change       │
│                                                                              │
└─────────────────────────────────────────────────────────────────────────────┘
```

### 6. Chart Update Strategy (Fallback Pattern)

```
┌─────────────────────────────────────────────────────────────────────────────┐
│  CHART DATA UPDATE STRATEGY                                                  │
├─────────────────────────────────────────────────────────────────────────────┤
│                                                                              │
│  STEP 1: Try Native Method (Best - preserves all formatting)               │
│  ┌─────────────────────────────────────────────────────────────────────┐    │
│  │  try:                                                                │    │
│  │      chart.replace_data(new_chart_data)                             │    │
│  │      return success                                                  │    │
│  │  except AttributeError:                                              │    │
│  │      # Older python-pptx version                                    │    │
│  │      pass                                                            │    │
│  └─────────────────────────────────────────────────────────────────────┘    │
│                                                                              │
│  STEP 2: Fallback Recreation (Preserves position/size/title only)          │
│  ┌─────────────────────────────────────────────────────────────────────┐    │
│  │  # Capture original properties                                       │    │
│  │  original_pos = get_shape_position(chart_shape)                      │    │
│  │  original_size = get_shape_size(chart_shape)                         │    │
│  │  original_title = chart.title                                        │    │
│  │                                                                      │    │
│  │  # Remove old chart                                                  │    │
│  │  remove_shape(chart_shape)                                           │    │
│  │                                                                      │    │
│  │  # Create new chart with same properties                             │    │
│  │  add_chart(..., position=original_pos, size=original_size)          │    │
│  │                                                                      │    │
│  │  ⚠️ Warning: Custom formatting (colors, fonts) will be RESET        │    │
│  └─────────────────────────────────────────────────────────────────────┘    │
│                                                                              │
└─────────────────────────────────────────────────────────────────────────────┘
```

### 7. Transient Slide Cleanup (Exact Implementation)

```python
# Complete transient slide pattern with exact cleanup mechanism

def analyze_layout_safe(prs, layout):
    """
    Create temporary slide for accurate geometry analysis.
    NEVER save() while transient slide exists.
    """
    slide = None
    added_index = -1
    
    try:
        # Create temporary slide
        slide = prs.slides.add_slide(layout)
        added_index = len(prs.slides) - 1
        
        # Analyze instantiated slide geometry
        # (Template positions are theoretical until instantiated)
        return extract_metrics(slide)
        
    finally:
        # ALWAYS cleanup (even on failure)
        if added_index != -1 and added_index < len(prs.slides):
            try:
                # Get relationship ID from slide ID list
                rId = prs.slides._sldIdLst[added_index].rId
                
                # Drop the relationship from the presentation part
                prs.part.drop_rel(rId)
                
                # Delete from the slide ID list
                del prs.slides._sldIdLst[added_index]
                
            except Exception:
                # Suppress cleanup errors to avoid masking analysis failures
                # File is not saved, so transient slide disappears anyway
                pass
```

### 8. Path Validation Details

```
┌─────────────────────────────────────────────────────────────────────────────┐
│  PATH VALIDATION SECURITY                                                    │
├─────────────────────────────────────────────────────────────────────────────┤
│                                                                              │
│  TRAVERSAL PROTECTION:                                                       │
│  ┌─────────────────────────────────────────────────────────────────────┐    │
│  │  if allowed_base_dirs is configured:                                 │    │
│  │      for base in allowed_base_dirs:                                  │    │
│  │          if not path.is_relative_to(base):                          │    │
│  │              raise PathValidationError("Path traversal detected")   │    │
│  └─────────────────────────────────────────────────────────────────────┘    │
│                                                                              │
│  EXTENSION ENFORCEMENT:                                                      │
│  ┌─────────────────────────────────────────────────────────────────────┐    │
│  │  ALLOWED_EXTENSIONS = {".pptx", ".pptm", ".potx"}                   │    │
│  │                                                                      │    │
│  │  if path.suffix.lower() not in ALLOWED_EXTENSIONS:                  │    │
│  │      raise PathValidationError(f"Invalid extension: {path.suffix}") │    │
│  └─────────────────────────────────────────────────────────────────────┘    │
│                                                                              │
└─────────────────────────────────────────────────────────────────────────────┘
```

### 9. Approval Token Generation (Exact Implementation)

```python
import hmac
import hashlib
import base64
import json
import time

def generate_approval_token(scope: str, user: str, secret_key: bytes) -> str:
    """
    Generate HMAC-based approval token for destructive operations.
    
    Args:
        scope: Required scope (e.g., "delete:slide", "remove:shape")
        user: User identifier
        secret_key: Secret key for HMAC signing (bytes)
        
    Returns:
        Token string in format "HMAC-SHA256:{base64_payload}.{signature}"
    """
    payload = {
        "scope": scope,
        "user": user,
        "issued": time.time(),
        "expiry": time.time() + 3600,  # 1 hour validity
        "single_use": True
    }
    
    # Serialize and Base64 encode
    json_payload = json.dumps(payload)
    b64_payload = base64.urlsafe_b64encode(json_payload.encode()).decode()
    
    # Sign with HMAC-SHA256
    signature = hmac.new(
        secret_key, 
        b64_payload.encode(), 
        hashlib.sha256
    ).hexdigest()
    
    # Combine into token format
    return f"HMAC-SHA256:{b64_payload}.{signature}"
```

### 10. Footer Mechanism Details

```
┌─────────────────────────────────────────────────────────────────────────────┐
│  FOOTER MECHANISM (The "Master Trap")                                        │
├─────────────────────────────────────────────────────────────────────────────┤
│                                                                              │
│  CORE BEHAVIOR:                                                              │
│  ┌─────────────────────────────────────────────────────────────────────┐    │
│  │  1. Iterate through ALL slides in presentation                       │    │
│  │  2. Search for placeholder types:                                    │    │
│  │     • Type 7: Date placeholder                                       │    │
│  │     • Type 6: Slide number placeholder                               │    │
│  │     • Type 5: Footer text placeholder                                │    │
│  │  3. If found: Set text/visibility as requested                       │    │
│  │  4. Return: slides_processed count                                   │    │
│  └─────────────────────────────────────────────────────────────────────┘    │
│                                                                              │
│  THE TRAP:                                                                   │
│  • prs.slide_masters[0].placeholders may show footer exists                 │
│  • BUT individual slides may have "Hide Background Graphics" enabled        │
│  • OR slides haven't instantiated the placeholder from master               │
│                                                                              │
│  IMPORTANT: Core does NOT create text boxes as fallback                     │
│  → That's the Tool layer's responsibility (ppt_set_footer.py)              │
│                                                                              │
│  TOOL LAYER FALLBACK:                                                        │
│  ┌─────────────────────────────────────────────────────────────────────┐    │
│  │  if result["slides_processed"] == 0:                                 │    │
│  │      # Fallback: Create text box overlay at footer position          │    │
│  │      add_text_box(slide_index, text, position=footer_pos, ...)       │    │
│  └─────────────────────────────────────────────────────────────────────┘    │
│                                                                              │
└─────────────────────────────────────────────────────────────────────────────┘
```

### 11. Backward Compatibility (v3.0 → v3.1)

```
┌─────────────────────────────────────────────────────────────────────────────┐
│  BACKWARD COMPATIBILITY MATRIX                                               │
├─────────────────────────────────────────────────────────────────────────────┤
│                                                                              │
│  ✅ PRESERVED (Additive Changes):                                           │
│  ┌─────────────────────────────────────────────────────────────────────┐    │
│  │  • Shape indices: Methods return dicts BUT preserve shape_index key  │    │
│  │  • Versioning: New presentation_version_before/after are ADDITIVE    │    │
│  │  • Transparency: Deprecated but supported with conversion + warning  │    │
│  └─────────────────────────────────────────────────────────────────────┘    │
│                                                                              │
│  ❌ BREAKING CHANGES (Intentional for Safety):                              │
│  ┌─────────────────────────────────────────────────────────────────────┐    │
│  │  • Silent index clamping: REMOVED                                    │    │
│  │    - v3.0: add_slide(index=99) → silently clamped to valid range   │    │
│  │    - v3.1: add_slide(index=99) → raises SlideNotFoundError         │    │
│  │                                                                      │    │
│  │  • Destructive operations: Now REQUIRE approval tokens               │    │
│  │    - delete_slide() without token → ApprovalTokenError              │    │
│  │    - remove_shape() without token → ApprovalTokenError              │    │
│  └─────────────────────────────────────────────────────────────────────┘    │
│                                                                              │
│  MIGRATION PATTERN:                                                          │
│  ┌─────────────────────────────────────────────────────────────────────┐    │
│  │  # v3.0 pattern (still works but deprecated)                         │    │
│  │  idx = agent.add_slide(layout_name="Blank")  # Returns int          │    │
│  │                                                                      │    │
│  │  # v3.1 pattern (recommended)                                        │    │
│  │  result = agent.add_slide(layout_name="Blank")  # Returns dict      │    │
│  │  idx = result["slide_index"]                                         │    │
│  │  version = result["presentation_version_after"]                      │    │
│  └─────────────────────────────────────────────────────────────────────┘    │
│                                                                              │
└─────────────────────────────────────────────────────────────────────────────┘
```

---

## 📋 Complete API Quick Reference

### File Operations

| Method | Arguments | Returns | Throws |
|--------|-----------|---------|--------|
| `open()` | `filepath`, `acquire_lock=True` | None | `PathValidationError`, `FileLockError` |
| `save()` | `filepath=None` | None | `PowerPointAgentError` |
| `clone_presentation()` | `output_path` | New `PowerPointAgent` instance | `PathValidationError` |
| `get_presentation_version()` | None | `str` (16-char SHA-256 prefix) | None |
| `get_presentation_info()` | None | `Dict` with metadata | None |

### Slide Operations

| Method | Arguments | Returns | Destructive? |
|--------|-----------|---------|--------------|
| `add_slide()` | `layout_name`, `index=None` | `Dict` with `slide_index`, versions | No |
| `delete_slide()` | `index`, `approval_token` | None | **Yes** (requires token) |
| `duplicate_slide()` | `index` | `int` (new index) | No |
| `reorder_slides()` | `from_index`, `to_index` | None | No |
| `get_slide_info()` | `slide_index` | `Dict` with shapes | No |

### Shape Operations

| Method | Arguments | Returns | Index Invalidating? |
|--------|-----------|---------|---------------------|
| `add_shape()` | `slide_index`, `shape_type`, `position`, `size`, `fill_opacity`, ... | `Dict` with `shape_index` | Yes (adds at end) |
| `format_shape()` | `slide_index`, `shape_index`, `fill_color`, `fill_opacity`, ... | `Dict` with changes | No |
| `remove_shape()` | `slide_index`, `shape_index`, `approval_token` | None | **Yes** (shifts indices) |
| `set_z_order()` | `slide_index`, `shape_index`, `action` | `Dict` with z-order change | **Yes** (reorders all) |

### Text & Content Operations

| Method | Arguments | Returns | Notes |
|--------|-----------|---------|-------|
| `add_text_box()` | `slide_index`, `text`, `position`, `size`, ... | `shape_index` | Auto-fit text |
| `add_bullet_list()` | `slide_index`, `items`, `position`, `size`, ... | `shape_index` | Styles: bullet, numbered, none |
| `replace_text()` | `find`, `replace`, `match_case=False` | `Dict` with replacements | Preserves formatting |
| `add_notes()` | `slide_index`, `text`, `mode="append"` | `Dict` with preview | Modes: append, prepend, overwrite |
| `set_footer()` | `text`, `show_number`, `show_date` | `slides_processed` count | No text box fallback |

### Chart & Image Operations

| Method | Arguments | Returns | Notes |
|--------|-----------|---------|-------|
| `add_chart()` | `chart_type`, `data`, `position`, `size`, ... | `shape_index` | Types: Column, Bar, Line, Pie, ... |
| `update_chart_data()` | `slide_index`, `chart_index`, `data` | None | Fallback recreation possible |
| `insert_image()` | `slide_index`, `image_path`, `position`, `size`, `alt_text` | `shape_index` | Auto aspect ratio |
| `replace_image()` | `slide_index`, `old_image_name`, `new_image_path` | None | Matches by name |

### Validation Operations

| Method | Arguments | Returns | Notes |
|--------|-----------|---------|-------|
| `validate_presentation()` | None | `Dict` with issues | Empty slides, missing assets |
| `check_accessibility()` | None | `Dict` with WCAG issues | Contrast, alt text, font size |

---

## 🎯 Ready State Confirmation

I now have **complete technical depth** on the PowerPointAgent core library, including:

| Aspect | Deep Understanding |
|--------|-------------------|
| **File Locking** | `O_CREAT\|O_EXCL` via `errno.EEXIST`, 10s timeout |
| **Version Hashing** | SHA-256 of geometry + content, 16-char prefix |
| **Performance** | Exact timing benchmarks per operation type |
| **XML Injection** | Opacity alpha tags, z-order element movement |
| **Chart Fallback** | `replace_data()` → recreation pattern |
| **Footer Trap** | Placeholder types 5/6/7, tool-layer fallback |
| **Token Generation** | HMAC-SHA256 with Base64 payload |
| **Path Validation** | `is_relative_to()`, extension whitelist |
| **Backward Compat** | Dict returns, deprecated transparency, removed clamping |

