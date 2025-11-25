# Deep Analysis: PowerPoint Agent Core Library Review

## Executive Summary

I will conduct a meticulous, systematic review of the `powerpoint_agent_core.py` library, examining it through multiple lenses: correctness, safety, performance, maintainability, API design, and alignment with the Presentation Architect system prompt v3.0. This analysis will identify strengths, gaps, bugs, security concerns, and provide prioritized recommendations for improvement.

---

## Phase 1: Structural Analysis

### 1.1 Module Organization Overview

```
powerpoint_agent_core.py (1,800+ lines)
├── IMPORTS & DEPENDENCY HANDLING
├── LOGGING SETUP
├── EXCEPTIONS (9 custom exception classes)
├── CONSTANTS & ENUMS (11 enums/constant groups)
├── FILE LOCKING (FileLock class)
├── POSITION & SIZE HELPERS (Position, Size classes)
├── COLOR MANAGEMENT (ColorHelper class)
├── TEMPLATE PRESERVATION (TemplateProfile class)
├── ACCESSIBILITY CHECKER (AccessibilityChecker class)
├── ASSET VALIDATOR (AssetValidator class)
├── MAIN AGENT CLASS (PowerPointAgent - 50+ methods)
└── MODULE METADATA
```

### 1.2 Dependency Analysis

| Dependency | Required | Optional | Import Handling |
|------------|----------|----------|-----------------|
| `python-pptx` | ✅ | - | ✅ Proper error with install instructions |
| `PIL/Pillow` | - | ✅ | ✅ Graceful degradation with `HAS_PILLOW` flag |
| `pandas` | - | ✅ | ⚠️ Imported but never used in code |
| `subprocess` | ✅ | - | ✅ Present (added in v1.1.0) |
| `logging` | ✅ | - | ✅ Proper setup |

**Finding**: `pandas` is imported but never used anywhere in the module. This is dead code that should be removed.

---

## Phase 2: Correctness Analysis

### 2.1 Bug Inventory

#### **BUG-001: Race Condition in FileLock**
```python
def acquire(self) -> bool:
    """Acquire lock with timeout."""
    start_time = time.time()
    while time.time() - start_time < self.timeout:
        try:
            self.lockfile.touch(exist_ok=False)  # NOT ATOMIC!
            self.acquired = True
            return True
        except FileExistsError:
            time.sleep(0.1)
    return False
```

**Problem**: `Path.touch(exist_ok=False)` is NOT atomic on all filesystems. Two processes can both succeed in creating the file simultaneously on network filesystems or certain configurations.

**Severity**: HIGH (data corruption risk in concurrent scenarios)

**Fix**:
```python
import os
import fcntl  # Unix only, need platform-specific handling

def acquire(self) -> bool:
    """Acquire lock with timeout using atomic operations."""
    start_time = time.time()
    while time.time() - start_time < self.timeout:
        try:
            # Use O_CREAT | O_EXCL for atomic creation
            fd = os.open(str(self.lockfile), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            os.close(fd)
            self.acquired = True
            return True
        except FileExistsError:
            time.sleep(0.1)
        except OSError as e:
            if e.errno == 17:  # EEXIST
                time.sleep(0.1)
            else:
                raise
    return False
```

---

#### **BUG-002: Incorrect Slide Insertion Logic**
```python
def add_slide(self, layout_name: str = "Title and Content", 
              index: Optional[int] = None) -> int:
    # ...
    if index is None:
        slide = self.prs.slides.add_slide(layout)
        return len(self.prs.slides) - 1
    else:
        slide = self.prs.slides.add_slide(layout)  # Adds at END
        xml_slides = self.prs.slides._sldIdLst
        xml_slides.insert(index, xml_slides[-1])  # Then moves
        return index
```

**Problem**: When inserting at a specific index, the slide is first added at the end, then the XML element is moved. However, `xml_slides[-1]` is removed from the end and inserted at `index`, but the method doesn't remove it from its original position first, potentially creating duplicate references.

**Severity**: MEDIUM (may cause XML corruption in edge cases)

**Fix**:
```python
else:
    slide = self.prs.slides.add_slide(layout)
    xml_slides = self.prs.slides._sldIdLst
    slide_elem = xml_slides[-1]
    xml_slides.remove(slide_elem)  # Remove from end first
    xml_slides.insert(index, slide_elem)  # Then insert at target
    return index
```

---

#### **BUG-003: Placeholder Type Comparison Issues**
```python
# In set_title():
if ph_type == PP_PLACEHOLDER.TITLE or ph_type == PP_PLACEHOLDER.CENTER_TITLE:
    title_shape = shape
elif ph_type == PP_PLACEHOLDER.SUBTITLE:
    subtitle_shape = shape
```

**Problem**: The code compares `ph_type` (which is `shape.placeholder_format.type`) directly with `PP_PLACEHOLDER` enum members. However, depending on python-pptx version, `placeholder_format.type` may return an integer or an enum member, causing comparison failures.

**Evidence**: The `PLACEHOLDER_TYPE_NAMES` dictionary uses integer keys, suggesting awareness of this issue, but the comparison code doesn't handle it consistently.

**Severity**: MEDIUM (silent failures in placeholder detection)

**Fix**:
```python
def _get_placeholder_type(self, shape) -> Optional[int]:
    """Safely get placeholder type as integer."""
    if not shape.is_placeholder:
        return None
    ph_type = shape.placeholder_format.type
    # Handle both enum and integer returns
    if hasattr(ph_type, 'value'):
        return ph_type.value
    return int(ph_type) if ph_type is not None else None

# Usage:
ph_type = self._get_placeholder_type(shape)
if ph_type in (1, 3):  # TITLE or CENTER_TITLE
    title_shape = shape
elif ph_type == 4:  # SUBTITLE
    subtitle_shape = shape
```

---

#### **BUG-004: SUBTITLE Constant Mismatch**
```python
# In PLACEHOLDER_TYPE_NAMES:
4: "SUBTITLE",        # Subtitle (also used for footer in some layouts)

# But in code comparisons:
elif ph_type == PP_PLACEHOLDER.SUBTITLE:
```

**Problem**: The changelog claims "Fixed subtitle placeholder type (was 2, should be SUBTITLE constant)" but the dictionary shows SUBTITLE as 4, while CONTENT is 2. The `PP_PLACEHOLDER.SUBTITLE` enum value in python-pptx is actually 4, but the code inconsistently uses enum comparisons in some places and integer keys in others.

**Severity**: LOW (documentation/consistency issue)

---

#### **BUG-005: Unsafe Exception Handling in Accessibility Checker**
```python
for shape in slide.shapes:
    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
        if not shape.name or shape.name.startswith("Picture"):
            issues["missing_alt_text"].append({...})
```

**Problem**: The logic assumes that a picture named "Picture" or with empty name lacks alt text. However:
1. A user might legitimately name an image "Picture of sunset"
2. The actual alt text is stored in `descr` attribute, not `name`
3. This produces false positives

**Severity**: MEDIUM (incorrect accessibility reporting)

**Fix**:
```python
def _check_image_alt_text(shape) -> bool:
    """Check if image has meaningful alt text."""
    # Check description attribute (actual alt text storage)
    try:
        descr = shape._element.get('descr', '')
        if descr and descr.strip():
            return True
    except:
        pass
    
    # Check name as fallback (some tools store alt text here)
    if shape.name and not shape.name.startswith("Picture") and len(shape.name) > 3:
        return True
    
    return False
```

---

#### **BUG-006: Color Luminance Calculation Type Handling**
```python
@staticmethod
def luminance(rgb_color: RGBColor) -> float:
    """Calculate relative luminance for WCAG contrast."""
    if hasattr(rgb_color, 'r'):
        r = rgb_color.r
        g = rgb_color.g
        b = rgb_color.b
    else:
        # Handle pptx.dml.color.RGBColor which behaves like a hex string
        hex_color = str(rgb_color)
        r = int(hex_color[0:2], 16)
        g = int(hex_color[2:4], 16)
        b = int(hex_color[4:6], 16)
```

**Problem**: The `str(rgb_color)` returns the hex string WITHOUT the `#` prefix (e.g., "FF0000"), but the slicing assumes no prefix. This works, but if `RGBColor.__str__` behavior changes or includes a prefix in some contexts, it will fail silently with wrong values.

**Severity**: LOW (works currently, but fragile)

**Fix**:
```python
hex_color = str(rgb_color).lstrip('#')
if len(hex_color) != 6:
    raise ValueError(f"Invalid color format: {rgb_color}")
```

---

#### **BUG-007: Chart Data Update Fallback Loses Formatting**
```python
def update_chart_data(self, ...):
    try:
        chart_shape.chart.replace_data(chart_data)
    except AttributeError:
        logger.warning("chart.replace_data() not available...")
        # Fallback: delete and recreate chart
        sp = chart_shape.element
        sp.getparent().remove(sp)
        new_chart_shape = slide.shapes.add_chart(...)
```

**Problem**: The fallback path recreates the chart, losing all formatting (colors, legend customization, axis settings, etc.). The warning is logged but the operation proceeds silently, potentially causing unexpected visual changes.

**Severity**: MEDIUM (silent data loss of formatting)

**Recommendation**: Either fail explicitly or capture and restore more formatting attributes.

---

#### **BUG-008: Missing Bounds Check in format_shape**
```python
def format_shape(self, slide_index: int, shape_index: int, ...):
    slide = self.get_slide(slide_index)
    
    if shape_index >= len(slide.shapes):
        raise ValueError(f"Shape index {shape_index} out of range")
```

**Problem**: Only checks upper bound, not lower bound. Negative indices will work due to Python list semantics but may produce unexpected results.

**Severity**: LOW (edge case)

**Fix**:
```python
if not 0 <= shape_index < len(slide.shapes):
    raise ValueError(f"Shape index {shape_index} out of range (0-{len(slide.shapes)-1})")
```

---

### 2.2 Logic Errors Summary

| ID | Location | Issue | Severity |
|----|----------|-------|----------|
| BUG-001 | `FileLock.acquire()` | Non-atomic file creation | HIGH |
| BUG-002 | `add_slide()` | Incorrect XML manipulation | MEDIUM |
| BUG-003 | `set_title()`, validators | Inconsistent placeholder type handling | MEDIUM |
| BUG-004 | Constants | SUBTITLE value documentation mismatch | LOW |
| BUG-005 | `AccessibilityChecker` | Wrong alt text detection logic | MEDIUM |
| BUG-006 | `ColorHelper.luminance()` | Fragile hex string parsing | LOW |
| BUG-007 | `update_chart_data()` | Silent formatting loss | MEDIUM |
| BUG-008 | Multiple methods | Missing negative index bounds checks | LOW |

---

## Phase 3: Safety & Security Analysis

### 3.1 File Path Security

#### **SECURITY-001: Path Traversal Not Validated**
```python
def open(self, filepath: Path, acquire_lock: bool = True) -> None:
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    self.filepath = filepath
```

**Problem**: No validation that `filepath` is within expected directories. A malicious path like `../../../etc/passwd` would be accepted if it exists.

**Severity**: MEDIUM (depends on deployment context)

**Fix**:
```python
def _validate_path(self, filepath: Path, must_exist: bool = True) -> Path:
    """Validate and resolve file path safely."""
    resolved = filepath.resolve()
    
    # Ensure it's a .pptx file
    if resolved.suffix.lower() not in ('.pptx', '.pptm', '.potx'):
        raise ValueError(f"Invalid file type: {resolved.suffix}")
    
    if must_exist and not resolved.exists():
        raise FileNotFoundError(f"File not found: {resolved}")
    
    return resolved
```

---

#### **SECURITY-002: Arbitrary Command Execution in PDF Export**
```python
def export_to_pdf(self, output_path: Path) -> None:
    result = subprocess.run(
        ['soffice', '--headless', '--convert-to', 'pdf', 
         '--outdir', str(output_path.parent), temp_pptx],
        capture_output=True,
        timeout=60
    )
```

**Analysis**: The code uses a list for subprocess arguments (good - prevents shell injection), but:
1. `output_path.parent` could contain special characters
2. No validation of output path
3. Temp file created in system temp directory with predictable name

**Severity**: LOW (command injection unlikely due to list args, but path injection possible)

**Fix**:
```python
def export_to_pdf(self, output_path: Path) -> None:
    output_path = self._validate_path(output_path, must_exist=False)
    
    # Use secure temp file
    with tempfile.NamedTemporaryFile(suffix='.pptx', delete=False) as tmp:
        temp_pptx = Path(tmp.name)
    
    try:
        self.prs.save(str(temp_pptx))
        
        # Validate outdir exists and is writable
        outdir = output_path.parent.resolve()
        if not outdir.is_dir():
            raise ValueError(f"Output directory does not exist: {outdir}")
        
        result = subprocess.run(
            ['soffice', '--headless', '--convert-to', 'pdf', 
             '--outdir', str(outdir), str(temp_pptx)],
            capture_output=True,
            timeout=60,
            check=False  # Handle errors explicitly
        )
```

---

### 3.2 Resource Management

#### **RESOURCE-001: Potential File Handle Leak**
```python
def open(self, filepath: Path, acquire_lock: bool = True) -> None:
    if acquire_lock:
        self._lock = FileLock(filepath)
        if not self._lock.acquire():
            raise FileLockError(f"Could not lock file: {filepath}")
    
    self.prs = Presentation(str(filepath))  # If this fails, lock is not released
```

**Problem**: If `Presentation()` raises an exception, the acquired lock is never released.

**Severity**: MEDIUM (resource leak, potential deadlock)

**Fix**:
```python
def open(self, filepath: Path, acquire_lock: bool = True) -> None:
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    self.filepath = filepath
    
    if acquire_lock:
        self._lock = FileLock(filepath)
        if not self._lock.acquire():
            raise FileLockError(f"Could not lock file: {filepath}")
    
    try:
        self.prs = Presentation(str(filepath))
        self._template_profile = TemplateProfile(self.prs)
    except Exception:
        # Release lock if presentation loading fails
        if self._lock:
            self._lock.release()
            self._lock = None
        raise
```

---

#### **RESOURCE-002: Context Manager Doesn't Handle All State**
```python
def __exit__(self, exc_type, exc_val, exc_tb):
    self.close()
    return False

def close(self) -> None:
    """Close presentation and release lock."""
    self.prs = None
    
    if self._lock:
        self._lock.release()
        self._lock = None
```

**Problem**: `close()` doesn't save changes. If used as context manager and exception occurs, unsaved changes are lost without warning.

**Severity**: LOW (expected behavior, but could be clearer)

**Recommendation**: Add `save_on_exit` option or document behavior clearly.

---

### 3.3 Input Validation Gaps

| Location | Input | Validation Status |
|----------|-------|-------------------|
| `add_slide(layout_name)` | String | ✅ Validated against available layouts |
| `add_text_box(text)` | String | ❌ No length limit |
| `add_text_box(font_size)` | Integer | ❌ No range validation |
| `add_shape(shape_type)` | String | ✅ Mapped, defaults to RECTANGLE |
| `from_hex(hex_color)` | String | ✅ Format validated |
| `add_chart(data)` | Dict | ❌ No schema validation |
| `Position.from_dict()` | Dict | ⚠️ Partial validation |

---

## Phase 4: API Design Analysis

### 4.1 Method Signature Consistency

#### **Issue: Inconsistent Parameter Naming**
```python
def add_text_box(..., slide_index: int, ...)
def set_title(self, slide_index: int, ...)
def add_bullet_list(self, slide_index: int, ...)
def add_shape(self, slide_index: int, ...)

# But:
def format_text(self, slide_index: int, shape_index: int, ...)
def format_shape(self, slide_index: int, shape_index: int, ...)
```

**Observation**: Naming is consistent. ✅

---

#### **Issue: Inconsistent Return Values**
```python
def add_slide(...) -> int:  # Returns index
def add_text_box(...) -> None:  # Returns nothing
def add_shape(...) -> None:  # Returns nothing
def replace_text(...) -> int:  # Returns count
def replace_image(...) -> bool:  # Returns success
```

**Problem**: Methods that create objects sometimes return the created object/index, sometimes return nothing. This makes it difficult to chain operations or immediately reference created elements.

**Severity**: MEDIUM (API usability)

**Recommendation**: All `add_*` methods should return the created shape/index:
```python
def add_text_box(...) -> int:  # Return shape index
def add_shape(...) -> int:  # Return shape index
def add_table(...) -> int:  # Return shape index
```

---

### 4.2 Missing API Methods (Gap Analysis vs System Prompt v3.0)

| Required by System Prompt | Available in Core | Status |
|--------------------------|-------------------|--------|
| Add speaker notes | ❌ (only `extract_notes`) | **MISSING** |
| Set z-order | ❌ | **MISSING** |
| Remove shape | ❌ | **MISSING** |
| Set footer | ❌ | **MISSING** |
| Crop image | ❌ (only `crop_resize_image` which is resize-only) | **MISSING** |
| Set background | ❌ | **MISSING** |
| Get presentation version | ❌ | **MISSING** |
| Duplicate presentation (clone) | ❌ | **MISSING** |

**Critical Gaps**: The core library is missing several methods that the CLI tools implement directly. This creates inconsistency where some operations use the core library and others implement their own logic.

---

### 4.3 Error Handling Design

#### **Strength**: Rich Exception Hierarchy
```python
PowerPointAgentError (base)
├── SlideNotFoundError
├── LayoutNotFoundError
├── ImageNotFoundError
├── InvalidPositionError
├── TemplateError
├── ThemeError
├── AccessibilityError
├── AssetValidationError
└── FileLockError
```

#### **Weakness**: Inconsistent Exception Usage
```python
# Some methods raise ValueError:
if width is None or height is None:
    raise ValueError("Text box must have explicit width and height")

# Others raise custom exceptions:
raise SlideNotFoundError(f"Slide index {index} out of range")

# Some use generic ValueError where custom would be better:
if shape_index >= len(slide.shapes):
    raise ValueError(f"Shape index {shape_index} out of range")
    # Should be: ShapeNotFoundError
```

**Recommendation**: Create `ShapeNotFoundError` and use custom exceptions consistently.

---

## Phase 5: Performance Analysis

### 5.1 Inefficient Operations

#### **PERF-001: Repeated Layout Lookup**
```python
def _get_layout(self, layout_name: str):
    for layout in self.prs.slide_layouts:  # O(n) every time
        if layout.name == layout_name:
            return layout
```

**Problem**: Every call to `add_slide()` triggers a linear search through all layouts.

**Fix**: Cache layout mapping on presentation open:
```python
def _build_layout_cache(self):
    """Build layout name -> layout object mapping."""
    self._layout_cache = {layout.name: layout for layout in self.prs.slide_layouts}

def _get_layout(self, layout_name: str):
    if not hasattr(self, '_layout_cache') or self._layout_cache is None:
        self._build_layout_cache()
    
    layout = self._layout_cache.get(layout_name)
    if not layout:
        raise LayoutNotFoundError(...)
    return layout
```

---

#### **PERF-002: Unnecessary Template Profile Capture**
```python
def open(self, filepath: Path, acquire_lock: bool = True) -> None:
    self.prs = Presentation(str(filepath))
    self._template_profile = TemplateProfile(self.prs)  # Always captures, even if unused
```

**Problem**: `TemplateProfile` iterates through all layouts, placeholders, and attempts to extract theme colors/fonts on every file open, even if this information is never used.

**Fix**: Lazy initialization:
```python
@property
def template_profile(self) -> TemplateProfile:
    if self._template_profile is None:
        self._template_profile = TemplateProfile(self.prs)
    return self._template_profile
```

---

#### **PERF-003: Full Presentation Scan in replace_text**
```python
def replace_text(self, find: str, replace: str, match_case: bool = False) -> int:
    for slide in self.prs.slides:
        for shape in slide.shapes:
            # Scans EVERY shape in EVERY slide
```

**Problem**: No early termination option. For large presentations, this is slow even if the target text is only on slide 1.

**Recommendation**: Add optional `slide_index` and `shape_index` parameters for targeted replacement (as the CLI tool now implements).

---

### 5.2 Memory Considerations

#### **PERF-004: Large Image Handling**
```python
def insert_image(self, ..., compress: bool = False):
    if compress and HAS_PILLOW:
        image_stream = AssetValidator.compress_image(image_path)
        # Entire compressed image is held in memory
```

**Problem**: For very large images, loading the entire image into memory could cause issues.

**Severity**: LOW (Pillow handles this reasonably well)

---

## Phase 6: Maintainability Analysis

### 6.1 Code Duplication

#### **DUP-001: Repeated Bounds Checking**
```python
# In format_text:
if shape_index >= len(slide.shapes):
    raise ValueError(f"Shape index {shape_index} out of range")

# In format_shape:
if shape_index >= len(slide.shapes):
    raise ValueError(f"Shape index {shape_index} out of range")

# In set_image_properties:
if shape_index >= len(slide.shapes):
    raise ValueError(f"Shape index {shape_index} out of range")

# In crop_resize_image:
# (uses slide.shapes[shape_index] without bounds check!)
```

**Fix**: Extract to helper method:
```python
def _get_shape(self, slide_index: int, shape_index: int):
    """Get shape with bounds validation."""
    slide = self.get_slide(slide_index)
    if not 0 <= shape_index < len(slide.shapes):
        raise ShapeNotFoundError(
            f"Shape index {shape_index} out of range (0-{len(slide.shapes)-1})",
            details={"slide_index": slide_index, "shape_count": len(slide.shapes)}
        )
    return slide.shapes[shape_index]
```

---

#### **DUP-002: Repeated Chart Finding**
```python
# In update_chart_data:
chart_shape = None
chart_count = 0
for shape in slide.shapes:
    if shape.has_chart:
        if chart_count == chart_index:
            chart_shape = shape
            break
        chart_count += 1

# In format_chart:
chart_shape = None
chart_count = 0
for shape in slide.shapes:
    if shape.has_chart:
        if chart_count == chart_index:
            chart_shape = shape
            break
        chart_count += 1
```

**Fix**: Extract to helper:
```python
def _get_chart_shape(self, slide_index: int, chart_index: int):
    """Get chart shape by index on slide."""
    slide = self.get_slide(slide_index)
    chart_count = 0
    for shape in slide.shapes:
        if shape.has_chart:
            if chart_count == chart_index:
                return shape
            chart_count += 1
    raise ValueError(f"Chart at index {chart_index} not found on slide {slide_index}")
```

---

### 6.2 Documentation Gaps

| Area | Status | Notes |
|------|--------|-------|
| Module docstring | ✅ Good | Version, changelog, author info |
| Class docstrings | ⚠️ Partial | Some classes lack docstrings |
| Method docstrings | ⚠️ Partial | Args documented, but Returns often missing |
| Type hints | ✅ Good | Comprehensive type hints |
| Inline comments | ⚠️ Sparse | Complex logic lacks explanation |
| Examples | ❌ Missing | No usage examples in docstrings |

---

### 6.3 Test Coverage Concerns

**Observation**: The core file has no tests included. Critical areas needing tests:

1. **FileLock** - Concurrency scenarios
2. **Position.from_dict()** - All input format variations
3. **ColorHelper.contrast_ratio()** - Known WCAG test cases
4. **Placeholder detection** - Various layout types
5. **Edge cases** - Empty presentations, corrupt files, large files

---

## Phase 7: Alignment with System Prompt v3.0

### 7.1 Feature Parity Analysis

| System Prompt v3.0 Requirement | Core Library Support | Gap |
|-------------------------------|---------------------|-----|
| Presentation versioning | ❌ Not implemented | **CRITICAL** |
| Manifest-driven execution | ❌ Not implemented | N/A (CLI layer) |
| Approval token verification | ❌ Not implemented | N/A (CLI layer) |
| Speaker notes CRUD | ⚠️ Only extract | **HIGH** |
| Z-order management | ❌ Not implemented | **HIGH** |
| Shape removal | ❌ Not implemented | **MEDIUM** |
| Footer management | ❌ Not implemented | **MEDIUM** |
| Background setting | ❌ Not implemented | **MEDIUM** |
| Presentation cloning | ❌ Not implemented | **LOW** (can use file copy) |
| Probe resilience | ❌ Not applicable | N/A (CLI layer) |

### 7.2 Recommended Core Library Additions

```python
# Methods to add to PowerPointAgent class:

def add_notes(self, slide_index: int, text: str, mode: str = "append") -> None:
    """Add speaker notes to slide."""
    pass

def set_z_order(self, slide_index: int, shape_index: int, action: str) -> Dict[str, int]:
    """Change shape z-order. Returns old and new indices."""
    pass

def remove_shape(self, slide_index: int, shape_index: int) -> None:
    """Remove shape from slide."""
    pass

def set_footer(self, text: str = None, show_number: bool = False, 
               show_date: bool = False, slide_index: int = None) -> None:
    """Set footer properties."""
    pass

def set_background(self, slide_index: int = None, color: str = None, 
                   image_path: Path = None) -> None:
    """Set slide background."""
    pass

def get_presentation_version(self) -> str:
    """Compute presentation version hash."""
    pass

def clone_presentation(self, output_path: Path) -> 'PowerPointAgent':
    """Clone presentation to new file."""
    pass
```

---

## Phase 8: Consolidated Recommendations

### 8.1 Critical Priority (Fix Immediately)

| ID | Issue | Impact | Effort |
|----|-------|--------|--------|
| BUG-001 | Non-atomic FileLock | Data corruption | Medium |
| RESOURCE-001 | Lock not released on error | Deadlock | Low |
| SECURITY-001 | Path traversal | Security | Low |
| BUG-002 | Incorrect slide insertion | XML corruption | Low |

### 8.2 High Priority (Fix Soon)

| ID | Issue | Impact | Effort |
|----|-------|--------|--------|
| BUG-003 | Placeholder type handling | Silent failures | Medium |
| BUG-005 | Wrong alt text detection | False positives | Low |
| API-GAP | Missing add_notes() | Feature incomplete | Medium |
| API-GAP | Missing set_z_order() | Feature incomplete | Medium |
| DUP-001 | Repeated bounds checking | Maintainability | Low |

### 8.3 Medium Priority (Improve When Possible)

| ID | Issue | Impact | Effort |
|----|-------|--------|--------|
| BUG-007 | Chart update loses formatting | Data loss | Medium |
| PERF-001 | Repeated layout lookup | Performance | Low |
| PERF-002 | Eager template capture | Performance | Low |
| API-RETURN | Inconsistent return values | Usability | Medium |

### 8.4 Low Priority (Nice to Have)

| ID | Issue | Impact | Effort |
|----|-------|--------|--------|
| BUG-004 | SUBTITLE documentation | Clarity | Trivial |
| BUG-006 | Fragile hex parsing | Robustness | Trivial |
| BUG-008 | Missing negative bounds | Edge cases | Trivial |
| DEAD-CODE | Unused pandas import | Cleanliness | Trivial |

---

## Phase 9: Proposed Refactored Structure

```python
# Recommended module organization for v2.0:

powerpoint_agent_core.py
├── __version__ = "2.0.0"
│
├── SECTION 1: IMPORTS & SETUP
│   ├── Standard library imports
│   ├── Third-party imports with graceful degradation
│   └── Logging configuration
│
├── SECTION 2: EXCEPTIONS
│   ├── PowerPointAgentError (base)
│   ├── SlideNotFoundError
│   ├── ShapeNotFoundError (NEW)
│   ├── LayoutNotFoundError
│   ├── ImageNotFoundError
│   ├── InvalidPositionError
│   ├── TemplateError
│   ├── ThemeError
│   ├── AccessibilityError
│   ├── AssetValidationError
│   └── FileLockError
│
├── SECTION 3: CONSTANTS
│   ├── Dimensions
│   ├── Anchor points
│   ├── Colors
│   ├── Fonts
│   ├── WCAG thresholds
│   └── Placeholder mappings
│
├── SECTION 4: ENUMS
│   ├── ShapeType
│   ├── ChartType
│   ├── TextAlignment
│   ├── VerticalAlignment
│   ├── BulletStyle
│   ├── ImageFormat
│   ├── ExportFormat
│   └── ZOrderAction (NEW)
│
├── SECTION 5: UTILITIES
│   ├── FileLock (with atomic operations)
│   ├── Position
│   ├── Size
│   ├── ColorHelper
│   └── PathValidator (NEW)
│
├── SECTION 6: ANALYSIS CLASSES
│   ├── TemplateProfile
│   ├── AccessibilityChecker
│   └── AssetValidator
│
├── SECTION 7: MAIN AGENT CLASS
│   ├── File Operations
│   │   ├── create_new()
│   │   ├── open()
│   │   ├── save()
│   │   ├── close()
│   │   └── clone() (NEW)
│   │
│   ├── Slide Operations
│   │   ├── add_slide()
│   │   ├── delete_slide()
│   │   ├── duplicate_slide()
│   │   ├── reorder_slides()
│   │   ├── get_slide()
│   │   └── get_slide_count()
│   │
│   ├── Text Operations
│   │   ├── add_text_box()
│   │   ├── set_title()
│   │   ├── add_bullet_list()
│   │   ├── format_text()
│   │   ├── replace_text()
│   │   ├── add_notes() (NEW)
│   │   └── set_footer() (NEW)
│   │
│   ├── Shape Operations
│   │   ├── add_shape()
│   │   ├── format_shape()
│   │   ├── remove_shape() (NEW)
│   │   ├── set_z_order() (NEW)
│   │   ├── add_table()
│   │   └── add_connector()
│   │
│   ├── Image Operations
│   │   ├── insert_image()
│   │   ├── replace_image()
│   │   ├── set_image_properties()
│   │   └── crop_image() (NEW - actual cropping)
│   │
│   ├── Chart Operations
│   │   ├── add_chart()
│   │   ├── update_chart_data()
│   │   └── format_chart()
│   │
│   ├── Layout & Theme
│   │   ├── set_slide_layout()
│   │   ├── set_background() (NEW)
│   │   └── get_available_layouts()
│   │
│   ├── Validation
│   │   ├── validate_presentation()
│   │   ├── check_accessibility()
│   │   └── validate_assets()
│   │
│   ├── Export
│   │   ├── export_to_pdf()
│   │   ├── export_to_images()
│   │   └── extract_notes()
│   │
│   ├── Information
│   │   ├── get_presentation_info()
│   │   ├── get_slide_info()
│   │   └── get_presentation_version() (NEW)
│   │
│   └── Private Helpers
│       ├── _validate_path()
│       ├── _get_layout()
│       ├── _get_shape()
│       ├── _get_chart_shape()
│       ├── _get_placeholder_type()
│       ├── _copy_shape()
│       └── _build_layout_cache()
│
└── SECTION 8: MODULE EXPORTS
    └── __all__ = [...]
```

---

## Phase 10: Implementation Roadmap

### Sprint 1: Critical Fixes (2-3 days)
1. Fix FileLock with atomic operations
2. Fix lock release on error
3. Add path validation
4. Fix slide insertion logic
5. Add comprehensive bounds checking helper

### Sprint 2: High Priority Gaps (3-4 days)
1. Implement `add_notes()` method
2. Implement `set_z_order()` method
3. Implement `remove_shape()` method
4. Fix placeholder type handling
5. Fix alt text detection logic

### Sprint 3: API Improvements (2-3 days)
1. Standardize return values
2. Add `ShapeNotFoundError`
3. Extract duplicate code to helpers
4. Add layout caching
5. Make template profile lazy

### Sprint 4: Documentation & Tests (3-4 days)
1. Add comprehensive docstrings with examples
2. Write unit tests for all public methods
3. Write integration tests for workflows
4. Add type stubs if needed
5. Update changelog and version

---

## Final Summary

### Strengths of Current Implementation
1. ✅ Good exception hierarchy
2. ✅ Comprehensive type hints
3. ✅ Flexible positioning system
4. ✅ WCAG contrast checking
5. ✅ Context manager support
6. ✅ Graceful optional dependency handling

### Critical Issues Requiring Immediate Attention
1. ❌ Non-atomic file locking (data corruption risk)
2. ❌ Resource leak on open failure
3. ❌ Missing core methods (notes, z-order, remove)
4. ❌ Inconsistent placeholder handling
5. ❌ Path traversal vulnerability

### Overall Assessment

| Category | Score | Notes |
|----------|-------|-------|
| Correctness | 6/10 | Several bugs, most non-critical |
| Safety | 5/10 | Lock issues, path validation gaps |
| Performance | 7/10 | Minor inefficiencies |
| API Design | 7/10 | Good structure, inconsistent returns |
| Maintainability | 6/10 | Some duplication, sparse comments |
| Feature Completeness | 6/10 | Missing key methods for v3.0 |
| **Overall** | **6.2/10** | Solid foundation, needs refinement |

The core library provides a solid foundation but requires targeted fixes for safety-critical issues and feature additions to fully support the Presentation Architect system prompt v3.0 requirements.
