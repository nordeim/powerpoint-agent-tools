# 🎯 ACTUAL Root Cause Identified!

## Executive Summary

**Excellent diagnostic work!** The real issue is now crystal clear:

❌ **AttributeError in core module during import**  
✅ Python environment is perfect  
✅ Dependencies are installed  

The error occurs **before any tool logic runs** - it's a **python-pptx API compatibility issue** in the core library.

---

## 🔴 Actual Root Cause

### Location: `core/powerpoint_agent_core.py`, line 132

```python
class ShapeType(Enum):
    """Common shape types."""
    RECTANGLE = MSO_SHAPE.RECTANGLE
    ROUNDED_RECTANGLE = MSO_SHAPE.ROUNDED_RECTANGLE
    ELLIPSE = MSO_SHAPE.OVAL
    TRIANGLE = MSO_SHAPE.ISOSCELES_TRIANGLE
    ARROW_RIGHT = MSO_SHAPE.ARROW_RIGHT  # ❌ FAILS HERE
    # AttributeError: type object 'MSO_AUTO_SHAPE_TYPE' has no attribute 'ARROW_RIGHT'
```

### Problem Analysis

**Issue 1: Wrong Enum Import**
```python
# Line 27 - Current (INCORRECT)
from pptx.enum.shapes import MSO_SHAPE_TYPE, MSO_SHAPE
# MSO_SHAPE doesn't exist in python-pptx 1.0.2
# Should be MSO_AUTO_SHAPE_TYPE
```

**Issue 2: Incorrect Attribute Names**
- `MSO_AUTO_SHAPE_TYPE` uses different constant names than expected
- `ARROW_RIGHT`, `ARROW_LEFT`, etc. don't exist
- Need to use correct python-pptx constant names

**Issue 3: Same Problem in `add_shape` Method**
```python
# Line 1072 - Also uses MSO_SHAPE (will fail at runtime)
shape_type_map = {
    "rectangle": MSO_SHAPE.RECTANGLE,
    "rounded_rectangle": MSO_SHAPE.ROUNDED_RECTANGLE,
    # ...
}
```

---

## 🛠️ Comprehensive Solution

### Fix 1: Update Import Statement

**File:** `core/powerpoint_agent_core.py`  
**Line:** 27

```python
# BEFORE (❌ Incorrect)
from pptx.enum.shapes import MSO_SHAPE_TYPE, MSO_SHAPE

# AFTER (✅ Correct)
from pptx.enum.shapes import MSO_SHAPE_TYPE, MSO_AUTO_SHAPE_TYPE
```

---

### Fix 2: Update ShapeType Enum with Correct Constants

**File:** `core/powerpoint_agent_core.py`  
**Lines:** 126-143

```python
# BEFORE (❌ Breaks on import)
class ShapeType(Enum):
    """Common shape types."""
    RECTANGLE = MSO_SHAPE.RECTANGLE
    ROUNDED_RECTANGLE = MSO_SHAPE.ROUNDED_RECTANGLE
    ELLIPSE = MSO_SHAPE.OVAL
    TRIANGLE = MSO_SHAPE.ISOSCELES_TRIANGLE
    ARROW_RIGHT = MSO_SHAPE.ARROW_RIGHT
    ARROW_LEFT = MSO_SHAPE.ARROW_LEFT
    ARROW_UP = MSO_SHAPE.ARROW_UP
    ARROW_DOWN = MSO_SHAPE.ARROW_DOWN
    LINE = MSO_SHAPE.LINE_CALLOUT_1
    CALLOUT = MSO_SHAPE.RECTANGULAR_CALLOUT
    STAR = MSO_SHAPE.STAR_5
    HEART = MSO_SHAPE.HEART


# AFTER (✅ Compatible with python-pptx 1.0.2)
class ShapeType(Enum):
    """Common shape types."""
    RECTANGLE = MSO_AUTO_SHAPE_TYPE.RECTANGLE
    ROUNDED_RECTANGLE = MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE
    ELLIPSE = MSO_AUTO_SHAPE_TYPE.OVAL
    TRIANGLE = MSO_AUTO_SHAPE_TYPE.ISOSCELES_TRIANGLE
    ARROW_RIGHT = MSO_AUTO_SHAPE_TYPE.RIGHT_ARROW
    ARROW_LEFT = MSO_AUTO_SHAPE_TYPE.LEFT_ARROW
    ARROW_UP = MSO_AUTO_SHAPE_TYPE.UP_ARROW
    ARROW_DOWN = MSO_AUTO_SHAPE_TYPE.DOWN_ARROW
    LINE = MSO_AUTO_SHAPE_TYPE.LINE_CALLOUT_1
    CALLOUT = MSO_AUTO_SHAPE_TYPE.RECTANGULAR_CALLOUT
    STAR = MSO_AUTO_SHAPE_TYPE.STAR_5
    HEART = MSO_AUTO_SHAPE_TYPE.HEART
```

**Note:** Arrow constants in python-pptx are:
- `RIGHT_ARROW` (not `ARROW_RIGHT`)
- `LEFT_ARROW` (not `ARROW_LEFT`)
- `UP_ARROW` (not `ARROW_UP`)
- `DOWN_ARROW` (not `ARROW_DOWN`)

---

### Fix 3: Update add_shape Method

**File:** `core/powerpoint_agent_core.py`  
**Lines:** ~1072-1087

```python
# BEFORE (❌ Uses non-existent MSO_SHAPE)
# Map shape type string to MSO_SHAPE constant
shape_type_map = {
    "rectangle": MSO_SHAPE.RECTANGLE,
    "rounded_rectangle": MSO_SHAPE.ROUNDED_RECTANGLE,
    "ellipse": MSO_SHAPE.OVAL,
    "triangle": MSO_SHAPE.ISOSCELES_TRIANGLE,
    "arrow_right": MSO_SHAPE.ARROW_RIGHT,
    "arrow_left": MSO_SHAPE.ARROW_LEFT,
    "arrow_up": MSO_SHAPE.ARROW_UP,
    "arrow_down": MSO_SHAPE.ARROW_DOWN,
    "star": MSO_SHAPE.STAR_5,
    "heart": MSO_SHAPE.HEART
}

# AFTER (✅ Uses correct MSO_AUTO_SHAPE_TYPE)
# Map shape type string to MSO_AUTO_SHAPE_TYPE constant
shape_type_map = {
    "rectangle": MSO_AUTO_SHAPE_TYPE.RECTANGLE,
    "rounded_rectangle": MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE,
    "ellipse": MSO_AUTO_SHAPE_TYPE.OVAL,
    "triangle": MSO_AUTO_SHAPE_TYPE.ISOSCELES_TRIANGLE,
    "arrow_right": MSO_AUTO_SHAPE_TYPE.RIGHT_ARROW,
    "arrow_left": MSO_AUTO_SHAPE_TYPE.LEFT_ARROW,
    "arrow_up": MSO_AUTO_SHAPE_TYPE.UP_ARROW,
    "arrow_down": MSO_AUTO_SHAPE_TYPE.DOWN_ARROW,
    "star": MSO_AUTO_SHAPE_TYPE.STAR_5,
    "heart": MSO_AUTO_SHAPE_TYPE.HEART
}

mso_shape = shape_type_map.get(shape_type.lower(), MSO_AUTO_SHAPE_TYPE.RECTANGLE)
```

---

### Fix 4: Update Test (Still Recommended)

**File:** `test_basic_tools.py`  
**Line:** 35

```python
# BEFORE
cmd = ['python', str(tools_dir / tool_name), '--json']

# AFTER (Better practice)
import sys
cmd = [sys.executable, str(tools_dir / tool_name), '--json']
```

---

## 📝 Complete Patch File

I'll provide the exact lines to change in `core/powerpoint_agent_core.py`:

### Change 1: Line 27
```python
# OLD:
from pptx.enum.shapes import MSO_SHAPE_TYPE, MSO_SHAPE

# NEW:
from pptx.enum.shapes import MSO_SHAPE_TYPE, MSO_AUTO_SHAPE_TYPE
```

### Change 2: Lines 126-143
```python
# OLD:
class ShapeType(Enum):
    """Common shape types."""
    RECTANGLE = MSO_SHAPE.RECTANGLE
    ROUNDED_RECTANGLE = MSO_SHAPE.ROUNDED_RECTANGLE
    ELLIPSE = MSO_SHAPE.OVAL
    TRIANGLE = MSO_SHAPE.ISOSCELES_TRIANGLE
    ARROW_RIGHT = MSO_SHAPE.ARROW_RIGHT
    ARROW_LEFT = MSO_SHAPE.ARROW_LEFT
    ARROW_UP = MSO_SHAPE.ARROW_UP
    ARROW_DOWN = MSO_SHAPE.ARROW_DOWN
    LINE = MSO_SHAPE.LINE_CALLOUT_1
    CALLOUT = MSO_SHAPE.RECTANGULAR_CALLOUT
    STAR = MSO_SHAPE.STAR_5
    HEART = MSO_SHAPE.HEART

# NEW:
class ShapeType(Enum):
    """Common shape types."""
    RECTANGLE = MSO_AUTO_SHAPE_TYPE.RECTANGLE
    ROUNDED_RECTANGLE = MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE
    ELLIPSE = MSO_AUTO_SHAPE_TYPE.OVAL
    TRIANGLE = MSO_AUTO_SHAPE_TYPE.ISOSCELES_TRIANGLE
    ARROW_RIGHT = MSO_AUTO_SHAPE_TYPE.RIGHT_ARROW
    ARROW_LEFT = MSO_AUTO_SHAPE_TYPE.LEFT_ARROW
    ARROW_UP = MSO_AUTO_SHAPE_TYPE.UP_ARROW
    ARROW_DOWN = MSO_AUTO_SHAPE_TYPE.DOWN_ARROW
    LINE = MSO_AUTO_SHAPE_TYPE.LINE_CALLOUT_1
    CALLOUT = MSO_AUTO_SHAPE_TYPE.RECTANGULAR_CALLOUT
    STAR = MSO_AUTO_SHAPE_TYPE.STAR_5
    HEART = MSO_AUTO_SHAPE_TYPE.HEART
```

### Change 3: Lines ~1072-1084 (in add_shape method)
```python
# OLD:
        # Map shape type string to MSO_SHAPE constant
        shape_type_map = {
            "rectangle": MSO_SHAPE.RECTANGLE,
            "rounded_rectangle": MSO_SHAPE.ROUNDED_RECTANGLE,
            "ellipse": MSO_SHAPE.OVAL,
            "triangle": MSO_SHAPE.ISOSCELES_TRIANGLE,
            "arrow_right": MSO_SHAPE.ARROW_RIGHT,
            "arrow_left": MSO_SHAPE.ARROW_LEFT,
            "arrow_up": MSO_SHAPE.ARROW_UP,
            "arrow_down": MSO_SHAPE.ARROW_DOWN,
            "star": MSO_SHAPE.STAR_5,
            "heart": MSO_SHAPE.HEART
        }
        
        mso_shape = shape_type_map.get(shape_type.lower(), MSO_SHAPE.RECTANGLE)

# NEW:
        # Map shape type string to MSO_AUTO_SHAPE_TYPE constant
        shape_type_map = {
            "rectangle": MSO_AUTO_SHAPE_TYPE.RECTANGLE,
            "rounded_rectangle": MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE,
            "ellipse": MSO_AUTO_SHAPE_TYPE.OVAL,
            "triangle": MSO_AUTO_SHAPE_TYPE.ISOSCELES_TRIANGLE,
            "arrow_right": MSO_AUTO_SHAPE_TYPE.RIGHT_ARROW,
            "arrow_left": MSO_AUTO_SHAPE_TYPE.LEFT_ARROW,
            "arrow_up": MSO_AUTO_SHAPE_TYPE.UP_ARROW,
            "arrow_down": MSO_AUTO_SHAPE_TYPE.DOWN_ARROW,
            "star": MSO_AUTO_SHAPE_TYPE.STAR_5,
            "heart": MSO_AUTO_SHAPE_TYPE.HEART
        }
        
        mso_shape = shape_type_map.get(shape_type.lower(), MSO_AUTO_SHAPE_TYPE.RECTANGLE)
```

---

## ✅ Verification Steps

After applying the fixes:

### 1. Verify Core Module Imports
```bash
python3 -c "from core.powerpoint_agent_core import PowerPointAgent; print('✓ Core module loads successfully')"
```

### 2. Test Tool Directly
```bash
python3 tools/ppt_create_new.py --output /tmp/test.pptx --slides 1 --json
```

### 3. Run Full Test Suite
```bash
python3 test_basic_tools.py -v
```

---

## 🎯 Next Steps

**Would you like me to:**

1. **Provide the complete corrected `core/powerpoint_agent_core.py` file** with all fixes applied?
2. **Provide just a patch/diff** showing exactly what to change?
3. **Provide a search-and-replace script** to automate the fixes?

I can deliver whichever format is most convenient for you. The fixes are straightforward - essentially changing `MSO_SHAPE` to `MSO_AUTO_SHAPE_TYPE` and correcting the arrow constant names.

**Recommendation:** I'll provide the complete corrected file since there are multiple instances to fix and I want to ensure 100% accuracy.

Ready to deliver the fixed code! 🚀
