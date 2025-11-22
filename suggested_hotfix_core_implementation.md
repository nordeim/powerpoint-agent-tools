# 🚨 **CRITICAL BUG FOUND - IMMEDIATE FIX REQUIRED**

## **Issue Analysis:**

The error shows that `PP_PLACEHOLDER.CLIP_ART` (and likely other placeholder types) don't exist in the version of python-pptx being used. I was too comprehensive in the `PLACEHOLDER_TYPE_NAMES` mapping without verifying which constants actually exist.

## **Root Cause:**
Line 233 in `core/powerpoint_agent_core.py`:
```python
PP_PLACEHOLDER.CLIP_ART: "CLIP_ART",  # ← This constant doesn't exist!
```

## **Quick Fix:**

I need to provide a corrected `core/powerpoint_agent_core.py` with only the placeholder types that definitely exist in python-pptx.

---

# 📦 **CORRECTED FILE: core/powerpoint_agent_core.py (v1.1.1 - Hotfix)**

The issue is in the `PLACEHOLDER_TYPE_NAMES` mapping (around line 167-183). Here's the corrected section:

Replace the `PLACEHOLDER_TYPE_NAMES` section with this safe version:

```python
# Placeholder type mapping for human-readable output
# Only includes commonly available placeholder types in python-pptx
# Uses numeric keys as fallback for compatibility across versions
PLACEHOLDER_TYPE_NAMES = {
    1: "TITLE",                    # PP_PLACEHOLDER.TITLE
    2: "CONTENT",                  # PP_PLACEHOLDER.BODY  
    3: "CENTER_TITLE",            # PP_PLACEHOLDER.CENTER_TITLE
    4: "SUBTITLE",                 # PP_PLACEHOLDER.SUBTITLE (also FOOTER in some contexts)
    7: "OBJECT",                   # PP_PLACEHOLDER.OBJECT
    8: "CHART",                    # PP_PLACEHOLDER.CHART
    9: "TABLE",                    # PP_PLACEHOLDER.TABLE
    11: "ORG_CHART",              # PP_PLACEHOLDER.ORG_CHART (if available)
    13: "SLIDE_NUMBER",           # PP_PLACEHOLDER.SLIDE_NUMBER
    15: "HEADER",                 # PP_PLACEHOLDER.HEADER (if available)
    16: "DATE",                   # PP_PLACEHOLDER.DATE
    18: "PICTURE",                # PP_PLACEHOLDER.PICTURE
}

# Helper function to safely get placeholder type name
def get_placeholder_type_name(ph_type_value):
    """
    Safely get human-readable name for placeholder type.
    Falls back to numeric value if type unknown.
    """
    return PLACEHOLDER_TYPE_NAMES.get(ph_type_value, f"UNKNOWN_{ph_type_value}")
```

Then update the `get_slide_info()` method to use the helper function (around line 1315):

```python
# In get_slide_info() method, around line 1315:
if shape.is_placeholder:
    ph_type = shape.placeholder_format.type
    ph_type_name = get_placeholder_type_name(ph_type)  # Use helper instead of direct lookup
    shape_type_str = f"PLACEHOLDER ({ph_type_name})"
```

---

# 🔧 **COMPLETE CORRECTED core/powerpoint_agent_core.py**

Since this is a critical fix, here's the **complete corrected section** to replace in your existing file:

**Find this section (around line 167-183):**
```python
# Placeholder type mapping for human-readable output
# Source: python-pptx documentation and MSO constants
PLACEHOLDER_TYPE_NAMES = {
    PP_PLACEHOLDER.TITLE: "TITLE",                          # 1
    PP_PLACEHOLDER.BODY: "CONTENT",                         # 2
    PP_PLACEHOLDER.CENTER_TITLE: "CENTER_TITLE",           # 3
    PP_PLACEHOLDER.SUBTITLE: "SUBTITLE",                    # 4
    PP_PLACEHOLDER.DATE: "DATE",                           # 16
    PP_PLACEHOLDER.SLIDE_NUMBER: "SLIDE_NUMBER",           # 13
    PP_PLACEHOLDER.FOOTER: "FOOTER",                       # 4
    PP_PLACEHOLDER.HEADER: "HEADER",                       # 15
    PP_PLACEHOLDER.OBJECT: "OBJECT",                       # 7
    PP_PLACEHOLDER.CHART: "CHART",                         # 8
    PP_PLACEHOLDER.TABLE: "TABLE",                         # 9
    PP_PLACEHOLDER.CLIP_ART: "CLIP_ART",                  # 10  ← ERROR HERE
    PP_PLACEHOLDER.PICTURE: "PICTURE",                     # 18
    PP_PLACEHOLDER.MEDIA_CLIP: "MEDIA_CLIP",              # 16
    PP_PLACEHOLDER.ORG_CHART: "ORG_CHART",                # 11
}
```

**Replace with this safer version:**
```python
# Placeholder type mapping for human-readable output
# Uses numeric keys for maximum compatibility across python-pptx versions
# Only includes commonly available placeholder types
PLACEHOLDER_TYPE_NAMES = {
    1: "TITLE",           # Title placeholder
    2: "CONTENT",         # Body/content placeholder
    3: "CENTER_TITLE",    # Centered title (title slides)
    4: "SUBTITLE",        # Subtitle (also used for footer in some layouts)
    7: "OBJECT",          # Object placeholder
    8: "CHART",           # Chart placeholder
    9: "TABLE",           # Table placeholder
    13: "SLIDE_NUMBER",   # Slide number placeholder
    16: "DATE",           # Date placeholder
    18: "PICTURE",        # Picture placeholder
}

def get_placeholder_type_name(ph_type_value):
    """
    Safely get human-readable name for placeholder type.
    
    Args:
        ph_type_value: Numeric placeholder type value
        
    Returns:
        Human-readable string name or "UNKNOWN_X" if not recognized
    """
    return PLACEHOLDER_TYPE_NAMES.get(ph_type_value, f"UNKNOWN_{ph_type_value}")
```

**Then find this in `get_slide_info()` method (around line 1315-1318):**
```python
if shape.is_placeholder:
    ph_type = shape.placeholder_format.type
    ph_type_name = PLACEHOLDER_TYPE_NAMES.get(ph_type, f"UNKNOWN_{ph_type}")
    shape_type_str = f"PLACEHOLDER ({ph_type_name})"
```

**Replace with:**
```python
if shape.is_placeholder:
    ph_type = shape.placeholder_format.type
    ph_type_name = get_placeholder_type_name(ph_type)
    shape_type_str = f"PLACEHOLDER ({ph_type_name})"
```

---

# ⚡ **FASTEST FIX (Minimal Change)**

If you want the absolute quickest fix, just **delete or comment out the problematic lines** in the existing file:

**In `PLACEHOLDER_TYPE_NAMES` dict (around line 167-183), remove these lines:**
```python
    # PP_PLACEHOLDER.CLIP_ART: "CLIP_ART",                  # 10  ← DELETE THIS
    # PP_PLACEHOLDER.MEDIA_CLIP: "MEDIA_CLIP",              # 16  ← DELETE THIS
    # PP_PLACEHOLDER.ORG_CHART: "ORG_CHART",                # 11  ← DELETE THIS (may not exist)
```

And change the dict to use numeric keys instead:

```python
PLACEHOLDER_TYPE_NAMES = {
    1: "TITLE",           # PP_PLACEHOLDER.TITLE
    2: "CONTENT",         # PP_PLACEHOLDER.BODY
    3: "CENTER_TITLE",    # PP_PLACEHOLDER.CENTER_TITLE
    4: "SUBTITLE",        # PP_PLACEHOLDER.SUBTITLE
    7: "OBJECT",          # PP_PLACEHOLDER.OBJECT
    8: "CHART",           # PP_PLACEHOLDER.CHART
    9: "TABLE",           # PP_PLACEHOLDER.TABLE
    13: "SLIDE_NUMBER",   # PP_PLACEHOLDER.SLIDE_NUMBER
    15: "HEADER",         # PP_PLACEHOLDER.HEADER
    16: "DATE",           # PP_PLACEHOLDER.DATE
    18: "PICTURE",        # PP_PLACEHOLDER.PICTURE
}
```

---

# 🎯 **RECOMMENDED ACTION**

**Option 1: Quick Patch (2 minutes)**
1. Open `core/powerpoint_agent_core.py`
2. Go to line ~167 (the `PLACEHOLDER_TYPE_NAMES` dict)
3. Replace the entire dict with the numeric-key version above
4. Re-run the script

**Option 2: Proper Fix (5 minutes)**
1. Replace the `PLACEHOLDER_TYPE_NAMES` section with the version that includes the helper function
2. Update `get_slide_info()` to use the helper
3. Re-run the script

**Option 3: I can provide a complete corrected file**
- Let me know and I'll generate the entire `core/powerpoint_agent_core.py` v1.1.1 with this fix

---

# 🔍 **Why This Happened**

I was being overly comprehensive and included placeholder type constants that:
1. Don't exist in all python-pptx versions
2. May have been deprecated
3. Were incorrectly documented

The safe approach is to use **numeric keys** instead of the enum constants in the dict, which makes it version-agnostic.

---

**Which fix would you like me to provide?**
1. ✅ **Quick patch** (just the dict replacement)
2. ✅ **Proper fix** (dict + helper function)  
3. ✅ **Complete new file** (entire corrected core library)

Let me know and I'll provide it immediately!
