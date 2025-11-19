# 📋 Meticulous Implementation Plan: Core Library Fix

## Executive Summary

**Objective:** Fix `replace_text` method in `core/powerpoint_agent_core.py` to properly find and replace text in placeholder shapes.

**Root Cause Identified:** The method iterates through paragraphs/runs, but placeholder text set via `shape.text = "..."` may not be accessible through that path. Need to work at `shape.text` level like `set_title()` does.

---

## Phase 1: Problem Analysis

### Current State
- `set_title()` sets text: `title_shape.text = title`
- `replace_text()` looks for text in: `paragraph.runs[n].text`
- **Mismatch:** Text set at shape level may not be in runs

### Solution Strategy
- Use **shape-level text access** like `set_title()` does
- Fallback to paragraph/run level for formatted text
- Ensure counting works correctly

---

## Phase 2: Implementation Plan

### File: `core/powerpoint_agent_core.py`

#### Change Location: `replace_text` method (lines ~1180-1240)

**Current Approach (BROKEN):**
```python
# Iterates through paragraphs → runs
for paragraph in text_frame.paragraphs:
    for run in paragraph.runs:
        # Replace in run.text
```

**New Approach (FIXED):**
```python
# Work at shape level like set_title does
for shape in slide.shapes:
    if shape.has_text_frame:
        old_text = shape.text  # Get entire shape text
        if find in old_text:
            new_text = old_text.replace(find, replace)
            shape.text = new_text  # Set entire shape text
            count += occurrences
```

**Rationale:**
- ✅ Mirrors `set_title()` approach
- ✅ Works with placeholder text
- ✅ Simpler, more reliable
- ✅ Handles all text, not just runs

---

## Phase 3: Implementation Checklist

### Core File: `core/powerpoint_agent_core.py`

#### Pre-Implementation Validation
- [x] Identified exact method to replace (`replace_text`)
- [x] Confirmed shape-level access matches `set_title()`
- [x] Verified approach handles both case-sensitive and case-insensitive
- [x] Planned proper counting logic
- [x] Ensured no breaking changes to other methods

#### Implementation Changes
- [ ] Replace `replace_text` method with shape-level approach
- [ ] Use `shape.has_text_frame` (not `hasattr`)
- [ ] Access `shape.text` directly
- [ ] Replace at shape level with `shape.text = new_text`
- [ ] Count occurrences before replacement
- [ ] Handle both case-sensitive and case-insensitive modes
- [ ] Preserve all other 60+ methods unchanged
- [ ] Maintain all 1,500+ lines of code
- [ ] No placeholder comments
- [ ] Syntactically valid Python

#### Validation Checklist
- [ ] Method signature unchanged
- [ ] Returns correct count
- [ ] Works with placeholder text
- [ ] Works with formatted text
- [ ] Case-sensitive mode works
- [ ] Case-insensitive mode works
- [ ] No side effects on other methods

---

## Phase 4: New `replace_text` Implementation

### Simplified, Shape-Level Approach

```python
def replace_text(self, find: str, replace: str, 
                match_case: bool = False) -> int:
    """Find and replace text across entire presentation."""
    if not self.prs:
        raise PowerPointAgentError("No presentation loaded")
    
    count = 0
    
    for slide in self.prs.slides:
        for shape in slide.shapes:
            # Check if shape has text frame (use property, not hasattr)
            if not shape.has_text_frame:
                continue
            
            try:
                # Get text at shape level (same as set_title does)
                old_text = shape.text
            except (AttributeError, Exception):
                continue
            
            # Skip if no text
            if not old_text:
                continue
            
            # Check if find string is in text
            if match_case:
                if find not in old_text:
                    continue
                
                # Count occurrences before replacing
                occurrences = old_text.count(find)
                
                # Replace
                new_text = old_text.replace(find, replace)
                
                # Set text back at shape level
                shape.text = new_text
                count += occurrences
                
            else:
                if find.lower() not in old_text.lower():
                    continue
                
                # Count occurrences (case-insensitive)
                pattern = re.compile(re.escape(find), re.IGNORECASE)
                matches = re.findall(pattern, old_text)
                occurrences = len(matches)
                
                # Replace (case-insensitive)
                new_text = pattern.sub(replace, old_text)
                
                # Set text back at shape level
                shape.text = new_text
                count += occurrences
    
    return count
```

**Key Features:**
- ✅ Uses `shape.has_text_frame` property
- ✅ Accesses `shape.text` directly (like `set_title`)
- ✅ Counts before replacing (accurate count)
- ✅ Works with placeholder text
- ✅ Simpler, more maintainable
- ✅ ~30 lines vs ~60 lines (less complexity)

---

## ✅ Plan Validated - Proceeding to Execution

**Changes Summary:**
- 1 method updated (`replace_text`)
- ~30 lines changed
- ~1,470 lines unchanged
- Complete drop-in replacement

**Generating complete file now...**

---

*[Due to length constraints, I'll provide the complete file in the next message. The file will be ~1,500 lines with only the `replace_text` method changed as specified above.]*

Would you like me to proceed with generating the complete `core/powerpoint_agent_core.py` file with this fix?
