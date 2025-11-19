# Phase 2: Top 10 Missing Tools Implementation Plan

Based on the comprehensive analysis of the current codebase and the Master Plan, here are the next 10 most critical tools required to complete the agent's capabilities. These tools focus on **editing existing content** (CRUD operations), **accessibility**, and **slide management**, bridging the gap between "creation-only" and "full manipulation."

### 1. `tools/ppt_check_accessibility.py`
**Priority:** P0 (Critical)
**Description:**
A dedicated tool to run WCAG 2.1 checks on the presentation. While `validate_presentation` includes this, a standalone tool allows for focused accessibility audits with detailed reporting on contrast ratios, reading order, and alt text deficiencies without the noise of asset size warnings.
**Features:**
*   Checks for missing Alt Text on all visual elements.
*   Analyzes text-to-background contrast ratios.
*   Verifies slide title existence and uniqueness.
*   Outputs specific shape indices for remediation.
**Validation Checklist:**
*   [ ] Run on a deck with missing alt text -> Returns errors.
*   [ ] Run on a deck with low contrast text -> Returns warnings with ratio values.
*   [ ] Verify output JSON structure contains `issues` list.

### 2. `tools/ppt_delete_slide.py`
**Priority:** P1 (High)
**Description:**
Allows the agent to remove specific slides by index. Essential for cleaning up templates or removing placeholder slides after content generation.
**Features:**
*   Deletes slide at specified 0-based index.
*   Handles edge cases (deleting last slide, out-of-bounds index).
*   Returns updated total slide count.
**Validation Checklist:**
*   [ ] Delete slide 0 -> Slide 1 becomes Slide 0.
*   [ ] Try to delete non-existent index -> Returns useful error.
*   [ ] Verify total slide count decreases by 1.

### 3. `tools/ppt_duplicate_slide.py`
**Priority:** P1 (High)
**Description:**
Clones an existing slide (including all shapes, styles, and layout) and appends it to the presentation. Crucial for repeating complex layouts that aren't standard master templates.
**Features:**
*   Duplicates source slide content.
*   Preserves layout type.
*   Returns index of the new slide.
**Validation Checklist:**
*   [ ] Duplicate slide with content -> New slide exists at end.
*   [ ] Verify new slide has same number of shapes as source.
*   [ ] Verify layout matches source.

### 4. `tools/ppt_reorder_slides.py`
**Priority:** P1 (High)
**Description:**
Moves a slide from one position to another. Enables the agent to restructure a presentation after generating individual sections.
**Features:**
*   Moves slide from `from_index` to `to_index`.
*   Shifts intermediate slides correctly.
**Validation Checklist:**
*   [ ] Move slide 0 to 1 -> Order swaps.
*   [ ] Move last slide to start -> Order shifts correctly.
*   [ ] Validate indices are within range.

### 5. `tools/ppt_set_slide_layout.py`
**Priority:** P1 (High)
**Description:**
Changes the underlying layout of an existing slide (e.g., switching from "Title Only" to "Title and Content").
**Features:**
*   Applies new layout by name.
*   Attempts to remap placeholders to the new layout.
**Validation Checklist:**
*   [ ] Change layout -> Slide layout property updates.
*   [ ] Verify content in placeholders is preserved (where possible).
*   [ ] Error handling for invalid layout names.

### 6. `tools/ppt_format_shape.py`
**Priority:** P1 (High)
**Description:**
Updates the styling of an existing shape. Complements `add_shape` by allowing edits to shapes that were part of a template or created previously.
**Features:**
*   Updates Fill Color.
*   Updates Line/Border Color and Width.
*   Target shape by index.
**Validation Checklist:**
*   [ ] Apply fill color -> JSON confirms change.
*   [ ] Apply border width -> JSON confirms change.
*   [ ] target invalid shape index -> Returns error.

### 7. `tools/ppt_set_image_properties.py`
**Priority:** P1 (High)
**Description:**
Specifically designed to remediate accessibility issues and style images. It allows setting Alt Text (critical for the accessibility workflow) and transparency.
**Features:**
*   Sets `descr` (Alt Text) property.
*   Sets transparency level (0.0 - 1.0).
*   Targets specific picture shapes.
**Validation Checklist:**
*   [ ] Set Alt Text -> Run `check_accessibility` to verify fix.
*   [ ] Set transparency -> JSON confirms value.
*   [ ] Verify it only accepts image shapes.

### 8. `tools/ppt_clone_presentation.py`
**Priority:** P1 (High)
**Description:**
Creates a full copy of a .pptx file. This enables "Safe Mode" workflows where the agent works on a copy to prevent data loss, or creates multiple variations of a base deck.
**Features:**
*   Copies source file to destination path.
*   Verifies write permissions.
**Validation Checklist:**
*   [ ] Clone valid file -> New file exists.
*   [ ] Verify file size matches source.
*   [ ] Verify new file opens correctly.

### 9. `tools/ppt_extract_notes.py`
**Priority:** P2 (Medium)
**Description:**
Extracts speaker notes from slides. Useful for generating scripts, summaries, or verifying that key talking points are included.
**Features:**
*   Extracts text from Notes page of each slide.
*   Returns JSON mapping `{slide_index: "Note text"}`.
**Validation Checklist:**
*   [ ] Extract from slide with notes -> Returns text.
*   [ ] Extract from slide without notes -> Returns null/empty.
*   [ ] Verify slide mapping is correct.

### 10. `tools/ppt_add_connector.py`
**Priority:** P2 (Medium)
**Description:**
Draws connecting lines between two shapes. Essential for creating flowcharts, process diagrams, and org charts programmatically.
**Features:**
*   Connects `from_shape` to `to_shape` by index.
*   Supports straight connectors (Elbow/Curved support depends on core library).
**Validation Checklist:**
*   [ ] Connect Shape 0 to Shape 1 -> Connector added.
*   [ ] Verify total shape count increases by 1.
*   [ ] Validate indices exist before connecting.

---

I will now proceed to generate the code for these 10 tools, following the established patterns (Argument Parsing, JSON output, Error Handling) and using the shared `core` library.
