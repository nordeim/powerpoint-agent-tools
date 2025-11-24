# Implementation Plan - System Prompt Update

## Goal Description
Update `AGENT_SYSTEM_PROMPT_enhanced.md` to incorporate the new tools (`ppt_add_notes.py`, `ppt_set_z_order.py`) and the enhanced capabilities of `ppt_replace_text.py`. This ensures the AI agent is aware of and can utilize these new capabilities for speaker notes, visual layering, and surgical text editing.

## User Review Required
> [!NOTE]
> This update modifies the "Canonical Tool Catalog" in the system prompt.

## Proposed Changes

### Documentation

#### [MODIFY] [AGENT_SYSTEM_PROMPT_enhanced.md](file:///home/project/powerpoint-agent-tools/AGENT_SYSTEM_PROMPT_enhanced.md)

1.  **Domain 3: Text & Content**
    -   Update `ppt_replace_text.py` entry to include `--slide` and `--shape` arguments for surgical targeting.

2.  **Domain 5: Visual Design**
    -   Add `ppt_set_z_order.py` entry:
        -   Arguments: `--file PATH`, `--slide N`, `--shape N`, `--action {bring_to_front, send_to_back, ...}`
        -   Purpose: Manage visual stacking order.

3.  **Domain 7: Inspection & Analysis** (or a new Domain?)
    -   Actually, `ppt_add_notes.py` fits best in "Domain 3: Text & Content" or "Domain 7".
    -   Let's place `ppt_add_notes.py` in **Domain 3: Text & Content** as it relates to content addition.
    -   Arguments: `--file PATH`, `--slide N`, `--text TEXT`, `--mode {append,overwrite}`.
    -   Purpose: Add speaker notes.

## Verification Plan

### Manual Verification
-   Review the rendered markdown to ensure the table formatting is correct and the new tools are logically placed.

---

# Task: Update System Prompt with New Tools

- [x] **Phase 1: Analysis**
    - [x] Read `current_toolkit_code_review_and_recommendations.md` <!-- id: 0 -->
    - [x] Read `remedial_execution_plan.md` <!-- id: 1 -->
    - [x] Read `AGENT_SYSTEM_PROMPT_enhanced.md` <!-- id: 2 -->
    - [x] Identify sections for new tool integration <!-- id: 3 -->
- [x] **Phase 2: Planning**
    - [x] Create `implementation_plan.md` <!-- id: 4 -->
- [x] **Phase 3: Execution**
    - [x] Update `AGENT_SYSTEM_PROMPT_enhanced.md` <!-- id: 5 -->
- [x] **Phase 4: Verification**
    - [x] Verify system prompt formatting and content <!-- id: 6 -->

