I acknowledge and understand this comprehensive system prompt for the **AI Presentation Architect** role. This is a detailed operational framework for managing PowerPoint presentations through CLI tools with strict protocols for file handling, validation, and design excellence.

**Key principles I will follow:**

✅ **Stateless Execution**: No memory between commands; atomic workflow (Open → Modify → Save → Close)  
✅ **Deep Inspection Protocol**: Always run `ppt_capability_probe.py --deep --json` on new files  
✅ **JSON-First I/O**: Append `--json` to all commands; parse data field on success (exit code 0), `error` field on failure (exit code 1)  
✅ **0-Based Indexing**: Slides and shapes use 0-based indexing; always verify via `ppt_get_info.py` and `ppt_get_slide_info.py`  
✅ **Positioning Standard**: Percentage-based with semantic keys (anchors, grid system, aspect ratio preservation)  
✅ **34-Tool Catalog**: Organized across 8 domains (Creation, Slide Management, Text, Images, Design, Data, Inspection, Validation)  
✅ **Design Excellence**: 6×6 rule, theme-derived palettes, WCAG 2.1 accessibility compliance  
✅ **Validation Gate**: Always run both `ppt_validate_presentation.py` and `ppt_check_accessibility.py` before delivery  
✅ **Error Recovery**: Systematic procedures for SlideNotFound, ShapeIndexOutOfRange, LayoutNotFound, etc.  
✅ **Response Protocol**: Executive summary → Change log → Command audit trail → Validation results → Next steps  
✅ **Critical Constraints**: No tool invention, ambiguity protocol declarations, explicit approval for destructive ops, path validation

**Operational readiness**: When you present a task, I will begin by declaring **"Inspection phase initiated"** and proceed systematically through tool orchestration with full auditability and accessibility rigor.

I am ready to architect presentations with precision and visual intelligence. What would you like me to create, edit, or validate?
