Comprehensive Implementation Plan: core/powerpoint_agent_core.py v3.1.0
Pre-Implementation Validation
Before proceeding, I'll validate each proposed change against the original code to ensure no regressions.

Validation Matrix: Proposed Changes vs. Original Code
Change	Original Behavior	Proposed Change	Backward Compatible?	Risk Level
Add approval tokens	No validation	Optional param with enforcement	✅ Yes (optional param)	Low
Add version tracking	Not returned	Add to return dict	✅ Yes (additional fields)	Low
Fix _log_warning()	print to stderr	logger.warning()	✅ Yes (same output target)	Low
Fix errno handling	Magic number 17	errno.EEXIST constant	✅ Yes (same logic)	Low
Add path traversal protection	No protection	Optional allowed_base_dirs	✅ Yes (optional param)	Low
Remove duplicate lxml import	Redundant import	Remove inline import	✅ Yes (no behavior change)	None
Fix silent index clamping	Silent correction	Raise error	⚠️ Breaking	Medium
Fix mixed hash algorithms	MD5 + SHA-256	SHA-256 only	⚠️ Version hash changes	Medium
Use NotesMode enum	String validation	Enum + string support	✅ Yes (accepts both)	Low
Consolidate helper method	Duplicate code	Single implementation	✅ Yes (same logic)	Low
Fix silent shape fallback	Silent rectangle	Raise error	⚠️ Breaking	Medium
Breaking Changes Mitigation Strategy
For the 3 potentially breaking changes:

Silent index clamping → Add strict_validation parameter (default: True)
Version hash change → Acceptable (version is for change detection, not persistence)
Silent shape fallback → Add warning before raising error in v3.1, make error mandatory in v4.0
Implementation Plan with Section Checklists
Section 1: Module Header, Docstring & Changelog
Checklist:

 Update __version__ to "3.1.0"
 Add v3.1.0 changelog entry
 Add errno to imports
 Verify all existing imports preserved
 Verify module docstring complete
Section 2: Logging Setup
Checklist:

 Verify logger configuration preserved
 No changes needed
Section 3: Exceptions
Checklist:

 Preserve all 12 existing exceptions
 Add new ApprovalTokenError exception
 Verify to_dict() and to_json() methods preserved
Section 4: Constants
Checklist:

 Preserve all existing constants
 Add DESTRUCTIVE_OPERATIONS set
 Add APPROVAL_SCOPE_DELETE_SLIDE constant
 Add APPROVAL_SCOPE_REMOVE_SHAPE constant
Section 5: Enums
Checklist:

 Preserve all 10 existing enums
 Verify NotesMode enum present for later use
Section 6: Utility Classes
6.1 FileLock
Checklist:

 Add import errno at module level
 Replace e.errno == 17 with e.errno == errno.EEXIST
 Verify all existing methods preserved
 Verify context manager behavior preserved
6.2 PathValidator
Checklist:

 Add allowed_base_dirs parameter to validate_pptx_path()
 Add path traversal check logic
 Preserve all existing validation logic
 Verify backward compatibility (param is optional)
6.3 Position, Size, ColorHelper
Checklist:

 No changes - verify all methods preserved exactly
Section 7: Analysis Classes
7.1 TemplateProfile
Checklist:

 Verify all methods preserved exactly
 No changes needed
7.2 AccessibilityChecker
Checklist:

 Remove duplicate _get_placeholder_type_int() method
 Update to use module-level helper function
 Verify all other methods preserved
7.3 AssetValidator
Checklist:

 Verify all methods preserved exactly
 No changes needed
Section 8: PowerPointAgent Class
8.1 __init__ and Context Management
Checklist:

 Verify __init__ preserved exactly
 Verify __enter__ preserved exactly
 Verify __exit__ preserved exactly
8.2 File Operations
Checklist:

 create_new() - no changes needed
 open() - verify lock release on failure preserved
 save() - no changes needed
 close() - no changes needed
 clone_presentation() - no changes needed
8.3 Slide Operations
Checklist:

 add_slide() - add version tracking, fix index validation
 delete_slide() - add approval token, add version tracking
 duplicate_slide() - add version tracking
 reorder_slides() - add version tracking
 get_slide_count() - no changes needed
8.4 Text Operations
Checklist:

 add_text_box() - add version tracking
 set_title() - add version tracking
 add_bullet_list() - add version tracking
 format_text() - add version tracking
 replace_text() - add version tracking
 _replace_text_in_shape() - no changes needed (private helper)
 add_notes() - add version tracking, use NotesMode enum
8.5 Footer Operation
Checklist:

 set_footer() - add version tracking
8.6 Shape Operations
Checklist:

 _set_fill_opacity() - no changes needed
 _set_line_opacity() - no changes needed
 _ensure_line_solid_fill() - no changes needed
 _log_warning() - fix to use logger.warning()
 add_shape() - add version tracking, add shape type validation
 format_shape() - add version tracking
 remove_shape() - add approval token, add version tracking
 set_z_order() - add version tracking
 add_table() - add version tracking
 add_connector() - add version tracking
8.7 Image Operations
Checklist:

 insert_image() - add version tracking
 replace_image() - add version tracking
 set_image_properties() - add version tracking
 crop_image() - remove duplicate import, add version tracking
 resize_image() - add version tracking
8.8 Chart Operations
Checklist:

 add_chart() - add version tracking
 update_chart_data() - add version tracking
 format_chart() - add version tracking
8.9 Layout & Theme Operations
Checklist:

 set_slide_layout() - add version tracking
 set_background() - add version tracking
 get_available_layouts() - no changes needed
8.10 Validation Operations
Checklist:

 validate_presentation() - no changes needed
 check_accessibility() - no changes needed
 validate_assets() - no changes needed
8.11 Export Operations
Checklist:

 export_to_pdf() - no changes needed
 extract_notes() - no changes needed
8.12 Information & Versioning
Checklist:

 get_presentation_info() - no changes needed
 get_slide_info() - no changes needed
 get_presentation_version() - fix to use SHA-256 consistently
8.13 Private Helper Methods
Checklist:

 _get_slide() - no changes needed
 _get_shape() - no changes needed
 _get_chart_shape() - no changes needed
 _get_layout() - no changes needed
 _ensure_layout_cache() - no changes needed
 _get_placeholder_type_int() - consolidate as single implementation
 _copy_shape() - no changes needed
 Add new _validate_approval_token() helper method
 Add new _capture_version() helper method for DRY
Section 9: Module Exports
Checklist:

 Add ApprovalTokenError to exports
 Verify all existing exports preserved
 Verify __version__ updated
