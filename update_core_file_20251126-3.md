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
Execution Plan
Now I will execute this plan section by section, producing the complete validated replacement file.

Complete Implementation: core/powerpoint_agent_core.py v3.1.0

```python
#!/usr/bin/env python3
"""
PowerPoint Agent Core Library v3.1.0
Production-grade PowerPoint manipulation with validation, accessibility, and full
alignment with Presentation Architect System Prompt v3.0.

This is the foundational library used by all CLI tools.
Designed for stateless, security-hardened PowerPoint operations.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Changelog v3.1.0 (Minor Release - Governance & Quality):
- NEW: ApprovalTokenError exception for governance enforcement
- NEW: Approval token validation for destructive operations (delete_slide, remove_shape)
- NEW: Version tracking (presentation_version_before/after) in all mutation methods
- NEW: Path traversal protection via allowed_base_dirs parameter
- NEW: _validate_approval_token() helper method
- NEW: _capture_version() helper method for DRY version tracking
- FIXED: _log_warning() now uses proper logger instead of print()
- FIXED: FileLock uses errno.EEXIST for cross-platform compatibility
- FIXED: crop_image() removed duplicate lxml import
- FIXED: get_presentation_version() now uses SHA-256 consistently (was MD5+SHA-256)
- FIXED: add_slide() now raises error for invalid index instead of silent clamping
- FIXED: add_shape() now validates shape_type and raises error for unknown types
- FIXED: add_notes() now accepts NotesMode enum in addition to strings
- IMPROVED: Consolidated duplicate _get_placeholder_type_int() implementations
- IMPROVED: All destructive operations require explicit approval_token parameter
- IMPROVED: All mutation methods return presentation_version_before and presentation_version_after

Changelog v3.0.0 (Major Release):
- NEW: add_notes() - Add/append/prepend/overwrite speaker notes
- NEW: set_z_order() - Control shape layering with 4 actions
- NEW: remove_shape() - Remove shapes from slides
- NEW: set_footer() - Configure footer text, numbers, date
- NEW: set_background() - Set slide/presentation background color or image
- NEW: crop_image() - True image cropping (not just resize)
- NEW: clone_presentation() - Clone presentation to new file
- NEW: get_presentation_version() - Compute deterministic version hash
- NEW: PathValidator class - Security-hardened path validation
- NEW: ShapeNotFoundError, ChartNotFoundError, PathValidationError exceptions
- NEW: ZOrderAction, NotesMode enums
- FIXED: FileLock now uses atomic os.open() with O_CREAT|O_EXCL
- FIXED: Lock released in finally block on open() failure
- FIXED: Slide insertion XML manipulation corrected
- FIXED: Placeholder type handling normalized to integers
- FIXED: Alt text detection checks 'descr' attribute
- FIXED: All bounds checks include negative index validation
- FIXED: Chart update error handling improved
- IMPROVED: All add_* methods return shape index for chaining
- IMPROVED: TemplateProfile uses lazy loading
- IMPROVED: Layout lookup cached for performance
- IMPROVED: Comprehensive docstrings with examples
- IMPROVED: Full alignment with System Prompt v3.0

Changelog v1.1.0:
- Added missing subprocess import for PDF export
- Added missing PP_PLACEHOLDER import and constants
- Replaced all magic numbers with named constants
- Removed text truncation in get_slide_info()
- Added position/size information to shape inspection
- Added placeholder subtype decoding
- Replaced print() with proper logging

Dependencies:
- python-pptx >= 0.6.21 (required)
- Pillow >= 9.0.0 (optional, for image operations)
"""

import os
import re
import sys
import json
import errno
import hashlib
import subprocess
import tempfile
import shutil
import time
import logging
import platform
from pathlib import Path
from typing import Any, Dict, List, Optional, Union, Tuple
from enum import Enum
from datetime import datetime
from io import BytesIO
from lxml import etree
from pptx.oxml.ns import qn

# qn() creates qualified names for XML namespace handling
# Example: qn('a:solidFill') -> '{http://schemas.openxmlformats.org/drawingml/2006/main}solidFill'

# ============================================================================
# THIRD-PARTY IMPORTS WITH GRACEFUL DEGRADATION
# ============================================================================

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu
    from pptx.enum.shapes import MSO_SHAPE_TYPE, MSO_AUTO_SHAPE_TYPE, MSO_CONNECTOR
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.enum.dml import MSO_THEME_COLOR
    from pptx.chart.data import CategoryChartData
    from pptx.dml.color import RGBColor
    PPTX_AVAILABLE = True
except ImportError:
    PPTX_AVAILABLE = False
    raise ImportError(
        "python-pptx is required. Install with:\n"
        "  pip install python-pptx\n"
        "  or: uv pip install python-pptx"
    )

try:
    from PIL import Image as PILImage
    HAS_PILLOW = True
except ImportError:
    HAS_PILLOW = False
    PILImage = None


# ============================================================================
# LOGGING SETUP
# ============================================================================

logger = logging.getLogger(__name__)
if not logger.handlers:
    handler = logging.StreamHandler()
    formatter = logging.Formatter('%(levelname)s:%(name)s:%(message)s')
    handler.setFormatter(formatter)
    logger.addHandler(handler)
    logger.setLevel(logging.WARNING)


# ============================================================================
# EXCEPTIONS
# ============================================================================

class PowerPointAgentError(Exception):
    """Base exception for all PowerPoint agent errors."""
    
    def __init__(self, message: str, details: Optional[Dict[str, Any]] = None):
        super().__init__(message)
        self.message = message
        self.details = details or {}
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert exception to JSON-serializable dict."""
        return {
            "error": self.__class__.__name__,
            "message": self.message,
            "details": self.details
        }
    
    def to_json(self) -> str:
        """Convert exception to JSON string."""
        return json.dumps(self.to_dict())


class SlideNotFoundError(PowerPointAgentError):
    """Raised when slide index is out of range."""
    pass


class ShapeNotFoundError(PowerPointAgentError):
    """Raised when shape index is out of range."""
    pass


class ChartNotFoundError(PowerPointAgentError):
    """Raised when chart is not found at specified index."""
    pass


class LayoutNotFoundError(PowerPointAgentError):
    """Raised when requested layout doesn't exist."""
    pass


class ImageNotFoundError(PowerPointAgentError):
    """Raised when image file is not found."""
    pass


class InvalidPositionError(PowerPointAgentError):
    """Raised when position specification is invalid."""
    pass


class TemplateError(PowerPointAgentError):
    """Raised when template operations fail."""
    pass


class ThemeError(PowerPointAgentError):
    """Raised when theme operations fail."""
    pass


class AccessibilityError(PowerPointAgentError):
    """Raised when accessibility validation fails."""
    pass


class AssetValidationError(PowerPointAgentError):
    """Raised when asset validation fails."""
    pass


class FileLockError(PowerPointAgentError):
    """Raised when file cannot be locked for exclusive access."""
    pass


class PathValidationError(PowerPointAgentError):
    """Raised when path validation fails (security)."""
    pass


class ApprovalTokenError(PowerPointAgentError):
    """Raised when approval token is missing or invalid for destructive operations."""
    pass


# ============================================================================
# CONSTANTS
# ============================================================================

__version__ = "3.1.0"
__author__ = "PowerPoint Agent Team"
__license__ = "MIT"

# Standard slide dimensions (16:9 widescreen) in inches
SLIDE_WIDTH_INCHES = 13.333
SLIDE_HEIGHT_INCHES = 7.5

# Alternative dimensions (4:3 standard) in inches
SLIDE_WIDTH_4_3_INCHES = 10.0
SLIDE_HEIGHT_4_3_INCHES = 7.5

# EMU conversion constant
EMU_PER_INCH = 914400

# Standard anchor points for positioning
ANCHOR_POINTS = {
    "top_left": (0.0, 0.0),
    "top_center": (0.5, 0.0),
    "top_right": (1.0, 0.0),
    "center_left": (0.0, 0.5),
    "center": (0.5, 0.5),
    "center_right": (1.0, 0.5),
    "bottom_left": (0.0, 1.0),
    "bottom_center": (0.5, 1.0),
    "bottom_right": (1.0, 1.0)
}

# Standard corporate colors (RGB tuples)
CORPORATE_COLORS = {
    "primary_blue": RGBColor(0, 112, 192),
    "secondary_gray": RGBColor(89, 89, 89),
    "accent_orange": RGBColor(237, 125, 49),
    "success_green": RGBColor(112, 173, 71),
    "warning_yellow": RGBColor(255, 192, 0),
    "danger_red": RGBColor(192, 0, 0),
    "white": RGBColor(255, 255, 255),
    "black": RGBColor(0, 0, 0)
}

# Standard fonts
STANDARD_FONTS = {
    "title": "Calibri Light",
    "body": "Calibri",
    "code": "Consolas"
}

# WCAG 2.1 color contrast ratios
WCAG_CONTRAST_NORMAL = 4.5
WCAG_CONTRAST_LARGE = 3.0

# Maximum recommended file size (MB)
MAX_RECOMMENDED_FILE_SIZE_MB = 50

# Valid PowerPoint extensions
VALID_PPTX_EXTENSIONS = {'.pptx', '.pptm', '.potx', '.potm'}

# Approval token scopes for destructive operations
APPROVAL_SCOPE_DELETE_SLIDE = "delete:slide"
APPROVAL_SCOPE_REMOVE_SHAPE = "remove:shape"
APPROVAL_SCOPE_REPLACE_ALL = "replace:all"

# Set of operations requiring approval tokens
DESTRUCTIVE_OPERATIONS = {
    "delete_slide",
    "remove_shape",
}

# Placeholder type mapping (integer keys for compatibility)
PLACEHOLDER_TYPE_NAMES = {
    0: "OBJECT",
    1: "TITLE",
    2: "BODY",
    3: "CENTER_TITLE",
    4: "SUBTITLE",
    5: "DATE",
    6: "SLIDE_NUMBER",
    7: "FOOTER",
    8: "HEADER",
    9: "OBJECT",
    10: "CHART",
    11: "TABLE",
    12: "CLIP_ART",
    13: "ORG_CHART",
    14: "MEDIA_CLIP",
    15: "BITMAP",
    16: "VERTICAL_TITLE",
    17: "VERTICAL_BODY",
    18: "PICTURE",
}

# Placeholder types that represent titles
TITLE_PLACEHOLDER_TYPES = {1, 3}  # TITLE and CENTER_TITLE

# Placeholder type for subtitle
SUBTITLE_PLACEHOLDER_TYPE = 4


def get_placeholder_type_int(ph_type: Any) -> int:
    """
    Convert placeholder type to integer safely.
    
    This is a module-level helper function used by multiple classes.
    
    Args:
        ph_type: Placeholder type (int, enum, or other)
        
    Returns:
        Integer representation of placeholder type
    """
    if ph_type is None:
        return 0
    if hasattr(ph_type, 'value'):
        return ph_type.value
    try:
        return int(ph_type)
    except (TypeError, ValueError):
        return 0


def get_placeholder_type_name(ph_type_value: Any) -> str:
    """
    Safely get human-readable name for placeholder type.
    
    Args:
        ph_type_value: Placeholder type (int or enum)
        
    Returns:
        Human-readable string name
    """
    if ph_type_value is None:
        return "NONE"
    
    int_value = get_placeholder_type_int(ph_type_value)
    return PLACEHOLDER_TYPE_NAMES.get(int_value, f"UNKNOWN_{int_value}")


# ============================================================================
# ENUMS
# ============================================================================

class ShapeType(Enum):
    """Common shape types supported by python-pptx."""
    RECTANGLE = "rectangle"
    ROUNDED_RECTANGLE = "rounded_rectangle"
    ELLIPSE = "ellipse"
    OVAL = "ellipse"
    TRIANGLE = "triangle"
    ARROW_RIGHT = "arrow_right"
    ARROW_LEFT = "arrow_left"
    ARROW_UP = "arrow_up"
    ARROW_DOWN = "arrow_down"
    STAR = "star"
    PENTAGON = "pentagon"
    HEXAGON = "hexagon"


class ChartType(Enum):
    """Supported chart types."""
    COLUMN = "column"
    COLUMN_CLUSTERED = "column"
    COLUMN_STACKED = "column_stacked"
    BAR = "bar"
    BAR_CLUSTERED = "bar"
    BAR_STACKED = "bar_stacked"
    LINE = "line"
    LINE_MARKERS = "line_markers"
    PIE = "pie"
    PIE_EXPLODED = "pie_exploded"
    AREA = "area"
    SCATTER = "scatter"


class TextAlignment(Enum):
    """Text alignment options."""
    LEFT = "left"
    CENTER = "center"
    RIGHT = "right"
    JUSTIFY = "justify"


class VerticalAlignment(Enum):
    """Vertical text alignment."""
    TOP = "top"
    MIDDLE = "middle"
    BOTTOM = "bottom"


class BulletStyle(Enum):
    """Bullet list styles."""
    BULLET = "bullet"
    NUMBERED = "numbered"
    NONE = "none"


class ImageFormat(Enum):
    """Supported image formats."""
    PNG = "png"
    JPG = "jpg"
    JPEG = "jpeg"
    GIF = "gif"
    BMP = "bmp"


class ExportFormat(Enum):
    """Export format options."""
    PDF = "pdf"
    PNG = "png"
    JPG = "jpg"
    PPTX = "pptx"


class ZOrderAction(Enum):
    """Z-order manipulation actions."""
    BRING_TO_FRONT = "bring_to_front"
    SEND_TO_BACK = "send_to_back"
    BRING_FORWARD = "bring_forward"
    SEND_BACKWARD = "send_backward"


class NotesMode(Enum):
    """Speaker notes insertion modes."""
    APPEND = "append"
    PREPEND = "prepend"
    OVERWRITE = "overwrite"


# ============================================================================
# UTILITY CLASSES
# ============================================================================

class FileLock:
    """
    Atomic file locking mechanism for concurrent access prevention.
    
    Uses OS-level atomic file creation to ensure only one process
    can hold the lock at a time.
    """
    
    def __init__(self, filepath: Path, timeout: float = 10.0):
        """
        Initialize file lock.
        
        Args:
            filepath: Path to file to lock
            timeout: Maximum seconds to wait for lock acquisition
        """
        self.filepath = Path(filepath)
        self.lockfile = self.filepath.parent / f".{self.filepath.name}.lock"
        self.timeout = timeout
        self.acquired = False
        self._fd: Optional[int] = None
    
    def acquire(self) -> bool:
        """
        Acquire lock with timeout using atomic file creation.
        
        Returns:
            True if lock acquired, False if timeout
        """
        start_time = time.time()
        
        while time.time() - start_time < self.timeout:
            try:
                # Use O_CREAT | O_EXCL for atomic creation
                # This is atomic on POSIX systems
                self._fd = os.open(
                    str(self.lockfile),
                    os.O_CREAT | os.O_EXCL | os.O_WRONLY,
                    0o644
                )
                self.acquired = True
                return True
            except FileExistsError:
                time.sleep(0.1)
            except OSError as e:
                # Use errno constant for cross-platform compatibility
                if e.errno == errno.EEXIST:
                    time.sleep(0.1)
                else:
                    raise
        
        return False
    
    def release(self) -> None:
        """Release lock and clean up lock file."""
        if self._fd is not None:
            try:
                os.close(self._fd)
            except OSError:
                pass
            self._fd = None
        
        if self.acquired:
            try:
                self.lockfile.unlink(missing_ok=True)
            except OSError:
                pass
            self.acquired = False
    
    def __enter__(self) -> 'FileLock':
        if not self.acquire():
            raise FileLockError(
                f"Could not acquire lock on {self.filepath} within {self.timeout}s",
                details={"filepath": str(self.filepath), "timeout": self.timeout}
            )
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb) -> bool:
        self.release()
        return False


class PathValidator:
    """
    Security-hardened path validation utility.
    
    Validates file paths to prevent path traversal attacks
    and ensure files are of expected types.
    """
    
    @staticmethod
    def validate_pptx_path(
        filepath: Union[str, Path],
        must_exist: bool = True,
        must_be_writable: bool = False,
        allowed_base_dirs: Optional[List[Path]] = None
    ) -> Path:
        """
        Validate a PowerPoint file path.
        
        Args:
            filepath: Path to validate
            must_exist: If True, file must exist
            must_be_writable: If True, parent directory must be writable
            allowed_base_dirs: Optional list of allowed base directories.
                              If provided, path must be within one of these directories.
            
        Returns:
            Resolved absolute Path
            
        Raises:
            PathValidationError: If validation fails
        """
        try:
            path = Path(filepath).resolve()
        except Exception as e:
            raise PathValidationError(
                f"Invalid path: {filepath}",
                details={"error": str(e)}
            )
        
        # Path traversal protection
        if allowed_base_dirs:
            is_within_allowed = False
            for base_dir in allowed_base_dirs:
                try:
                    resolved_base = Path(base_dir).resolve()
                    # Check if path is relative to (within) the base directory
                    path.relative_to(resolved_base)
                    is_within_allowed = True
                    break
                except ValueError:
                    # path.relative_to raises ValueError if not relative
                    continue
            
            if not is_within_allowed:
                raise PathValidationError(
                    f"Path escapes allowed directories: {path}",
                    details={
                        "path": str(path),
                        "allowed_dirs": [str(d) for d in allowed_base_dirs]
                    }
                )
        
        # Check extension
        if path.suffix.lower() not in VALID_PPTX_EXTENSIONS:
            raise PathValidationError(
                f"Invalid file extension: {path.suffix}",
                details={
                    "path": str(path),
                    "valid_extensions": list(VALID_PPTX_EXTENSIONS)
                }
            )
        
        # Check existence
        if must_exist and not path.exists():
            raise PathValidationError(
                f"File does not exist: {path}",
                details={"path": str(path)}
            )
        
        # Check if it's a file (not directory)
        if must_exist and not path.is_file():
            raise PathValidationError(
                f"Path is not a file: {path}",
                details={"path": str(path)}
            )
        
        # Check writability
        if must_be_writable:
            parent = path.parent
            if not parent.exists():
                raise PathValidationError(
                    f"Parent directory does not exist: {parent}",
                    details={"path": str(path), "parent": str(parent)}
                )
            if not os.access(str(parent), os.W_OK):
                raise PathValidationError(
                    f"Parent directory is not writable: {parent}",
                    details={"path": str(path), "parent": str(parent)}
                )
        
        return path
    
    @staticmethod
    def validate_image_path(filepath: Union[str, Path]) -> Path:
        """
        Validate an image file path.
        
        Args:
            filepath: Path to validate
            
        Returns:
            Resolved absolute Path
            
        Raises:
            ImageNotFoundError: If validation fails
        """
        try:
            path = Path(filepath).resolve()
        except Exception as e:
            raise ImageNotFoundError(
                f"Invalid image path: {filepath}",
                details={"error": str(e)}
            )
        
        if not path.exists():
            raise ImageNotFoundError(
                f"Image file does not exist: {path}",
                details={"path": str(path)}
            )
        
        if not path.is_file():
            raise ImageNotFoundError(
                f"Image path is not a file: {path}",
                details={"path": str(path)}
            )
        
        valid_image_extensions = {'.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.webp'}
        if path.suffix.lower() not in valid_image_extensions:
            raise ImageNotFoundError(
                f"Invalid image extension: {path.suffix}",
                details={"path": str(path), "valid_extensions": list(valid_image_extensions)}
            )
        
        return path


class Position:
    """Flexible position system supporting multiple input formats."""
    
    @staticmethod
    def from_dict(
        pos_dict: Dict[str, Any],
        slide_width: float = SLIDE_WIDTH_INCHES,
        slide_height: float = SLIDE_HEIGHT_INCHES
    ) -> Tuple[float, float]:
        """
        Convert position dict to (left, top) in inches.
        
        Supports multiple formats:
        1. Absolute inches: {"left": 1.5, "top": 2.0}
        2. Percentage: {"left": "20%", "top": "30%"}
        3. Anchor-based: {"anchor": "center", "offset_x": 0.5, "offset_y": -1.0}
        4. Grid system: {"grid_row": 2, "grid_col": 3, "grid_size": 12}
        
        Args:
            pos_dict: Position specification dictionary
            slide_width: Slide width in inches (for percentage calculations)
            slide_height: Slide height in inches (for percentage calculations)
            
        Returns:
            Tuple of (left, top) in inches
            
        Raises:
            InvalidPositionError: If format is invalid
        """
        if not isinstance(pos_dict, dict):
            raise InvalidPositionError(
                f"Position must be a dictionary, got {type(pos_dict).__name__}",
                details={"value": str(pos_dict)}
            )
        
        # Format 1 & 2: Absolute or percentage with left/top
        if "left" in pos_dict and "top" in pos_dict:
            left = Position._parse_dimension(pos_dict["left"], slide_width)
            top = Position._parse_dimension(pos_dict["top"], slide_height)
            return (left, top)
        
        # Format 3: Anchor-based
        if "anchor" in pos_dict:
            anchor_name = pos_dict["anchor"].lower().replace("-", "_").replace(" ", "_")
            anchor = ANCHOR_POINTS.get(anchor_name)
            
            if anchor is None:
                raise InvalidPositionError(
                    f"Unknown anchor: {pos_dict['anchor']}",
                    details={"available_anchors": list(ANCHOR_POINTS.keys())}
                )
            
            # Anchor is in relative coordinates (0-1), convert to inches
            base_left = anchor[0] * slide_width
            base_top = anchor[1] * slide_height
            
            offset_x = float(pos_dict.get("offset_x", 0))
            offset_y = float(pos_dict.get("offset_y", 0))
            
            return (base_left + offset_x, base_top + offset_y)
        
        # Format 4: Grid system
        if "grid_row" in pos_dict and "grid_col" in pos_dict:
            grid_size = int(pos_dict.get("grid_size", 12))
            cell_width = slide_width / grid_size
            cell_height = slide_height / grid_size
            
            col = int(pos_dict["grid_col"])
            row = int(pos_dict["grid_row"])
            
            left = col * cell_width
            top = row * cell_height
            
            return (left, top)
        
        raise InvalidPositionError(
            "Invalid position format",
            details={
                "provided": pos_dict,
                "expected_formats": [
                    {"left": "value", "top": "value"},
                    {"anchor": "center", "offset_x": 0, "offset_y": 0},
                    {"grid_row": 0, "grid_col": 0, "grid_size": 12}
                ]
            }
        )
    
    @staticmethod
    def _parse_dimension(value: Union[str, float, int], max_dimension: float) -> float:
        """
        Parse dimension value (supports percentages or absolute values).
        
        Args:
            value: Dimension value (e.g., "50%", 2.5, "2.5")
            max_dimension: Maximum dimension for percentage calculation
            
        Returns:
            Dimension in inches
        """
        if isinstance(value, str):
            value = value.strip()
            if value.endswith('%'):
                percent = float(value[:-1]) / 100.0
                return percent * max_dimension
            else:
                return float(value)
        return float(value)


class Size:
    """Flexible size system supporting multiple input formats."""
    
    @staticmethod
    def from_dict(
        size_dict: Dict[str, Any],
        slide_width: float = SLIDE_WIDTH_INCHES,
        slide_height: float = SLIDE_HEIGHT_INCHES,
        aspect_ratio: Optional[float] = None
    ) -> Tuple[Optional[float], Optional[float]]:
        """
        Convert size dict to (width, height) in inches.
        
        Supports:
        - {"width": 5.0, "height": 3.0}  # Absolute inches
        - {"width": "50%", "height": "30%"}  # Percentage of slide
        - {"width": "auto", "height": 3.0}  # Maintain aspect ratio
        - {"width": 5.0, "height": "auto"}  # Maintain aspect ratio
        
        Args:
            size_dict: Size specification dictionary
            slide_width: Slide width in inches
            slide_height: Slide height in inches
            aspect_ratio: Optional aspect ratio (width/height) for "auto" calculations
            
        Returns:
            Tuple of (width, height) in inches, either can be None for "auto"
        """
        if not isinstance(size_dict, dict):
            raise ValueError(f"Size must be a dictionary, got {type(size_dict).__name__}")
        
        if "width" not in size_dict and "height" not in size_dict:
            raise ValueError("Size must have at least 'width' or 'height'")
        
        width_spec = size_dict.get("width")
        height_spec = size_dict.get("height")
        
        # Parse width
        if width_spec == "auto" or width_spec is None:
            width = None
        else:
            width = Position._parse_dimension(width_spec, slide_width)
        
        # Parse height
        if height_spec == "auto" or height_spec is None:
            height = None
        else:
            height = Position._parse_dimension(height_spec, slide_height)
        
        # Apply aspect ratio if one dimension is auto
        if aspect_ratio is not None:
            if width is None and height is not None:
                width = height * aspect_ratio
            elif height is None and width is not None:
                height = width / aspect_ratio
        
        return (width, height)


class ColorHelper:
    """Utilities for color conversion and validation."""
    
    @staticmethod
    def from_hex(hex_color: str) -> RGBColor:
        """
        Convert hex color string to RGBColor.
        
        Args:
            hex_color: Hex color string (e.g., "#FF0000" or "FF0000")
            
        Returns:
            RGBColor object
            
        Raises:
            ValueError: If hex color format is invalid
        """
        hex_color = hex_color.strip().lstrip('#')
        
        if len(hex_color) != 6:
            raise ValueError(f"Invalid hex color: {hex_color}. Must be 6 hex digits.")
        
        if not all(c in '0123456789ABCDEFabcdef' for c in hex_color):
            raise ValueError(f"Invalid hex color: {hex_color}. Contains non-hex characters.")
        
        r = int(hex_color[0:2], 16)
        g = int(hex_color[2:4], 16)
        b = int(hex_color[4:6], 16)
        
        return RGBColor(r, g, b)
    
    @staticmethod
    def to_hex(rgb_color: RGBColor) -> str:
        """
        Convert RGBColor to hex string.
        
        Args:
            rgb_color: RGBColor object
            
        Returns:
            Hex color string with # prefix
        """
        if hasattr(rgb_color, '__iter__') and len(rgb_color) == 3:
            r, g, b = rgb_color
        elif hasattr(rgb_color, 'r'):
            r, g, b = rgb_color.r, rgb_color.g, rgb_color.b
        else:
            # Handle string representation
            hex_str = str(rgb_color).lstrip('#')
            return f"#{hex_str}"
        
        return f"#{r:02x}{g:02x}{b:02x}"
    
    @staticmethod
    def luminance(rgb_color: Union[RGBColor, Tuple[int, int, int]]) -> float:
        """
        Calculate relative luminance for WCAG contrast calculations.
        
        Args:
            rgb_color: RGBColor or (r, g, b) tuple
            
        Returns:
            Relative luminance value (0.0 to 1.0)
        """
        # Extract RGB values
        if hasattr(rgb_color, 'r'):
            r, g, b = rgb_color.r, rgb_color.g, rgb_color.b
        elif hasattr(rgb_color, '__iter__'):
            r, g, b = rgb_color
        else:
            # Handle string representation
            hex_str = str(rgb_color).lstrip('#')
            if len(hex_str) == 6:
                r = int(hex_str[0:2], 16)
                g = int(hex_str[2:4], 16)
                b = int(hex_str[4:6], 16)
            else:
                raise ValueError(f"Cannot parse color: {rgb_color}")
        
        def _linearize(channel: int) -> float:
            c = channel / 255.0
            if c <= 0.03928:
                return c / 12.92
            return ((c + 0.055) / 1.055) ** 2.4
        
        r_lin = _linearize(r)
        g_lin = _linearize(g)
        b_lin = _linearize(b)
        
        return 0.2126 * r_lin + 0.7152 * g_lin + 0.0722 * b_lin
    
    @staticmethod
    def contrast_ratio(color1: RGBColor, color2: RGBColor) -> float:
        """
        Calculate WCAG contrast ratio between two colors.
        
        Args:
            color1: First color
            color2: Second color
            
        Returns:
            Contrast ratio (1.0 to 21.0)
        """
        lum1 = ColorHelper.luminance(color1)
        lum2 = ColorHelper.luminance(color2)
        
        lighter = max(lum1, lum2)
        darker = min(lum1, lum2)
        
        return (lighter + 0.05) / (darker + 0.05)
    
    @staticmethod
    def meets_wcag(
        foreground: RGBColor,
        background: RGBColor,
        is_large_text: bool = False
    ) -> bool:
        """
        Check if color combination meets WCAG 2.1 AA standards.
        
        Args:
            foreground: Text/foreground color
            background: Background color
            is_large_text: True if text is 18pt+ or 14pt+ bold
            
        Returns:
            True if contrast is sufficient
        """
        ratio = ColorHelper.contrast_ratio(foreground, background)
        threshold = WCAG_CONTRAST_LARGE if is_large_text else WCAG_CONTRAST_NORMAL
        return ratio >= threshold


# ============================================================================
# ANALYSIS CLASSES
# ============================================================================

class TemplateProfile:
    """
    Captures and provides access to PowerPoint template formatting.
    
    Uses lazy loading to avoid performance penalty when profile is not needed.
    """
    
    def __init__(self, prs: Optional['Presentation'] = None):
        """
        Initialize template profile.
        
        Args:
            prs: Optional Presentation to analyze immediately
        """
        self._prs = prs
        self._captured = False
        self._slide_layouts: List[Dict[str, Any]] = []
        self._theme_colors: Dict[str, str] = {}
        self._theme_fonts: Dict[str, str] = {}
    
    def _ensure_captured(self) -> None:
        """Ensure template data has been captured (lazy loading)."""
        if self._captured or self._prs is None:
            return
        
        self._capture_layouts()
        self._capture_theme()
        self._captured = True
    
    def _capture_layouts(self) -> None:
        """Capture layout information from presentation."""
        for layout in self._prs.slide_layouts:
            layout_info = {
                "name": layout.name,
                "placeholders": []
            }
            
            for ph in layout.placeholders:
                try:
                    ph_info = {
                        "type": get_placeholder_type_int(ph.placeholder_format.type),
                        "idx": ph.placeholder_format.idx
                    }
                    if hasattr(ph, 'left') and ph.left is not None:
                        ph_info["position"] = {
                            "left": ph.left / EMU_PER_INCH,
                            "top": ph.top / EMU_PER_INCH
                        }
                    if hasattr(ph, 'width') and ph.width is not None:
                        ph_info["size"] = {
                            "width": ph.width / EMU_PER_INCH,
                            "height": ph.height / EMU_PER_INCH
                        }
                    layout_info["placeholders"].append(ph_info)
                except Exception:
                    continue
            
            self._slide_layouts.append(layout_info)
    
    def _capture_theme(self) -> None:
        """Capture theme colors and fonts from presentation."""
        try:
            # Attempt to extract theme colors
            if hasattr(self._prs, 'slide_master') and self._prs.slide_master:
                master = self._prs.slide_master
                
                # Extract fonts from shapes
                for shape in master.shapes:
                    if hasattr(shape, 'text_frame'):
                        try:
                            for para in shape.text_frame.paragraphs:
                                if para.font.name:
                                    font_key = f"font_{len(self._theme_fonts)}"
                                    if para.font.name not in self._theme_fonts.values():
                                        self._theme_fonts[font_key] = para.font.name
                        except Exception:
                            continue
        except Exception:
            pass
    
    @property
    def slide_layouts(self) -> List[Dict[str, Any]]:
        """Get slide layout information."""
        self._ensure_captured()
        return self._slide_layouts
    
    @property
    def theme_colors(self) -> Dict[str, str]:
        """Get theme colors."""
        self._ensure_captured()
        return self._theme_colors
    
    @property
    def theme_fonts(self) -> Dict[str, str]:
        """Get theme fonts."""
        self._ensure_captured()
        return self._theme_fonts
    
    def get_layout_names(self) -> List[str]:
        """Get list of available layout names."""
        self._ensure_captured()
        return [layout["name"] for layout in self._slide_layouts]
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert profile to JSON-serializable dict."""
        self._ensure_captured()
        return {
            "slide_layouts": self._slide_layouts,
            "theme_colors": self._theme_colors,
            "theme_fonts": self._theme_fonts
        }


class AccessibilityChecker:
    """WCAG 2.1 compliance checker for presentations."""
    
    @staticmethod
    def check_presentation(prs: 'Presentation') -> Dict[str, Any]:
        """
        Comprehensive accessibility check.
        
        Args:
            prs: Presentation to check
            
        Returns:
            Dict containing:
            - status: "accessible" or "issues_found"
            - total_issues: Count of all issues
            - issues: Detailed issue breakdown
            - wcag_level: "AA" if passing, "fail" otherwise
        """
        issues = {
            "missing_alt_text": [],
            "low_contrast": [],
            "missing_titles": [],
            "small_text": [],
            "reading_order_warnings": []
        }
        
        for slide_idx, slide in enumerate(prs.slides):
            # Check for title
            has_title = AccessibilityChecker._check_slide_has_title(slide)
            if not has_title:
                issues["missing_titles"].append({
                    "slide": slide_idx,
                    "message": "Slide lacks a title for screen reader navigation"
                })
            
            # Check each shape
            for shape_idx, shape in enumerate(slide.shapes):
                # Check images for alt text
                if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    if not AccessibilityChecker._has_alt_text(shape):
                        issues["missing_alt_text"].append({
                            "slide": slide_idx,
                            "shape": shape_idx,
                            "shape_name": shape.name,
                            "message": "Image lacks alternative text"
                        })
                
                # Check text for contrast and size
                if hasattr(shape, 'text_frame') and shape.has_text_frame:
                    AccessibilityChecker._check_text_accessibility(
                        shape, slide_idx, shape_idx, issues
                    )
        
        total_issues = sum(len(v) for v in issues.values())
        
        return {
            "status": "issues_found" if total_issues > 0 else "accessible",
            "total_issues": total_issues,
            "issues": issues,
            "wcag_level": "AA" if total_issues == 0 else "fail",
            "checked_slides": len(prs.slides)
        }
    
    @staticmethod
    def _check_slide_has_title(slide) -> bool:
        """Check if slide has a non-empty title."""
        for shape in slide.shapes:
            if shape.is_placeholder:
                ph_type = get_placeholder_type_int(shape.placeholder_format.type)
                if ph_type in TITLE_PLACEHOLDER_TYPES:
                    if shape.has_text_frame and shape.text_frame.text.strip():
                        return True
        return False
    
    @staticmethod
    def _has_alt_text(shape) -> bool:
        """
        Check if image shape has meaningful alt text.
        
        Checks both the description attribute (proper alt text)
        and the shape name as fallback.
        """
        # Check description attribute (the actual alt text storage)
        try:
            element = shape._element
            # Check for description in various possible locations
            descr = element.get('descr')
            if descr and descr.strip() and len(descr.strip()) > 3:
                return True
            
            # Check nvPicPr/cNvPr for descr
            for child in element.iter():
                if child.get('descr'):
                    descr = child.get('descr')
                    if descr and descr.strip() and len(descr.strip()) > 3:
                        return True
        except Exception:
            pass
        
        # Fallback: check name (not ideal, but some tools use this)
        if shape.name:
            name = shape.name.strip()
            # Reject generic names
            if name.lower().startswith('picture'):
                return False
            if name.lower().startswith('image'):
                return False
            if len(name) > 5:  # Meaningful name
                return True
        
        return False
    
    @staticmethod
    def _check_text_accessibility(
        shape,
        slide_idx: int,
        shape_idx: int,
        issues: Dict[str, Any]
    ) -> None:
        """Check text shape for accessibility issues."""
        try:
            text_frame = shape.text_frame
            for para in text_frame.paragraphs:
                # Check font size
                if para.font.size is not None:
                    size_pt = para.font.size.pt
                    if size_pt < 10:
                        issues["small_text"].append({
                            "slide": slide_idx,
                            "shape": shape_idx,
                            "size_pt": size_pt,
                            "text_preview": para.text[:50] if para.text else "",
                            "message": f"Text size {size_pt}pt is below minimum 10pt"
                        })
        except Exception:
            pass


class AssetValidator:
    """Validates and provides information about presentation assets."""
    
    @staticmethod
    def validate_presentation_assets(
        prs: 'Presentation',
        filepath: Optional[Path] = None
    ) -> Dict[str, Any]:
        """
        Validate all assets in presentation.
        
        Args:
            prs: Presentation to validate
            filepath: Optional file path for size check
            
        Returns:
            Validation report dict
        """
        issues = {
            "large_images": [],
            "total_embedded_size_bytes": 0,
            "image_count": 0
        }
        
        for slide_idx, slide in enumerate(prs.slides):
            for shape_idx, shape in enumerate(slide.shapes):
                if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    issues["image_count"] += 1
                    try:
                        image_blob = shape.image.blob
                        image_size = len(image_blob)
                        issues["total_embedded_size_bytes"] += image_size
                        
                        # Flag images over 2MB
                        if image_size > 2 * 1024 * 1024:
                            issues["large_images"].append({
                                "slide": slide_idx,
                                "shape": shape_idx,
                                "size_bytes": image_size,
                                "size_mb": round(image_size / (1024 * 1024), 2)
                            })
                    except Exception:
                        pass
        
        # Check total file size
        if filepath and Path(filepath).exists():
            file_size = Path(filepath).stat().st_size
            issues["file_size_bytes"] = file_size
            issues["file_size_mb"] = round(file_size / (1024 * 1024), 2)
            
            if file_size > MAX_RECOMMENDED_FILE_SIZE_MB * 1024 * 1024:
                issues["large_file_warning"] = {
                    "size_mb": issues["file_size_mb"],
                    "recommended_max_mb": MAX_RECOMMENDED_FILE_SIZE_MB
                }
        
        total_issues = len(issues["large_images"])
        if "large_file_warning" in issues:
            total_issues += 1
        
        return {
            "status": "issues_found" if total_issues > 0 else "valid",
            "total_issues": total_issues,
            "issues": issues
        }
    
    @staticmethod
    def compress_image(
        image_path: Path,
        max_width: int = 1920,
        quality: int = 85
    ) -> BytesIO:
        """
        Compress image for PowerPoint embedding.
        
        Args:
            image_path: Path to source image
            max_width: Maximum width in pixels
            quality: JPEG quality (1-100)
            
        Returns:
            BytesIO containing compressed image
            
        Raises:
            ImportError: If Pillow is not available
        """
        if not HAS_PILLOW:
            raise ImportError("Pillow is required for image compression")
        
        with PILImage.open(image_path) as img:
            # Resize if needed
            if img.width > max_width:
                ratio = max_width / img.width
                new_height = int(img.height * ratio)
                img = img.resize((max_width, new_height), PILImage.LANCZOS)
            
            # Convert to RGB if necessary
            if img.mode in ('RGBA', 'LA', 'P'):
                background = PILImage.new('RGB', img.size, (255, 255, 255))
                if img.mode == 'P':
                    img = img.convert('RGBA')
                if img.mode in ('RGBA', 'LA'):
                    background.paste(img, mask=img.split()[-1])
                else:
                    background.paste(img)
                img = background
            
            # Save to BytesIO
            output = BytesIO()
            img.save(output, format='JPEG', quality=quality, optimize=True)
            output.seek(0)
            
            return output


# ============================================================================
# MAIN POWERPOINT AGENT CLASS
# ============================================================================

class PowerPointAgent:
    """
    Core PowerPoint manipulation class for stateless tool operations.
    
    Provides comprehensive PowerPoint editing capabilities optimized for
    AI agent consumption through simple, composable operations.
    
    Features:
    - Stateless design for tool-based workflows
    - Comprehensive validation and accessibility checking
    - Atomic file locking for concurrent access safety
    - Approval token enforcement for destructive operations
    - Version tracking for all mutation operations
    - Full alignment with Presentation Architect System Prompt v3.0
    
    Example:
        with PowerPointAgent() as agent:
            agent.open(Path("presentation.pptx"))
            agent.add_slide("Title and Content")
            agent.set_title(0, "My Presentation")
            agent.save()
    """
    
    # Shape type mapping - class level for reuse
    SHAPE_TYPE_MAP = {
        "rectangle": MSO_AUTO_SHAPE_TYPE.RECTANGLE,
        "rounded_rectangle": MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE,
        "ellipse": MSO_AUTO_SHAPE_TYPE.OVAL,
        "oval": MSO_AUTO_SHAPE_TYPE.OVAL,
        "triangle": MSO_AUTO_SHAPE_TYPE.ISOSCELES_TRIANGLE,
        "arrow_right": MSO_AUTO_SHAPE_TYPE.RIGHT_ARROW,
        "arrow_left": MSO_AUTO_SHAPE_TYPE.LEFT_ARROW,
        "arrow_up": MSO_AUTO_SHAPE_TYPE.UP_ARROW,
        "arrow_down": MSO_AUTO_SHAPE_TYPE.DOWN_ARROW,
        "diamond": MSO_AUTO_SHAPE_TYPE.DIAMOND,
        "pentagon": MSO_AUTO_SHAPE_TYPE.PENTAGON,
        "hexagon": MSO_AUTO_SHAPE_TYPE.HEXAGON,
        "star": MSO_AUTO_SHAPE_TYPE.STAR_5_POINT,
        "heart": MSO_AUTO_SHAPE_TYPE.HEART,
        "lightning": MSO_AUTO_SHAPE_TYPE.LIGHTNING_BOLT,
        "sun": MSO_AUTO_SHAPE_TYPE.SUN,
        "moon": MSO_AUTO_SHAPE_TYPE.MOON,
        "cloud": MSO_AUTO_SHAPE_TYPE.CLOUD,
    }
    
    def __init__(self, filepath: Optional[Union[str, Path]] = None):
        """
        Initialize PowerPoint agent.
        
        Args:
            filepath: Optional path to open immediately
        """
        self.filepath: Optional[Path] = None
        self.prs: Optional[Presentation] = None
        self._lock: Optional[FileLock] = None
        self._template_profile: Optional[TemplateProfile] = None
        self._layout_cache: Optional[Dict[str, Any]] = None
        
        if filepath:
            self.filepath = Path(filepath)
    
    # ========================================================================
    # CONTEXT MANAGEMENT
    # ========================================================================
    
    def __enter__(self) -> 'PowerPointAgent':
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb) -> bool:
        self.close()
        return False
    
    # ========================================================================
    # VERSION TRACKING HELPERS
    # ========================================================================
    
    def _capture_version(self) -> str:
        """
        Capture current presentation version hash.
        
        Returns:
            Version hash string (16 characters)
            
        Raises:
            PowerPointAgentError: If no presentation loaded
        """
        return self.get_presentation_version()
    
    def _validate_approval_token(
        self,
        operation: str,
        approval_token: Optional[str],
        required_scope: str
    ) -> None:
        """
        Validate approval token for destructive operations.
        
        Args:
            operation: Name of the operation (for error messages)
            approval_token: The approval token string (or None)
            required_scope: Required scope for this operation
            
        Raises:
            ApprovalTokenError: If token is missing or invalid
        """
        if approval_token is None:
            raise ApprovalTokenError(
                f"Approval token required for {operation} operation",
                details={
                    "operation": operation,
                    "required_scope": required_scope,
                    "suggestion": f"Generate approval token with scope '{required_scope}' and retry"
                }
            )
        
        # Basic token format validation
        # In production, this would verify cryptographic signature
        if not isinstance(approval_token, str) or len(approval_token) < 10:
            raise ApprovalTokenError(
                f"Invalid approval token format for {operation}",
                details={
                    "operation": operation,
                    "required_scope": required_scope,
                    "token_provided": bool(approval_token)
                }
            )
        
        # Token is considered valid if it passes basic checks
        # Full validation would include signature verification, expiry check, etc.
        logger.debug(f"Approval token validated for {operation}")
    
    # ========================================================================
    # FILE OPERATIONS
    # ========================================================================
    
    def create_new(self, template: Optional[Union[str, Path]] = None) -> None:
        """
        Create new presentation, optionally from template.
        
        Args:
            template: Optional path to template .pptx file
            
        Raises:
            FileNotFoundError: If template doesn't exist
            TemplateError: If template cannot be loaded
        """
        if template:
            template_path = PathValidator.validate_pptx_path(template, must_exist=True)
            try:
                self.prs = Presentation(str(template_path))
            except Exception as e:
                raise TemplateError(
                    f"Failed to load template: {template_path}",
                    details={"error": str(e)}
                )
        else:
            self.prs = Presentation()
        
        self._template_profile = TemplateProfile(self.prs)
        self._layout_cache = None
    
    def open(
        self,
        filepath: Union[str, Path],
        acquire_lock: bool = True
    ) -> None:
        """
        Open existing presentation.
        
        Args:
            filepath: Path to .pptx file
            acquire_lock: Whether to acquire exclusive file lock
            
        Raises:
            PathValidationError: If path is invalid
            FileLockError: If lock cannot be acquired
            PowerPointAgentError: If file cannot be opened
        """
        validated_path = PathValidator.validate_pptx_path(filepath, must_exist=True)
        self.filepath = validated_path
        
        # Acquire lock if requested
        if acquire_lock:
            self._lock = FileLock(validated_path)
            if not self._lock.acquire():
                raise FileLockError(
                    f"Could not acquire lock on {validated_path}",
                    details={"filepath": str(validated_path)}
                )
        
        # Load presentation (with lock release on failure)
        try:
            self.prs = Presentation(str(validated_path))
            self._template_profile = TemplateProfile(self.prs)
            self._layout_cache = None
        except Exception as e:
            # Release lock on failure
            if self._lock:
                self._lock.release()
                self._lock = None
            raise PowerPointAgentError(
                f"Failed to open presentation: {validated_path}",
                details={"error": str(e)}
            )
    
    def save(self, filepath: Optional[Union[str, Path]] = None) -> None:
        """
        Save presentation.
        
        Args:
            filepath: Output path (uses original path if None)
            
        Raises:
            PowerPointAgentError: If no presentation loaded
            PathValidationError: If output path is invalid
        """
        if not self.prs:
            raise PowerPointAgentError("No presentation loaded")
        
        target = filepath or self.filepath
        if not target:
            raise PowerPointAgentError("No output path specified")
        
        target_path = PathValidator.validate_pptx_path(
            target,
            must_exist=False,
            must_be_writable=True
        )
        
        # Ensure parent directory exists
        target_path.parent.mkdir(parents=True, exist_ok=True)
        
        self.prs.save(str(target_path))
        self.filepath = target_path
    
    def close(self) -> None:
        """Close presentation and release resources."""
        self.prs = None
        self._template_profile = None
        self._layout_cache = None
        
        if self._lock:
            self._lock.release()
            self._lock = None
    
    def clone_presentation(self, output_path: Union[str, Path]) -> 'PowerPointAgent':
        """
        Clone current presentation to a new file.
        
        Args:
            output_path: Path for the cloned presentation
            
        Returns:
            New PowerPointAgent instance with cloned presentation
            
        Raises:
            PowerPointAgentError: If no presentation loaded
            PathValidationError: If output path is invalid
        """
        if not self.prs:
            raise PowerPointAgentError("No presentation loaded")
        
        output = PathValidator.validate_pptx_path(
            output_path,
            must_exist=False,
            must_be_writable=True
        )
        
        # Save to new location
        output.parent.mkdir(parents=True, exist_ok=True)
        self.prs.save(str(output))
        
        # Create new agent with cloned file
        new_agent = PowerPointAgent()
        new_agent.open(output)
        
        return new_agent
    
    # ========================================================================
    # SLIDE OPERATIONS
    # ========================================================================
    
    def add_slide(
        self,
        layout_name: str = "Title and Content",
        index: Optional[int] = None
    ) -> Dict[str, Any]:
        """
        Add new slide with specified layout.
        
        Args:
            layout_name: Name of layout to use
            index: Position to insert (None = append at end)
            
        Returns:
            Dict with slide_index, layout_name, and version tracking
            
        Raises:
            PowerPointAgentError: If no presentation loaded
            LayoutNotFoundError: If layout doesn't exist
            SlideNotFoundError: If index is invalid
        """
        if not self.prs:
            raise PowerPointAgentError("No presentation loaded")
        
        version_before = self._capture_version()
        
        layout = self._get_layout(layout_name)
        
        # Validate index before adding slide
        if index is not None:
            max_valid_index = len(self.prs.slides)
            if not 0 <= index <= max_valid_index:
                raise SlideNotFoundError(
                    f"Insert index {index} out of range (0-{max_valid_index})",
                    details={
                        "index": index,
                        "valid_range": f"0-{max_valid_index}",
                        "current_slide_count": len(self.prs.slides)
                    }
                )
        
        slide = self.prs.slides.add_slide(layout)
        
        if index is not None:
            # Move slide from end to target position
            xml_slides = self.prs.slides._sldIdLst
            slide_elem = xml_slides[-1]
            xml_slides.remove(slide_elem)
            xml_slides.insert(index, slide_elem)
            result_index = index
        else:
            result_index = len(self.prs.slides) - 1
        
        version_after = self._capture_version()
        
        return {
            "slide_index": result_index,
            "layout_name": layout_name,
            "total_slides": len(self.prs.slides),
            "presentation_version_before": version_before,
            "presentation_version_after": version_after
        }
    
    def delete_slide(
        self,
        index: int,
        approval_token: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Delete slide at index.
        
        ⚠️ DESTRUCTIVE OPERATION - Requires approval token.
        
        Args:
            index: Slide index (0-based)
            approval_token: Required approval token for this destructive operation
            
        Returns:
            Dict with deleted index, new slide count, and version tracking
            
        Raises:
            ApprovalTokenError: If approval token is missing or invalid
            SlideNotFoundError: If index is out of range
        """
        if not self.prs:
            raise PowerPointAgentError("No presentation loaded")
        
        # Validate approval token
        self._validate_approval_token(
            "delete_slide",
            approval_token,
            APPROVAL_SCOPE_DELETE_SLIDE
        )
        
        version_before = self._capture_version()
        
        slide_count = len(self.prs.slides)
        if not 0 <= index < slide_count:
            raise SlideNotFoundError(
                f"Slide index {index} out of range",
                details={"index": index, "slide_count": slide_count}
            )
        
        # Get slide relationship ID and remove
        rId = self.prs.slides._sldIdLst[index].rId
        self.prs.part.drop_rel(rId)
        del self.prs.slides._sldIdLst[index]
        
        version_after = self._capture_version()
        
        return {
            "deleted_index": index,
            "previous_count": slide_count,
            "new_count": len(self.prs.slides),
            "presentation_version_before": version_before,
            "presentation_version_after": version_after
        }
    
    def duplicate_slide(self, index: int) -> Dict[str, Any]:
        """
        Duplicate slide at index.
        
        Args:
            index: Slide index to duplicate
            
        Returns:
            Dict with new slide index and version tracking
            
        Raises:
            SlideNotFoundError: If index is out of range
        """
        version_before = self._capture_version()
        
        source_slide = self._get_slide(index)
        
        # Add new slide with same layout
        layout = source_slide.slide_layout
        new_slide = self.prs.slides.add_slide(layout)
        new_index = len(self.prs.slides) - 1
        
        # Copy shapes
        for shape in source_slide.shapes:
            try:
                self._copy_shape(shape, new_slide)
            except Exception as e:
                logger.warning(f"Could not copy shape: {e}")
        
        version_after = self._capture_version()
        
        return {
            "source_index": index,
            "new_index": new_index,
            "total_slides": len(self.prs.slides),
            "presentation_version_before": version_before,
            "presentation_version_after": version_after
        }
    
    def reorder_slides(self, from_index: int, to_index: int) -> Dict[str, Any]:
        """
        Move slide from one position to another.
        
        Args:
            from_index: Current position
            to_index: Desired position
            
        Returns:
            Dict with movement details and version tracking
            
        Raises:
            SlideNotFoundError: If either index is out of range
        """
        if not self.prs:
            raise PowerPointAgentError("No presentation loaded")
        
        version_before = self._capture_version()
        
        slide_count = len(self.prs.slides)
        
        if not 0 <= from_index < slide_count:
            raise SlideNotFoundError(
                f"Source index {from_index} out of range",
                details={"from_index": from_index, "slide_count": slide_count}
            )
        
        if not 0 <= to_index < slide_count:
            raise SlideNotFoundError(
                f"Target index {to_index} out of range",
                details={"to_index": to_index, "slide_count": slide_count}
            )
        
        xml_slides = self.prs.slides._sldIdLst
        slide_elem = xml_slides[from_index]
        xml_slides.remove(slide_elem)
        xml_slides.insert(to_index, slide_elem)
        
        version_after = self._capture_version()
        
        return {
            "from_index": from_index,
            "to_index": to_index,
            "total_slides": slide_count,
            "presentation_version_before": version_before,
            "presentation_version_after": version_after
        }
    
    def get_slide_count(self) -> int:
        """
        Get total number of slides.
        
        Returns:
            Number of slides
        """
        if not self.prs:
            raise PowerPointAgentError("No presentation loaded")
        return len(self.prs.slides)
    
    # ========================================================================
    # TEXT OPERATIONS
    # ========================================================================
    
    def add_text_box(
        self,
        slide_index: int,
        text: str,
        position: Dict[str, Any],
        size: Dict[str, Any],
        font_name: Optional[str] = None,
        font_size: int = 18,
        bold: bool = False,
        italic: bool = False,
        color: Optional[str] = None,
        alignment: str = "left"
    ) -> Dict[str, Any]:
        """
        Add text box to slide.
        
        Args:
            slide_index: Target slide index
            text: Text content
            position: Position dict (see Position.from_dict)
            size: Size dict (see Size.from_dict)
            font_name: Font name (None uses theme font)
            font_size: Font size in points
            bold: Bold text
            italic: Italic text
            color: Text color hex (e.g., "#FF0000")
            alignment: Text alignment ("left", "center", "right", "justify")
            
        Returns:
            Dict with shape_index, details, and version tracking
            
        Raises:
            SlideNotFoundError: If slide index is invalid
            InvalidPositionError: If position is invalid
        """
        version_before = self._capture_version()
        
        slide = self._get_slide(slide_index)
        
        # Parse position and size
        left, top = Position.from_dict(position)
        width, height = Size.from_dict(size)
        
        if width is None or height is None:
            raise ValueError("Text box must have explicit width and height")
        
        # Create text box
        text_box = slide.shapes.add_textbox(
            Inches(left), Inches(top),
            Inches(width), Inches(height)
        )
        
        # Configure text frame
        text_frame = text_box.text_frame
        text_frame.text = text
        text_frame.word_wrap = True
        
        # Apply formatting
        paragraph = text_frame.paragraphs[0]
        if font_name:
            paragraph.font.name = font_name
        paragraph.font.size = Pt(font_size)
        paragraph.font.bold = bold
        paragraph.font.italic = italic
        
        if color:
            paragraph.font.color.rgb = ColorHelper.from_hex(color)
        
        # Set alignment
        alignment_map = {
            "left": PP_ALIGN.LEFT,
            "center": PP_ALIGN.CENTER,
            "right": PP_ALIGN.RIGHT,
            "justify": PP_ALIGN.JUSTIFY
        }
        paragraph.alignment = alignment_map.get(alignment.lower(), PP_ALIGN.LEFT)
        
        # Find shape index
        shape_index = len(slide.shapes) - 1
        
        version_after = self._capture_version()
        
        return {
            "slide_index": slide_index,
            "shape_index": shape_index,
            "text_length": len(text),
            "position": {"left": left, "top": top},
            "size": {"width": width, "height": height},
            "presentation_version_before": version_before,
            "presentation_version_after": version_after
        }
    
    def set_title(
        self,
        slide_index: int,
        title: str,
        subtitle: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Set slide title and optional subtitle.
        
        Args:
            slide_index: Target slide index
            title: Title text
            subtitle: Optional subtitle text
            
        Returns:
            Dict with title/subtitle set status and version tracking
            
        Raises:
            SlideNotFoundError: If slide index is invalid
        """
        version_before = self._capture_version()
        
        slide = self._get_slide(slide_index)
        
        title_set = False
        subtitle_set = False
        title_shape_index = None
        subtitle_shape_index = None
        
        for idx, shape in enumerate(slide.shapes):
            if shape.is_placeholder:
                ph_type = get_placeholder_type_int(shape.placeholder_format.type)
                
                # Check for title placeholder
                if ph_type in TITLE_PLACEHOLDER_TYPES:
                    if shape.has_text_frame:
                        shape.text_frame.text = title
                        title_set = True
                        title_shape_index = idx
                
                # Check for subtitle placeholder
                elif ph_type == SUBTITLE_PLACEHOLDER_TYPE:
                    if subtitle and shape.has_text_frame:
                        shape.text_frame.text = subtitle
                        subtitle_set = True
                        subtitle_shape_index = idx
        
        version_after = self._capture_version()
        
        return {
            "slide_index": slide_index,
            "title_set": title_set,
            "subtitle_set": subtitle_set,
            "title_shape_index": title_shape_index,
            "subtitle_shape_index": subtitle_shape_index,
            "presentation_version_before": version_before,
            "presentation_version_after": version_after
        }
    
    def add_bullet_list(
        self,
        slide_index: int,
        items: List[str],
        position: Dict[str, Any],
        size: Dict[str, Any],
        bullet_style: str = "bullet",
        font_size: int = 18,
        font_name: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Add bullet list to slide.
        
        Args:
            slide_index: Target slide index
            items: List of bullet items
            position: Position dict
            size: Size dict
            bullet_style: "bullet", "numbered", or "none"
            font_size: Font size in points
            font_name: Optional font name
            
        Returns:
            Dict with shape_index, item count, and version tracking
        """
        version_before = self._capture_version()
        
        slide = self._get_slide(slide_index)
        
        left, top = Position.from_dict(position)
        width, height = Size.from_dict(size)
        
        if width is None or height is None:
            raise ValueError("Bullet list must have explicit width and height")
        
        # Create text box for bullets
        text_box = slide.shapes.add_textbox(
            Inches(left), Inches(top),
            Inches(width), Inches(height)
        )
        
        text_frame = text_box.text_frame
        text_frame.word_wrap = True
        
        for idx, item in enumerate(items):
            if idx == 0:
                p = text_frame.paragraphs[0]
            else:
                p = text_frame.add_paragraph()
            
            if bullet_style == "numbered":
                p.text = f"{idx + 1}. {item}"
            else:
                p.text = item
            
            p.level = 0
            p.font.size = Pt(font_size)
            if font_name:
                p.font.name = font_name
        
        shape_index = len(slide.shapes) - 1
        
        version_after = self._capture_version()
        
        return {
            "slide_index": slide_index,
            "shape_index": shape_index,
            "item_count": len(items),
            "bullet_style": bullet_style,
            "presentation_version_before": version_before,
            "presentation_version_after": version_after
        }
    
    def format_text(
        self,
        slide_index: int,
        shape_index: int,
        font_name: Optional[str] = None,
        font_size: Optional[int] = None,
        bold: Optional[bool] = None,
        italic: Optional[bool] = None,
        color: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Format existing text shape.
        
        Args:
            slide_index: Target slide index
            shape_index: Shape index on slide
            font_name: Optional font name
            font_size: Optional font size in points
            bold: Optional bold setting
            italic: Optional italic setting
            color: Optional color hex
            
        Returns:
            Dict with formatting applied and version tracking
        """
        version_before = self._capture_version()
        
        shape = self._get_shape(slide_index, shape_index)
        
        if not hasattr(shape, 'text_frame') or not shape.has_text_frame:
            raise ValueError(f"Shape at index {shape_index} does not have text")
        
        changes = []
        
        for paragraph in shape.text_frame.paragraphs:
            if font_name is not None:
                paragraph.font.name = font_name
                changes.append("font_name")
            if font_size is not None:
                paragraph.font.size = Pt(font_size)
                changes.append("font_size")
            if bold is not None:
                paragraph.font.bold = bold
                changes.append("bold")
            if italic is not None:
                paragraph.font.italic = italic
                changes.append("italic")
            if color is not None:
                paragraph.font.color.rgb = ColorHelper.from_hex(color)
                changes.append("color")
        
        version_after = self._capture_version()
        
        return {
            "slide_index": slide_index,
            "shape_index": shape_index,
            "changes_applied": list(set(changes)),
            "presentation_version_before": version_before,
            "presentation_version_after": version_after
        }
    
    def replace_text(
        self,
        find: str,
        replace: str,
        slide_index: Optional[int] = None,
        shape_index: Optional[int] = None,
        match_case: bool = False
    ) -> Dict[str, Any]:
        """
        Find and replace text in presentation.
        
        Args:
            find: Text to find
            replace: Replacement text
            slide_index: Optional specific slide (None = all slides)
            shape_index: Optional specific shape (requires slide_index)
            match_case: Case-sensitive matching
            
        Returns:
            Dict with replacement count, locations, and version tracking
        """
        if not self.prs:
            raise PowerPointAgentError("No presentation loaded")
        
        if shape_index is not None and slide_index is None:
            raise ValueError("shape_index requires slide_index to be specified")
        
        version_before = self._capture_version()
        
        replacements = []
        total_count = 0
        
        # Determine slides to process
        if slide_index is not None:
            slides_to_process = [(slide_index, self._get_slide(slide_index))]
        else:
            slides_to_process = list(enumerate(self.prs.slides))
        
        for s_idx, slide in slides_to_process:
            # Determine shapes to process
            if shape_index is not None:
                shapes_to_process = [(shape_index, self._get_shape(s_idx, shape_index))]
            else:
                shapes_to_process = list(enumerate(slide.shapes))
            
            for sh_idx, shape in shapes_to_process:
                if not hasattr(shape, 'text_frame') or not shape.has_text_frame:
                    continue
                
                count = self._replace_text_in_shape(shape, find, replace, match_case)
                if count > 0:
                    total_count += count
                    replacements.append({
                        "slide": s_idx,
                        "shape": sh_idx,
                        "count": count
                    })
        
        version_after = self._capture_version()
        
        return {
            "find": find,
            "replace": replace,
            "match_case": match_case,
            "total_replacements": total_count,
            "locations": replacements,
            "presentation_version_before": version_before,
            "presentation_version_after": version_after
        }
    
    def _replace_text_in_shape(
        self,
        shape,
        find: str,
        replace: str,
        match_case: bool
    ) -> int:
        """Replace text within a single shape, preserving formatting where possible."""
        count = 0
        
        try:
            text_frame = shape.text_frame
        except (AttributeError, TypeError):
            return 0
        
        # Strategy 1: Replace in runs (preserves formatting)
        for paragraph in text_frame.paragraphs:
            for run in paragraph.runs:
                if match_case:
                    if find in run.text:
                        occurrences = run.text.count(find)
                        run.text = run.text.replace(find, replace)
                        count += occurrences
                else:
                    if find.lower() in run.text.lower():
                        pattern = re.compile(re.escape(find), re.IGNORECASE)
                        matches = pattern.findall(run.text)
                        run.text = pattern.sub(replace, run.text)
                        count += len(matches)
        
        if count > 0:
            return count
        
        # Strategy 2: Full text replacement (if text spans runs)
        try:
            full_text = shape.text
            if not full_text:
                return 0
            
            if 
```
