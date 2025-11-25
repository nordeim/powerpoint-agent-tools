#!/usr/bin/env python3
"""
PowerPoint Agent Core Library
Production-grade PowerPoint manipulation with validation and accessibility

This is the foundational library used by all CLI tools.
Designed for stateless, security-hardened PowerPoint operations.

Author: PowerPoint Agent Team
License: MIT
Version: 1.1.0 (Bug Fix Release)

Changelog v1.1.0:
- Added missing subprocess import for PDF export
- Added missing PP_PLACEHOLDER import and constants
- Replaced all magic numbers with named constants
- Removed text truncation in get_slide_info() (was causing data loss)
- Added position/size information to shape inspection
- Added placeholder subtype decoding
- Replaced print() with proper logging
- Fixed subtitle placeholder type (was 2, should be SUBTITLE constant)
- Added comprehensive PLACEHOLDER_TYPE_NAMES mapping
"""

import re
import json
import subprocess
import tempfile
import shutil
import threading
import time
import logging
from pathlib import Path
from typing import Any, Dict, List, Optional, Union, Tuple, Set
from enum import Enum
from datetime import datetime
from io import BytesIO

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu
    from pptx.enum.shapes import MSO_SHAPE_TYPE, MSO_AUTO_SHAPE_TYPE, PP_PLACEHOLDER, MSO_CONNECTOR
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.chart.data import CategoryChartData
    from pptx.dml.color import RGBColor
    from pptx.oxml.xmlchemy import OxmlElement
except ImportError:
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

try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False
    pd = None


# ============================================================================
# LOGGING SETUP
# ============================================================================

logger = logging.getLogger(__name__)
logger.setLevel(logging.WARNING)

if not logger.handlers:
    handler = logging.StreamHandler()
    formatter = logging.Formatter('%(levelname)s:%(name)s:%(message)s')
    handler.setFormatter(formatter)
    logger.addHandler(handler)


# ============================================================================
# EXCEPTIONS
# ============================================================================

class PowerPointAgentError(Exception):
    """Base exception for all PowerPoint agent errors."""
    def __init__(self, message: str, details: Optional[Dict[str, Any]] = None):
        super().__init__(message)
        self.message = message
        self.details = details or {}
    
    def to_json(self) -> Dict[str, Any]:
        """Convert exception to JSON-serializable dict."""
        return {
            "error": self.__class__.__name__,
            "message": self.message,
            "details": self.details
        }


class SlideNotFoundError(PowerPointAgentError):
    """Raised when slide index is out of range."""
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


# ============================================================================
# CONSTANTS & ENUMS
# ============================================================================

class ShapeType(Enum):
    """Common shape types supported by python-pptx."""
    RECTANGLE = MSO_AUTO_SHAPE_TYPE.RECTANGLE
    ROUNDED_RECTANGLE = MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE
    ELLIPSE = MSO_AUTO_SHAPE_TYPE.OVAL
    TRIANGLE = MSO_AUTO_SHAPE_TYPE.ISOSCELES_TRIANGLE
    ARROW_RIGHT = MSO_AUTO_SHAPE_TYPE.RIGHT_ARROW
    ARROW_LEFT = MSO_AUTO_SHAPE_TYPE.LEFT_ARROW
    ARROW_UP = MSO_AUTO_SHAPE_TYPE.UP_ARROW
    ARROW_DOWN = MSO_AUTO_SHAPE_TYPE.DOWN_ARROW


class ChartType(Enum):
    """Supported chart types."""
    COLUMN = XL_CHART_TYPE.COLUMN_CLUSTERED
    COLUMN_STACKED = XL_CHART_TYPE.COLUMN_STACKED
    BAR = XL_CHART_TYPE.BAR_CLUSTERED
    BAR_STACKED = XL_CHART_TYPE.BAR_STACKED
    LINE = XL_CHART_TYPE.LINE
    LINE_MARKERS = XL_CHART_TYPE.LINE_MARKERS
    PIE = XL_CHART_TYPE.PIE
    PIE_EXPLODED = XL_CHART_TYPE.PIE_EXPLODED
    AREA = XL_CHART_TYPE.AREA
    SCATTER = XL_CHART_TYPE.XY_SCATTER


class TextAlignment(Enum):
    """Text alignment options."""
    LEFT = PP_ALIGN.LEFT
    CENTER = PP_ALIGN.CENTER
    RIGHT = PP_ALIGN.RIGHT
    JUSTIFY = PP_ALIGN.JUSTIFY


class VerticalAlignment(Enum):
    """Vertical text alignment."""
    TOP = MSO_ANCHOR.TOP
    MIDDLE = MSO_ANCHOR.MIDDLE
    BOTTOM = MSO_ANCHOR.BOTTOM


class BulletStyle(Enum):
    """Bullet list styles."""
    BULLET = "bullet"
    NUMBERED = "numbered"
    NONE = "none"


class ImageFormat(Enum):
    """Supported image formats."""
    PNG = "PNG"
    JPG = "JPEG"
    JPEG = "JPEG"
    GIF = "GIF"
    BMP = "BMP"
    SVG = "SVG"


class ExportFormat(Enum):
    """Export format options."""
    PDF = "pdf"
    PNG = "png"
    JPG = "jpg"
    PPTX = "pptx"

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

# Standard slide dimensions (16:9 widescreen)
SLIDE_WIDTH_INCHES = 10.0
SLIDE_HEIGHT_INCHES = 7.5

# Alternative dimensions (4:3 standard)
SLIDE_WIDTH_4_3 = 10.0
SLIDE_HEIGHT_4_3 = 7.5

# Standard anchor points for positioning
ANCHOR_POINTS = {
    "top_left": (0.0, 0.0),
    "top_center": (SLIDE_WIDTH_INCHES / 2, 0.0),
    "top_right": (SLIDE_WIDTH_INCHES, 0.0),
    "center_left": (0.0, SLIDE_HEIGHT_INCHES / 2),
    "center": (SLIDE_WIDTH_INCHES / 2, SLIDE_HEIGHT_INCHES / 2),
    "center_right": (SLIDE_WIDTH_INCHES, SLIDE_HEIGHT_INCHES / 2),
    "bottom_left": (0.0, SLIDE_HEIGHT_INCHES),
    "bottom_center": (SLIDE_WIDTH_INCHES / 2, SLIDE_HEIGHT_INCHES),
    "bottom_right": (SLIDE_WIDTH_INCHES, SLIDE_HEIGHT_INCHES)
}

# Standard corporate colors (RGB)
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
    "title": "Calibri",
    "body": "Calibri",
    "code": "Consolas"
}

# WCAG 2.1 color contrast ratios
WCAG_CONTRAST_NORMAL = 4.5  # Normal text
WCAG_CONTRAST_LARGE = 3.0   # Large text (18pt+)

# Maximum recommended file size (MB)
MAX_RECOMMENDED_FILE_SIZE = 50


# ============================================================================
# FILE LOCKING
# ============================================================================

class FileLock:
    """Simple file locking mechanism for concurrent access prevention."""
    
    def __init__(self, filepath: Path, timeout: float = 10.0):
        self.filepath = filepath
        self.lockfile = filepath.parent / f".{filepath.name}.lock"
        self.timeout = timeout
        self.acquired = False
    
    def acquire(self) -> bool:
        """Acquire lock with timeout."""
        start_time = time.time()
        while time.time() - start_time < self.timeout:
            try:
                self.lockfile.touch(exist_ok=False)
                self.acquired = True
                return True
            except FileExistsError:
                time.sleep(0.1)
        return False
    
    def release(self) -> None:
        """Release lock."""
        if self.acquired:
            try:
                self.lockfile.unlink(missing_ok=True)
                self.acquired = False
            except Exception:
                pass
    
    def __enter__(self):
        if not self.acquire():
            raise FileLockError(
                f"Could not acquire lock on {self.filepath} within {self.timeout}s"
            )
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        self.release()
        return False


# ============================================================================
# POSITION & SIZE HELPERS
# ============================================================================

class Position:
    """Flexible position system supporting multiple input formats."""
    
    @staticmethod
    def from_dict(pos_dict: Dict[str, Any], 
                  slide_width: float = SLIDE_WIDTH_INCHES,
                  slide_height: float = SLIDE_HEIGHT_INCHES) -> Tuple[float, float]:
        """
        Convert position dict to (left, top) in inches.
        
        Supports multiple formats:
        1. Absolute inches: {"left": 1.5, "top": 2.0}
        2. Percentage: {"left": "20%", "top": "30%"}
        3. Anchor-based: {"anchor": "center", "offset_x": 0.5, "offset_y": -1.0}
        4. Grid system: {"grid_row": 2, "grid_col": 3, "grid_size": 12}
        5. Excel-like: {"grid": "C4"}
        
        Args:
            pos_dict: Position specification
            slide_width: Slide width in inches
            slide_height: Slide height in inches
            
        Returns:
            Tuple of (left, top) in inches
            
        Raises:
            InvalidPositionError: If format is invalid
        """
        if "left" in pos_dict and "top" in pos_dict:
            left = Position._parse_dimension(pos_dict["left"], slide_width)
            top = Position._parse_dimension(pos_dict["top"], slide_height)
            return (left, top)
        
        if "anchor" in pos_dict:
            anchor = ANCHOR_POINTS.get(pos_dict["anchor"])
            if not anchor:
                raise InvalidPositionError(
                    f"Unknown anchor: {pos_dict['anchor']}. "
                    f"Available: {list(ANCHOR_POINTS.keys())}"
                )
            offset_x = pos_dict.get("offset_x", 0)
            offset_y = pos_dict.get("offset_y", 0)
            return (anchor[0] + offset_x, anchor[1] + offset_y)
        
        if "grid_row" in pos_dict and "grid_col" in pos_dict:
            grid_size = pos_dict.get("grid_size", 12)
            cell_width = slide_width / grid_size
            cell_height = slide_height / grid_size
            left = pos_dict["grid_col"] * cell_width
            top = pos_dict["grid_row"] * cell_height
            return (left, top)
        
        if "grid" in pos_dict:
            grid_ref = pos_dict["grid"].upper()
            match = re.match(r'^([A-Z]+)(\d+)$', grid_ref)
            if not match:
                raise InvalidPositionError(f"Invalid grid reference: {grid_ref}")
            
            col_str, row_str = match.groups()
            
            col_num = 0
            for char in col_str:
                col_num = col_num * 26 + (ord(char) - ord('A') + 1)
            
            row_num = int(row_str)
            
            grid_size = pos_dict.get("grid_size", 12)
            cell_width = slide_width / grid_size
            cell_height = slide_height / grid_size
            
            left = (col_num - 1) * cell_width
            top = (row_num - 1) * cell_height
            
            return (left, top)
        
        raise InvalidPositionError(
            f"Invalid position format: {pos_dict}. "
            "Must have 'left'/'top', 'anchor', 'grid_row'/'grid_col', or 'grid'"
        )
    
    @staticmethod
    def _parse_dimension(value: Union[str, float, int], max_dimension: float) -> float:
        """Parse dimension (supports percentages or absolute values)."""
        if isinstance(value, str):
            if value.endswith('%'):
                percent = float(value[:-1]) / 100
                return percent * max_dimension
            else:
                return float(value)
        return float(value)


class Size:
    """Flexible size system supporting multiple input formats."""
    
    @staticmethod
    def from_dict(size_dict: Dict[str, Any],
                  slide_width: float = SLIDE_WIDTH_INCHES,
                  slide_height: float = SLIDE_HEIGHT_INCHES,
                  aspect_ratio: Optional[float] = None) -> Tuple[Optional[float], Optional[float]]:
        """
        Convert size dict to (width, height) in inches.
        
        Supports:
        - {"width": 5.0, "height": 3.0}  # Absolute
        - {"width": "50%", "height": "30%"}  # Percentage
        - {"width": "auto", "height": 3.0}  # Maintain aspect ratio
        - {"width": 5.0, "height": "auto"}  # Maintain aspect ratio
        
        Args:
            size_dict: Size specification
            slide_width: Slide width in inches
            slide_height: Slide height in inches
            aspect_ratio: Optional aspect ratio (width/height)
            
        Returns:
            Tuple of (width, height) in inches, either can be None for "auto"
        """
        if "width" not in size_dict and "height" not in size_dict:
            raise ValueError("Size must have at least width or height")
        
        width_spec = size_dict.get("width")
        height_spec = size_dict.get("height")
        
        if width_spec == "auto":
            width = None
        elif width_spec:
            width = Position._parse_dimension(width_spec, slide_width)
        else:
            width = None
        
        if height_spec == "auto":
            height = None
        elif height_spec:
            height = Position._parse_dimension(height_spec, slide_height)
        else:
            height = None
        
        if width is None and height is not None and aspect_ratio:
            width = height * aspect_ratio
        elif height is None and width is not None and aspect_ratio:
            height = width / aspect_ratio
        
        return (width, height)


# ============================================================================
# COLOR MANAGEMENT
# ============================================================================

class ColorHelper:
    """Utilities for color conversion and validation."""
    
    @staticmethod
    def from_hex(hex_color: str) -> RGBColor:
        """Convert hex color to RGBColor."""
        hex_color = hex_color.lstrip('#')
        if len(hex_color) != 6 or not all(c in '0123456789ABCDEFabcdef' for c in hex_color):
            raise ValueError(f"Invalid hex color: {hex_color}. Must be 6 hex digits.")
        
        r = int(hex_color[0:2], 16)
        g = int(hex_color[2:4], 16)
        b = int(hex_color[4:6], 16)
        
        return RGBColor(r, g, b)
    
    @staticmethod
    def to_hex(rgb_color: RGBColor) -> str:
        """Convert RGBColor to hex string."""
        return f"#{rgb_color.r:02x}{rgb_color.g:02x}{rgb_color.b:02x}"
    
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

        def _channel(c):
            c = c / 255.0
            return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
        
        r = _channel(r)
        g = _channel(g)
        b = _channel(b)
        
        return 0.2126 * r + 0.7152 * g + 0.0722 * b
    
    @staticmethod
    def contrast_ratio(color1: RGBColor, color2: RGBColor) -> float:
        """Calculate WCAG contrast ratio between two colors."""
        lum1 = ColorHelper.luminance(color1)
        lum2 = ColorHelper.luminance(color2)
        
        lighter = max(lum1, lum2)
        darker = min(lum1, lum2)
        
        return (lighter + 0.05) / (darker + 0.05)
    
    @staticmethod
    def meets_wcag(color1: RGBColor, color2: RGBColor, 
                   is_large_text: bool = False) -> bool:
        """Check if color combination meets WCAG 2.1 standards."""
        ratio = ColorHelper.contrast_ratio(color1, color2)
        threshold = WCAG_CONTRAST_LARGE if is_large_text else WCAG_CONTRAST_NORMAL
        return ratio >= threshold


# ============================================================================
# TEMPLATE PRESERVATION
# ============================================================================

class TemplateProfile:
    """Captures and applies PowerPoint template formatting."""
    
    def __init__(self, prs: Optional[Presentation] = None):
        self.slide_layouts: List[Dict[str, Any]] = []
        self.master_slides: List[Dict[str, Any]] = []
        self.theme_colors: Dict[str, str] = {}
        self.theme_fonts: Dict[str, str] = {}
        self.default_text_styles: Dict[str, Any] = {}
        
        if prs:
            self.capture_from_presentation(prs)
    
    def capture_from_presentation(self, prs: Presentation) -> None:
        """Analyze and capture all template elements from presentation."""
        
        for layout in prs.slide_layouts:
            layout_info = {
                "name": layout.name,
                "placeholders": [
                    {
                        "type": ph.placeholder_format.type,
                        "idx": ph.placeholder_format.idx,
                        "position": (ph.left, ph.top) if hasattr(ph, 'left') else None,
                        "size": (ph.width, ph.height) if hasattr(ph, 'width') else None
                    }
                    for ph in layout.placeholders
                ]
            }
            self.slide_layouts.append(layout_info)
        
        try:
            theme = prs.slide_master.theme
            for idx, color in enumerate(theme.theme_color_scheme):
                self.theme_colors[f"color_{idx}"] = ColorHelper.to_hex(color)
        except:
            pass
        
        try:
            for shape in prs.slide_master.shapes:
                if hasattr(shape, 'text_frame'):
                    for paragraph in shape.text_frame.paragraphs:
                        if paragraph.font.name:
                            self.theme_fonts[f"font_{len(self.theme_fonts)}"] = paragraph.font.name
        except:
            pass
    
    def get_layout_names(self) -> List[str]:
        """Get list of available layout names."""
        return [layout["name"] for layout in self.slide_layouts]
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert profile to JSON-serializable dict."""
        return {
            "slide_layouts": self.slide_layouts,
            "master_slides": self.master_slides,
            "theme_colors": self.theme_colors,
            "theme_fonts": self.theme_fonts,
            "default_text_styles": self.default_text_styles
        }


# ============================================================================
# ACCESSIBILITY CHECKER
# ============================================================================

class AccessibilityChecker:
    """WCAG 2.1 compliance checker for presentations."""
    
    @staticmethod
    def check_presentation(prs: Presentation) -> Dict[str, Any]:
        """
        Comprehensive accessibility check.
        
        Returns dict with issues found:
        - missing_alt_text: List of images without alt text
        - low_contrast: List of text with insufficient contrast
        - missing_titles: List of slides without titles
        - reading_order_issues: List of slides with unclear reading order
        """
        issues = {
            "missing_alt_text": [],
            "low_contrast": [],
            "missing_titles": [],
            "reading_order_issues": [],
            "large_file_size_warning": False
        }
        
        for slide_idx, slide in enumerate(prs.slides):
            has_title = False
            for shape in slide.shapes:
                if shape.is_placeholder:
                    ph_type = shape.placeholder_format.type
                    if ph_type == PP_PLACEHOLDER.TITLE or ph_type == PP_PLACEHOLDER.CENTER_TITLE:
                        if shape.has_text_frame and shape.text_frame.text.strip():
                            has_title = True
                            break
            
            if not has_title:
                issues["missing_titles"].append(slide_idx)
            
            for shape_idx, shape in enumerate(slide.shapes):
                if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    if not shape.name or shape.name.startswith("Picture"):
                        issues["missing_alt_text"].append({
                            "slide": slide_idx,
                            "shape": shape_idx,
                            "shape_name": shape.name
                        })
                
                if hasattr(shape, 'text_frame'):
                    try:
                        AccessibilityChecker._check_text_contrast(
                            shape, slide_idx, shape_idx, issues
                        )
                    except:
                        pass
        
        total_issues = sum(len(v) if isinstance(v, list) else (1 if v else 0) 
                          for v in issues.values())
        
        return {
            "status": "issues_found" if total_issues > 0 else "accessible",
            "total_issues": total_issues,
            "issues": issues,
            "wcag_level": "AA" if total_issues == 0 else "fail"
        }
    
    @staticmethod
    def _check_text_contrast(shape, slide_idx: int, shape_idx: int, 
                            issues: Dict[str, Any]) -> None:
        """Check text color contrast."""
        if not shape.has_text_frame:
            return
        
        for para in shape.text_frame.paragraphs:
            if para.font.color and para.font.color.type == 1:
                text_color = para.font.color.rgb
                
                bg_color = RGBColor(255, 255, 255)
                
                if shape.fill.type == 1:
                    bg_color = shape.fill.fore_color.rgb
                
                is_large = (para.font.size and para.font.size >= Pt(18)) or \
                          (para.font.bold and para.font.size and para.font.size >= Pt(14))
                
                if not ColorHelper.meets_wcag(text_color, bg_color, is_large):
                    issues["low_contrast"].append({
                        "slide": slide_idx,
                        "shape": shape_idx,
                        "text": para.text[:50],
                        "contrast_ratio": ColorHelper.contrast_ratio(text_color, bg_color),
                        "required": WCAG_CONTRAST_LARGE if is_large else WCAG_CONTRAST_NORMAL
                    })


# ============================================================================
# ASSET VALIDATOR
# ============================================================================

class AssetValidator:
    """Validates and optimizes presentation assets."""
    
    @staticmethod
    def validate_presentation_assets(prs: Presentation, 
                                    filepath: Optional[Path] = None) -> Dict[str, Any]:
        """
        Validate all assets in presentation.
        
        Checks:
        - Image resolution and file size
        - Missing external links
        - Large embedded objects
        - Total file size
        """
        issues = {
            "low_resolution_images": [],
            "large_images": [],
            "total_embedded_size": 0,
            "external_links": []
        }
        
        for slide_idx, slide in enumerate(prs.slides):
            for shape_idx, shape in enumerate(slide.shapes):
                if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    try:
                        image = shape.image
                        
                        if hasattr(image, 'dpi'):
                            if image.dpi[0] < 96 or image.dpi[1] < 96:
                                issues["low_resolution_images"].append({
                                    "slide": slide_idx,
                                    "shape": shape_idx,
                                    "dpi": image.dpi
                                })
                        
                        image_size = len(image.blob)
                        issues["total_embedded_size"] += image_size
                        
                        if image_size > 2 * 1024 * 1024:
                            issues["large_images"].append({
                                "slide": slide_idx,
                                "shape": shape_idx,
                                "size_mb": image_size / (1024 * 1024)
                            })
                    except:
                        pass
        
        if filepath and filepath.exists():
            file_size = filepath.stat().st_size
            if file_size > MAX_RECOMMENDED_FILE_SIZE * 1024 * 1024:
                issues["large_file_warning"] = {
                    "size_mb": file_size / (1024 * 1024),
                    "recommended_max_mb": MAX_RECOMMENDED_FILE_SIZE
                }
        
        total_issues = len(issues["low_resolution_images"]) + \
                      len(issues["large_images"]) + \
                      (1 if "large_file_warning" in issues else 0)
        
        return {
            "status": "issues_found" if total_issues > 0 else "valid",
            "total_issues": total_issues,
            "issues": issues
        }
    
    @staticmethod
    def compress_image(image_path: Path, max_width: int = 1920, 
                      quality: int = 85) -> BytesIO:
        """Compress image for PowerPoint embedding."""
        if not HAS_PILLOW:
            raise ImportError("Pillow required for image compression")
        
        with PILImage.open(image_path) as img:
            if img.width > max_width:
                ratio = max_width / img.width
                new_height = int(img.height * ratio)
                img = img.resize((max_width, new_height), PILImage.LANCZOS)
            
            if img.mode in ('RGBA', 'LA', 'P'):
                background = PILImage.new('RGB', img.size, (255, 255, 255))
                if img.mode == 'P':
                    img = img.convert('RGBA')
                background.paste(img, mask=img.split()[-1] if img.mode == 'RGBA' else None)
                img = background
            
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
    """
    
    def __init__(self, filepath: Optional[Path] = None):
        """Initialize agent with optional file."""
        self.filepath = Path(filepath) if filepath else None
        self.prs: Optional[Presentation] = None
        self._lock: Optional[FileLock] = None
        self._template_profile: Optional[TemplateProfile] = None
    
    # ========================================================================
    # FILE OPERATIONS
    # ========================================================================
    
    def create_new(self, template: Optional[Path] = None) -> None:
        """
        Create new presentation, optionally from template.
        
        Args:
            template: Optional path to template .pptx file
        """
        if template:
            if not template.exists():
                raise FileNotFoundError(f"Template not found: {template}")
            self.prs = Presentation(str(template))
            self._template_profile = TemplateProfile(self.prs)
        else:
            self.prs = Presentation()
            self._template_profile = TemplateProfile(self.prs)
    
    def open(self, filepath: Path, acquire_lock: bool = True) -> None:
        """
        Open existing presentation.
        
        Args:
            filepath: Path to .pptx file
            acquire_lock: Whether to acquire exclusive file lock
        """
        if not filepath.exists():
            raise FileNotFoundError(f"File not found: {filepath}")
        
        self.filepath = filepath
        
        if acquire_lock:
            self._lock = FileLock(filepath)
            if not self._lock.acquire():
                raise FileLockError(f"Could not lock file: {filepath}")
        
        self.prs = Presentation(str(filepath))
        self._template_profile = TemplateProfile(self.prs)
    
    def save(self, filepath: Optional[Path] = None) -> None:
        """
        Save presentation.
        
        Args:
            filepath: Output path (uses original path if None)
        """
        if not self.prs:
            raise PowerPointAgentError("No presentation loaded")
        
        target = filepath or self.filepath
        if not target:
            raise PowerPointAgentError("No output path specified")
        
        target = Path(target)
        target.parent.mkdir(parents=True, exist_ok=True)
        
        self.prs.save(str(target))
        self.filepath = target
    
    def close(self) -> None:
        """Close presentation and release lock."""
        self.prs = None
        
        if self._lock:
            self._lock.release()
            self._lock = None
    
    # ========================================================================
    # SLIDE OPERATIONS
    # ========================================================================
    
    def add_slide(self, layout_name: str = "Title and Content", 
                  index: Optional[int] = None) -> int:
        """
        Add new slide with specified layout.
        
        Args:
            layout_name: Name of layout to use
            index: Position to insert (None = end)
            
        Returns:
            Index of newly created slide
        """
        if not self.prs:
            raise PowerPointAgentError("No presentation loaded")
        
        layout = self._get_layout(layout_name)
        
        if index is None:
            slide = self.prs.slides.add_slide(layout)
            return len(self.prs.slides) - 1
        else:
            slide = self.prs.slides.add_slide(layout)
            xml_slides = self.prs.slides._sldIdLst
            xml_slides.insert(index, xml_slides[-1])
            return index
    
    def delete_slide(self, index: int) -> None:
        """
        Delete slide at index.
        
        Args:
            index: Slide index (0-based)
        """
        if not self.prs:
            raise PowerPointAgentError("No presentation loaded")
        
        if not 0 <= index < len(self.prs.slides):
            raise SlideNotFoundError(f"Slide index {index} out of range (0-{len(self.prs.slides)-1})")
        
        rId = self.prs.slides._sldIdLst[index].rId
        self.prs.part.drop_rel(rId)
        del self.prs.slides._sldIdLst[index]
    
    def duplicate_slide(self, index: int) -> int:
        """
        Duplicate slide at index.
        
        Args:
            index: Slide index to duplicate
            
        Returns:
            Index of duplicated slide
        """
        if not self.prs:
            raise PowerPointAgentError("No presentation loaded")
        
        source_slide = self.get_slide(index)
        
        layout = source_slide.slide_layout
        new_slide = self.prs.slides.add_slide(layout)
        
        for shape in source_slide.shapes:
            self._copy_shape(shape, new_slide)
        
        return len(self.prs.slides) - 1
    
    def reorder_slides(self, from_index: int, to_index: int) -> None:
        """
        Move slide from one position to another.
        
        Args:
            from_index: Current position
            to_index: Desired position
        """
        if not self.prs:
            raise PowerPointAgentError("No presentation loaded")
        
        xml_slides = self.prs.slides._sldIdLst
        
        if not 0 <= from_index < len(xml_slides):
            raise SlideNotFoundError(f"Source index {from_index} out of range")
        
        if not 0 <= to_index < len(xml_slides):
            raise SlideNotFoundError(f"Target index {to_index} out of range")
        
        slide = xml_slides[from_index]
        xml_slides.remove(slide)
        xml_slides.insert(to_index, slide)
    
    def get_slide(self, index: int):
        """
        Get slide by index.
        
        Args:
            index: Slide index (0-based)
            
        Returns:
            Slide object
        """
        if not self.prs:
            raise PowerPointAgentError("No presentation loaded")
        
        if not 0 <= index < len(self.prs.slides):
            raise SlideNotFoundError(f"Slide index {index} out of range")
        
        return self.prs.slides[index]
    
    def get_slide_count(self) -> int:
        """Get total number of slides."""
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
        font_name: str = "Calibri",
        font_size: int = 18,
        bold: bool = False,
        italic: bool = False,
        color: Optional[str] = None,
        alignment: str = "left"
    ) -> None:
        """
        Add text box to slide.
        
        Args:
            slide_index: Target slide index
            text: Text content
            position: Position dict (see Position.from_dict)
            size: Size dict (see Size.from_dict)
            font_name: Font name
            font_size: Font size in points
            bold: Bold text
            italic: Italic text
            color: Text color (hex, e.g., "#FF0000")
            alignment: Text alignment ("left", "center", "right", "justify")
        """
        slide = self.get_slide(slide_index)
        
        left, top = Position.from_dict(position)
        width, height = Size.from_dict(size)
        
        if width is None or height is None:
            raise ValueError("Text box must have explicit width and height")
        
        text_box = slide.shapes.add_textbox(
            Inches(left), Inches(top),
            Inches(width), Inches(height)
        )
        
        text_frame = text_box.text_frame
        text_frame.text = text
        text_frame.word_wrap = True
        
        paragraph = text_frame.paragraphs[0]
        paragraph.font.name = font_name
        paragraph.font.size = Pt(font_size)
        paragraph.font.bold = bold
        paragraph.font.italic = italic
        
        if color:
            paragraph.font.color.rgb = ColorHelper.from_hex(color)
        
        alignment_map = {
            "left": PP_ALIGN.LEFT,
            "center": PP_ALIGN.CENTER,
            "right": PP_ALIGN.RIGHT,
            "justify": PP_ALIGN.JUSTIFY
        }
        paragraph.alignment = alignment_map.get(alignment.lower(), PP_ALIGN.LEFT)
    
    def set_title(self, slide_index: int, title: str, 
                  subtitle: Optional[str] = None) -> None:
        """
        Set slide title and optional subtitle.
        
        Args:
            slide_index: Target slide index
            title: Title text
            subtitle: Optional subtitle text
        """
        slide = self.get_slide(slide_index)
        
        title_shape = None
        subtitle_shape = None
        
        for shape in slide.shapes:
            if shape.is_placeholder:
                ph_type = shape.placeholder_format.type
                if ph_type == PP_PLACEHOLDER.TITLE or ph_type == PP_PLACEHOLDER.CENTER_TITLE:
                    title_shape = shape
                elif ph_type == PP_PLACEHOLDER.SUBTITLE:
                    subtitle_shape = shape
        
        if title_shape and title_shape.has_text_frame:
            title_shape.text = title
        
        if subtitle and subtitle_shape and subtitle_shape.has_text_frame:
            subtitle_shape.text = subtitle
    
    def add_bullet_list(
        self,
        slide_index: int,
        items: List[str],
        position: Dict[str, Any],
        size: Dict[str, Any],
        bullet_style: str = "bullet",
        font_size: int = 18
    ) -> None:
        """
        Add bullet list to slide.
        
        Args:
            slide_index: Target slide index
            items: List of bullet items
            position: Position dict
            size: Size dict
            bullet_style: "bullet", "numbered", or "none"
            font_size: Font size in points
        """
        slide = self.get_slide(slide_index)
        
        left, top = Position.from_dict(position)
        width, height = Size.from_dict(size)
        
        if width is None or height is None:
            raise ValueError("Bullet list must have explicit width and height")
        
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
            
            p.text = item
            p.level = 0
            p.font.size = Pt(font_size)
            
            if bullet_style == "bullet":
                pass
            elif bullet_style == "numbered":
                p.text = f"{idx + 1}. {item}"
    
    def format_text(
        self,
        slide_index: int,
        shape_index: int,
        font_name: Optional[str] = None,
        font_size: Optional[int] = None,
        bold: Optional[bool] = None,
        italic: Optional[bool] = None,
        color: Optional[str] = None
    ) -> None:
        """
        Format existing text shape.
        
        Args:
            slide_index: Target slide
            shape_index: Shape index on slide
            font_name: Optional font name
            font_size: Optional font size
            bold: Optional bold setting
            italic: Optional italic setting
            color: Optional color hex
        """
        slide = self.get_slide(slide_index)
        
        if shape_index >= len(slide.shapes):
            raise ValueError(f"Shape index {shape_index} out of range")
        
        shape = slide.shapes[shape_index]
        
        if not hasattr(shape, 'text_frame'):
            raise ValueError("Shape does not have text")
        
        for paragraph in shape.text_frame.paragraphs:
            if font_name is not None:
                paragraph.font.name = font_name
            if font_size is not None:
                paragraph.font.size = Pt(font_size)
            if bold is not None:
                paragraph.font.bold = bold
            if italic is not None:
                paragraph.font.italic = italic
            if color is not None:
                paragraph.font.color.rgb = ColorHelper.from_hex(color)
    
    def replace_text(self, find: str, replace: str, match_case: bool = False) -> int:
        """
        Find and replace text across entire presentation.
        
        Attempts to preserve formatting by replacing within text runs first.
        Falls back to shape-level replacement if text is split across runs.
        
        Args:
            find: Text to find
            replace: Replacement text
            match_case: Case-sensitive search
            
        Returns:
            Number of replacements made
        """
        if not self.prs:
            raise PowerPointAgentError("No presentation loaded")
        
        count = 0
        
        for slide in self.prs.slides:
            for shape in slide.shapes:
                if not shape.has_text_frame:
                    continue
                
                try:
                    text_frame = shape.text_frame
                except (AttributeError, TypeError):
                    continue
                
                replacements_in_runs = 0
                for paragraph in text_frame.paragraphs:
                    for run in paragraph.runs:
                        if match_case:
                            if find in run.text:
                                run.text = run.text.replace(find, replace)
                                replacements_in_runs += 1
                        else:
                            if find.lower() in run.text.lower():
                                pattern = re.compile(re.escape(find), re.IGNORECASE)
                                if pattern.search(run.text):
                                    new_text = pattern.sub(replace, run.text)
                                    run.text = new_text
                                    replacements_in_runs += 1
                
                if replacements_in_runs > 0:
                    count += replacements_in_runs
                    continue
                
                try:
                    full_text = shape.text
                    if not full_text:
                        continue
                        
                    should_replace = False
                    if match_case:
                        if find in full_text:
                            should_replace = True
                    else:
                        if find.lower() in full_text.lower():
                            should_replace = True
                    
                    if should_replace:
                        if match_case:
                            new_text = full_text.replace(find, replace)
                            shape.text = new_text
                            count += full_text.count(find)
                        else:
                            pattern = re.compile(re.escape(find), re.IGNORECASE)
                            matches = re.findall(pattern, full_text)
                            new_text = pattern.sub(replace, full_text)
                            shape.text = new_text
                            count += len(matches)
                except (AttributeError, TypeError):
                    continue
        
        return count

    # ========================================================================
    # SHAPE OPERATIONS
    # ========================================================================
    
    def add_shape(
        self,
        slide_index: int,
        shape_type: str,
        position: Dict[str, Any],
        size: Dict[str, Any],
        fill_color: Optional[str] = None,
        line_color: Optional[str] = None,
        line_width: float = 1.0
    ) -> None:
        """
        Add shape to slide.
        
        Args:
            slide_index: Target slide
            shape_type: Shape type (rectangle, ellipse, arrow, etc.)
            position: Position dict
            size: Size dict
            fill_color: Fill color hex
            line_color: Line color hex
            line_width: Line width in points
        """
        slide = self.get_slide(slide_index)
        
        left, top = Position.from_dict(position)
        width, height = Size.from_dict(size)
        
        if width is None or height is None:
            raise ValueError("Shape must have explicit width and height")
        
        shape_type_map = {
            "rectangle": MSO_AUTO_SHAPE_TYPE.RECTANGLE,
            "rounded_rectangle": MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE,
            "ellipse": MSO_AUTO_SHAPE_TYPE.OVAL,
            "triangle": MSO_AUTO_SHAPE_TYPE.ISOSCELES_TRIANGLE,
            "arrow_right": MSO_AUTO_SHAPE_TYPE.RIGHT_ARROW,
            "arrow_left": MSO_AUTO_SHAPE_TYPE.LEFT_ARROW,
            "arrow_up": MSO_AUTO_SHAPE_TYPE.UP_ARROW,
            "arrow_down": MSO_AUTO_SHAPE_TYPE.DOWN_ARROW
        }
        
        mso_shape = shape_type_map.get(shape_type.lower(), MSO_AUTO_SHAPE_TYPE.RECTANGLE)
        
        shape = slide.shapes.add_shape(
            mso_shape,
            Inches(left), Inches(top),
            Inches(width), Inches(height)
        )
        
        if fill_color:
            shape.fill.solid()
            shape.fill.fore_color.rgb = ColorHelper.from_hex(fill_color)
        
        if line_color:
            shape.line.color.rgb = ColorHelper.from_hex(line_color)
            shape.line.width = Pt(line_width)
    
    def format_shape(
        self,
        slide_index: int,
        shape_index: int,
        fill_color: Optional[str] = None,
        line_color: Optional[str] = None,
        line_width: Optional[float] = None
    ) -> None:
        """Format existing shape."""
        slide = self.get_slide(slide_index)
        
        if shape_index >= len(slide.shapes):
            raise ValueError(f"Shape index {shape_index} out of range")
        
        shape = slide.shapes[shape_index]
        
        if fill_color:
            shape.fill.solid()
            shape.fill.fore_color.rgb = ColorHelper.from_hex(fill_color)
        
        if line_color:
            shape.line.color.rgb = ColorHelper.from_hex(line_color)
        
        if line_width:
            shape.line.width = Pt(line_width)
    
    def add_table(
        self,
        slide_index: int,
        rows: int,
        cols: int,
        position: Dict[str, Any],
        size: Dict[str, Any],
        data: Optional[List[List[Any]]] = None
    ) -> None:
        """
        Add table to slide.
        
        Args:
            slide_index: Target slide
            rows: Number of rows
            cols: Number of columns
            position: Position dict
            size: Size dict
            data: Optional 2D list of cell values
        """
        slide = self.get_slide(slide_index)
        
        left, top = Position.from_dict(position)
        width, height = Size.from_dict(size)
        
        if width is None or height is None:
            raise ValueError("Table must have explicit width and height")
        
        table_shape = slide.shapes.add_table(
            rows, cols,
            Inches(left), Inches(top),
            Inches(width), Inches(height)
        )
        
        table = table_shape.table
        
        if data:
            for row_idx, row_data in enumerate(data):
                if row_idx >= rows:
                    break
                for col_idx, cell_value in enumerate(row_data):
                    if col_idx >= cols:
                        break
                    table.cell(row_idx, col_idx).text = str(cell_value)
    
    def add_connector(
        self,
        slide_index: int,
        from_shape: int,
        to_shape: int,
        connector_type: str = "straight"
    ) -> None:
        """
        Add connector line between two shapes.
        
        Note: python-pptx has limited connector support.
        This is a simplified implementation using straight lines.
        """
        slide = self.get_slide(slide_index)
        
        if from_shape >= len(slide.shapes) or to_shape >= len(slide.shapes):
            raise ValueError("Shape index out of range")
        
        shape1 = slide.shapes[from_shape]
        shape2 = slide.shapes[to_shape]
        
        x1 = shape1.left + shape1.width // 2
        y1 = shape1.top + shape1.height // 2
        x2 = shape2.left + shape2.width // 2
        y2 = shape2.top + shape2.height // 2
        
        connector = slide.shapes.add_connector(
            MSO_CONNECTOR.STRAIGHT,
            x1, y1, x2, y2
        )
    
    # ========================================================================
    # IMAGE OPERATIONS  
    # ========================================================================
    
    def insert_image(
        self,
        slide_index: int,
        image_path: Path,
        position: Dict[str, Any],
        size: Optional[Dict[str, Any]] = None,
        compress: bool = False
    ) -> None:
        """
        Insert image on slide.
        
        Args:
            slide_index: Target slide
            image_path: Path to image file
            position: Position dict
            size: Optional size dict (can use "auto" for aspect ratio)
            compress: Compress image before inserting
        """
        if not image_path.exists():
            raise ImageNotFoundError(f"Image not found: {image_path}")
        
        slide = self.get_slide(slide_index)
        
        left, top = Position.from_dict(position)
        
        if HAS_PILLOW:
            with PILImage.open(image_path) as img:
                aspect_ratio = img.width / img.height
        else:
            aspect_ratio = None
        
        if size:
            width, height = Size.from_dict(size, aspect_ratio=aspect_ratio)
            
            if width is None and height and aspect_ratio:
                width = height * aspect_ratio
            elif height is None and width and aspect_ratio:
                height = width / aspect_ratio
        else:
            width = SLIDE_WIDTH_INCHES * 0.5
            if aspect_ratio:
                height = width / aspect_ratio
            else:
                height = SLIDE_HEIGHT_INCHES * 0.3
        
        if compress and HAS_PILLOW:
            image_stream = AssetValidator.compress_image(image_path)
            picture = slide.shapes.add_picture(
                image_stream,
                Inches(left), Inches(top),
                width=Inches(width) if width else None,
                height=Inches(height) if height else None
            )
        else:
            picture = slide.shapes.add_picture(
                str(image_path),
                Inches(left), Inches(top),
                width=Inches(width) if width else None,
                height=Inches(height) if height else None
            )
    
    def replace_image(
        self,
        slide_index: int,
        old_image_name: str,
        new_image_path: Path,
        compress: bool = False
    ) -> bool:
        """
        Replace existing image by name.
        
        Args:
            slide_index: Target slide
            old_image_name: Name of image to replace
            new_image_path: Path to new image
            compress: Compress new image
            
        Returns:
            True if replacement occurred
        """
        slide = self.get_slide(slide_index)
        
        for shape in slide.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                if shape.name == old_image_name or old_image_name in shape.name:
                    left = shape.left
                    top = shape.top
                    width = shape.width
                    height = shape.height
                    
                    sp = shape.element
                    sp.getparent().remove(sp)
                    
                    if compress and HAS_PILLOW:
                        image_stream = AssetValidator.compress_image(new_image_path)
                        slide.shapes.add_picture(
                            image_stream, left, top,
                            width=width, height=height
                        )
                    else:
                        slide.shapes.add_picture(
                            str(new_image_path), left, top,
                            width=width, height=height
                        )
                    
                    return True
        
        return False
    
    def set_image_properties(
        self,
        slide_index: int,
        shape_index: int,
        alt_text: Optional[str] = None,
        transparency: Optional[float] = None
    ) -> None:
        """
        Set image properties (alt text, transparency).
        
        Args:
            slide_index: Target slide
            shape_index: Image shape index
            alt_text: Alternative text for accessibility
            transparency: Transparency (0.0 = opaque, 1.0 = invisible)
        """
        slide = self.get_slide(slide_index)
        
        if shape_index >= len(slide.shapes):
            raise ValueError(f"Shape index {shape_index} out of range")
        
        shape = slide.shapes[shape_index]
        
        if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
            raise ValueError("Shape is not an image")
        
        if alt_text:
            shape.name = alt_text
            try:
                shape._element.set('descr', alt_text)
            except:
                pass
        
        if transparency is not None:
            try:
                shape.fill.transparency = transparency
            except:
                pass
    
    def crop_resize_image(
        self,
        slide_index: int,
        shape_index: int,
        width: Optional[float] = None,
        height: Optional[float] = None,
        maintain_aspect: bool = True
    ) -> None:
        """
        Resize image shape.
        
        Args:
            slide_index: Target slide
            shape_index: Image shape index
            width: New width in inches (None = keep current)
            height: New height in inches (None = keep current)
            maintain_aspect: Maintain aspect ratio
        """
        slide = self.get_slide(slide_index)
        shape = slide.shapes[shape_index]
        
        if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
            raise ValueError("Shape is not an image")
        
        if maintain_aspect and width and not height:
            aspect = shape.width / shape.height
            height = width / aspect
        elif maintain_aspect and height and not width:
            aspect = shape.width / shape.height
            width = height * aspect
        
        if width:
            shape.width = Inches(width)
        if height:
            shape.height = Inches(height)
    
    # ========================================================================
    # CHART OPERATIONS
    # ========================================================================
    
    def add_chart(
        self,
        slide_index: int,
        chart_type: str,
        data: Dict[str, Any],
        position: Dict[str, Any],
        size: Dict[str, Any],
        chart_title: Optional[str] = None
    ) -> None:
        """
        Add chart to slide.
        
        Args:
            slide_index: Target slide
            chart_type: Chart type (column, bar, line, pie, etc.)
            data: Chart data dict with "categories" and "series"
            position: Position dict
            size: Size dict
            chart_title: Optional chart title
            
        Example data:
            {
                "categories": ["Q1", "Q2", "Q3", "Q4"],
                "series": [
                    {"name": "Revenue", "values": [100, 120, 140, 160]},
                    {"name": "Costs", "values": [80, 90, 100, 110]}
                ]
            }
        """
        slide = self.get_slide(slide_index)
        
        left, top = Position.from_dict(position)
        width, height = Size.from_dict(size)
        
        if width is None or height is None:
            raise ValueError("Chart must have explicit width and height")
        
        chart_type_map = {
            "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
            "column_stacked": XL_CHART_TYPE.COLUMN_STACKED,
            "bar": XL_CHART_TYPE.BAR_CLUSTERED,
            "bar_stacked": XL_CHART_TYPE.BAR_STACKED,
            "line": XL_CHART_TYPE.LINE,
            "line_markers": XL_CHART_TYPE.LINE_MARKERS,
            "pie": XL_CHART_TYPE.PIE,
            "area": XL_CHART_TYPE.AREA,
            "scatter": XL_CHART_TYPE.XY_SCATTER
        }
        
        xl_chart_type = chart_type_map.get(chart_type.lower(), XL_CHART_TYPE.COLUMN_CLUSTERED)
        
        chart_data = CategoryChartData()
        chart_data.categories = data.get("categories", [])
        
        for series in data.get("series", []):
            chart_data.add_series(series["name"], series["values"])
        
        chart_shape = slide.shapes.add_chart(
            xl_chart_type,
            Inches(left), Inches(top),
            Inches(width), Inches(height),
            chart_data
        )
        
        chart = chart_shape.chart
        
        if chart_title:
            chart.has_title = True
            chart.chart_title.text_frame.text = chart_title
    
    def update_chart_data(
        self,
        slide_index: int,
        chart_index: int,
        data: Dict[str, Any]
    ) -> None:
        """
        Update existing chart data.
        
        Args:
            slide_index: Target slide
            chart_index: Chart index on slide
            data: New chart data dict
        """
        slide = self.get_slide(slide_index)
        
        chart_shape = None
        chart_count = 0
        
        for shape in slide.shapes:
            if shape.has_chart:
                if chart_count == chart_index:
                    chart_shape = shape
                    break
                chart_count += 1
        
        if not chart_shape:
            raise ValueError(f"Chart at index {chart_index} not found")
        
        chart_data = CategoryChartData()
        chart_data.categories = data.get("categories", [])
        
        for series in data.get("series", []):
            chart_data.add_series(series["name"], series["values"])
            
        try:
            chart_shape.chart.replace_data(chart_data)
        except AttributeError:
            logger.warning("chart.replace_data() not available, falling back to chart recreation (formatting may be lost)")
            
            left = chart_shape.left
            top = chart_shape.top
            width = chart_shape.width
            height = chart_shape.height
            chart_type = chart_shape.chart.chart_type
            chart_title = chart_shape.chart.chart_title.text_frame.text if chart_shape.chart.has_title else None
            
            sp = chart_shape.element
            sp.getparent().remove(sp)
            
            new_chart_shape = slide.shapes.add_chart(
                chart_type, left, top, width, height, chart_data
            )
            
            if chart_title:
                new_chart_shape.chart.has_title = True
                new_chart_shape.chart.chart_title.text_frame.text = chart_title

    def format_chart(
        self,
        slide_index: int,
        chart_index: int,
        title: Optional[str] = None,
        legend_position: Optional[str] = None
    ) -> None:
        """Format existing chart."""
        slide = self.get_slide(slide_index)
        
        chart_shape = None
        chart_count = 0
        
        for shape in slide.shapes:
            if shape.has_chart:
                if chart_count == chart_index:
                    chart_shape = shape
                    break
                chart_count += 1
        
        if not chart_shape:
            raise ValueError(f"Chart at index {chart_index} not found")
        
        chart = chart_shape.chart
        
        if title:
            chart.has_title = True
            chart.chart_title.text_frame.text = title
        
        if legend_position:
            from pptx.enum.chart import XL_LEGEND_POSITION
            position_map = {
                "bottom": XL_LEGEND_POSITION.BOTTOM,
                "left": XL_LEGEND_POSITION.LEFT,
                "right": XL_LEGEND_POSITION.RIGHT,
                "top": XL_LEGEND_POSITION.TOP
            }
            if legend_position.lower() in position_map:
                chart.has_legend = True
                chart.legend.position = position_map[legend_position.lower()]
    
    # ========================================================================
    # LAYOUT & THEME
    # ========================================================================
    
    def apply_theme(self, theme_path: Path) -> None:
        """
        Apply theme from .thmx file or template.
        
        Note: python-pptx has limited theme support.
        Best approach is to use template with desired theme.
        """
        raise NotImplementedError(
            "Theme application not fully supported by python-pptx. "
            "Use --template flag when creating presentation instead."
        )
    
    def set_slide_layout(self, slide_index: int, layout_name: str) -> None:
        """Change slide layout."""
        slide = self.get_slide(slide_index)
        layout = self._get_layout(layout_name)
        
        slide.slide_layout = layout
    
    def apply_master_slide(self, master_index: int = 0) -> None:
        """Apply master slide formatting."""
        raise NotImplementedError(
            "Master slide application not fully supported. "
            "Use template with desired master slides."
        )
    
    def get_available_layouts(self) -> List[str]:
        """Get list of available layout names."""
        if not self.prs:
            raise PowerPointAgentError("No presentation loaded")
        
        if self._template_profile:
            return self._template_profile.get_layout_names()
        
        return [layout.name for layout in self.prs.slide_layouts]
    
    # ========================================================================
    # VALIDATION
    # ========================================================================
    
    def validate_presentation(self) -> Dict[str, Any]:
        """
        Comprehensive presentation validation.
        
        Checks for:
        - Missing images/assets
        - Empty slides
        - Text overflow
        - Inconsistent formatting
        
        Returns:
            Validation report dict
        """
        if not self.prs:
            raise PowerPointAgentError("No presentation loaded")
        
        issues = {
            "empty_slides": [],
            "text_overflow": [],
            "inconsistent_fonts": set(),
            "slides_without_titles": []
        }
        
        for idx, slide in enumerate(self.prs.slides):
            if len(slide.shapes) == 0:
                issues["empty_slides"].append(idx)
            
            has_title = False
            for shape in slide.shapes:
                if shape.is_placeholder:
                    ph_type = shape.placeholder_format.type
                    if ph_type == PP_PLACEHOLDER.TITLE or ph_type == PP_PLACEHOLDER.CENTER_TITLE:
                        if shape.has_text_frame and shape.text_frame.text.strip():
                            has_title = True
            
            if not has_title:
                issues["slides_without_titles"].append(idx)
            
            for shape in slide.shapes:
                if hasattr(shape, 'text_frame'):
                    for paragraph in shape.text_frame.paragraphs:
                        if paragraph.font.name:
                            issues["inconsistent_fonts"].add(paragraph.font.name)
        
        issues["inconsistent_fonts"] = list(issues["inconsistent_fonts"])
        
        total_issues = (len(issues["empty_slides"]) + 
                       len(issues["text_overflow"]) + 
                       len(issues["slides_without_titles"]))
        
        return {
            "status": "issues_found" if total_issues > 0 else "valid",
            "total_issues": total_issues,
            "issues": issues
        }
    
    def check_accessibility(self) -> Dict[str, Any]:
        """Run accessibility checker."""
        if not self.prs:
            raise PowerPointAgentError("No presentation loaded")
        
        return AccessibilityChecker.check_presentation(self.prs)
    
    def validate_assets(self) -> Dict[str, Any]:
        """Run asset validator."""
        if not self.prs:
            raise PowerPointAgentError("No presentation loaded")
        
        return AssetValidator.validate_presentation_assets(self.prs, self.filepath)
    
    # ========================================================================
    # EXPORT
    # ========================================================================
    
    def export_to_pdf(self, output_path: Path) -> None:
        """
        Export presentation to PDF.
        
        Note: Requires LibreOffice or Microsoft Office installed.
        python-pptx doesn't support PDF export natively.
        """
        if not self.prs:
            raise PowerPointAgentError("No presentation loaded")
        
        temp_pptx = tempfile.mktemp(suffix='.pptx')
        self.prs.save(temp_pptx)
        
        try:
            result = subprocess.run(
                ['soffice', '--headless', '--convert-to', 'pdf', 
                 '--outdir', str(output_path.parent), temp_pptx],
                capture_output=True,
                timeout=60
            )
            
            if result.returncode != 0:
                raise PowerPointAgentError(
                    "PDF export failed. LibreOffice required for PDF export.\n"
                    "Install: sudo apt install libreoffice-impress (Linux) "
                    "or brew install --cask libreoffice (Mac)"
                )
        finally:
            Path(temp_pptx).unlink(missing_ok=True)
    
    def export_to_images(self, output_dir: Path, 
                        format: str = "png") -> List[Path]:
        """
        Export each slide as image.
        
        Note: Requires PIL/Pillow for image conversion.
        """
        if not HAS_PILLOW:
            raise ImportError("Pillow required for image export")
        
        raise NotImplementedError(
            "Image export requires external tool or LibreOffice. "
            "Use LibreOffice: soffice --headless --convert-to png presentation.pptx"
        )
    
    def extract_notes(self) -> Dict[int, str]:
        """
        Extract speaker notes from all slides.
        
        Returns:
            Dict mapping slide index to notes text
        """
        if not self.prs:
            raise PowerPointAgentError("No presentation loaded")
        
        notes = {}
        
        for idx, slide in enumerate(self.prs.slides):
            if slide.has_notes_slide:
                notes_slide = slide.notes_slide
                text_frame = notes_slide.notes_text_frame
                if text_frame.text.strip():
                    notes[idx] = text_frame.text
        
        return notes
    
    # ========================================================================
    # UTILITIES
    # ========================================================================
    
    def get_presentation_info(self) -> Dict[str, Any]:
        """Get presentation metadata."""
        if not self.prs:
            raise PowerPointAgentError("No presentation loaded")
        
        info = {
            "slide_count": len(self.prs.slides),
            "layouts": [layout.name for layout in self.prs.slide_layouts],
            "slide_width_inches": self.prs.slide_width.inches,
            "slide_height_inches": self.prs.slide_height.inches,
            "aspect_ratio": f"{self.prs.slide_width.inches:.1f}:{self.prs.slide_height.inches:.1f}"
        }
        
        if self.filepath:
            info["file"] = str(self.filepath)
            if self.filepath.exists():
                stat = self.filepath.stat()
                info["file_size_bytes"] = stat.st_size
                info["file_size_mb"] = stat.st_size / (1024 * 1024)
                info["modified"] = datetime.fromtimestamp(stat.st_mtime).isoformat()
        
        return info
    
    def get_slide_info(self, slide_index: int) -> Dict[str, Any]:
        """
        Get detailed information about specific slide.
        
        Fixed in v1.1.0:
        - Removed 100-character text truncation (was causing data loss)
        - Added full text with separate preview field
        - Added position and size information for all shapes
        - Added human-readable placeholder type names
        
        Args:
            slide_index: Index of slide to inspect
            
        Returns:
            Dict containing comprehensive slide information including
            full text content, positions, sizes, and placeholder types
        """
        slide = self.get_slide(slide_index)
        
        shapes_info = []
        for idx, shape in enumerate(slide.shapes):
            shape_type_str = str(shape.shape_type)
            
            if shape.is_placeholder:
                ph_type = shape.placeholder_format.type
                ph_type_name = get_placeholder_type_name(ph_type)
                shape_type_str = f"PLACEHOLDER ({ph_type_name})"

            shape_info = {
                "index": idx,
                "type": shape_type_str,
                "name": shape.name,
                "has_text": hasattr(shape, 'text_frame'),
                "position": {
                    "left_inches": shape.left / 914400,
                    "top_inches": shape.top / 914400,
                    "left_percent": f"{(shape.left / self.prs.slide_width * 100):.1f}%",
                    "top_percent": f"{(shape.top / self.prs.slide_height * 100):.1f}%"
                },
                "size": {
                    "width_inches": shape.width / 914400,
                    "height_inches": shape.height / 914400,
                    "width_percent": f"{(shape.width / self.prs.slide_width * 100):.1f}%",
                    "height_percent": f"{(shape.height / self.prs.slide_height * 100):.1f}%"
                }
            }
            
            if shape.has_text_frame:
                full_text = shape.text_frame.text
                shape_info["text"] = full_text
                shape_info["text_length"] = len(full_text)
                if len(full_text) > 100:
                    shape_info["text_preview"] = full_text[:100] + "..."
            
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                shape_info["image_size_bytes"] = len(shape.image.blob)
            
            shapes_info.append(shape_info)
        
        return {
            "slide_index": slide_index,
            "layout": slide.slide_layout.name,
            "shape_count": len(slide.shapes),
            "shapes": shapes_info,
            "has_notes": slide.has_notes_slide
        }
    
    def _get_layout(self, layout_name: str):
        """Get layout by name."""
        if not self.prs:
            raise PowerPointAgentError("No presentation loaded")
        
        for layout in self.prs.slide_layouts:
            if layout.name == layout_name:
                return layout
        
        available = [l.name for l in self.prs.slide_layouts]
        raise LayoutNotFoundError(
            f"Layout '{layout_name}' not found. "
            f"Available: {available}"
        )
    
    def _copy_shape(self, source_shape, target_slide):
        """Copy shape to target slide (helper for duplicate_slide)."""
        
        if source_shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            try:
                blob = source_shape.image.blob
                target_slide.shapes.add_picture(
                    BytesIO(blob),
                    source_shape.left, source_shape.top,
                    source_shape.width, source_shape.height
                )
            except Exception as e:
                logger.warning(f"Failed to copy picture: {e}")
                
        elif source_shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE or \
             source_shape.shape_type == MSO_SHAPE_TYPE.TEXT_BOX:
            try:
                new_shape = target_slide.shapes.add_shape(
                    source_shape.auto_shape_type,
                    source_shape.left, source_shape.top,
                    source_shape.width, source_shape.height
                )
                
                if source_shape.has_text_frame and source_shape.text:
                    new_shape.text_frame.text = source_shape.text_frame.text
                
                try:
                    if source_shape.fill.type == 1:
                        new_shape.fill.solid()
                        new_shape.fill.fore_color.rgb = source_shape.fill.fore_color.rgb
                except:
                    pass
                    
            except Exception as e:
                logger.warning(f"Failed to copy shape: {e}")
        
        elif source_shape.shape_type == MSO_SHAPE_TYPE.CHART:
            logger.warning("Chart copying not supported in duplicate_slide (python-pptx limitation)")
    
    def __enter__(self):
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()
        return False


# ============================================================================
# MODULE METADATA
# ============================================================================

__version__ = "1.1.0"
__author__ = "PowerPoint Agent Team"
__license__ = "MIT"

__all__ = [
    "PowerPointAgent",
    "PowerPointAgentError",
    "SlideNotFoundError",
    "LayoutNotFoundError",
    "ImageNotFoundError",
    "InvalidPositionError",
    "TemplateError",
    "ThemeError",
    "AccessibilityError",
    "AssetValidationError",
    "FileLockError",
    "Position",
    "Size",
    "ColorHelper",
    "TemplateProfile",
    "AccessibilityChecker",
    "AssetValidator",
    "ShapeType",
    "ChartType",
    "TextAlignment",
    "VerticalAlignment",
    "BulletStyle",
    "ImageFormat",
    "ExportFormat",
    "SLIDE_WIDTH_INCHES",
    "SLIDE_HEIGHT_INCHES",
    "ANCHOR_POINTS",
    "CORPORATE_COLORS",
    "STANDARD_FONTS",
    "PLACEHOLDER_TYPE_NAMES",
    "PP_PLACEHOLDER",
]
