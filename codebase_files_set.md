# core/__init__.py
```py
"""
PowerPoint Agent Core Package
"""

from .powerpoint_agent_core import (
    # Core class
    PowerPointAgent,
    
    # Exceptions
    PowerPointAgentError,
    SlideNotFoundError,
    LayoutNotFoundError,
    ImageNotFoundError,
    InvalidPositionError,
    TemplateError,
    ThemeError,
    AccessibilityError,
    AssetValidationError,
    FileLockError,
    
    # Helpers
    Position,
    Size,
    ColorHelper,
    TemplateProfile,
    AccessibilityChecker,
    AssetValidator,
    
    # Enums
    ShapeType,
    ChartType,
    TextAlignment,
    VerticalAlignment,
    BulletStyle,
    ImageFormat,
    ExportFormat,
    
    # Constants
    SLIDE_WIDTH_INCHES,
    SLIDE_HEIGHT_INCHES,
    ANCHOR_POINTS,
    CORPORATE_COLORS,
    STANDARD_FONTS,
)

__version__ = "1.0.0"
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
]

```

# core/powerpoint_agent_core.py
```py
#!/usr/bin/env python3
"""
PowerPoint Agent Core Library
Production-grade PowerPoint manipulation with validation and accessibility

This is the foundational library used by all CLI tools.
Designed for stateless, security-hardened PowerPoint operations.

Author: PowerPoint Agent Team
License: MIT
Version: 1.0.0
"""

import re
import json
import tempfile
import shutil
import threading
import time
from pathlib import Path
from typing import Any, Dict, List, Optional, Union, Tuple, Set
from enum import Enum
from datetime import datetime
from io import BytesIO

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu
    from pptx.enum.shapes import MSO_SHAPE_TYPE, MSO_AUTO_SHAPE_TYPE
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
        # Format 1: Absolute inches
        if "left" in pos_dict and "top" in pos_dict:
            left = Position._parse_dimension(pos_dict["left"], slide_width)
            top = Position._parse_dimension(pos_dict["top"], slide_height)
            return (left, top)
        
        # Format 3: Anchor-based
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
        
        # Format 4: Grid system
        if "grid_row" in pos_dict and "grid_col" in pos_dict:
            grid_size = pos_dict.get("grid_size", 12)
            cell_width = slide_width / grid_size
            cell_height = slide_height / grid_size
            left = pos_dict["grid_col"] * cell_width
            top = pos_dict["grid_row"] * cell_height
            return (left, top)
        
        # Format 5: Excel-like grid
        if "grid" in pos_dict:
            grid_ref = pos_dict["grid"].upper()
            match = re.match(r'^([A-Z]+)(\d+)$', grid_ref)
            if not match:
                raise InvalidPositionError(f"Invalid grid reference: {grid_ref}")
            
            col_str, row_str = match.groups()
            
            # Convert column letters to number
            col_num = 0
            for char in col_str:
                col_num = col_num * 26 + (ord(char) - ord('A') + 1)
            
            row_num = int(row_str)
            
            # Default 12x12 grid
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
        
        # Parse width
        if width_spec == "auto":
            width = None
        elif width_spec:
            width = Position._parse_dimension(width_spec, slide_width)
        else:
            width = None
        
        # Parse height
        if height_spec == "auto":
            height = None
        elif height_spec:
            height = Position._parse_dimension(height_spec, slide_height)
        else:
            height = None
        
        # Handle "auto" with aspect ratio
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
        if len(hex_color) != 6:
            raise ValueError(f"Invalid hex color: {hex_color}")
        
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
        def _channel(c):
            c = c / 255.0
            return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
        
        r = _channel(rgb_color.r)
        g = _channel(rgb_color.g)
        b = _channel(rgb_color.b)
        
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
        
        # Capture slide layouts
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
        
        # Capture theme colors
        try:
            theme = prs.slide_master.theme
            for idx, color in enumerate(theme.theme_color_scheme):
                self.theme_colors[f"color_{idx}"] = ColorHelper.to_hex(color)
        except:
            pass
        
        # Capture default fonts
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
            # Check for title
            has_title = False
            for shape in slide.shapes:
                if shape.is_placeholder:
                    # Check for TITLE (1) or CENTER_TITLE (3)
                    if shape.placeholder_format.type == 1 or shape.placeholder_format.type == 3:
                        if shape.has_text_frame and shape.text_frame.text.strip():
                            has_title = True
                            break
            
            if not has_title:
                issues["missing_titles"].append(slide_idx)
            
            # Check shapes
            for shape_idx, shape in enumerate(slide.shapes):
                # Check images for alt text
                if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                    if not shape.name or shape.name.startswith("Picture"):
                        issues["missing_alt_text"].append({
                            "slide": slide_idx,
                            "shape": shape_idx,
                            "shape_name": shape.name
                        })
                
                # Check text contrast
                if hasattr(shape, 'text_frame'):
                    try:
                        AccessibilityChecker._check_text_contrast(
                            shape, slide_idx, shape_idx, issues
                        )
                    except:
                        pass
        
        # Calculate severity
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
            if para.font.color and para.font.color.type == 1:  # RGB
                text_color = para.font.color.rgb
                
                # Assume white background if no fill
                bg_color = RGBColor(255, 255, 255)
                
                if shape.fill.type == 1:  # Solid fill
                    bg_color = shape.fill.fore_color.rgb
                
                # Check if large text (18pt = 24px or 14pt bold = 18.67px)
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
                        
                        # Check resolution (should be at least 96 DPI)
                        # DPI = pixels / inches
                        if hasattr(image, 'dpi'):
                            if image.dpi[0] < 96 or image.dpi[1] < 96:
                                issues["low_resolution_images"].append({
                                    "slide": slide_idx,
                                    "shape": shape_idx,
                                    "dpi": image.dpi
                                })
                        
                        # Check file size (warn if >2MB)
                        image_size = len(image.blob)
                        issues["total_embedded_size"] += image_size
                        
                        if image_size > 2 * 1024 * 1024:  # 2MB
                            issues["large_images"].append({
                                "slide": slide_idx,
                                "shape": shape_idx,
                                "size_mb": image_size / (1024 * 1024)
                            })
                    except:
                        pass
        
        # Check total file size if filepath provided
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
            # Resize if too large
            if img.width > max_width:
                ratio = max_width / img.width
                new_height = int(img.height * ratio)
                img = img.resize((max_width, new_height), PILImage.LANCZOS)
            
            # Convert to RGB if necessary
            if img.mode in ('RGBA', 'LA', 'P'):
                background = PILImage.new('RGB', img.size, (255, 255, 255))
                if img.mode == 'P':
                    img = img.convert('RGBA')
                background.paste(img, mask=img.split()[-1] if img.mode == 'RGBA' else None)
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
            # Insert at specific position (requires XML manipulation)
            slide = self.prs.slides.add_slide(layout)
            # Move to desired position
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
        
        # Create new slide with same layout
        layout = source_slide.slide_layout
        new_slide = self.prs.slides.add_slide(layout)
        
        # Copy all shapes from source
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
        
        # Format paragraph
        paragraph = text_frame.paragraphs[0]
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
                # Title can be TITLE (1) or CENTER_TITLE (3)
                if ph_type == 1 or ph_type == 3:
                    title_shape = shape
                elif ph_type == 2:  # Subtitle
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
            
            # Set bullet style
            if bullet_style == "bullet":
                # Use bullet character (default)
                pass
            elif bullet_style == "numbered":
                # Note: python-pptx has limited numbered list support
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
                # Skip shapes that definitely don't have text
                if not shape.has_text_frame:
                    continue
                
                try:
                    text_frame = shape.text_frame
                except Exception:
                    continue
                
                # Strategy 1: Try to replace in individual runs (Preserves Formatting)
                replacements_in_runs = 0
                for paragraph in text_frame.paragraphs:
                    for run in paragraph.runs:
                        # Check match based on case sensitivity
                        if match_case:
                            if find in run.text:
                                run.text = run.text.replace(find, replace)
                                replacements_in_runs += 1
                        else:
                            if find.lower() in run.text.lower():
                                pattern = re.compile(re.escape(find), re.IGNORECASE)
                                # Check if it actually changes anything
                                if pattern.search(run.text):
                                    new_text = pattern.sub(replace, run.text)
                                    run.text = new_text
                                    replacements_in_runs += 1
                
                if replacements_in_runs > 0:
                    count += replacements_in_runs
                    continue
                
                # Strategy 2: Fallback to Shape-level replacement (May lose formatting)
                # Only if we didn't find it in runs, but it might exist across runs
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
                        # This is destructive to formatting, but ensures replacement happens
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
                except Exception:
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
        
        # Map shape type string to MSO_AUTO_SHAPE_TYPE constant
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
        
        # Format shape
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
        
        # Populate data if provided
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
        This is a simplified implementation.
        """
        slide = self.get_slide(slide_index)
        
        if from_shape >= len(slide.shapes) or to_shape >= len(slide.shapes):
            raise ValueError("Shape index out of range")
        
        shape1 = slide.shapes[from_shape]
        shape2 = slide.shapes[to_shape]
        
        # Calculate center points
        x1 = shape1.left + shape1.width // 2
        y1 = shape1.top + shape1.height // 2
        x2 = shape2.left + shape2.width // 2
        y2 = shape2.top + shape2.height // 2
        
        # Add line connector
        connector = slide.shapes.add_connector(
            1,  # Straight connector
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
        
        # Get image dimensions for aspect ratio
        if HAS_PILLOW:
            with PILImage.open(image_path) as img:
                aspect_ratio = img.width / img.height
        else:
            aspect_ratio = None
        
        # Handle size
        if size:
            width, height = Size.from_dict(size, aspect_ratio=aspect_ratio)
            
            # If one dimension is None (auto), calculate from aspect ratio
            if width is None and height and aspect_ratio:
                width = height * aspect_ratio
            elif height is None and width and aspect_ratio:
                height = width / aspect_ratio
        else:
            # Default to 50% of slide width
            width = SLIDE_WIDTH_INCHES * 0.5
            if aspect_ratio:
                height = width / aspect_ratio
            else:
                height = SLIDE_HEIGHT_INCHES * 0.3
        
        # Compress if requested
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
                    # Get current position and size
                    left = shape.left
                    top = shape.top
                    width = shape.width
                    height = shape.height
                    
                    # Remove old image
                    sp = shape.element
                    sp.getparent().remove(sp)
                    
                    # Insert new image
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
            # Set alt text (accessible name)
            shape.name = alt_text
            # Also try to set description if available
            try:
                shape._element.set('descr', alt_text)
            except:
                pass
        
        if transparency is not None:
            # Note: python-pptx has limited transparency support
            # This may not work in all cases
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
        
        # Map chart type
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
        
        # Prepare chart data
        chart_data = CategoryChartData()
        chart_data.categories = data.get("categories", [])
        
        for series in data.get("series", []):
            chart_data.add_series(series["name"], series["values"])
        
        # Add chart
        chart_shape = slide.shapes.add_chart(
            xl_chart_type,
            Inches(left), Inches(top),
            Inches(width), Inches(height),
            chart_data
        )
        
        chart = chart_shape.chart
        
        # Set title if provided
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
        
        # Prepare new chart data
        chart_data = CategoryChartData()
        chart_data.categories = data.get("categories", [])
        
        for series in data.get("series", []):
            chart_data.add_series(series["name"], series["values"])
            
        # Update chart data
        # Note: replace_data is available in newer python-pptx versions
        # and is preferred over recreation as it preserves formatting.
        try:
            chart_shape.chart.replace_data(chart_data)
        except AttributeError:
            # Fallback for older versions or if replace_data fails: Recreate chart
            # This is a destructive operation and will lose custom formatting
            print("WARNING: chart.replace_data failed, falling back to recreation (formatting may be lost)")
            
            # 1. Extract properties
            left = chart_shape.left
            top = chart_shape.top
            width = chart_shape.width
            height = chart_shape.height
            chart_type = chart_shape.chart.chart_type
            chart_title = chart_shape.chart.chart_title.text_frame.text if chart_shape.chart.has_title else None
            
            # 2. Delete existing
            sp = chart_shape.element
            sp.getparent().remove(sp)
            
            # 3. Create new
            new_chart_shape = slide.shapes.add_chart(
                chart_type, left, top, width, height, chart_data
            )
            
            # 4. Restore basic properties
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
        
        # Note: Changing layout can lose content
        # This is a limitation of python-pptx
        slide.slide_layout = layout
    
    def apply_master_slide(self, master_index: int = 0) -> None:
        """Apply master slide formatting."""
        # Note: Master slide manipulation is limited in python-pptx
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
            # Check for empty slides
            if len(slide.shapes) == 0:
                issues["empty_slides"].append(idx)
            
            # Check for title
            has_title = False
            for shape in slide.shapes:
                if shape.is_placeholder:
                    ph_type = shape.placeholder_format.type
                    # Check for TITLE (1) or CENTER_TITLE (3)
                    if ph_type == 1 or ph_type == 3:
                        if shape.has_text_frame and shape.text_frame.text.strip():
                            has_title = True
            
            if not has_title:
                issues["slides_without_titles"].append(idx)
            
            # Collect fonts
            for shape in slide.shapes:
                if hasattr(shape, 'text_frame'):
                    for paragraph in shape.text_frame.paragraphs:
                        if paragraph.font.name:
                            issues["inconsistent_fonts"].add(paragraph.font.name)
        
        # Convert set to list for JSON serialization
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
        
        # Save current presentation
        temp_pptx = tempfile.mktemp(suffix='.pptx')
        self.prs.save(temp_pptx)
        
        try:
            # Try LibreOffice conversion
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
        """Get detailed information about specific slide."""
        slide = self.get_slide(slide_index)
        
        shapes_info = []
        for idx, shape in enumerate(slide.shapes):
            shape_info = {
                "index": idx,
                "type": str(shape.shape_type),
                "name": shape.name,
                "has_text": hasattr(shape, 'text_frame')
            }
            
            if shape.has_text_frame:
                shape_info["text"] = shape.text_frame.text[:100]  # First 100 chars
            
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
        # Basic implementation for common shape types
        
        # 1. Pictures
        if source_shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            try:
                blob = source_shape.image.blob
                target_slide.shapes.add_picture(
                    BytesIO(blob),
                    source_shape.left, source_shape.top,
                    source_shape.width, source_shape.height
                )
            except Exception as e:
                print(f"WARNING: Failed to copy picture: {e}")
                
        # 2. AutoShapes (Rectangles, Ellipses, etc.) & TextBoxes
        elif source_shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE or \
             source_shape.shape_type == MSO_SHAPE_TYPE.TEXT_BOX:
            try:
                # Create new shape
                new_shape = target_slide.shapes.add_shape(
                    source_shape.auto_shape_type,
                    source_shape.left, source_shape.top,
                    source_shape.width, source_shape.height
                )
                
                # Copy text
                if source_shape.has_text_frame and source_shape.text:
                    new_shape.text_frame.text = source_shape.text_frame.text
                    # Note: This loses run-level formatting. 
                    # A full implementation would iterate paragraphs/runs.
                
                # Copy fill (basic)
                try:
                    if source_shape.fill.type == 1: # Solid fill
                        new_shape.fill.solid()
                        new_shape.fill.fore_color.rgb = source_shape.fill.fore_color.rgb
                except:
                    pass
                    
            except Exception as e:
                print(f"WARNING: Failed to copy shape: {e}")
        
        # 3. Charts (Complex - Try to recreate if possible, else skip)
        elif source_shape.shape_type == MSO_SHAPE_TYPE.CHART:
            # Cloning charts is very hard in python-pptx without low-level XML
            # For now, we skip to avoid errors
            print("WARNING: Chart copying not supported in duplicate_slide")
            
        else:
            # Try generic copy for other shapes if possible, or skip
            pass
    
    def __enter__(self):
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()
        return False


# ============================================================================
# MODULE METADATA
# ============================================================================

__version__ = "1.0.0"
__author__ = "PowerPoint Agent Team"
__license__ = "MIT"

__all__ = [
    # Core class
    "PowerPointAgent",
    
    # Exceptions
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
    
    # Helpers
    "Position",
    "Size",
    "ColorHelper",
    "TemplateProfile",
    "AccessibilityChecker",
    "AssetValidator",
    
    # Enums
    "ShapeType",
    "ChartType",
    "TextAlignment",
    "VerticalAlignment",
    "BulletStyle",
    "ImageFormat",
    "ExportFormat",
    
    # Constants
    "SLIDE_WIDTH_INCHES",
    "SLIDE_HEIGHT_INCHES",
    "ANCHOR_POINTS",
    "CORPORATE_COLORS",
    "STANDARD_FONTS",
]

```

# tools/ppt_add_bullet_list.py
```py
#!/usr/bin/env python3
"""
PowerPoint Add Bullet List Tool
Add bullet or numbered list to slide

Usage:
    uv python ppt_add_bullet_list.py --file presentation.pptx --slide 0 --items "Item 1,Item 2,Item 3" --position '{"left":"10%","top":"25%"}' --size '{"width":"80%","height":"60%"}' --json

Exit Codes:
    0: Success
    1: Error occurred
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any, List

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError, SlideNotFoundError
)


def add_bullet_list(
    filepath: Path,
    slide_index: int,
    items: List[str],
    position: Dict[str, Any],
    size: Dict[str, Any],
    bullet_style: str = "bullet",
    font_size: int = 18,
    font_name: str = "Calibri",
    color: str = None,
    line_spacing: float = 1.0
) -> Dict[str, Any]:
    """Add bullet or numbered list to slide."""
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if not items:
        raise ValueError("At least one item required")
    
    if len(items) > 20:
        raise ValueError("Maximum 20 items per list (readability limit)")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Validate slide index
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides-1})"
            )
        
        # Add bullet list
        agent.add_bullet_list(
            slide_index=slide_index,
            items=items,
            position=position,
            size=size,
            bullet_style=bullet_style,
            font_size=font_size
        )
        
        # Get the last added shape for additional formatting if needed
        slide_info = agent.get_slide_info(slide_index)
        last_shape_idx = slide_info["shape_count"] - 1
        
        # Apply color if specified
        if color:
            agent.format_text(
                slide_index=slide_index,
                shape_index=last_shape_idx,
                color=color
            )
        
        # Save
        agent.save()
    
    return {
        "status": "success",
        "file": str(filepath),
        "slide_index": slide_index,
        "items_added": len(items),
        "items": items,
        "bullet_style": bullet_style,
        "formatting": {
            "font_size": font_size,
            "font_name": font_name,
            "color": color,
            "line_spacing": line_spacing
        }
    }


def main():
    parser = argparse.ArgumentParser(
        description="Add bullet or numbered list to PowerPoint slide",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
List Formats:
  1. Comma-separated: "Item 1,Item 2,Item 3"
  2. Newline-separated: "Item 1\\nItem 2\\nItem 3"
  3. JSON array from file: --items-file items.json

Bullet Styles:
  - bullet: Traditional bullet points (default)
  - numbered: 1. 2. 3. numbering
  - none: Plain list without bullets

Examples:
  # Simple bullet list
  uv python ppt_add_bullet_list.py \\
    --file presentation.pptx \\
    --slide 1 \\
    --items "Revenue up 45%,Customer growth 60%,Market share 23%" \\
    --position '{"left":"10%","top":"25%"}' \\
    --size '{"width":"80%","height":"60%"}' \\
    --json
  
  # Numbered list with custom formatting
  uv python ppt_add_bullet_list.py \\
    --file presentation.pptx \\
    --slide 2 \\
    --items "Define objectives,Analyze market,Develop strategy,Execute plan" \\
    --bullet-style numbered \\
    --position '{"left":"15%","top":"30%"}' \\
    --size '{"width":"70%","height":"50%"}' \\
    --font-size 20 \\
    --color "#0070C0" \\
    --json
  
  # Multi-line items (agenda)
  uv python ppt_add_bullet_list.py \\
    --file presentation.pptx \\
    --slide 0 \\
    --items "Introduction and welcome,Q4 financial results,2024 strategic priorities,Q&A session" \\
    --position '{"grid":"B3"}' \\
    --size '{"width":"60%","height":"50%"}' \\
    --font-size 22 \\
    --json
  
  # Key takeaways (centered)
  uv python ppt_add_bullet_list.py \\
    --file presentation.pptx \\
    --slide 10 \\
    --items "Strong Q4 performance,Exceeded revenue targets,Positive market reception" \\
    --position '{"anchor":"center","offset_x":0,"offset_y":0}' \\
    --size '{"width":"70%","height":"40%"}' \\
    --font-size 24 \\
    --color "#00B050" \\
    --json
  
  # From JSON file
  echo '["First point", "Second point", "Third point"]' > items.json
  uv python ppt_add_bullet_list.py \\
    --file presentation.pptx \\
    --slide 3 \\
    --items-file items.json \\
    --position '{"left":"10%","top":"25%"}' \\
    --size '{"width":"80%","height":"60%"}' \\
    --json

Best Practices:
  - Keep items concise (max 2 lines per bullet)
  - Use 3-7 items per slide for readability
  - Start each item with action verb for impact
  - Use parallel structure (all items same format)
  - Font size 18-24pt for body text
  - Leave white space (don't fill entire slide)

Formatting Tips:
  - Use color for emphasis (sparingly)
  - Increase font size for key points
  - Use numbered lists for sequential steps
  - Use bullet lists for unordered points
  - Consistent spacing improves readability
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file path'
    )
    
    parser.add_argument(
        '--slide',
        required=True,
        type=int,
        help='Slide index (0-based)'
    )
    
    parser.add_argument(
        '--items',
        help='Comma-separated list items'
    )
    
    parser.add_argument(
        '--items-file',
        type=Path,
        help='JSON file with array of items'
    )
    
    parser.add_argument(
        '--position',
        required=True,
        type=json.loads,
        help='Position dict (JSON string)'
    )
    
    parser.add_argument(
        '--size',
        type=json.loads,
        help='Size dict (JSON string)'
    )
    
    parser.add_argument(
        '--bullet-style',
        choices=['bullet', 'numbered', 'none'],
        default='bullet',
        help='Bullet style (default: bullet)'
    )
    
    parser.add_argument(
        '--font-size',
        type=int,
        default=18,
        help='Font size in points (default: 18)'
    )
    
    parser.add_argument(
        '--font-name',
        default='Calibri',
        help='Font name (default: Calibri)'
    )
    
    parser.add_argument(
        '--color',
        help='Text color (hex, e.g., #0070C0)'
    )
    
    parser.add_argument(
        '--line-spacing',
        type=float,
        default=1.0,
        help='Line spacing multiplier (default: 1.0)'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        help='Output JSON response'
    )
    
    args = parser.parse_args()
    
    try:
        # Parse items
        if args.items_file:
            if not args.items_file.exists():
                raise FileNotFoundError(f"Items file not found: {args.items_file}")
            with open(args.items_file, 'r') as f:
                items = json.load(f)
            if not isinstance(items, list):
                raise ValueError("Items file must contain JSON array")
        elif args.items:
            # Split by comma or newline
            if '\\n' in args.items:
                items = args.items.split('\\n')
            else:
                items = args.items.split(',')
            items = [item.strip() for item in items if item.strip()]
        else:
            raise ValueError("Either --items or --items-file required")
            
        # Handle optional size and merge from position
        size = args.size if args.size else {}
        position = args.position
        
        if "width" in position and "width" not in size:
            size["width"] = position["width"]
        if "height" in position and "height" not in size:
            size["height"] = position["height"]
            
        # Apply defaults if still missing
        if "width" not in size:
            size["width"] = "80%"
        if "height" not in size:
            size["height"] = "50%"
        
        result = add_bullet_list(
            filepath=args.file,
            slide_index=args.slide,
            items=items,
            position=position,
            size=size,
            bullet_style=args.bullet_style,
            font_size=args.font_size,
            font_name=args.font_name,
            color=args.color,
            line_spacing=args.line_spacing
        )
        
        if args.json:
            print(json.dumps(result, indent=2))
        else:
            print(f"✅ Added {result['items_added']}-item {result['bullet_style']} list to slide {result['slide_index']}")
            print(f"   Items:")
            for i, item in enumerate(result['items'][:5], 1):
                prefix = f"{i}." if args.bullet_style == 'numbered' else "•"
                print(f"     {prefix} {item}")
            if len(result['items']) > 5:
                print(f"     ... and {len(result['items']) - 5} more")
        
        sys.exit(0)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__
        }
        
        if args.json:
            print(json.dumps(error_result, indent=2))
        else:
            print(f"❌ Error: {e}", file=sys.stderr)
        
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_add_chart.py
```py
#!/usr/bin/env python3
"""
PowerPoint Add Chart Tool
Add data visualization chart to slide

Usage:
    uv python ppt_add_chart.py --file presentation.pptx --slide 1 --chart-type column --data chart_data.json --position '{"left":"10%","top":"20%"}' --size '{"width":"80%","height":"60%"}' --json

Exit Codes:
    0: Success
    1: Error occurred
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any, List

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError, SlideNotFoundError
)


def add_chart(
    filepath: Path,
    slide_index: int,
    chart_type: str,
    data: Dict[str, Any],
    position: Dict[str, Any],
    size: Dict[str, Any],
    chart_title: str = None
) -> Dict[str, Any]:
    """Add chart to slide."""
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    # Validate data structure
    if "categories" not in data:
        raise ValueError("Data must contain 'categories' key")
    
    if "series" not in data or not data["series"]:
        raise ValueError("Data must contain at least one series")
    
    # Validate all series have same length as categories
    cat_len = len(data["categories"])
    for series in data["series"]:
        if len(series.get("values", [])) != cat_len:
            raise ValueError(
                f"Series '{series.get('name', 'unnamed')}' has {len(series['values'])} values, "
                f"but {cat_len} categories. Must match."
            )
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Validate slide index
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides-1})"
            )
        
        # Add chart
        agent.add_chart(
            slide_index=slide_index,
            chart_type=chart_type,
            data=data,
            position=position,
            size=size,
            chart_title=chart_title
        )
        
        # Save
        agent.save()
    
    return {
        "status": "success",
        "file": str(filepath),
        "slide_index": slide_index,
        "chart_type": chart_type,
        "chart_title": chart_title,
        "categories": len(data["categories"]),
        "series": len(data["series"]),
        "data_points": sum(len(s["values"]) for s in data["series"])
    }


def main():
    parser = argparse.ArgumentParser(
        description="Add data visualization chart to PowerPoint slide",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Chart Types:
  - column: Vertical bars (compare across categories)
  - column_stacked: Stacked vertical bars (show composition)
  - bar: Horizontal bars (compare items)
  - bar_stacked: Stacked horizontal bars
  - line: Line chart (show trends over time)
  - line_markers: Line with data point markers
  - pie: Pie chart (show proportions, single series only)
  - area: Area chart (emphasize magnitude of change)
  - scatter: Scatter plot (show relationships)

Data Format (JSON):
{
  "categories": ["Q1", "Q2", "Q3", "Q4"],
  "series": [
    {"name": "Revenue", "values": [100, 120, 140, 160]},
    {"name": "Costs", "values": [80, 90, 100, 110]}
  ]
}

Examples:
  # Revenue growth chart
  cat > revenue_data.json << 'EOF'
{
  "categories": ["Q1 2023", "Q2 2023", "Q3 2023", "Q4 2023", "Q1 2024"],
  "series": [
    {"name": "Revenue ($M)", "values": [12.5, 15.2, 18.7, 22.1, 25.8]},
    {"name": "Target ($M)", "values": [15, 16, 18, 20, 24]}
  ]
}
EOF
  
  uv python ppt_add_chart.py \\
    --file presentation.pptx \\
    --slide 1 \\
    --chart-type column \\
    --data revenue_data.json \\
    --position '{"left":"10%","top":"20%"}' \\
    --size '{"width":"80%","height":"60%"}' \\
    --title "Revenue Growth Trajectory" \\
    --json
  
  # Market share pie chart
  cat > market_data.json << 'EOF'
{
  "categories": ["Our Company", "Competitor A", "Competitor B", "Others"],
  "series": [
    {"name": "Market Share", "values": [35, 28, 22, 15]}
  ]
}
EOF
  
  uv python ppt_add_chart.py \\
    --file presentation.pptx \\
    --slide 2 \\
    --chart-type pie \\
    --data market_data.json \\
    --position '{"anchor":"center"}' \\
    --size '{"width":"60%","height":"60%"}' \\
    --title "Market Share Distribution" \\
    --json
  
  # Year-over-year comparison (bar chart)
  cat > yoy_data.json << 'EOF'
{
  "categories": ["Revenue", "Profit", "Customers", "Employees"],
  "series": [
    {"name": "2023", "values": [100, 25, 1000, 150]},
    {"name": "2024", "values": [145, 38, 1450, 200]}
  ]
}
EOF
  
  uv python ppt_add_chart.py \\
    --file presentation.pptx \\
    --slide 3 \\
    --chart-type bar \\
    --data yoy_data.json \\
    --position '{"left":"5%","top":"25%"}' \\
    --size '{"width":"90%","height":"60%"}' \\
    --title "Year-over-Year Growth" \\
    --json
  
  # Line chart (trends)
  cat > trend_data.json << 'EOF'
{
  "categories": ["Jan", "Feb", "Mar", "Apr", "May", "Jun"],
  "series": [
    {"name": "Website Traffic", "values": [12000, 13500, 15200, 16800, 18500, 21000]},
    {"name": "Conversions", "values": [240, 270, 304, 336, 370, 420]}
  ]
}
EOF
  
  uv python ppt_add_chart.py \\
    --file presentation.pptx \\
    --slide 4 \\
    --chart-type line_markers \\
    --data trend_data.json \\
    --position '{"left":"10%","top":"20%"}' \\
    --size '{"width":"80%","height":"65%"}' \\
    --title "Traffic & Conversion Trends" \\
    --json
  
  # Inline data (short example)
  uv python ppt_add_chart.py \\
    --file presentation.pptx \\
    --slide 5 \\
    --chart-type column \\
    --data-string '{"categories":["A","B","C"],"series":[{"name":"Sales","values":[10,20,15]}]}' \\
    --position '{"left":"20%","top":"25%"}' \\
    --size '{"width":"60%","height":"50%"}' \\
    --json

Best Practices:
  - Use column charts for comparing categories (most common)
  - Use line charts for showing trends over time
  - Use pie charts for proportions (max 5-7 slices)
  - Use bar charts when category names are long
  - Keep data series count to 3-5 max for clarity
  - Use consistent colors across presentation
  - Always include a descriptive title
  - Round numbers for readability

Chart Selection Guide:
  - Compare values: Column or Bar chart
  - Show trends: Line chart
  - Show proportions: Pie chart
  - Show composition: Stacked column/bar
  - Show correlation: Scatter plot
  - Emphasize change: Area chart
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file path'
    )
    
    parser.add_argument(
        '--slide',
        required=True,
        type=int,
        help='Slide index (0-based)'
    )
    
    parser.add_argument(
        '--chart-type',
        required=True,
        choices=['column', 'column_stacked', 'bar', 'bar_stacked', 
                'line', 'line_markers', 'pie', 'area', 'scatter'],
        help='Chart type'
    )
    
    parser.add_argument(
        '--data',
        type=Path,
        help='JSON file with chart data'
    )
    
    parser.add_argument(
        '--data-string',
        help='Inline JSON data string'
    )
    
    parser.add_argument(
        '--position',
        required=True,
        type=json.loads,
        help='Position dict (JSON string)'
    )
    
    parser.add_argument(
        '--size',
        type=json.loads,
        help='Size dict (JSON string)'
    )
    
    parser.add_argument(
        '--title',
        help='Chart title'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        help='Output JSON response'
    )
    
    args = parser.parse_args()
    
    try:
        # Load chart data
        if args.data:
            if not args.data.exists():
                raise FileNotFoundError(f"Data file not found: {args.data}")
            with open(args.data, 'r') as f:
                data = json.load(f)
        elif args.data_string:
            data = json.loads(args.data_string)
        else:
            raise ValueError("Either --data or --data-string required")
            
        # Handle optional size and merge from position
        size = args.size if args.size else {}
        position = args.position
        
        if "width" in position and "width" not in size:
            size["width"] = position["width"]
        if "height" in position and "height" not in size:
            size["height"] = position["height"]
            
        # Apply defaults if still missing
        if "width" not in size:
            size["width"] = "50%"
        if "height" not in size:
            size["height"] = "50%"
        
        result = add_chart(
            filepath=args.file,
            slide_index=args.slide,
            chart_type=args.chart_type,
            data=data,
            position=position,
            size=size,
            chart_title=args.title
        )
        
        if args.json:
            print(json.dumps(result, indent=2))
        else:
            print(f"✅ Added {result['chart_type']} chart to slide {result['slide_index']}")
            if args.title:
                print(f"   Title: {result['chart_title']}")
            print(f"   Categories: {result['categories']}")
            print(f"   Series: {result['series']}")
            print(f"   Total data points: {result['data_points']}")
        
        sys.exit(0)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__
        }
        
        if args.json:
            print(json.dumps(error_result, indent=2))
        else:
            print(f"❌ Error: {e}", file=sys.stderr)
        
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_add_connector.py
```py
#!/usr/bin/env python3
"""
PowerPoint Add Connector Tool
Draw a line between two shapes

Usage:
    uv python ppt_add_connector.py --file deck.pptx --slide 0 --from-shape 0 --to-shape 1 --json
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError, SlideNotFoundError
)


def add_connector(
    filepath: Path,
    slide_index: int,
    from_shape: int,
    to_shape: int,
    connector_type: str = "straight"
) -> Dict[str, Any]:
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        total = agent.get_slide_count()
        if not 0 <= slide_index < total:
            raise SlideNotFoundError(f"Slide index {slide_index} out of range")
            
        agent.add_connector(
            slide_index=slide_index,
            from_shape=from_shape,
            to_shape=to_shape,
            connector_type=connector_type
        )
        agent.save()
        
    return {
        "status": "success",
        "file": str(filepath),
        "slide_index": slide_index,
        "connection": {
            "from": from_shape,
            "to": to_shape,
            "type": connector_type
        }
    }


def main():
    parser = argparse.ArgumentParser(description="Add connector between shapes")
    parser.add_argument('--file', required=True, type=Path, help='PowerPoint file path')
    parser.add_argument('--slide', required=True, type=int, help='Slide index (0-based)')
    parser.add_argument('--from-shape', required=True, type=int, help='Start shape index')
    parser.add_argument('--to-shape', required=True, type=int, help='End shape index')
    parser.add_argument('--type', default='straight', help='Connector type (straight)')
    parser.add_argument('--json', action='store_true', default=True, help='Output JSON')
    
    args = parser.parse_args()
    
    try:
        result = add_connector(
            filepath=args.file,
            slide_index=args.slide,
            from_shape=args.from_shape,
            to_shape=args.to_shape,
            connector_type=args.type
        )
        print(json.dumps(result, indent=2))
        sys.exit(0)
    except Exception as e:
        print(json.dumps({"status": "error", "error": str(e), "error_type": type(e).__name__}, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_add_shape.py
```py
#!/usr/bin/env python3
"""
PowerPoint Add Shape Tool
Add shape (rectangle, circle, arrow, etc.) to slide

Usage:
    uv python ppt_add_shape.py --file presentation.pptx --slide 0 --shape rectangle --position '{"left":"20%","top":"30%"}' --size '{"width":"60%","height":"40%"}' --fill-color "#0070C0" --json

Exit Codes:
    0: Success
    1: Error occurred
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError, SlideNotFoundError
)


def add_shape(
    filepath: Path,
    slide_index: int,
    shape_type: str,
    position: Dict[str, Any],
    size: Dict[str, Any],
    fill_color: str = None,
    line_color: str = None,
    line_width: float = 1.0
) -> Dict[str, Any]:
    """Add shape to slide."""
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Validate slide index
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides-1})"
            )
        
        # Add shape
        agent.add_shape(
            slide_index=slide_index,
            shape_type=shape_type,
            position=position,
            size=size,
            fill_color=fill_color,
            line_color=line_color,
            line_width=line_width
        )
        
        # Save
        agent.save()
    
    return {
        "status": "success",
        "file": str(filepath),
        "slide_index": slide_index,
        "shape_type": shape_type,
        "position": position,
        "size": size,
        "styling": {
            "fill_color": fill_color,
            "line_color": line_color,
            "line_width": line_width
        }
    }


def main():
    parser = argparse.ArgumentParser(
        description="Add shape to PowerPoint slide",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Available Shapes:
  - rectangle: Standard rectangle/box
  - rounded_rectangle: Rectangle with rounded corners
  - ellipse: Circle or oval
  - triangle: Triangle (isosceles)
  - arrow_right: Right-pointing arrow
  - arrow_left: Left-pointing arrow
  - arrow_up: Up-pointing arrow
  - arrow_down: Down-pointing arrow
  - star: 5-point star
  - heart: Heart shape

Common Uses:
  - rectangle: Callout boxes, containers, dividers
  - rounded_rectangle: Buttons, soft containers
  - ellipse: Emphasis, icons, venn diagrams
  - arrows: Process flows, directional indicators
  - star: Highlights, ratings, attention
  - triangle: Warning indicators, play buttons

Examples:
  # Blue callout box
  uv python ppt_add_shape.py \\
    --file presentation.pptx \\
    --slide 1 \\
    --shape rounded_rectangle \\
    --position '{"left":"10%","top":"15%"}' \\
    --size '{"width":"30%","height":"15%"}' \\
    --fill-color "#0070C0" \\
    --line-color "#FFFFFF" \\
    --line-width 2 \\
    --json
  
  # Process flow arrows
  uv python ppt_add_shape.py \\
    --file presentation.pptx \\
    --slide 2 \\
    --shape arrow_right \\
    --position '{"left":"30%","top":"40%"}' \\
    --size '{"width":"15%","height":"8%"}' \\
    --fill-color "#00B050" \\
    --json
  
  # Emphasis circle
  uv python ppt_add_shape.py \\
    --file presentation.pptx \\
    --slide 3 \\
    --shape ellipse \\
    --position '{"anchor":"center"}' \\
    --size '{"width":"20%","height":"20%"}' \\
    --fill-color "#FFC000" \\
    --line-color "#C65911" \\
    --line-width 3 \\
    --json
  
  # Warning triangle
  uv python ppt_add_shape.py \\
    --file presentation.pptx \\
    --slide 4 \\
    --shape triangle \\
    --position '{"left":"5%","top":"5%"}' \\
    --size '{"width":"8%","height":"8%"}' \\
    --fill-color "#FF0000" \\
    --json
  
  # Transparent overlay (no fill)
  uv python ppt_add_shape.py \\
    --file presentation.pptx \\
    --slide 5 \\
    --shape rectangle \\
    --position '{"left":"0%","top":"0%"}' \\
    --size '{"width":"100%","height":"100%"}' \\
    --line-color "#0070C0" \\
    --line-width 5 \\
    --json

Design Tips:
  - Use shapes to organize content visually
  - Consistent colors across shapes
  - Align shapes to grid for professional look
  - Use subtle colors for backgrounds
  - Bold colors for emphasis
  - Combine shapes to create diagrams
  - Layer shapes for depth (background first)

Color Palette (Corporate):
  - Primary Blue: #0070C0
  - Secondary Gray: #595959
  - Accent Orange: #ED7D31
  - Success Green: #70AD47
  - Warning Yellow: #FFC000
  - Danger Red: #C00000
  - White: #FFFFFF
  - Black: #000000

Shape Layering:
  - Shapes are added in order (first = back, last = front)
  - Use transparent shapes for overlays
  - Group related shapes together
  - Send to back: create shape first
  - Bring to front: create shape last
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file path'
    )
    
    parser.add_argument(
        '--slide',
        required=True,
        type=int,
        help='Slide index (0-based)'
    )
    
    parser.add_argument(
        '--shape',
        required=True,
        choices=['rectangle', 'rounded_rectangle', 'ellipse', 'triangle',
                'arrow_right', 'arrow_left', 'arrow_up', 'arrow_down',
                'star', 'heart'],
        help='Shape type'
    )
    
    parser.add_argument(
        '--position',
        required=True,
        type=json.loads,
        help='Position dict (JSON string)'
    )
    
    parser.add_argument(
        '--size',
        type=json.loads,
        help='Size dict (JSON string)'
    )
    
    parser.add_argument(
        '--fill-color',
        help='Fill color (hex, e.g., #0070C0)'
    )
    
    parser.add_argument(
        '--line-color',
        help='Line/border color (hex)'
    )
    
    parser.add_argument(
        '--line-width',
        type=float,
        default=1.0,
        help='Line width in points (default: 1.0)'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        help='Output JSON response'
    )
    
    args = parser.parse_args()
    
    try:
        # Handle optional size and merge from position
        size = args.size if args.size else {}
        position = args.position
        
        if "width" in position and "width" not in size:
            size["width"] = position["width"]
        if "height" in position and "height" not in size:
            size["height"] = position["height"]

        # Apply defaults if still missing
        if "width" not in size:
            size["width"] = "20%"
        if "height" not in size:
            size["height"] = "20%"

        result = add_shape(
            filepath=args.file,
            slide_index=args.slide,
            shape_type=args.shape,
            position=position,
            size=size,
            fill_color=args.fill_color,
            line_color=args.line_color,
            line_width=args.line_width
        )
        
        if args.json:
            print(json.dumps(result, indent=2))
        else:
            print(f"✅ Added {result['shape_type']} to slide {result['slide_index']}")
            if args.fill_color:
                print(f"   Fill: {result['styling']['fill_color']}")
            if args.line_color:
                print(f"   Line: {result['styling']['line_color']} ({result['styling']['line_width']}pt)")
        
        sys.exit(0)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__
        }
        
        if args.json:
            print(json.dumps(error_result, indent=2))
        else:
            print(f"❌ Error: {e}", file=sys.stderr)
        
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_add_slide.py
```py
#!/usr/bin/env python3
"""
PowerPoint Add Slide Tool
Add new slide to existing presentation with specific layout

Usage:
    uv python ppt_add_slide.py --file presentation.pptx --layout "Title and Content" --index 2 --json

Exit Codes:
    0: Success
    1: Error occurred
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError
)


def add_slide(
    filepath: Path,
    layout: str,
    index: int = None,
    set_title: str = None
) -> Dict[str, Any]:
    """Add slide to presentation."""
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Get available layouts
        available_layouts = agent.get_available_layouts()
        
        # Validate layout
        if layout not in available_layouts:
            # Try fuzzy match
            layout_lower = layout.lower()
            for avail in available_layouts:
                if layout_lower in avail.lower():
                    layout = avail
                    break
            else:
                raise ValueError(
                    f"Layout '{layout}' not found. "
                    f"Available: {available_layouts}"
                )
        
        # Add slide
        slide_index = agent.add_slide(layout_name=layout, index=index)
        
        # Set title if provided
        if set_title:
            agent.set_title(slide_index, set_title)
        
        # Get slide info before saving
        slide_info = agent.get_slide_info(slide_index)
        
        # Save
        agent.save()
        
        # Get updated presentation info
        prs_info = agent.get_presentation_info()
    
    return {
        "status": "success",
        "file": str(filepath),
        "slide_index": slide_index,
        "layout": layout,
        "title_set": set_title,
        "total_slides": prs_info["slide_count"],
        "slide_info": {
            "shape_count": slide_info["shape_count"],
            "has_notes": slide_info["has_notes"]
        }
    }


def main():
    parser = argparse.ArgumentParser(
        description="Add new slide to PowerPoint presentation",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Add slide at end
  uv python ppt_add_slide.py --file presentation.pptx --layout "Title and Content" --json
  
  # Add slide at specific position
  uv python ppt_add_slide.py --file deck.pptx --layout "Section Header" --index 2 --json
  
  # Add slide with title
  uv python ppt_add_slide.py --file presentation.pptx --layout "Title Slide" --title "Q4 Results" --json
  
  # Add blank slide
  uv python ppt_add_slide.py --file presentation.pptx --layout "Blank" --json

Common Layouts:
  - Title Slide: For presentation opening
  - Title and Content: Most common layout
  - Section Header: For section breaks
  - Two Content: Side-by-side content
  - Comparison: Compare two items
  - Title Only: Maximum content space
  - Blank: Complete freedom
  - Picture with Caption: Image-focused

Tips:
  - Use --index 0 to insert at beginning
  - Omit --index to add at end
  - Get available layouts: uv python ppt_get_info.py --file your.pptx --json
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file path'
    )
    
    parser.add_argument(
        '--layout',
        required=True,
        help='Layout name for new slide'
    )
    
    parser.add_argument(
        '--index',
        type=int,
        help='Position to insert slide (0-based, default: end)'
    )
    
    parser.add_argument(
        '--title',
        help='Optional title text to set'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        help='Output JSON response'
    )
    
    args = parser.parse_args()
    
    try:
        result = add_slide(
            filepath=args.file,
            layout=args.layout,
            index=args.index,
            set_title=args.title
        )
        
        if args.json:
            print(json.dumps(result, indent=2))
        else:
            print(f"✅ Added slide to: {result['file']}")
            print(f"   Layout: {result['layout']}")
            print(f"   Position: {result['slide_index']}")
            print(f"   Total slides: {result['total_slides']}")
            if args.title:
                print(f"   Title: {result['title_set']}")
        
        sys.exit(0)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__
        }
        
        if args.json:
            print(json.dumps(error_result, indent=2))
        else:
            print(f"❌ Error: {e}", file=sys.stderr)
        
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_add_table.py
```py
#!/usr/bin/env python3
"""
PowerPoint Add Table Tool
Add data table to slide

Usage:
    uv python ppt_add_table.py --file presentation.pptx --slide 1 --rows 5 --cols 3 --data table_data.json --position '{"left":"10%","top":"25%"}' --size '{"width":"80%","height":"50%"}' --json

Exit Codes:
    0: Success
    1: Error occurred
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any, List

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError, SlideNotFoundError
)


def add_table(
    filepath: Path,
    slide_index: int,
    rows: int,
    cols: int,
    position: Dict[str, Any],
    size: Dict[str, Any],
    data: List[List[Any]] = None,
    headers: List[str] = None
) -> Dict[str, Any]:
    """Add table to slide."""
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if rows < 1 or cols < 1:
        raise ValueError("Table must have at least 1 row and 1 column")
    
    if rows > 50 or cols > 20:
        raise ValueError("Maximum table size: 50 rows × 20 columns (readability limit)")
    
    # Prepare data with headers
    table_data = []
    
    if headers:
        if len(headers) != cols:
            raise ValueError(f"Headers count ({len(headers)}) must match columns ({cols})")
        table_data.append(headers)
        data_rows = rows - 1  # One row used for headers
    else:
        data_rows = rows
    
    # Add data rows
    if data:
        if len(data) > data_rows:
            raise ValueError(f"Too many data rows ({len(data)}) for table size ({data_rows} data rows)")
        
        for row in data:
            if len(row) != cols:
                raise ValueError(f"Data row has {len(row)} items, expected {cols}")
            table_data.append(row)
        
        # Pad with empty rows if needed
        while len(table_data) < rows:
            table_data.append([""] * cols)
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Validate slide index
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides-1})"
            )
        
        # Add table
        agent.add_table(
            slide_index=slide_index,
            rows=rows,
            cols=cols,
            position=position,
            size=size,
            data=table_data if table_data else None
        )
        
        # Save
        agent.save()
    
    return {
        "status": "success",
        "file": str(filepath),
        "slide_index": slide_index,
        "rows": rows,
        "cols": cols,
        "has_headers": headers is not None,
        "data_rows_filled": len(data) if data else 0,
        "total_cells": rows * cols
    }


def main():
    parser = argparse.ArgumentParser(
        description="Add data table to PowerPoint slide",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Data Format (JSON):
  - 2D array: [["A1","B1","C1"], ["A2","B2","C2"]]
  - CSV file: converted to 2D array
  - Pandas DataFrame: exported to JSON array

Examples:
  # Simple pricing table
  cat > pricing.json << 'EOF'
[
  ["Starter", "$9/mo", "Basic features"],
  ["Pro", "$29/mo", "Advanced features"],
  ["Enterprise", "$99/mo", "All features + support"]
]
EOF
  
  uv python ppt_add_table.py \\
    --file presentation.pptx \\
    --slide 3 \\
    --rows 4 \\
    --cols 3 \\
    --headers "Plan,Price,Features" \\
    --data pricing.json \\
    --position '{"left":"15%","top":"25%"}' \\
    --size '{"width":"70%","height":"50%"}' \\
    --json
  
  # Quarterly results table
  cat > results.json << 'EOF'
[
  ["Q1", "10.5", "8.2", "2.3"],
  ["Q2", "12.8", "9.1", "3.7"],
  ["Q3", "15.2", "10.5", "4.7"],
  ["Q4", "18.6", "12.1", "6.5"]
]
EOF
  
  uv python ppt_add_table.py \\
    --file presentation.pptx \\
    --slide 4 \\
    --rows 5 \\
    --cols 4 \\
    --headers "Quarter,Revenue,Costs,Profit" \\
    --data results.json \\
    --position '{"left":"10%","top":"20%"}' \\
    --size '{"width":"80%","height":"55%"}' \\
    --json
  
  # Comparison table (centered)
  cat > comparison.json << 'EOF'
[
  ["Speed", "Fast", "Very Fast"],
  ["Security", "Standard", "Enterprise"],
  ["Support", "Email", "24/7 Phone"]
]
EOF
  
  uv python ppt_add_table.py \\
    --file presentation.pptx \\
    --slide 5 \\
    --rows 4 \\
    --cols 3 \\
    --headers "Feature,Basic,Premium" \\
    --data comparison.json \\
    --position '{"anchor":"center"}' \\
    --size '{"width":"60%","height":"40%"}' \\
    --json
  
  # Empty table (for manual filling)
  uv python ppt_add_table.py \\
    --file presentation.pptx \\
    --slide 6 \\
    --rows 6 \\
    --cols 4 \\
    --headers "Name,Role,Department,Email" \\
    --position '{"left":"10%","top":"25%"}' \\
    --size '{"width":"80%","height":"60%"}' \\
    --json
  
  # From CSV file (convert first)
  # cat data.csv | python -c "import csv, json, sys; print(json.dumps(list(csv.reader(sys.stdin))))" > data.json
  # Then use --data data.json

Best Practices:
  - Keep tables under 10 rows for readability
  - Use headers for all tables
  - Align numbers right, text left
  - Use consistent decimal places
  - Highlight key values with color
  - Leave white space around table
  - Use alternating row colors for large tables

Table Size Guidelines:
  - 3-5 columns: Optimal for most presentations
  - 6-10 rows: Maximum for comfortable reading
  - Font size: 12-16pt for body, 14-18pt for headers
  - Cell padding: Leave breathing room

When to Use Tables vs Charts:
  - Use tables: Exact values matter, detailed data
  - Use charts: Show trends, comparisons, patterns
  - Use both: Table with summary chart
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file path'
    )
    
    parser.add_argument(
        '--slide',
        required=True,
        type=int,
        help='Slide index (0-based)'
    )
    
    parser.add_argument(
        '--rows',
        required=True,
        type=int,
        help='Number of rows (including header if present)'
    )
    
    parser.add_argument(
        '--cols',
        required=True,
        type=int,
        help='Number of columns'
    )
    
    parser.add_argument(
        '--position',
        required=True,
        type=json.loads,
        help='Position dict (JSON string)'
    )
    
    parser.add_argument(
        '--size',
        required=True,
        type=json.loads,
        help='Size dict (JSON string)'
    )
    
    parser.add_argument(
        '--data',
        type=Path,
        help='JSON file with 2D array of cell values'
    )
    
    parser.add_argument(
        '--data-string',
        help='Inline JSON 2D array string'
    )
    
    parser.add_argument(
        '--headers',
        help='Comma-separated header row (will be row 0)'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        help='Output JSON response'
    )
    
    args = parser.parse_args()
    
    try:
        # Parse headers
        headers = None
        if args.headers:
            headers = [h.strip() for h in args.headers.split(',')]
        
        # Load table data
        data = None
        if args.data:
            if not args.data.exists():
                raise FileNotFoundError(f"Data file not found: {args.data}")
            with open(args.data, 'r') as f:
                data = json.load(f)
        elif args.data_string:
            data = json.loads(args.data_string)
        
        # Validate data is 2D array
        if data is not None:
            if not isinstance(data, list):
                raise ValueError("Data must be a 2D array (list of lists)")
            if data and not isinstance(data[0], list):
                raise ValueError("Data must be a 2D array (list of lists)")
        
        result = add_table(
            filepath=args.file,
            slide_index=args.slide,
            rows=args.rows,
            cols=args.cols,
            position=args.position,
            size=args.size,
            data=data,
            headers=headers
        )
        
        if args.json:
            print(json.dumps(result, indent=2))
        else:
            print(f"✅ Added {result['rows']}×{result['cols']} table to slide {result['slide_index']}")
            if result['has_headers']:
                print(f"   Headers: Yes")
            print(f"   Data rows filled: {result['data_rows_filled']}")
            print(f"   Total cells: {result['total_cells']}")
        
        sys.exit(0)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__
        }
        
        if args.json:
            print(json.dumps(error_result, indent=2))
        else:
            print(f"❌ Error: {e}", file=sys.stderr)
        
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_add_text_box.py
```py
#!/usr/bin/env python3
"""
PowerPoint Add Text Box Tool
Add text box to slide with flexible positioning

Usage:
    uv python ppt_add_text_box.py --file presentation.pptx --slide 0 --text "Revenue: $1.5M" --position '{"left":"20%","top":"30%"}' --size '{"width":"60%","height":"10%"}' --json

Exit Codes:
    0: Success
    1: Error occurred
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError, SlideNotFoundError,
    InvalidPositionError
)


def add_text_box(
    filepath: Path,
    slide_index: int,
    text: str,
    position: Dict[str, Any],
    size: Dict[str, Any],
    font_name: str = "Calibri",
    font_size: int = 18,
    bold: bool = False,
    italic: bool = False,
    color: str = None,
    alignment: str = "left"
) -> Dict[str, Any]:
    """Add text box to slide."""
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Validate slide index
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides-1})"
            )
        
        # Add text box
        agent.add_text_box(
            slide_index=slide_index,
            text=text,
            position=position,
            size=size,
            font_name=font_name,
            font_size=font_size,
            bold=bold,
            italic=italic,
            color=color,
            alignment=alignment
        )
        
        # Get updated slide info
        slide_info = agent.get_slide_info(slide_index)
        
        # Save
        agent.save()
    
    return {
        "status": "success",
        "file": str(filepath),
        "slide_index": slide_index,
        "text": text[:50] + "..." if len(text) > 50 else text,
        "position": position,
        "size": size,
        "formatting": {
            "font_name": font_name,
            "font_size": font_size,
            "bold": bold,
            "italic": italic,
            "color": color,
            "alignment": alignment
        },
        "slide_shape_count": slide_info["shape_count"]
    }


def main():
    parser = argparse.ArgumentParser(
        description="Add text box to PowerPoint slide",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Position Formats:
  1. Percentage: {"left": "20%", "top": "30%"}
  2. Absolute inches: {"left": 2.0, "top": 3.0}
  3. Anchor-based: {"anchor": "center", "offset_x": 0, "offset_y": -1.0}
  4. Grid system: {"grid_row": 2, "grid_col": 3, "grid_size": 12}
  5. Excel-like: {"grid": "C4"}

Size Formats:
  - {"width": "60%", "height": "10%"}
  - {"width": 5.0, "height": 2.0}  (inches)

Anchor Points:
  top_left, top_center, top_right,
  center_left, center, center_right,
  bottom_left, bottom_center, bottom_right

Examples:
  # Percentage positioning (easiest for AI)
  uv python ppt_add_text_box.py \\
    --file presentation.pptx \\
    --slide 0 \\
    --text "Revenue: $1.5M" \\
    --position '{"left":"20%","top":"30%"}' \\
    --size '{"width":"60%","height":"10%"}' \\
    --font-size 24 \\
    --bold \\
    --json
  
  # Grid positioning (Excel-like)
  uv python ppt_add_text_box.py \\
    --file presentation.pptx \\
    --slide 1 \\
    --text "Q4 Summary" \\
    --position '{"grid":"C4"}' \\
    --size '{"width":"25%","height":"8%"}' \\
    --json
  
  # Anchor-based (centered text)
  uv python ppt_add_text_box.py \\
    --file presentation.pptx \\
    --slide 2 \\
    --text "Thank You!" \\
    --position '{"anchor":"center","offset_x":0,"offset_y":0}' \\
    --size '{"width":"80%","height":"15%"}' \\
    --font-size 48 \\
    --bold \\
    --alignment center \\
    --json
  
  # Bottom right copyright
  uv python ppt_add_text_box.py \\
    --file presentation.pptx \\
    --slide 0 \\
    --text "© 2024 Company Inc." \\
    --position '{"anchor":"bottom_right","offset_x":-0.5,"offset_y":-0.3}' \\
    --size '{"width":"2.5","height":"0.3"}' \\
    --font-size 10 \\
    --color "#808080" \\
    --json
  
  # Colored headline
  uv python ppt_add_text_box.py \\
    --file presentation.pptx \\
    --slide 1 \\
    --text "Key Takeaways" \\
    --position '{"left":"5%","top":"15%"}' \\
    --size '{"width":"90%","height":"8%"}' \\
    --font-name "Arial" \\
    --font-size 36 \\
    --bold \\
    --color "#0070C0" \\
    --json

Tips:
  - Use percentages for responsive layouts
  - Grid system ("C4") is intuitive for structured content
  - Anchor points great for headers/footers
  - Keep font size 18pt+ for readability
  - Use hex colors: #FF0000 (red), #0070C0 (blue), #00B050 (green)
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file path'
    )
    
    parser.add_argument(
        '--slide',
        required=True,
        type=int,
        help='Slide index (0-based)'
    )
    
    parser.add_argument(
        '--text',
        required=True,
        help='Text content'
    )
    
    parser.add_argument(
        '--position',
        required=True,
        type=json.loads,
        help='Position dict (JSON string)'
    )
    
    parser.add_argument(
        '--size',
        type=json.loads,
        help='Size dict (JSON string)'
    )
    
    parser.add_argument(
        '--font-name',
        default='Calibri',
        help='Font name (default: Calibri)'
    )
    
    parser.add_argument(
        '--font-size',
        type=int,
        default=18,
        help='Font size in points (default: 18)'
    )
    
    parser.add_argument(
        '--bold',
        nargs='?',
        const='true',
        default='false',
        help='Bold text (optional: true/false)'
    )
    
    parser.add_argument(
        '--italic',
        nargs='?',
        const='true',
        default='false',
        help='Italic text (optional: true/false)'
    )
    
    parser.add_argument(
        '--color',
        help='Text color (hex, e.g., #FF0000)'
    )
    
    parser.add_argument(
        '--alignment',
        choices=['left', 'center', 'right', 'justify'],
        default='left',
        help='Text alignment (default: left)'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        help='Output JSON response'
    )
    
    args = parser.parse_args()
    
    try:
        # Handle optional size and merge from position
        size = args.size if args.size else {}
        position = args.position
        
        if "width" in position and "width" not in size:
            size["width"] = position["width"]
        if "height" in position and "height" not in size:
            size["height"] = position["height"]
            
        # Apply defaults if still missing
        if "width" not in size:
            size["width"] = "40%"
        if "height" not in size:
            size["height"] = "20%"
            
        # Helper to parse boolean string/flag
        def parse_bool(val):
            if isinstance(val, bool): return val
            if val is None: return False
            return str(val).lower() in ('true', 'yes', '1', 'on')

        result = add_text_box(
            filepath=args.file,
            slide_index=args.slide,
            text=args.text,
            position=position,
            size=size,
            font_name=args.font_name,
            font_size=args.font_size,
            bold=parse_bool(args.bold),
            italic=parse_bool(args.italic),
            color=args.color,
            alignment=args.alignment
        )
        
        if args.json:
            print(json.dumps(result, indent=2))
        else:
            print(f"✅ Added text box to slide {result['slide_index']}")
            print(f"   Text: {result['text']}")
            print(f"   Font: {result['formatting']['font_name']} {result['formatting']['font_size']}pt")
            if args.bold or args.italic:
                style = []
                if args.bold:
                    style.append("bold")
                if args.italic:
                    style.append("italic")
                print(f"   Style: {', '.join(style)}")
        
        sys.exit(0)
        
    except json.JSONDecodeError as e:
        error_result = {
            "status": "error",
            "error": f"Invalid JSON in position or size argument: {e}",
            "error_type": "JSONDecodeError",
            "hint": "Use single quotes around JSON and double quotes inside: '{\"left\":\"20%\",\"top\":\"30%\"}'"
        }
        
        if args.json:
            print(json.dumps(error_result, indent=2))
        else:
            print(f"❌ Error: {error_result['error']}", file=sys.stderr)
            print(f"   Hint: {error_result['hint']}", file=sys.stderr)
        
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__
        }
        
        if args.json:
            print(json.dumps(error_result, indent=2))
        else:
            print(f"❌ Error: {e}", file=sys.stderr)
        
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_check_accessibility.py
```py
#!/usr/bin/env python3
"""
PowerPoint Check Accessibility Tool
Run WCAG 2.1 accessibility checks on presentation

Usage:
    uv python ppt_check_accessibility.py --file presentation.pptx --json

Exit Codes:
    0: Success (even if issues found - check 'status')
    1: Error occurred (file not found, etc)
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError
)


def check_accessibility(filepath: Path) -> Dict[str, Any]:
    """Run accessibility checks."""
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath, acquire_lock=False)  # Read-only check
        result = agent.check_accessibility()
        result["file"] = str(filepath)
    
    return result


def main():
    parser = argparse.ArgumentParser(
        description="Check PowerPoint accessibility (WCAG 2.1)",
        formatter_class=argparse.RawDescriptionHelpFormatter
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file path'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default)'
    )
    
    args = parser.parse_args()
    
    try:
        result = check_accessibility(filepath=args.file)
        print(json.dumps(result, indent=2))
        sys.exit(0)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_clone_presentation.py
```py
#!/usr/bin/env python3
"""
PowerPoint Clone Presentation Tool
Create an exact copy of a presentation

Usage:
    uv python ppt_clone_presentation.py --source base.pptx --output new_deck.pptx --json
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import PowerPointAgent


def clone_presentation(source: Path, output: Path) -> Dict[str, Any]:
    
    if not source.exists():
        raise FileNotFoundError(f"Source file not found: {source}")
    
    # Validate output path
    if not output.suffix.lower() == '.pptx':
        output = output.with_suffix('.pptx')

    # Uses agent to open and save-as, effectively cloning
    with PowerPointAgent(source) as agent:
        agent.open(source, acquire_lock=False) # Read-only open
        agent.save(output)
        
    return {
        "status": "success",
        "source": str(source),
        "output": str(output),
        "size_bytes": output.stat().st_size
    }


def main():
    parser = argparse.ArgumentParser(description="Clone PowerPoint presentation")
    parser.add_argument('--source', required=True, type=Path, help='Source file')
    parser.add_argument('--output', required=True, type=Path, help='Destination file')
    parser.add_argument('--json', action='store_true', default=True, help='Output JSON')
    
    args = parser.parse_args()
    
    try:
        result = clone_presentation(source=args.source, output=args.output)
        print(json.dumps(result, indent=2))
        sys.exit(0)
    except Exception as e:
        print(json.dumps({"status": "error", "error": str(e), "error_type": type(e).__name__}, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_create_from_structure.py
```py
#!/usr/bin/env python3
"""
PowerPoint Create From Structure Tool
Create presentation from JSON structure definition

Usage:
    uv python ppt_create_from_structure.py --structure deck_structure.json --output presentation.pptx --json

Exit Codes:
    0: Success
    1: Error occurred
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any, List

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError
)


def validate_structure(structure: Dict[str, Any]) -> None:
    """Validate JSON structure schema."""
    if "slides" not in structure:
        raise ValueError("Structure must contain 'slides' array")
    
    if not isinstance(structure["slides"], list):
        raise ValueError("'slides' must be an array")
    
    if len(structure["slides"]) == 0:
        raise ValueError("Must have at least one slide")
    
    if len(structure["slides"]) > 100:
        raise ValueError("Maximum 100 slides (performance limit)")


def create_from_structure(
    structure: Dict[str, Any],
    output: Path
) -> Dict[str, Any]:
    """Create presentation from structure definition."""
    
    # Validate structure
    validate_structure(structure)
    
    # Track created items
    stats = {
        "slides_created": 0,
        "text_boxes_added": 0,
        "images_inserted": 0,
        "charts_added": 0,
        "tables_added": 0,
        "shapes_added": 0,
        "errors": []
    }
    
    with PowerPointAgent() as agent:
        # Create base presentation
        template = structure.get("template")
        if template and Path(template).exists():
            agent.create_new(template=Path(template))
        else:
            agent.create_new()
        
        # Process each slide
        for slide_idx, slide_def in enumerate(structure["slides"]):
            try:
                # Add slide
                layout = slide_def.get("layout", "Title and Content")
                agent.add_slide(layout_name=layout)
                stats["slides_created"] += 1
                
                # Set title if provided
                if "title" in slide_def:
                    agent.set_title(
                        slide_index=slide_idx,
                        title=slide_def["title"],
                        subtitle=slide_def.get("subtitle")
                    )
                
                # Process content items
                for item in slide_def.get("content", []):
                    try:
                        item_type = item.get("type")
                        
                        if item_type == "text_box":
                            agent.add_text_box(
                                slide_index=slide_idx,
                                text=item["text"],
                                position=item["position"],
                                size=item["size"],
                                font_name=item.get("font_name", "Calibri"),
                                font_size=item.get("font_size", 18),
                                bold=item.get("bold", False),
                                italic=item.get("italic", False),
                                color=item.get("color"),
                                alignment=item.get("alignment", "left")
                            )
                            stats["text_boxes_added"] += 1
                        
                        elif item_type == "image":
                            image_path = Path(item["path"])
                            if image_path.exists():
                                agent.insert_image(
                                    slide_index=slide_idx,
                                    image_path=image_path,
                                    position=item["position"],
                                    size=item.get("size"),
                                    compress=item.get("compress", False)
                                )
                                stats["images_inserted"] += 1
                            else:
                                stats["errors"].append(f"Image not found: {item['path']}")
                        
                        elif item_type == "chart":
                            agent.add_chart(
                                slide_index=slide_idx,
                                chart_type=item["chart_type"],
                                data=item["data"],
                                position=item["position"],
                                size=item["size"],
                                chart_title=item.get("title")
                            )
                            stats["charts_added"] += 1
                        
                        elif item_type == "table":
                            agent.add_table(
                                slide_index=slide_idx,
                                rows=item["rows"],
                                cols=item["cols"],
                                position=item["position"],
                                size=item["size"],
                                data=item.get("data")
                            )
                            stats["tables_added"] += 1
                        
                        elif item_type == "shape":
                            agent.add_shape(
                                slide_index=slide_idx,
                                shape_type=item["shape_type"],
                                position=item["position"],
                                size=item["size"],
                                fill_color=item.get("fill_color"),
                                line_color=item.get("line_color"),
                                line_width=item.get("line_width", 1.0)
                            )
                            stats["shapes_added"] += 1
                        
                        elif item_type == "bullet_list":
                            agent.add_bullet_list(
                                slide_index=slide_idx,
                                items=item["items"],
                                position=item["position"],
                                size=item["size"],
                                bullet_style=item.get("bullet_style", "bullet"),
                                font_size=item.get("font_size", 18)
                            )
                            stats["text_boxes_added"] += 1
                        
                        else:
                            stats["errors"].append(f"Unknown content type: {item_type}")
                    
                    except Exception as e:
                        stats["errors"].append(f"Error adding {item_type}: {str(e)}")
            
            except Exception as e:
                stats["errors"].append(f"Error processing slide {slide_idx}: {str(e)}")
        
        # Save
        agent.save(output)
        
        # Get final info
        info = agent.get_presentation_info()
    
    return {
        "status": "success" if len(stats["errors"]) == 0 else "success_with_errors",
        "file": str(output),
        "slides_created": stats["slides_created"],
        "content_added": {
            "text_boxes": stats["text_boxes_added"],
            "images": stats["images_inserted"],
            "charts": stats["charts_added"],
            "tables": stats["tables_added"],
            "shapes": stats["shapes_added"]
        },
        "errors": stats["errors"],
        "error_count": len(stats["errors"]),
        "file_size_bytes": output.stat().st_size if output.exists() else 0
    }


def main():
    parser = argparse.ArgumentParser(
        description="Create PowerPoint from JSON structure definition",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
JSON Structure Schema:
{
  "template": "optional_template.pptx",
  "slides": [
    {
      "layout": "Title Slide",
      "title": "Presentation Title",
      "subtitle": "Subtitle",
      "content": [
        {
          "type": "text_box",
          "text": "Content here",
          "position": {"left": "10%", "top": "20%"},
          "size": {"width": "80%", "height": "10%"},
          "font_size": 18,
          "color": "#000000"
        },
        {
          "type": "image",
          "path": "image.png",
          "position": {"left": "20%", "top": "30%"},
          "size": {"width": "60%", "height": "auto"}
        },
        {
          "type": "chart",
          "chart_type": "column",
          "data": {
            "categories": ["Q1", "Q2", "Q3"],
            "series": [{"name": "Revenue", "values": [100, 120, 140]}]
          },
          "position": {"left": "10%", "top": "20%"},
          "size": {"width": "80%", "height": "60%"}
        },
        {
          "type": "table",
          "rows": 3,
          "cols": 3,
          "position": {"left": "10%", "top": "20%"},
          "size": {"width": "80%", "height": "50%"},
          "data": [["A", "B", "C"], ["1", "2", "3"]]
        },
        {
          "type": "shape",
          "shape_type": "rectangle",
          "position": {"left": "10%", "top": "10%"},
          "size": {"width": "30%", "height": "15%"},
          "fill_color": "#0070C0"
        },
        {
          "type": "bullet_list",
          "items": ["Item 1", "Item 2", "Item 3"],
          "position": {"left": "10%", "top": "25%"},
          "size": {"width": "80%", "height": "60%"}
        }
      ]
    }
  ]
}

Examples:
  # Create simple presentation
  cat > structure.json << 'EOF'
{
  "slides": [
    {
      "layout": "Title Slide",
      "title": "My Presentation",
      "subtitle": "Created from Structure"
    },
    {
      "layout": "Title and Content",
      "title": "Agenda",
      "content": [
        {
          "type": "bullet_list",
          "items": ["Introduction", "Main Content", "Conclusion"],
          "position": {"left": "10%", "top": "25%"},
          "size": {"width": "80%", "height": "60%"}
        }
      ]
    }
  ]
}
EOF
  
  uv python ppt_create_from_structure.py \\
    --structure structure.json \\
    --output presentation.pptx \\
    --json

  # Create complex presentation with charts
  cat > complex.json << 'EOF'
{
  "slides": [
    {
      "layout": "Title Slide",
      "title": "Q4 Report"
    },
    {
      "layout": "Title and Content",
      "title": "Revenue Growth",
      "content": [
        {
          "type": "chart",
          "chart_type": "column",
          "data": {
            "categories": ["Q1", "Q2", "Q3", "Q4"],
            "series": [
              {"name": "2023", "values": [100, 110, 120, 130]},
              {"name": "2024", "values": [120, 135, 145, 160]}
            ]
          },
          "position": {"left": "10%", "top": "20%"},
          "size": {"width": "80%", "height": "65%"},
          "title": "Year over Year Comparison"
        }
      ]
    }
  ]
}
EOF
  
  uv python ppt_create_from_structure.py \\
    --structure complex.json \\
    --output q4_report.pptx \\
    --json

Use Cases:
  - Automated report generation
  - Template-based presentations from data
  - Batch presentation creation
  - AI-generated presentations
  - Programmatic deck building

Best Practices:
  - Validate JSON structure before processing
  - Use templates for consistent branding
  - Keep image paths relative or absolute
  - Handle missing images gracefully
  - Test structure files before production use
        """
    )
    
    parser.add_argument(
        '--structure',
        required=True,
        type=Path,
        help='JSON structure file'
    )
    
    parser.add_argument(
        '--output',
        required=True,
        type=Path,
        help='Output presentation path'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        help='Output JSON response'
    )
    
    args = parser.parse_args()
    
    try:
        # Load structure
        if not args.structure.exists():
            raise FileNotFoundError(f"Structure file not found: {args.structure}")
        
        with open(args.structure, 'r') as f:
            structure = json.load(f)
        
        # Validate output path
        if not args.output.suffix.lower() == '.pptx':
            args.output = args.output.with_suffix('.pptx')
        
        result = create_from_structure(
            structure=structure,
            output=args.output
        )
        
        if args.json:
            print(json.dumps(result, indent=2))
        else:
            print(f"✅ Created presentation: {result['file']}")
            print(f"   Slides: {result['slides_created']}")
            print(f"   Content: {sum(result['content_added'].values())} items")
            if result['error_count'] > 0:
                print(f"   ⚠️  Errors: {result['error_count']}")
                for error in result['errors'][:5]:
                    print(f"      - {error}")
        
        sys.exit(0)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__
        }
        
        if args.json:
            print(json.dumps(error_result, indent=2))
        else:
            print(f"❌ Error: {e}", file=sys.stderr)
        
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_create_from_template.py
```py
#!/usr/bin/env python3
"""
PowerPoint Create From Template Tool
Create new presentation from existing .pptx template

Usage:
    uv python ppt_create_from_template.py --template corporate_template.pptx --output new_presentation.pptx --slides 10 --json

Exit Codes:
    0: Success
    1: Error occurred
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError
)


def create_from_template(
    template: Path,
    output: Path,
    slides: int = 1,
    layout: str = "Title and Content"
) -> Dict[str, Any]:
    """Create presentation from template."""
    
    if not template.exists():
        raise FileNotFoundError(f"Template file not found: {template}")
    
    if not template.suffix.lower() == '.pptx':
        raise ValueError(f"Template must be .pptx file, got: {template.suffix}")
    
    if slides < 1:
        raise ValueError("Must create at least 1 slide")
    
    if slides > 100:
        raise ValueError("Maximum 100 slides per creation (performance limit)")
    
    with PowerPointAgent() as agent:
        # Create from template
        agent.create_new(template=template)
        
        # Get available layouts from template
        available_layouts = agent.get_available_layouts()
        
        # Validate layout
        if layout not in available_layouts:
            # Try to find closest match
            layout_lower = layout.lower()
            for avail in available_layouts:
                if layout_lower in avail.lower():
                    layout = avail
                    break
            else:
                # Use first available layout as fallback
                layout = available_layouts[0] if available_layouts else "Title Slide"
        
        # Template comes with at least 1 slide usually, check current count
        current_slides = agent.get_slide_count()
        
        # Add additional slides if needed
        slides_to_add = max(0, slides - current_slides)
        slide_indices = list(range(current_slides))
        
        for i in range(slides_to_add):
            idx = agent.add_slide(layout_name=layout)
            slide_indices.append(idx)
        
        # Save
        agent.save(output)
        
        # Get presentation info
        info = agent.get_presentation_info()
    
    return {
        "status": "success",
        "file": str(output),
        "template_used": str(template),
        "total_slides": info["slide_count"],
        "slides_requested": slides,
        "template_slides": current_slides,
        "slides_added": slides_to_add,
        "layout_used": layout,
        "available_layouts": info["layouts"],
        "file_size_bytes": output.stat().st_size if output.exists() else 0,
        "slide_dimensions": {
            "width_inches": info["slide_width_inches"],
            "height_inches": info["slide_height_inches"],
            "aspect_ratio": info["aspect_ratio"]
        }
    }


def main():
    parser = argparse.ArgumentParser(
        description="Create PowerPoint presentation from template",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Create from corporate template
  uv python ppt_create_from_template.py \\
    --template templates/corporate.pptx \\
    --output q4_report.pptx \\
    --slides 15 \\
    --json
  
  # Create presentation matching template's default layout
  uv python ppt_create_from_template.py \\
    --template templates/minimal.pptx \\
    --output demo.pptx \\
    --slides 5 \\
    --layout "Section Header" \\
    --json
  
  # Quick presentation from template
  uv python ppt_create_from_template.py \\
    --template templates/branded.pptx \\
    --output quick_deck.pptx \\
    --json

Use Cases:
  - Corporate presentations with branding
  - Consistent theme across team presentations
  - Pre-formatted layouts (fonts, colors, logos)
  - Department-specific templates
  - Client-specific branded decks

Template Benefits:
  - Consistent branding across organization
  - Pre-configured master slides
  - Corporate colors and fonts
  - Logo placements
  - Standard layouts
  - Accessibility features

Creating Templates:
  1. Design in PowerPoint with desired theme
  2. Configure master slides
  3. Set up color scheme
  4. Define standard layouts
  5. Save as .pptx template
  6. Use with this tool

Best Practices:
  - Maintain template library for different purposes
  - Version control templates
  - Document template usage guidelines
  - Test templates before distribution
  - Include variety of layouts in template
        """
    )
    
    parser.add_argument(
        '--template',
        required=True,
        type=Path,
        help='Path to template .pptx file'
    )
    
    parser.add_argument(
        '--output',
        required=True,
        type=Path,
        help='Output presentation path'
    )
    
    parser.add_argument(
        '--slides',
        type=int,
        default=1,
        help='Total number of slides desired (default: 1)'
    )
    
    parser.add_argument(
        '--layout',
        default='Title and Content',
        help='Layout for additional slides (default: "Title and Content")'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        help='Output JSON response'
    )
    
    args = parser.parse_args()
    
    try:
        # Validate output path
        if not args.output.suffix.lower() == '.pptx':
            args.output = args.output.with_suffix('.pptx')
        
        result = create_from_template(
            template=args.template,
            output=args.output,
            slides=args.slides,
            layout=args.layout
        )
        
        if args.json:
            print(json.dumps(result, indent=2))
        else:
            print(f"✅ Created presentation from template: {result['file']}")
            print(f"   Template: {result['template_used']}")
            print(f"   Total slides: {result['total_slides']}")
            print(f"   Template had: {result['template_slides']} slides")
            print(f"   Added: {result['slides_added']} slides")
            print(f"   Layout: {result['layout_used']}")
        
        sys.exit(0)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__
        }
        
        if args.json:
            print(json.dumps(error_result, indent=2))
        else:
            print(f"❌ Error: {e}", file=sys.stderr)
        
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_create_new.py
```py
#!/usr/bin/env python3
"""
PowerPoint Create New Tool
Create a new PowerPoint presentation with specified slides

Usage:
    uv python ppt_create_new.py --output presentation.pptx --slides 5 --layout "Title and Content" --json

Exit Codes:
    0: Success
    1: Error occurred
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any

# Add parent directory to path for core import
sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError
)


def create_new_presentation(
    output: Path,
    slides: int,
    template: Path = None,
    layout: str = "Title and Content"
) -> Dict[str, Any]:
    """Create new PowerPoint presentation."""
    
    if slides < 1:
        raise ValueError("Must create at least 1 slide")
    
    if slides > 100:
        raise ValueError("Maximum 100 slides per creation (performance limit)")
    
    with PowerPointAgent() as agent:
        # Create from template or blank
        agent.create_new(template=template)
        
        # Get available layouts
        available_layouts = agent.get_available_layouts()
        
        # Validate layout
        if layout not in available_layouts:
            # Try to find closest match
            layout_lower = layout.lower()
            for avail in available_layouts:
                if layout_lower in avail.lower():
                    layout = avail
                    break
            else:
                # Use first available layout as fallback
                layout = available_layouts[0] if available_layouts else "Title Slide"
        
        # Add requested number of slides
        slide_indices = []
        for i in range(slides):
            # First slide uses "Title Slide" if available, others use specified layout
            if i == 0 and "Title Slide" in available_layouts:
                slide_layout = "Title Slide"
            else:
                slide_layout = layout
            
            idx = agent.add_slide(layout_name=slide_layout)
            slide_indices.append(idx)
        
        # Save
        agent.save(output)
        
        # Get presentation info
        info = agent.get_presentation_info()
    
    return {
        "status": "success",
        "file": str(output),
        "slides_created": slides,
        "slide_indices": slide_indices,
        "file_size_bytes": output.stat().st_size if output.exists() else 0,
        "slide_dimensions": {
            "width_inches": info["slide_width_inches"],
            "height_inches": info["slide_height_inches"],
            "aspect_ratio": info["aspect_ratio"]
        },
        "available_layouts": info["layouts"],
        "layout_used": layout,
        "template_used": str(template) if template else None
    }


def main():
    parser = argparse.ArgumentParser(
        description="Create new PowerPoint presentation with specified slides",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Create presentation with 5 blank slides
  uv python ppt_create_new.py --output presentation.pptx --slides 5 --json
  
  # Create with specific layout
  uv python ppt_create_new.py --output pitch_deck.pptx --slides 10 --layout "Title and Content" --json
  
  # Create from template
  uv python ppt_create_new.py --output new_deck.pptx --slides 3 --template corporate_template.pptx --json
  
  # Create single title slide
  uv python ppt_create_new.py --output title.pptx --slides 1 --layout "Title Slide" --json

Available Layouts (typical):
  - Title Slide
  - Title and Content
  - Section Header
  - Two Content
  - Comparison
  - Title Only
  - Blank
  - Content with Caption
  - Picture with Caption

Output Format:
  {
    "status": "success",
    "file": "presentation.pptx",
    "slides_created": 5,
    "file_size_bytes": 28432,
    "slide_dimensions": {
      "width_inches": 10.0,
      "height_inches": 7.5,
      "aspect_ratio": "16:9"
    },
    "available_layouts": ["Title Slide", "Title and Content", ...],
    "layout_used": "Title and Content"
  }
        """
    )
    
    parser.add_argument(
        '--output',
        required=True,
        type=Path,
        help='Output PowerPoint file path (.pptx)'
    )
    
    parser.add_argument(
        '--slides',
        type=int,
        default=1,
        help='Number of slides to create (default: 1)'
    )
    
    parser.add_argument(
        '--template',
        type=Path,
        help='Optional template file to use (.pptx)'
    )
    
    parser.add_argument(
        '--layout',
        default='Title and Content',
        help='Layout to use for slides (default: "Title and Content")'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        help='Output JSON response'
    )
    
    args = parser.parse_args()
    
    try:
        # Validate template if specified
        if args.template:
            if not args.template.exists():
                raise FileNotFoundError(f"Template file not found: {args.template}")
            if not args.template.suffix.lower() == '.pptx':
                raise ValueError(f"Template must be .pptx file, got: {args.template.suffix}")
        
        # Validate output path
        if not args.output.suffix.lower() == '.pptx':
            args.output = args.output.with_suffix('.pptx')
        
        # Create presentation
        result = create_new_presentation(
            output=args.output,
            slides=args.slides,
            template=args.template,
            layout=args.layout
        )
        
        if args.json:
            print(json.dumps(result, indent=2))
        else:
            print(f"✅ Created presentation: {result['file']}")
            print(f"   Slides: {result['slides_created']}")
            print(f"   Layout: {result['layout_used']}")
            print(f"   Dimensions: {result['slide_dimensions']['aspect_ratio']}")
            if args.template:
                print(f"   Template: {result['template_used']}")
        
        sys.exit(0)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__
        }
        
        if args.json:
            print(json.dumps(error_result, indent=2))
        else:
            print(f"❌ Error: {e}", file=sys.stderr)
        
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_crop_image.py
```py
#!/usr/bin/env python3
"""
PowerPoint Crop Image Tool
Crop an existing image on a slide

Usage:
    uv python ppt_crop_image.py --file deck.pptx --slide 0 --shape 1 --left 0.1 --right 0.1 --json
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError, SlideNotFoundError
)
from pptx.enum.shapes import MSO_SHAPE_TYPE

def crop_image(
    filepath: Path,
    slide_index: int,
    shape_index: int,
    left: float = 0.0,
    right: float = 0.0,
    top: float = 0.0,
    bottom: float = 0.0
) -> Dict[str, Any]:
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
        
    if not (0.0 <= left <= 1.0 and 0.0 <= right <= 1.0 and 0.0 <= top <= 1.0 and 0.0 <= bottom <= 1.0):
        raise ValueError("Crop values must be between 0.0 and 1.0")

    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        total = agent.get_slide_count()
        if not 0 <= slide_index < total:
            raise SlideNotFoundError(f"Slide index {slide_index} out of range")
            
        slide = agent.prs.slides[slide_index]
        
        if not 0 <= shape_index < len(slide.shapes):
             raise ValueError(f"Shape index {shape_index} out of range")
             
        shape = slide.shapes[shape_index]
        
        if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
            raise ValueError(f"Shape {shape_index} is not a picture")
            
        # Apply crop
        # python-pptx handles crop as a percentage of original size trimmed from edges
        if left > 0: shape.crop_left = left
        if right > 0: shape.crop_right = right
        if top > 0: shape.crop_top = top
        if bottom > 0: shape.crop_bottom = bottom
        
        agent.save()
        
    return {
        "status": "success",
        "file": str(filepath),
        "slide_index": slide_index,
        "shape_index": shape_index,
        "crop_applied": {
            "left": left,
            "right": right,
            "top": top,
            "bottom": bottom
        }
    }

def main():
    parser = argparse.ArgumentParser(description="Crop PowerPoint image")
    parser.add_argument('--file', required=True, type=Path, help='PowerPoint file path')
    parser.add_argument('--slide', required=True, type=int, help='Slide index (0-based)')
    parser.add_argument('--shape', required=True, type=int, help='Shape index (0-based)')
    parser.add_argument('--left', type=float, default=0.0, help='Crop from left (0.0-1.0)')
    parser.add_argument('--right', type=float, default=0.0, help='Crop from right (0.0-1.0)')
    parser.add_argument('--top', type=float, default=0.0, help='Crop from top (0.0-1.0)')
    parser.add_argument('--bottom', type=float, default=0.0, help='Crop from bottom (0.0-1.0)')
    parser.add_argument('--json', action='store_true', default=True, help='Output JSON')
    
    args = parser.parse_args()
    
    try:
        result = crop_image(
            filepath=args.file, 
            slide_index=args.slide, 
            shape_index=args.shape,
            left=args.left,
            right=args.right,
            top=args.top,
            bottom=args.bottom
        )
        print(json.dumps(result, indent=2))
        sys.exit(0)
    except Exception as e:
        print(json.dumps({"status": "error", "error": str(e), "error_type": type(e).__name__}, indent=2))
        sys.exit(1)

if __name__ == "__main__":
    main()

```

# tools/ppt_delete_slide.py
```py
#!/usr/bin/env python3
"""
PowerPoint Delete Slide Tool
Remove a slide from the presentation

Usage:
    uv python ppt_delete_slide.py --file presentation.pptx --index 1 --json
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError, SlideNotFoundError
)


def delete_slide(filepath: Path, index: int) -> Dict[str, Any]:
    """Delete slide at index."""
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        total_slides = agent.get_slide_count()
        if not 0 <= index < total_slides:
            raise SlideNotFoundError(f"Index {index} out of range (0-{total_slides-1})")
            
        agent.delete_slide(index)
        agent.save()
        
        new_count = agent.get_slide_count()
    
    return {
        "status": "success",
        "file": str(filepath),
        "deleted_index": index,
        "remaining_slides": new_count
    }


def main():
    parser = argparse.ArgumentParser(description="Delete PowerPoint slide")
    parser.add_argument('--file', required=True, type=Path, help='PowerPoint file path')
    parser.add_argument('--index', required=True, type=int, help='Slide index to delete (0-based)')
    parser.add_argument('--json', action='store_true', default=True, help='Output JSON')
    
    args = parser.parse_args()
    
    try:
        result = delete_slide(filepath=args.file, index=args.index)
        print(json.dumps(result, indent=2))
        sys.exit(0)
    except Exception as e:
        print(json.dumps({"status": "error", "error": str(e), "error_type": type(e).__name__}, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_duplicate_slide.py
```py
#!/usr/bin/env python3
"""
PowerPoint Duplicate Slide Tool
Clone an existing slide

Usage:
    uv python ppt_duplicate_slide.py --file presentation.pptx --index 0 --json
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError, SlideNotFoundError
)


def duplicate_slide(filepath: Path, index: int) -> Dict[str, Any]:
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        total = agent.get_slide_count()
        if not 0 <= index < total:
            raise SlideNotFoundError(f"Index {index} out of range (0-{total-1})")
            
        new_index = agent.duplicate_slide(index)
        agent.save()
        
        # Get info about new slide
        info = agent.get_slide_info(new_index)
        
    return {
        "status": "success",
        "file": str(filepath),
        "source_index": index,
        "new_slide_index": new_index,
        "total_slides": new_index + 1,
        "layout": info["layout"]
    }


def main():
    parser = argparse.ArgumentParser(description="Duplicate PowerPoint slide")
    parser.add_argument('--file', required=True, type=Path, help='PowerPoint file path')
    parser.add_argument('--index', required=True, type=int, help='Source slide index (0-based)')
    parser.add_argument('--json', action='store_true', default=True, help='Output JSON')
    
    args = parser.parse_args()
    
    try:
        result = duplicate_slide(filepath=args.file, index=args.index)
        print(json.dumps(result, indent=2))
        sys.exit(0)
    except Exception as e:
        print(json.dumps({"status": "error", "error": str(e), "error_type": type(e).__name__}, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_export_images.py
```py
#!/usr/bin/env python3
"""
PowerPoint Export Images Tool
Export each slide as PNG or JPG image

Usage:
    uv python ppt_export_images.py --file presentation.pptx --output-dir output/ --format png --json

Exit Codes:
    0: Success
    1: Error occurred

Requirements:
    LibreOffice must be installed for image export
    - Linux: sudo apt install libreoffice-impress
    - macOS: brew install --cask libreoffice
    - Windows: Download from https://www.libreoffice.org/
"""

import sys
import json
import argparse
import subprocess
import shutil
from pathlib import Path
from typing import Dict, Any, List

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError
)


def check_libreoffice() -> bool:
    """Check if LibreOffice is installed."""
    return shutil.which('soffice') is not None or shutil.which('libreoffice') is not None


def export_images(
    filepath: Path,
    output_dir: Path,
    format: str = "png",
    prefix: str = "slide_"
) -> Dict[str, Any]:
    """Export slides as images."""
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if not filepath.suffix.lower() == '.pptx':
        raise ValueError(f"Input must be .pptx file, got: {filepath.suffix}")
    
    if format.lower() not in ['png', 'jpg', 'jpeg']:
        raise ValueError(f"Format must be png or jpg, got: {format}")
    
    # Normalize format
    format_ext = 'png' if format.lower() == 'png' else 'jpg'
    
    # Check LibreOffice
    if not check_libreoffice():
        raise RuntimeError(
            "LibreOffice not found. Image export requires LibreOffice.\n"
            "Install:\n"
            "  Linux: sudo apt install libreoffice-impress\n"
            "  macOS: brew install --cask libreoffice\n"
            "  Windows: https://www.libreoffice.org/download/"
        )
    
    # Create output directory
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # Get slide count first
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath, acquire_lock=False)
        slide_count = agent.get_slide_count()
    
    # Use PDF-intermediate workflow for better reliability
    # 1. Export to PDF
    base_name = filepath.stem
    pdf_path = output_dir / f"{base_name}.pdf"
    
    # Use LibreOffice to export to PDF
    cmd_pdf = [
        'soffice' if shutil.which('soffice') else 'libreoffice',
        '--headless',
        '--convert-to', 'pdf',
        '--outdir', str(output_dir),
        str(filepath)
    ]
    
    result_pdf = subprocess.run(cmd_pdf, capture_output=True, text=True, timeout=120)
    
    if result_pdf.returncode != 0:
        raise PowerPointAgentError(
            f"PDF export failed: {result_pdf.stderr}\n"
            f"Command: {' '.join(cmd_pdf)}"
        )
        
    # 2. Convert PDF to Images using pdftoppm (if available)
    if shutil.which('pdftoppm'):
        # pdftoppm -png -r 150 input.pdf output_prefix
        cmd_img = [
            'pdftoppm',
            f"-{format_ext}",
            '-r', '150',  # 150 DPI
            str(pdf_path),
            str(output_dir / base_name)
        ]
        
        result_img = subprocess.run(cmd_img, capture_output=True, text=True, timeout=120)
        
        if result_img.returncode != 0:
             # Fallback to direct export if pdftoppm fails
             print(f"Warning: pdftoppm failed, falling back to LibreOffice direct export: {result_img.stderr}", file=sys.stderr)
             _export_direct(filepath, output_dir, format_ext)
        
        # Clean up PDF
        # Clean up PDF
        if pdf_path.exists():
            pdf_path.unlink()
            
    else:
        # Fallback to direct export if pdftoppm not installed
        print("Warning: pdftoppm not found, using LibreOffice direct export (may be incomplete)", file=sys.stderr)
        _export_direct(filepath, output_dir, format_ext)

    # 3. Process and rename files
    return _scan_and_process_results(filepath, output_dir, format_ext, prefix)


def _export_direct(filepath: Path, output_dir: Path, format_ext: str):
    """Direct export using LibreOffice (legacy method)."""
    cmd = [
        'soffice' if shutil.which('soffice') else 'libreoffice',
        '--headless',
        '--convert-to', format_ext,
        '--outdir', str(output_dir),
        str(filepath)
    ]
    
    result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
    
    if result.returncode != 0:
        raise PowerPointAgentError(
            f"Image export failed: {result.stderr}\n"
            f"Command: {' '.join(cmd)}"
        )


def _scan_and_process_results(
    filepath: Path,
    output_dir: Path,
    format_ext: str,
    prefix: str
) -> Dict[str, Any]:
    """Find, rename, and report exported images."""
    # LibreOffice/pdftoppm creates files named: presentation.png or presentation-1.png
    base_name = filepath.stem
    
    # Find exported images using a robust pattern match
    # Match base_name*.ext
    candidates = sorted(output_dir.glob(f"{base_name}*.{format_ext}"))
    
    exported_files = []
    
    # Rename found files sequentially
    for i, old_file in enumerate(candidates):
        new_file = output_dir / f"{prefix}{i+1:03d}.{format_ext}"
        
        # Handle case where source and dest are same
        if old_file != new_file:
            # If target exists (from previous run), remove it
            if new_file.exists():
                new_file.unlink()
            old_file.rename(new_file)
            exported_files.append(new_file)
        else:
            exported_files.append(old_file)
    
    if len(exported_files) == 0:
        raise PowerPointAgentError(
            "Export completed but no image files found. "
            f"Expected files in: {output_dir}"
        )
    
    # Get file sizes
    total_size = sum(f.stat().st_size for f in exported_files)
    
    return {
        "status": "success",
        "input_file": str(filepath),
        "output_dir": str(output_dir),
        "format": format_ext,
        "slides_exported": len(exported_files),
        "files": [str(f) for f in exported_files],
        "total_size_bytes": total_size,
        "total_size_mb": round(total_size / (1024 * 1024), 2),
        "average_size_mb": round(total_size / (1024 * 1024) / len(exported_files), 2) if exported_files else 0
    }


def main():
    parser = argparse.ArgumentParser(
        description="Export PowerPoint slides as images",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Export as PNG
  uv python ppt_export_images.py \\
    --file presentation.pptx \\
    --output-dir slides/ \\
    --format png \\
    --json
  
  # Export as JPG with custom prefix
  uv python ppt_export_images.py \\
    --file presentation.pptx \\
    --output-dir images/ \\
    --format jpg \\
    --prefix deck_ \\
    --json

Output Files:
  Files are named: <prefix><number>.<format>
  Examples:
    slide_001.png
    slide_002.png
    deck_001.jpg
    deck_002.jpg

Requirements:
  LibreOffice must be installed:
  
  Linux:
    sudo apt install libreoffice-impress
  
  macOS:
    brew install --cask libreoffice
  
  Windows:
    Download from https://www.libreoffice.org/download/

Use Cases:
  - Website screenshots
  - Social media sharing
  - Email attachments
  - Documentation
  - Thumbnails for preview
  - Archive for reference

Format Comparison:
  PNG:
    - Lossless compression
    - Better for text/diagrams
    - Larger file size
    - Supports transparency
  
  JPG:
    - Lossy compression
    - Better for photos
    - Smaller file size
    - No transparency

Performance:
  - Export time: ~2-3 seconds per slide
  - File size: 200-500 KB per slide (PNG)
  - File size: 100-300 KB per slide (JPG)

Troubleshooting:
  If export fails:
  1. Verify LibreOffice: soffice --version
  2. Check disk space
  3. Ensure file not corrupted
  4. Try shorter timeout for small files
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file to export'
    )
    
    parser.add_argument(
        '--output-dir',
        required=True,
        type=Path,
        help='Output directory for images'
    )
    
    parser.add_argument(
        '--format',
        choices=['png', 'jpg', 'jpeg'],
        default='png',
        help='Image format (default: png)'
    )
    
    parser.add_argument(
        '--prefix',
        default='slide_',
        help='Filename prefix (default: slide_)'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        help='Output JSON response'
    )
    
    args = parser.parse_args()
    
    try:
        result = export_images(
            filepath=args.file,
            output_dir=args.output_dir,
            format=args.format,
            prefix=args.prefix
        )
        
        if args.json:
            print(json.dumps(result, indent=2))
        else:
            print(f"✅ Exported {result['slides_exported']} slides to {result['output_dir']}")
            print(f"   Format: {result['format'].upper()}")
            print(f"   Total size: {result['total_size_mb']} MB")
            print(f"   Average: {result['average_size_mb']} MB per slide")
        
        sys.exit(0)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__
        }
        
        if args.json:
            print(json.dumps(error_result, indent=2))
        else:
            print(f"❌ Error: {e}", file=sys.stderr)
        
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_export_pdf.py
```py
#!/usr/bin/env python3
"""
PowerPoint Export PDF Tool
Export presentation to PDF format

Usage:
    uv python ppt_export_pdf.py --file presentation.pptx --output presentation.pdf --json

Exit Codes:
    0: Success
    1: Error occurred

Requirements:
    LibreOffice must be installed for PDF export
    - Linux: sudo apt install libreoffice-impress
    - macOS: brew install --cask libreoffice
    - Windows: Download from https://www.libreoffice.org/
"""

import sys
import json
import argparse
import subprocess
import shutil
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError
)


def check_libreoffice() -> bool:
    """Check if LibreOffice is installed."""
    return shutil.which('soffice') is not None or shutil.which('libreoffice') is not None


def export_pdf(
    filepath: Path,
    output: Path
) -> Dict[str, Any]:
    """Export presentation to PDF."""
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if not filepath.suffix.lower() == '.pptx':
        raise ValueError(f"Input must be .pptx file, got: {filepath.suffix}")
    
    # Check LibreOffice
    if not check_libreoffice():
        raise RuntimeError(
            "LibreOffice not found. PDF export requires LibreOffice.\n"
            "Install:\n"
            "  Linux: sudo apt install libreoffice-impress\n"
            "  macOS: brew install --cask libreoffice\n"
            "  Windows: https://www.libreoffice.org/download/"
        )
    
    # Ensure output directory exists
    output.parent.mkdir(parents=True, exist_ok=True)
    
    # Use LibreOffice to convert
    # Note: --headless runs without GUI, --convert-to pdf exports to PDF
    cmd = [
        'soffice' if shutil.which('soffice') else 'libreoffice',
        '--headless',
        '--convert-to', 'pdf',
        '--outdir', str(output.parent),
        str(filepath)
    ]
    
    result = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
    
    if result.returncode != 0:
        raise PowerPointAgentError(
            f"PDF export failed: {result.stderr}\n"
            f"Command: {' '.join(cmd)}"
        )
    
    # LibreOffice names output file based on input, rename if needed
    expected_pdf = filepath.parent / f"{filepath.stem}.pdf"
    if output.parent != filepath.parent:
        expected_pdf = output.parent / f"{filepath.stem}.pdf"
    
    if expected_pdf != output and expected_pdf.exists():
        if output.exists():
            output.unlink()
        expected_pdf.rename(output)
    elif not output.exists() and expected_pdf.exists():
        expected_pdf.rename(output)
    
    if not output.exists():
        raise PowerPointAgentError("PDF export completed but output file not found")
    
    # Get file sizes
    input_size = filepath.stat().st_size
    output_size = output.stat().st_size
    
    return {
        "status": "success",
        "input_file": str(filepath),
        "output_file": str(output),
        "input_size_bytes": input_size,
        "input_size_mb": round(input_size / (1024 * 1024), 2),
        "output_size_bytes": output_size,
        "output_size_mb": round(output_size / (1024 * 1024), 2),
        "size_ratio": round(output_size / input_size, 2) if input_size > 0 else 0
    }


def main():
    parser = argparse.ArgumentParser(
        description="Export PowerPoint presentation to PDF",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Basic export
  uv python ppt_export_pdf.py \\
    --file presentation.pptx \\
    --output presentation.pdf \\
    --json
  
  # Export with automatic naming
  uv python ppt_export_pdf.py \\
    --file quarterly_report.pptx \\
    --output reports/q4_report.pdf \\
    --json

Requirements:
  LibreOffice must be installed:
  
  Linux:
    sudo apt update
    sudo apt install libreoffice-impress
  
  macOS:
    brew install --cask libreoffice
  
  Windows:
    Download from https://www.libreoffice.org/download/

Verification:
  # Check LibreOffice installation
  soffice --version
  # or
  libreoffice --version

Use Cases:
  - Share presentations as PDFs
  - Archive presentations
  - Print-ready versions
  - Email distribution (smaller, universal format)
  - Document repositories

PDF Benefits:
  - Universal compatibility
  - Prevents editing
  - Smaller file size typically
  - Better for printing
  - Preserves layout exactly

Limitations:
  - Animations not preserved
  - Embedded videos become static
  - No speaker notes in output
  - Transitions removed
  - Interactive elements static

Performance:
  - Export time: ~2-5 seconds per slide
  - Large presentations (100+ slides): ~3-5 minutes
  - File size typically 30-50% of .pptx

Troubleshooting:
  If export fails:
  1. Verify LibreOffice installed: soffice --version
  2. Check file not corrupted: open in PowerPoint
  3. Ensure disk space available
  4. Try shorter timeout for small files
  5. Check LibreOffice logs
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file to export'
    )
    
    parser.add_argument(
        '--output',
        required=True,
        type=Path,
        help='Output PDF file path'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        help='Output JSON response'
    )
    
    args = parser.parse_args()
    
    try:
        # Ensure output has .pdf extension
        if not args.output.suffix.lower() == '.pdf':
            args.output = args.output.with_suffix('.pdf')
        
        result = export_pdf(
            filepath=args.file,
            output=args.output
        )
        
        if args.json:
            print(json.dumps(result, indent=2))
        else:
            print(f"✅ Exported to PDF: {result['output_file']}")
            print(f"   Input: {result['input_size_mb']} MB")
            print(f"   Output: {result['output_size_mb']} MB")
            print(f"   Ratio: {int(result['size_ratio'] * 100)}%")
        
        sys.exit(0)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__
        }
        
        if args.json:
            print(json.dumps(error_result, indent=2))
        else:
            print(f"❌ Error: {e}", file=sys.stderr)
        
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_extract_notes.py
```py
#!/usr/bin/env python3
"""
PowerPoint Extract Notes Tool
Get speaker notes from all slides

Usage:
    uv python ppt_extract_notes.py --file presentation.pptx --json
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import PowerPointAgent


def extract_notes(filepath: Path) -> Dict[str, Any]:
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath, acquire_lock=False)
        notes = agent.extract_notes()
        
    return {
        "status": "success",
        "file": str(filepath),
        "notes_found": len(notes),
        "notes": notes # Dict {slide_index: text}
    }


def main():
    parser = argparse.ArgumentParser(description="Extract speaker notes")
    parser.add_argument('--file', required=True, type=Path, help='PowerPoint file path')
    parser.add_argument('--json', action='store_true', default=True, help='Output JSON')
    
    args = parser.parse_args()
    
    try:
        result = extract_notes(filepath=args.file)
        print(json.dumps(result, indent=2))
        sys.exit(0)
    except Exception as e:
        print(json.dumps({"status": "error", "error": str(e), "error_type": type(e).__name__}, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_format_chart.py
```py
#!/usr/bin/env python3
"""
PowerPoint Format Chart Tool
Format existing chart (title, legend position)

Usage:
    uv python ppt_format_chart.py --file presentation.pptx --slide 1 --chart 0 --title "Revenue Growth" --legend bottom --json

Exit Codes:
    0: Success
    1: Error occurred
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError, SlideNotFoundError
)


def format_chart(
    filepath: Path,
    slide_index: int,
    chart_index: int = 0,
    title: str = None,
    legend_position: str = None
) -> Dict[str, Any]:
    """Format chart on slide."""
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if title is None and legend_position is None:
        raise ValueError("At least one formatting option (title or legend) must be specified")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Validate slide index
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides-1})"
            )
        
        # Format chart
        agent.format_chart(
            slide_index=slide_index,
            chart_index=chart_index,
            title=title,
            legend_position=legend_position
        )
        
        # Save
        agent.save()
    
    return {
        "status": "success",
        "file": str(filepath),
        "slide_index": slide_index,
        "chart_index": chart_index,
        "formatting_applied": {
            "title": title,
            "legend_position": legend_position
        }
    }


def main():
    parser = argparse.ArgumentParser(
        description="Format PowerPoint chart",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Set chart title
  uv python ppt_format_chart.py \\
    --file presentation.pptx \\
    --slide 1 \\
    --chart 0 \\
    --title "Revenue Growth Trend" \\
    --json
  
  # Position legend at bottom
  uv python ppt_format_chart.py \\
    --file presentation.pptx \\
    --slide 2 \\
    --chart 0 \\
    --legend bottom \\
    --json
  
  # Set title and legend
  uv python ppt_format_chart.py \\
    --file presentation.pptx \\
    --slide 3 \\
    --chart 0 \\
    --title "Q4 Performance" \\
    --legend right \\
    --json

Legend Positions:
  - bottom: Below chart (common)
  - right: Right side of chart (default)
  - top: Above chart
  - left: Left side of chart

Finding Charts:
  Charts are indexed in order they appear on the slide.
  First chart = 0, second = 1, etc.
  
  To find charts:
  uv python ppt_get_slide_info.py --file presentation.pptx --slide 1 --json

Best Practices:
  - Keep titles concise and descriptive
  - Place legend where it doesn't obscure data
  - Bottom legend works well for wide charts
  - Right legend works well for tall charts
  - Consider removing legend if only 1 series

Chart Formatting Limitations:
  Note: python-pptx has limited chart formatting support.
  This tool handles:
  - Chart title text
  - Legend position
  
  Not supported (requires PowerPoint):
  - Individual series colors
  - Axis formatting
  - Data labels
  - Chart styles/templates
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file path'
    )
    
    parser.add_argument(
        '--slide',
        required=True,
        type=int,
        help='Slide index (0-based)'
    )
    
    parser.add_argument(
        '--chart',
        type=int,
        default=0,
        help='Chart index on slide (default: 0)'
    )
    
    parser.add_argument(
        '--title',
        help='Chart title text'
    )
    
    parser.add_argument(
        '--legend',
        choices=['bottom', 'left', 'right', 'top'],
        help='Legend position'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        help='Output JSON response'
    )
    
    args = parser.parse_args()
    
    try:
        result = format_chart(
            filepath=args.file,
            slide_index=args.slide,
            chart_index=args.chart,
            title=args.title,
            legend_position=args.legend
        )
        
        if args.json:
            print(json.dumps(result, indent=2))
        else:
            print(f"✅ Formatted chart on slide {result['slide_index']}")
            if args.title:
                print(f"   Title: {result['formatting_applied']['title']}")
            if args.legend:
                print(f"   Legend: {result['formatting_applied']['legend_position']}")
        
        sys.exit(0)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__
        }
        
        if args.json:
            print(json.dumps(error_result, indent=2))
        else:
            print(f"❌ Error: {e}", file=sys.stderr)
        
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_format_shape.py
```py
#!/usr/bin/env python3
"""
PowerPoint Format Shape Tool
Update fill color, line color, and line width of an existing shape

Usage:
    uv python ppt_format_shape.py --file presentation.pptx --slide 0 --shape 1 --fill-color "#FF0000" --json
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError, SlideNotFoundError
)


def format_shape(
    filepath: Path,
    slide_index: int,
    shape_index: int,
    fill_color: str = None,
    line_color: str = None,
    line_width: float = None
) -> Dict[str, Any]:
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if all(v is None for v in [fill_color, line_color, line_width]):
        raise ValueError("At least one formatting option required")

    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        total = agent.get_slide_count()
        if not 0 <= slide_index < total:
            raise SlideNotFoundError(f"Slide index {slide_index} out of range")
            
        agent.format_shape(
            slide_index=slide_index,
            shape_index=shape_index,
            fill_color=fill_color,
            line_color=line_color,
            line_width=line_width
        )
        agent.save()
        
    return {
        "status": "success",
        "file": str(filepath),
        "slide_index": slide_index,
        "shape_index": shape_index,
        "formatting": {
            "fill_color": fill_color,
            "line_color": line_color,
            "line_width": line_width
        }
    }


def main():
    parser = argparse.ArgumentParser(description="Format PowerPoint shape")
    parser.add_argument('--file', required=True, type=Path, help='PowerPoint file path')
    parser.add_argument('--slide', required=True, type=int, help='Slide index (0-based)')
    parser.add_argument('--shape', required=True, type=int, help='Shape index (0-based)')
    parser.add_argument('--fill-color', help='Fill hex color')
    parser.add_argument('--line-color', help='Line hex color')
    parser.add_argument('--line-width', type=float, help='Line width in points')
    parser.add_argument('--json', action='store_true', default=True, help='Output JSON')
    
    args = parser.parse_args()
    
    try:
        result = format_shape(
            filepath=args.file, 
            slide_index=args.slide, 
            shape_index=args.shape,
            fill_color=args.fill_color,
            line_color=args.line_color,
            line_width=args.line_width
        )
        print(json.dumps(result, indent=2))
        sys.exit(0)
    except Exception as e:
        print(json.dumps({"status": "error", "error": str(e), "error_type": type(e).__name__}, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_format_text.py
```py
#!/usr/bin/env python3
"""
PowerPoint Format Text Tool
Format existing text (font, size, color, bold, italic)

Usage:
    uv python ppt_format_text.py --file presentation.pptx --slide 0 --shape 0 --font-name Arial --font-size 24 --color "#FF0000" --bold --json

Exit Codes:
    0: Success
    1: Error occurred
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError, SlideNotFoundError
)


def format_text(
    filepath: Path,
    slide_index: int,
    shape_index: int,
    font_name: str = None,
    font_size: int = None,
    color: str = None,
    bold: bool = None,
    italic: bool = None
) -> Dict[str, Any]:
    """Format text in specified shape."""
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    # Check that at least one formatting option is provided
    if all(v is None for v in [font_name, font_size, color, bold, italic]):
        raise ValueError("At least one formatting option must be specified")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Validate slide index
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides-1})"
            )
        
        # Get slide info to validate shape index
        slide_info = agent.get_slide_info(slide_index)
        if shape_index >= slide_info["shape_count"]:
            raise ValueError(
                f"Shape index {shape_index} out of range (0-{slide_info['shape_count']-1})"
            )
        
        # Format text
        agent.format_text(
            slide_index=slide_index,
            shape_index=shape_index,
            font_name=font_name,
            font_size=font_size,
            bold=bold,
            italic=italic,
            color=color
        )
        
        # Save
        agent.save()
    
    return {
        "status": "success",
        "file": str(filepath),
        "slide_index": slide_index,
        "shape_index": shape_index,
        "formatting_applied": {
            "font_name": font_name,
            "font_size": font_size,
            "color": color,
            "bold": bold,
            "italic": italic
        }
    }


def main():
    parser = argparse.ArgumentParser(
        description="Format text in PowerPoint shape",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Change font and size
  uv python ppt_format_text.py \\
    --file presentation.pptx \\
    --slide 0 \\
    --shape 0 \\
    --font-name "Arial" \\
    --font-size 24 \\
    --json
  
  # Make text bold and red
  uv python ppt_format_text.py \\
    --file presentation.pptx \\
    --slide 1 \\
    --shape 2 \\
    --bold \\
    --color "#FF0000" \\
    --json
  
  # Comprehensive formatting
  uv python ppt_format_text.py \\
    --file presentation.pptx \\
    --slide 0 \\
    --shape 1 \\
    --font-name "Calibri" \\
    --font-size 18 \\
    --bold \\
    --italic \\
    --color "#0070C0" \\
    --json

Common Fonts:
  - Calibri (default Office)
  - Arial
  - Times New Roman
  - Helvetica
  - Georgia
  - Verdana
  - Tahoma

Color Examples:
  - Black: #000000
  - White: #FFFFFF
  - Red: #FF0000
  - Blue: #0070C0
  - Green: #00B050
  - Orange: #FFC000

Finding Shape Index:
  # Use ppt_get_slide_info.py to list shapes
  uv python ppt_get_slide_info.py --file presentation.pptx --slide 0 --json

Best Practices:
  - Use standard fonts for compatibility
  - Keep font sizes 12pt or larger
  - Ensure sufficient color contrast
  - Test on actual presentation display
  - Use bold for emphasis sparingly
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file path'
    )
    
    parser.add_argument(
        '--slide',
        required=True,
        type=int,
        help='Slide index (0-based)'
    )
    
    parser.add_argument(
        '--shape',
        required=True,
        type=int,
        help='Shape index (0-based)'
    )
    
    parser.add_argument(
        '--font-name',
        help='Font name (e.g., Arial, Calibri)'
    )
    
    parser.add_argument(
        '--font-size',
        type=int,
        help='Font size in points'
    )
    
    parser.add_argument(
        '--color',
        help='Text color (hex, e.g., #FF0000)'
    )
    
    parser.add_argument(
        '--bold',
        action='store_true',
        help='Make text bold'
    )
    
    parser.add_argument(
        '--italic',
        action='store_true',
        help='Make text italic'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        help='Output JSON response'
    )
    
    args = parser.parse_args()
    
    try:
        result = format_text(
            filepath=args.file,
            slide_index=args.slide,
            shape_index=args.shape,
            font_name=args.font_name,
            font_size=args.font_size,
            color=args.color,
            bold=args.bold if args.bold else None,
            italic=args.italic if args.italic else None
        )
        
        if args.json:
            print(json.dumps(result, indent=2))
        else:
            print(f"✅ Formatted text in slide {result['slide_index']}, shape {result['shape_index']}")
            formatting = result['formatting_applied']
            if formatting['font_name']:
                print(f"   Font: {formatting['font_name']}")
            if formatting['font_size']:
                print(f"   Size: {formatting['font_size']}pt")
            if formatting['color']:
                print(f"   Color: {formatting['color']}")
            if formatting['bold']:
                print(f"   Bold: Yes")
            if formatting['italic']:
                print(f"   Italic: Yes")
        
        sys.exit(0)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__
        }
        
        if args.json:
            print(json.dumps(error_result, indent=2))
        else:
            print(f"❌ Error: {e}", file=sys.stderr)
        
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_get_info.py
```py
#!/usr/bin/env python3
"""
PowerPoint Get Info Tool
Get presentation metadata (slide count, dimensions, file size)

Usage:
    uv python ppt_get_info.py --file presentation.pptx --json

Exit Codes:
    0: Success
    1: Error occurred
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError
)


def get_info(filepath: Path) -> Dict[str, Any]:
    """Get presentation information."""
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath, acquire_lock=False)
        
        # Get presentation info
        info = agent.get_presentation_info()
    
    return {
        "status": "success",
        "file": info["file"],
        "slide_count": info["slide_count"],
        "file_size_bytes": info.get("file_size_bytes", 0),
        "file_size_mb": info.get("file_size_mb", 0),
        "slide_dimensions": {
            "width_inches": info["slide_width_inches"],
            "height_inches": info["slide_height_inches"],
            "aspect_ratio": info["aspect_ratio"]
        },
        "layouts": info["layouts"],
        "layout_count": len(info["layouts"]),
        "modified": info.get("modified")
    }


def main():
    parser = argparse.ArgumentParser(
        description="Get PowerPoint presentation information",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Get presentation info
  uv python ppt_get_info.py \\
    --file presentation.pptx \\
    --json

Output Information:
  - File path and size
  - Total slide count
  - Slide dimensions (width x height)
  - Aspect ratio (16:9, 4:3, etc.)
  - Available layouts
  - Last modified date

Use Cases:
  - Verify presentation structure
  - Check aspect ratio before editing
  - List available layouts
  - File size checking
  - Metadata inspection

Example Output:
{
  "file": "presentation.pptx",
  "slide_count": 15,
  "file_size_mb": 2.45,
  "slide_dimensions": {
    "width_inches": 10.0,
    "height_inches": 7.5,
    "aspect_ratio": "16:9"
  },
  "layouts": [
    "Title Slide",
    "Title and Content",
    "Section Header"
  ],
  "layout_count": 11
}

Aspect Ratios:
  - 16:9 (Widescreen): Most common, modern standard
  - 4:3 (Standard): Traditional, older format
  - 16:10: Some displays, between 16:9 and 4:3

Layout Information:
  The layouts list shows all slide layouts available in the presentation.
  Use these names with:
  - ppt_create_new.py --layout "Title Slide"
  - ppt_add_slide.py --layout "Title and Content"
  - ppt_set_slide_layout.py --layout "Section Header"
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file path'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default)'
    )
    
    args = parser.parse_args()
    
    try:
        result = get_info(filepath=args.file)
        
        print(json.dumps(result, indent=2))
        sys.exit(0)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__
        }
        
        print(json.dumps(error_result, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_get_slide_info.py
```py
#!/usr/bin/env python3
"""
PowerPoint Get Slide Info Tool
Get detailed information about slide content (shapes, images, text)

Usage:
    uv python ppt_get_slide_info.py --file presentation.pptx --slide 0 --json

Exit Codes:
    0: Success
    1: Error occurred
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError, SlideNotFoundError
)


def get_slide_info(
    filepath: Path,
    slide_index: int
) -> Dict[str, Any]:
    """Get detailed slide information."""
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath, acquire_lock=False)
        
        # Validate slide index
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides-1})"
            )
        
        # Get slide info
        slide_info = agent.get_slide_info(slide_index)
    
    return {
        "status": "success",
        "file": str(filepath),
        "slide_index": slide_index,
        "layout": slide_info["layout"],
        "shape_count": slide_info["shape_count"],
        "shapes": slide_info["shapes"],
        "has_notes": slide_info["has_notes"]
    }


def main():
    parser = argparse.ArgumentParser(
        description="Get PowerPoint slide information",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Get info for first slide
  uv python ppt_get_slide_info.py \\
    --file presentation.pptx \\
    --slide 0 \\
    --json
  
  # Get info for specific slide
  uv python ppt_get_slide_info.py \\
    --file presentation.pptx \\
    --slide 5 \\
    --json

Output Information:
  - Slide layout name
  - Total shape count
  - List of all shapes with:
    - Shape index (for targeting with other tools)
    - Shape type (PLACEHOLDER, PICTURE, TEXT_BOX, etc.)
    - Shape name
    - Whether it contains text
    - Text preview (first 100 chars)
    - Image size (for pictures)

Use Cases:
  - Find shape indices for ppt_format_text.py
  - Locate images for ppt_replace_image.py
  - Inspect slide layout
  - Audit slide content
  - Debug presentation structure

Finding Shape Indices:
  Use this tool before:
  - ppt_format_text.py (need shape index)
  - ppt_replace_image.py (need image name)
  - ppt_format_shape.py (need shape index)

Example Output:
{
  "slide_index": 0,
  "layout": "Title Slide",
  "shape_count": 3,
  "shapes": [
    {
      "index": 0,
      "type": "PLACEHOLDER",
      "name": "Title 1",
      "has_text": true,
      "text": "My Presentation"
    },
    {
      "index": 1,
      "type": "PICTURE",
      "name": "company_logo",
      "has_text": false,
      "image_size_bytes": 45678
    }
  ]
}
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file path'
    )
    
    parser.add_argument(
        '--slide',
        required=True,
        type=int,
        help='Slide index (0-based)'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default)'
    )
    
    args = parser.parse_args()
    
    try:
        result = get_slide_info(
            filepath=args.file,
            slide_index=args.slide
        )
        
        print(json.dumps(result, indent=2))
        sys.exit(0)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__
        }
        
        print(json.dumps(error_result, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_insert_image.py
```py
#!/usr/bin/env python3
"""
PowerPoint Insert Image Tool
Insert image into slide with automatic aspect ratio handling

Usage:
    uv python ppt_insert_image.py --file presentation.pptx --slide 0 --image logo.png --position '{"left":"10%","top":"10%"}' --size '{"width":"20%","height":"auto"}' --json

Exit Codes:
    0: Success
    1: Error occurred
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError, SlideNotFoundError,
    ImageNotFoundError, InvalidPositionError
)


def insert_image(
    filepath: Path,
    slide_index: int,
    image_path: Path,
    position: Dict[str, Any],
    size: Dict[str, Any] = None,
    compress: bool = False,
    alt_text: str = None
) -> Dict[str, Any]:
    """Insert image into slide."""
    
    if not filepath.exists():
        raise FileNotFoundError(f"Presentation file not found: {filepath}")
    
    if not image_path.exists():
        raise ImageNotFoundError(f"Image file not found: {image_path}")
    
    # Validate image format
    valid_extensions = {'.png', '.jpg', '.jpeg', '.gif', '.bmp'}
    if image_path.suffix.lower() not in valid_extensions:
        raise ValueError(
            f"Unsupported image format: {image_path.suffix}. "
            f"Supported: {valid_extensions}"
        )
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Validate slide index
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides-1})"
            )
        
        # Insert image
        agent.insert_image(
            slide_index=slide_index,
            image_path=image_path,
            position=position,
            size=size,
            compress=compress
        )
        
        # Set alt text if provided
        if alt_text:
            slide_info = agent.get_slide_info(slide_index)
            # Find the last shape (just added image)
            last_shape_idx = slide_info["shape_count"] - 1
            agent.set_image_properties(
                slide_index=slide_index,
                shape_index=last_shape_idx,
                alt_text=alt_text
            )
        
        # Get updated slide info
        slide_info = agent.get_slide_info(slide_index)
        
        # Save
        agent.save()
    
    # Get image file info
    image_size = image_path.stat().st_size
    
    return {
        "status": "success",
        "file": str(filepath),
        "slide_index": slide_index,
        "image_file": str(image_path),
        "image_size_bytes": image_size,
        "image_size_mb": round(image_size / (1024 * 1024), 2),
        "position": position,
        "size": size,
        "compressed": compress,
        "alt_text": alt_text,
        "slide_shape_count": slide_info["shape_count"]
    }


def main():
    parser = argparse.ArgumentParser(
        description="Insert image into PowerPoint slide",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Size Options:
  - {"width": "50%", "height": "auto"} - Auto-calculate height (recommended)
  - {"width": "auto", "height": "40%"} - Auto-calculate width
  - {"width": "30%", "height": "20%"} - Fixed dimensions
  - {"width": 3.0, "height": 2.0} - Absolute inches

Position Options:
  Same as ppt_add_text_box.py (percentage, anchor, grid, Excel-like)

Examples:
  # Insert logo (top-left, auto height)
  uv python ppt_insert_image.py \\
    --file presentation.pptx \\
    --slide 0 \\
    --image company_logo.png \\
    --position '{"left":"5%","top":"5%"}' \\
    --size '{"width":"15%","height":"auto"}' \\
    --alt-text "Company Logo" \\
    --json
  
  # Insert centered hero image
  uv python ppt_insert_image.py \\
    --file presentation.pptx \\
    --slide 1 \\
    --image product_photo.jpg \\
    --position '{"anchor":"center","offset_x":0,"offset_y":0}' \\
    --size '{"width":"80%","height":"auto"}' \\
    --compress \\
    --json
  
  # Insert screenshot (grid positioning)
  uv python ppt_insert_image.py \\
    --file presentation.pptx \\
    --slide 2 \\
    --image screenshot.png \\
    --position '{"grid":"B3"}' \\
    --size '{"width":"60%","height":"auto"}' \\
    --alt-text "Dashboard Screenshot - Q4 Metrics" \\
    --json
  
  # Insert chart export
  uv python ppt_insert_image.py \\
    --file presentation.pptx \\
    --slide 3 \\
    --image revenue_chart.png \\
    --position '{"left":"10%","top":"25%"}' \\
    --size '{"width":"80%","height":"auto"}' \\
    --compress \\
    --alt-text "Revenue Growth Chart 2020-2024" \\
    --json

Supported Formats:
  - PNG (recommended for logos, diagrams)
  - JPG/JPEG (recommended for photos)
  - GIF (animated not supported, will show first frame)
  - BMP (not recommended, large file size)

Best Practices:
  - Always use --alt-text for accessibility
  - Use "auto" for height/width to maintain aspect ratio
  - Use --compress for large images (>1MB)
  - Keep images under 2MB for best performance
  - Use PNG for transparency, JPG for photos
  - Recommended resolution: 1920x1080 max

Compression:
  --compress flag reduces image size by:
  - Resizing to max 1920px width
  - Converting RGBA to RGB
  - Optimizing JPEG quality to 85%
  - Typically reduces file size by 50-70%
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file path'
    )
    
    parser.add_argument(
        '--slide',
        required=True,
        type=int,
        help='Slide index (0-based)'
    )
    
    parser.add_argument(
        '--image',
        required=True,
        type=Path,
        help='Image file path'
    )
    
    parser.add_argument(
        '--position',
        required=True,
        type=json.loads,
        help='Position dict (JSON string)'
    )
    
    parser.add_argument(
        '--size',
        type=json.loads,
        help='Size dict (JSON string, optional - defaults to 50%% width with auto height)'
    )
    
    parser.add_argument(
        '--compress',
        action='store_true',
        help='Compress image before inserting (recommended for large images)'
    )
    
    parser.add_argument(
        '--alt-text',
        help='Alternative text for accessibility (highly recommended)'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        help='Output JSON response'
    )
    
    args = parser.parse_args()
    
    try:
        # Default size if not provided
        if not args.size:
            args.size = {"width": "50%", "height": "auto"}
        
        result = insert_image(
            filepath=args.file,
            slide_index=args.slide,
            image_path=args.image,
            position=args.position,
            size=args.size,
            compress=args.compress,
            alt_text=args.alt_text
        )
        
        if args.json:
            print(json.dumps(result, indent=2))
        else:
            print(f"✅ Inserted image into slide {result['slide_index']}")
            print(f"   Image: {result['image_file']}")
            print(f"   Size: {result['image_size_mb']} MB")
            if args.compress:
                print(f"   Compressed: Yes")
            if args.alt_text:
                print(f"   Alt text: {result['alt_text']}")
        
        sys.exit(0)
        
    except json.JSONDecodeError as e:
        error_result = {
            "status": "error",
            "error": f"Invalid JSON in position or size argument: {e}",
            "error_type": "JSONDecodeError",
            "hint": "Use single quotes around JSON: '{\"left\":\"20%\",\"top\":\"30%\"}'"
        }
        
        if args.json:
            print(json.dumps(error_result, indent=2))
        else:
            print(f"❌ Error: {error_result['error']}", file=sys.stderr)
            print(f"   Hint: {error_result['hint']}", file=sys.stderr)
        
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__
        }
        
        if args.json:
            print(json.dumps(error_result, indent=2))
        else:
            print(f"❌ Error: {e}", file=sys.stderr)
        
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_reorder_slides.py
```py
#!/usr/bin/env python3
"""
PowerPoint Reorder Slides Tool
Move a slide to a new position

Usage:
    uv python ppt_reorder_slides.py --file presentation.pptx --from-index 3 --to-index 1 --json
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError, SlideNotFoundError
)


def reorder_slides(filepath: Path, from_index: int, to_index: int) -> Dict[str, Any]:
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        total = agent.get_slide_count()
        if not 0 <= from_index < total:
            raise SlideNotFoundError(f"Source index {from_index} out of range")
        if not 0 <= to_index < total:
            raise SlideNotFoundError(f"Target index {to_index} out of range")
            
        agent.reorder_slides(from_index, to_index)
        agent.save()
        
    return {
        "status": "success",
        "file": str(filepath),
        "moved_from": from_index,
        "moved_to": to_index,
        "total_slides": total
    }


def main():
    parser = argparse.ArgumentParser(description="Reorder PowerPoint slides")
    parser.add_argument('--file', required=True, type=Path, help='PowerPoint file path')
    parser.add_argument('--from-index', required=True, type=int, help='Current slide index')
    parser.add_argument('--to-index', required=True, type=int, help='New slide index')
    parser.add_argument('--json', action='store_true', default=True, help='Output JSON')
    
    args = parser.parse_args()
    
    try:
        result = reorder_slides(filepath=args.file, from_index=args.from_index, to_index=args.to_index)
        print(json.dumps(result, indent=2))
        sys.exit(0)
    except Exception as e:
        print(json.dumps({"status": "error", "error": str(e), "error_type": type(e).__name__}, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_replace_image.py
```py
#!/usr/bin/env python3
"""
PowerPoint Replace Image Tool
Replace existing image (useful for logo/photo updates)

Usage:
    uv python ppt_replace_image.py --file presentation.pptx --slide 0 --old-image "logo" --new-image new_logo.png --json

Exit Codes:
    0: Success
    1: Error occurred
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError, SlideNotFoundError,
    ImageNotFoundError
)


def replace_image(
    filepath: Path,
    slide_index: int,
    old_image: str,
    new_image: Path,
    compress: bool = False
) -> Dict[str, Any]:
    """Replace image in presentation."""
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if not new_image.exists():
        raise ImageNotFoundError(f"New image not found: {new_image}")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Validate slide index
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides-1})"
            )
        
        # Try to replace by name
        replaced = agent.replace_image(
            slide_index=slide_index,
            old_image_name=old_image,
            new_image_path=new_image,
            compress=compress
        )
        
        if not replaced:
            raise ImageNotFoundError(
                f"Image '{old_image}' not found on slide {slide_index}. "
                "Use ppt_get_slide_info.py to list images."
            )
        
        # Save
        agent.save()
    
    # Get new image size
    new_size = new_image.stat().st_size
    
    return {
        "status": "success",
        "file": str(filepath),
        "slide_index": slide_index,
        "old_image": old_image,
        "new_image": str(new_image),
        "new_image_size_bytes": new_size,
        "new_image_size_mb": round(new_size / (1024 * 1024), 2),
        "compressed": compress,
        "replaced": True
    }


def main():
    parser = argparse.ArgumentParser(
        description="Replace image in PowerPoint presentation",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Replace logo by name
  uv python ppt_replace_image.py \\
    --file presentation.pptx \\
    --slide 0 \\
    --old-image "company_logo" \\
    --new-image new_logo.png \\
    --json
  
  # Replace and compress
  uv python ppt_replace_image.py \\
    --file presentation.pptx \\
    --slide 5 \\
    --old-image "product_photo" \\
    --new-image updated_photo.jpg \\
    --compress \\
    --json
  
  # Replace image with partial name match
  uv python ppt_replace_image.py \\
    --file presentation.pptx \\
    --slide 0 \\
    --old-image "logo" \\
    --new-image rebrand_logo.png \\
    --json

Finding Images:
  # List images on slide
  uv python ppt_get_slide_info.py \\
    --file presentation.pptx \\
    --slide 0 \\
    --json

Use Cases:
  - Logo updates (rebranding)
  - Product photo updates
  - Team photo updates
  - Chart/diagram updates
  - Screenshot updates

Search Strategy:
  The tool searches for images by:
  1. Exact name match
  2. Partial name match (contains)
  3. First match wins

Image Compression:
  --compress flag reduces size by:
  - Resizing to max 1920px width
  - Converting to JPEG at 85% quality
  - Typically reduces size 50-70%

Best Practices:
  - Use descriptive image names in PowerPoint
  - Keep new images similar dimensions to old
  - Use --compress for large images (>1MB)
  - Test on a copy first
  - Verify aspect ratios match
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file path'
    )
    
    parser.add_argument(
        '--slide',
        required=True,
        type=int,
        help='Slide index (0-based)'
    )
    
    parser.add_argument(
        '--old-image',
        required=True,
        help='Name or pattern of image to replace'
    )
    
    parser.add_argument(
        '--new-image',
        required=True,
        type=Path,
        help='Path to new image file'
    )
    
    parser.add_argument(
        '--compress',
        action='store_true',
        help='Compress new image before inserting'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        help='Output JSON response'
    )
    
    args = parser.parse_args()
    
    try:
        result = replace_image(
            filepath=args.file,
            slide_index=args.slide,
            old_image=args.old_image,
            new_image=args.new_image,
            compress=args.compress
        )
        
        if args.json:
            print(json.dumps(result, indent=2))
        else:
            print(f"✅ Replaced image on slide {result['slide_index']}")
            print(f"   Old: {result['old_image']}")
            print(f"   New: {result['new_image']}")
            print(f"   Size: {result['new_image_size_mb']} MB")
            if args.compress:
                print(f"   Compressed: Yes")
        
        sys.exit(0)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__
        }
        
        if args.json:
            print(json.dumps(error_result, indent=2))
        else:
            print(f"❌ Error: {e}", file=sys.stderr)
        
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_replace_text.py
```py
#!/usr/bin/env python3
"""
PowerPoint Replace Text Tool
Find and replace text across entire presentation

Usage:
    uv python ppt_replace_text.py --file presentation.pptx --find "Company Inc." --replace "Company LLC" --json

Exit Codes:
    0: Success
    1: Error occurred
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError
)


def replace_text(
    filepath: Path,
    find: str,
    replace: str,
    match_case: bool = False,
    whole_words: bool = False,
    dry_run: bool = False
) -> Dict[str, Any]:
    """Find and replace text across presentation."""
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if not find:
        raise ValueError("Find text cannot be empty")
    
    # For dry run, just scan without modifying
    if dry_run:
        with PowerPointAgent(filepath) as agent:
            agent.open(filepath, acquire_lock=False)  # Read-only
            
            # Count occurrences
            count = 0
            locations = []
            
            for slide_idx, slide in enumerate(agent.prs.slides):
                for shape_idx, shape in enumerate(slide.shapes):
                    if hasattr(shape, 'text_frame'):
                        text = shape.text_frame.text
                        
                        if match_case:
                            occurrences = text.count(find)
                        else:
                            occurrences = text.lower().count(find.lower())
                        
                        if occurrences > 0:
                            count += occurrences
                            locations.append({
                                "slide": slide_idx,
                                "shape": shape_idx,
                                "occurrences": occurrences,
                                "preview": text[:100]
                            })
        
        return {
            "status": "dry_run",
            "file": str(filepath),
            "find": find,
            "replace": replace,
            "matches_found": count,
            "locations": locations[:10],  # First 10 locations
            "total_locations": len(locations),
            "match_case": match_case
        }
    
    # Actual replacement
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Perform replacement
        count = agent.replace_text(
            find=find,
            replace=replace,
            match_case=match_case
        )
        
        # Save
        agent.save()
    
    return {
        "status": "success",
        "file": str(filepath),
        "find": find,
        "replace": replace,
        "replacements_made": count,
        "match_case": match_case
    }


def main():
    parser = argparse.ArgumentParser(
        description="Find and replace text across PowerPoint presentation",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Simple replacement
  uv python ppt_replace_text.py \\
    --file presentation.pptx \\
    --find "2023" \\
    --replace "2024" \\
    --json
  
  # Case-sensitive replacement
  uv python ppt_replace_text.py \\
    --file presentation.pptx \\
    --find "Company Inc." \\
    --replace "Company LLC" \\
    --match-case \\
    --json
  
  # Dry run to preview changes
  uv python ppt_replace_text.py \\
    --file presentation.pptx \\
    --find "old_term" \\
    --replace "new_term" \\
    --dry-run \\
    --json
  
  # Update product name
  uv python ppt_replace_text.py \\
    --file product_deck.pptx \\
    --find "Product X" \\
    --replace "Product Y" \\
    --json
  
  # Fix typo across all slides
  uv python ppt_replace_text.py \\
    --file presentation.pptx \\
    --find "recieve" \\
    --replace "receive" \\
    --json

Common Use Cases:
  - Update dates (2023 → 2024)
  - Change company names (rebranding)
  - Fix recurring typos
  - Update product names
  - Change terminology
  - Update prices/numbers
  - Localization (English → Spanish)
  - Template customization

Best Practices:
  1. Always use --dry-run first to preview changes
  2. Create backup before bulk replacements
  3. Use --match-case for proper nouns
  4. Test on a copy first
  5. Review results after replacement
  6. Be specific with find text to avoid unwanted matches

Safety Tips:
  - Backup file before major changes
  - Use dry-run to verify matches
  - Check match count makes sense
  - Review a few slides manually after
  - Use case-sensitive for precision
  - Avoid replacing common words

Limitations:
  - Only replaces visible text (not in images)
  - Does not replace text in charts/tables (text only)
  - Preserves original formatting
  - Cannot use regex patterns (exact match only)
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file path'
    )
    
    parser.add_argument(
        '--find',
        required=True,
        help='Text to find'
    )
    
    parser.add_argument(
        '--replace',
        required=True,
        help='Replacement text'
    )
    
    parser.add_argument(
        '--match-case',
        action='store_true',
        help='Case-sensitive matching'
    )
    
    parser.add_argument(
        '--whole-words',
        action='store_true',
        help='Match whole words only (not yet implemented)'
    )
    
    parser.add_argument(
        '--dry-run',
        action='store_true',
        help='Preview changes without modifying file'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        help='Output JSON response'
    )
    
    args = parser.parse_args()
    
    try:
        result = replace_text(
            filepath=args.file,
            find=args.find,
            replace=args.replace,
            match_case=args.match_case,
            whole_words=args.whole_words,
            dry_run=args.dry_run
        )
        
        if args.json:
            print(json.dumps(result, indent=2))
        else:
            if args.dry_run:
                print(f"🔍 Dry run - no changes made")
                print(f"   Found: {result['matches_found']} occurrences")
                print(f"   In: {result['total_locations']} locations")
                if result['locations']:
                    print(f"   Sample locations:")
                    for loc in result['locations'][:3]:
                        print(f"     - Slide {loc['slide']}: {loc['occurrences']} matches")
            else:
                print(f"✅ Replaced '{args.find}' with '{args.replace}'")
                print(f"   Replacements: {result['replacements_made']}")
                if args.match_case:
                    print(f"   Case-sensitive: Yes")
        
        sys.exit(0)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__
        }
        
        if args.json:
            print(json.dumps(error_result, indent=2))
        else:
            print(f"❌ Error: {e}", file=sys.stderr)
        
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_set_background.py
```py
#!/usr/bin/env python3
"""
PowerPoint Set Background Tool
Set slide background to color or image

Usage:
    uv python ppt_set_background.py --file deck.pptx --color "#FFFFFF" --json
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError, ColorHelper
)

def set_background(
    filepath: Path,
    color: str = None,
    image: Path = None,
    slide_index: int = None
) -> Dict[str, Any]:
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
        
    if not color and not image:
        raise ValueError("Must specify either --color or --image")

    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Determine target slides
        if slide_index is not None:
             target_slides = [agent.prs.slides[slide_index]]
        else:
             target_slides = agent.prs.slides
             
        for slide in target_slides:
            bg = slide.background
            fill = bg.fill
            
            if color:
                fill.solid()
                fill.fore_color.rgb = ColorHelper.from_hex(color)
            elif image:
                if not image.exists():
                    raise FileNotFoundError(f"Image not found: {image}")
                # Note: python-pptx background image support is limited in some versions
                # but user_picture is the standard method
                fill.user_picture(str(image))
                
        agent.save()
        
    return {
        "status": "success",
        "file": str(filepath),
        "slides_affected": len(target_slides),
        "type": "color" if color else "image"
    }

def main():
    parser = argparse.ArgumentParser(description="Set PowerPoint background")
    parser.add_argument('--file', required=True, type=Path, help='PowerPoint file path')
    parser.add_argument('--slide', type=int, help='Slide index (optional, defaults to all)')
    parser.add_argument('--color', help='Hex color code')
    parser.add_argument('--image', type=Path, help='Background image path')
    parser.add_argument('--json', action='store_true', default=True, help='Output JSON')
    
    args = parser.parse_args()
    
    try:
        result = set_background(
            filepath=args.file, 
            slide_index=args.slide, 
            color=args.color,
            image=args.image
        )
        print(json.dumps(result, indent=2))
        sys.exit(0)
    except Exception as e:
        print(json.dumps({"status": "error", "error": str(e), "error_type": type(e).__name__}, indent=2))
        sys.exit(1)

if __name__ == "__main__":
    main()

```

# tools/ppt_set_footer.py
```py
#!/usr/bin/env python3
"""
PowerPoint Set Footer Tool
Configure slide footer, date, and slide number

Usage:
    uv python ppt_set_footer.py --file deck.pptx --text "Confidential" --show-number --json
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import PowerPointAgent

def set_footer(
    filepath: Path,
    text: str = None,
    show_number: bool = False,
    show_date: bool = False,
    apply_to_master: bool = True
) -> Dict[str, Any]:
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")

    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # In python-pptx, footer visibility is often controlled via the master
        # or individual slide layouts.
        
        if apply_to_master:
             masters = agent.prs.slide_masters
             for master in masters:
                 # Update layouts in master
                 for layout in master.slide_layouts:
                     # Iterate shapes to find placeholders
                     for shape in layout.placeholders:
                         if shape.is_placeholder:
                             # Footer type is 15, Slide Number is 16, Date is 14
                             if shape.placeholder_format.type == 15 and text: # Footer
                                 shape.text = text
                             
        # Also attempt to set on individual slides for immediate visibility
        count = 0
        for slide in agent.prs.slides:
            # This is simplified; robustness varies by template
            # We try to find standard placeholders
            for shape in slide.placeholders:
                 if shape.placeholder_format.type == 15 and text:
                     shape.text = text
                     count += 1
                     
        agent.save()
        
    return {
        "status": "success",
        "file": str(filepath),
        "footer_text": text,
        "settings": {
            "show_number": show_number,
            "show_date": show_date
        },
        "slides_updated": count
    }

def main():
    parser = argparse.ArgumentParser(description="Set PowerPoint footer")
    parser.add_argument('--file', required=True, type=Path, help='PowerPoint file path')
    parser.add_argument('--text', help='Footer text')
    parser.add_argument('--show-number', nargs='?', const='true', default='false', help='Show slide number (optional: true/false)')
    parser.add_argument('--show-date', nargs='?', const='true', default='false', help='Show date (optional: true/false)')
    parser.add_argument('--json', action='store_true', default=True, help='Output JSON')
    
    args = parser.parse_args()
    
    try:
        # Helper to parse boolean string/flag
        def parse_bool(val):
            if isinstance(val, bool): return val
            if val is None: return False
            return str(val).lower() in ('true', 'yes', '1', 'on')

        result = set_footer(
            filepath=args.file, 
            text=args.text, 
            show_number=parse_bool(args.show_number),
            show_date=parse_bool(args.show_date)
        )
        print(json.dumps(result, indent=2))
        sys.exit(0)
    except Exception as e:
        print(json.dumps({"status": "error", "error": str(e), "error_type": type(e).__name__}, indent=2))
        sys.exit(1)

if __name__ == "__main__":
    main()

```

# tools/ppt_set_image_properties.py
```py
#!/usr/bin/env python3
"""
PowerPoint Set Image Properties Tool
Set alt text and transparency for images

Usage:
    uv python ppt_set_image_properties.py --file deck.pptx --slide 0 --shape 1 --alt-text "Logo" --json
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError, SlideNotFoundError
)


def set_image_properties(
    filepath: Path,
    slide_index: int,
    shape_index: int,
    alt_text: str = None,
    transparency: float = None
) -> Dict[str, Any]:
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if alt_text is None and transparency is None:
        raise ValueError("At least alt-text or transparency must be set")

    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        total = agent.get_slide_count()
        if not 0 <= slide_index < total:
            raise SlideNotFoundError(f"Slide index {slide_index} out of range")
            
        agent.set_image_properties(
            slide_index=slide_index,
            shape_index=shape_index,
            alt_text=alt_text,
            transparency=transparency
        )
        agent.save()
        
    return {
        "status": "success",
        "file": str(filepath),
        "slide_index": slide_index,
        "shape_index": shape_index,
        "properties": {
            "alt_text": alt_text,
            "transparency": transparency
        }
    }


def main():
    parser = argparse.ArgumentParser(description="Set PowerPoint image properties")
    parser.add_argument('--file', required=True, type=Path, help='PowerPoint file path')
    parser.add_argument('--slide', required=True, type=int, help='Slide index (0-based)')
    parser.add_argument('--shape', required=True, type=int, help='Shape index (0-based)')
    parser.add_argument('--alt-text', help='Alt text for accessibility')
    parser.add_argument('--transparency', type=float, help='Transparency (0.0-1.0)')
    parser.add_argument('--json', action='store_true', default=True, help='Output JSON')
    
    args = parser.parse_args()
    
    try:
        result = set_image_properties(
            filepath=args.file, 
            slide_index=args.slide, 
            shape_index=args.shape,
            alt_text=args.alt_text,
            transparency=args.transparency
        )
        print(json.dumps(result, indent=2))
        sys.exit(0)
    except Exception as e:
        print(json.dumps({"status": "error", "error": str(e), "error_type": type(e).__name__}, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_set_slide_layout.py
```py
#!/usr/bin/env python3
"""
PowerPoint Set Slide Layout Tool
Change the layout of an existing slide

Usage:
    uv python ppt_set_slide_layout.py --file presentation.pptx --slide 0 --layout "Title Only" --json
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError, SlideNotFoundError, LayoutNotFoundError
)


def set_slide_layout(filepath: Path, slide_index: int, layout_name: str) -> Dict[str, Any]:
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        total = agent.get_slide_count()
        if not 0 <= slide_index < total:
            raise SlideNotFoundError(f"Slide index {slide_index} out of range")
            
        # Get available layouts to validate/fuzzy match
        available = agent.get_available_layouts()
        if layout_name not in available:
            # Try fuzzy match
            layout_lower = layout_name.lower()
            for avail in available:
                if layout_lower in avail.lower():
                    layout_name = avail
                    break
            else:
                 raise LayoutNotFoundError(f"Layout '{layout_name}' not found. Available: {available}")

        agent.set_slide_layout(slide_index, layout_name)
        agent.save()
        
    return {
        "status": "success",
        "file": str(filepath),
        "slide_index": slide_index,
        "new_layout": layout_name
    }


def main():
    parser = argparse.ArgumentParser(description="Set PowerPoint slide layout")
    parser.add_argument('--file', required=True, type=Path, help='PowerPoint file path')
    parser.add_argument('--slide', required=True, type=int, help='Slide index (0-based)')
    parser.add_argument('--layout', required=True, help='New layout name')
    parser.add_argument('--json', action='store_true', default=True, help='Output JSON')
    
    args = parser.parse_args()
    
    try:
        result = set_slide_layout(filepath=args.file, slide_index=args.slide, layout_name=args.layout)
        print(json.dumps(result, indent=2))
        sys.exit(0)
    except Exception as e:
        print(json.dumps({"status": "error", "error": str(e), "error_type": type(e).__name__}, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_set_title.py
```py
#!/usr/bin/env python3
"""
PowerPoint Set Title Tool
Set slide title and optional subtitle

Usage:
    uv python ppt_set_title.py --file presentation.pptx --slide 0 --title "Q4 Results" --subtitle "Financial Review" --json

Exit Codes:
    0: Success
    1: Error occurred
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError, SlideNotFoundError
)


def set_title(
    filepath: Path,
    slide_index: int,
    title: str,
    subtitle: str = None
) -> Dict[str, Any]:
    """Set slide title and subtitle."""
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Validate slide index
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides-1})"
            )
        
        # Set title
        agent.set_title(slide_index, title, subtitle)
        
        # Get slide info
        slide_info = agent.get_slide_info(slide_index)
        
        # Save
        agent.save()
    
    return {
        "status": "success",
        "file": str(filepath),
        "slide_index": slide_index,
        "title": title,
        "subtitle": subtitle,
        "layout": slide_info["layout"],
        "shape_count": slide_info["shape_count"]
    }


def main():
    parser = argparse.ArgumentParser(
        description="Set PowerPoint slide title and subtitle",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Set title only
  uv python ppt_set_title.py --file presentation.pptx --slide 0 --title "Q4 Financial Results" --json
  
  # Set title and subtitle
  uv python ppt_set_title.py --file deck.pptx --slide 0 --title "2024 Strategy" --subtitle "Driving Growth & Innovation" --json
  
  # Update existing title
  uv python ppt_set_title.py --file presentation.pptx --slide 5 --title "Updated Section Title" --json
  
  # Set title on last slide
  uv python ppt_set_title.py --file presentation.pptx --slide -1 --title "Thank You" --json

Best Practices:
  - Keep titles concise (max 60 characters)
  - Use title case: "This Is Title Case"
  - Subtitles provide context, not repetition
  - First slide (index 0) should use "Title Slide" layout
  - Section headers benefit from clear, bold titles

Slide Index:
  - 0 = first slide
  - 1 = second slide
  - -1 = last slide (not yet supported in this version)
  
  To find total slides: uv python ppt_get_info.py --file your.pptx --json
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file path'
    )
    
    parser.add_argument(
        '--slide',
        required=True,
        type=int,
        help='Slide index (0-based)'
    )
    
    parser.add_argument(
        '--title',
        required=True,
        help='Title text'
    )
    
    parser.add_argument(
        '--subtitle',
        help='Optional subtitle text'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        help='Output JSON response'
    )
    
    args = parser.parse_args()
    
    try:
        result = set_title(
            filepath=args.file,
            slide_index=args.slide,
            title=args.title,
            subtitle=args.subtitle
        )
        
        if args.json:
            print(json.dumps(result, indent=2))
        else:
            print(f"✅ Set title on slide {result['slide_index']}")
            print(f"   Title: {result['title']}")
            if args.subtitle:
                print(f"   Subtitle: {result['subtitle']}")
            print(f"   Layout: {result['layout']}")
        
        sys.exit(0)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__
        }
        
        if args.json:
            print(json.dumps(error_result, indent=2))
        else:
            print(f"❌ Error: {e}", file=sys.stderr)
        
        sys.exit(1)


if __name__ == "__main__":
    main()

```

# tools/ppt_update_chart_data.py
```py
#!/usr/bin/env python3
"""
PowerPoint Update Chart Data Tool
Update the data of an existing chart

Usage:
    uv python ppt_update_chart_data.py --file deck.pptx --slide 0 --chart 0 --data new_data.json --json
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError, SlideNotFoundError
)
from pptx.chart.data import CategoryChartData

def update_chart_data(
    filepath: Path,
    slide_index: int,
    chart_index: int,
    data: Dict[str, Any]
) -> Dict[str, Any]:
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
        
    if "categories" not in data or "series" not in data:
        raise ValueError("Data JSON must contain 'categories' and 'series'")

    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        total = agent.get_slide_count()
        if not 0 <= slide_index < total:
            raise SlideNotFoundError(f"Slide index {slide_index} out of range")
            
        slide = agent.prs.slides[slide_index]
        
        # Find the chart
        charts = [shape for shape in slide.shapes if shape.has_chart]
        if not 0 <= chart_index < len(charts):
             raise ValueError(f"Chart index {chart_index} out of range. Slide has {len(charts)} charts.")
             
        chart = charts[chart_index].chart
        
        # Create new chart data
        chart_data = CategoryChartData()
        chart_data.categories = data["categories"]
        
        for series in data["series"]:
            chart_data.add_series(series["name"], series["values"])
            
        # Replace data
        chart.replace_data(chart_data)
        
        agent.save()
        
    return {
        "status": "success",
        "file": str(filepath),
        "slide_index": slide_index,
        "chart_index": chart_index,
        "data_points": sum(len(s["values"]) for s in data["series"])
    }

def main():
    parser = argparse.ArgumentParser(description="Update PowerPoint chart data")
    parser.add_argument('--file', required=True, type=Path, help='PowerPoint file path')
    parser.add_argument('--slide', required=True, type=int, help='Slide index (0-based)')
    parser.add_argument('--chart', required=True, type=int, help='Chart index (0-based)')
    parser.add_argument('--data', required=True, type=Path, help='JSON data file')
    parser.add_argument('--json', action='store_true', default=True, help='Output JSON')
    
    args = parser.parse_args()
    
    try:
        with open(args.data, 'r') as f:
            data_content = json.load(f)
            
        result = update_chart_data(
            filepath=args.file, 
            slide_index=args.slide, 
            chart_index=args.chart,
            data=data_content
        )
        print(json.dumps(result, indent=2))
        sys.exit(0)
    except Exception as e:
        print(json.dumps({"status": "error", "error": str(e), "error_type": type(e).__name__}, indent=2))
        sys.exit(1)

if __name__ == "__main__":
    main()

```

# tools/ppt_validate_presentation.py
```py
#!/usr/bin/env python3
"""
PowerPoint Validate Presentation Tool
Comprehensive validation for missing assets, errors, quality issues

Usage:
    uv python ppt_validate_presentation.py --file presentation.pptx --json

Exit Codes:
    0: Success (no issues or only warnings)
    1: Error occurred or critical issues found
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError
)


def validate_presentation(filepath: Path) -> Dict[str, Any]:
    """Validate presentation for issues."""
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath, acquire_lock=False)
        
        # Run all validation checks
        validation = agent.validate_presentation()
        accessibility = agent.check_accessibility()
        assets = agent.validate_assets()
    
    # Combine results
    all_issues = {
        "validation": validation,
        "accessibility": accessibility,
        "assets": assets
    }
    
    # Calculate totals
    total_issues = (
        validation.get("total_issues", 0) +
        accessibility.get("total_issues", 0) +
        assets.get("total_issues", 0)
    )
    
    # Determine severity
    has_critical = (
        len(validation.get("issues", {}).get("empty_slides", [])) > 0 or
        len(accessibility.get("issues", {}).get("missing_alt_text", [])) > 5
    )
    
    return {
        "status": "critical" if has_critical else ("warnings" if total_issues > 0 else "valid"),
        "file": str(filepath),
        "total_issues": total_issues,
        "summary": {
            "empty_slides": len(validation.get("issues", {}).get("empty_slides", [])),
            "slides_without_titles": len(validation.get("issues", {}).get("slides_without_titles", [])),
            "missing_alt_text": len(accessibility.get("issues", {}).get("missing_alt_text", [])),
            "low_contrast": len(accessibility.get("issues", {}).get("low_contrast", [])),
            "low_resolution_images": len(assets.get("issues", {}).get("low_resolution_images", [])),
            "large_images": len(assets.get("issues", {}).get("large_images", []))
        },
        "details": all_issues,
        "recommendations": generate_recommendations(all_issues)
    }


def generate_recommendations(issues: Dict[str, Any]) -> list:
    """Generate actionable recommendations based on issues."""
    recommendations = []
    
    validation = issues.get("validation", {}).get("issues", {})
    accessibility = issues.get("accessibility", {}).get("issues", {})
    assets = issues.get("assets", {}).get("issues", {})
    
    # Empty slides
    if validation.get("empty_slides"):
        recommendations.append({
            "priority": "high",
            "issue": "Empty slides found",
            "action": "Remove empty slides or add content",
            "affected_slides": validation["empty_slides"]
        })
    
    # Missing titles
    if validation.get("slides_without_titles"):
        recommendations.append({
            "priority": "medium",
            "issue": "Slides without titles",
            "action": "Add titles using ppt_set_title.py",
            "affected_slides": validation["slides_without_titles"][:5]
        })
    
    # Missing alt text
    if len(accessibility.get("missing_alt_text", [])) > 0:
        recommendations.append({
            "priority": "high",
            "issue": "Images without alt text",
            "action": "Add alt text using ppt_set_image_properties.py",
            "count": len(accessibility["missing_alt_text"])
        })
    
    # Low contrast
    if len(accessibility.get("low_contrast", [])) > 0:
        recommendations.append({
            "priority": "medium",
            "issue": "Low color contrast text",
            "action": "Improve text/background contrast for readability",
            "count": len(accessibility["low_contrast"])
        })
    
    # Low resolution images
    if assets.get("low_resolution_images"):
        recommendations.append({
            "priority": "medium",
            "issue": "Low resolution images",
            "action": "Replace with higher resolution images (96 DPI minimum)",
            "count": len(assets["low_resolution_images"])
        })
    
    # Large images
    if assets.get("large_images"):
        recommendations.append({
            "priority": "low",
            "issue": "Large images (>2MB)",
            "action": "Compress images using ppt_replace_image.py --compress",
            "count": len(assets["large_images"])
        })
    
    return recommendations


def main():
    parser = argparse.ArgumentParser(
        description="Validate PowerPoint presentation for quality issues",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Validate presentation
  uv python ppt_validate_presentation.py \\
    --file presentation.pptx \\
    --json

Validation Checks:
  1. Structure Validation:
     - Empty slides
     - Slides without titles
     - Text overflow
     - Inconsistent formatting
  
  2. Accessibility (WCAG 2.1):
     - Missing alt text on images
     - Low color contrast
     - Missing slide titles
     - Reading order issues
  
  3. Asset Quality:
     - Low resolution images (<96 DPI)
     - Large images (>2MB)
     - Total embedded file size
     - Missing external links

Output Status:
  - "valid": No issues found
  - "warnings": Minor issues that should be fixed
  - "critical": Major issues that must be fixed

Example Output:
{
  "status": "warnings",
  "total_issues": 5,
  "summary": {
    "empty_slides": 0,
    "slides_without_titles": 2,
    "missing_alt_text": 3,
    "low_contrast": 0
  },
  "recommendations": [
    {
      "priority": "high",
      "issue": "Images without alt text",
      "action": "Add alt text using ppt_set_image_properties.py",
      "count": 3
    }
  ]
}

Use Cases:
  - Pre-publication checks
  - Accessibility compliance
  - Quality assurance
  - Best practices enforcement
  - File size optimization

Fixing Issues:
  Use these tools to fix found issues:
  - Empty slides: ppt_delete_slide.py
  - Missing titles: ppt_set_title.py
  - Missing alt text: ppt_set_image_properties.py
  - Large images: ppt_replace_image.py --compress
  - Low contrast: ppt_format_text.py --color

Best Practices:
  - Run validation before sharing
  - Fix critical issues first
  - Address accessibility for compliance
  - Optimize large presentations
  - Regular validation during development
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file to validate'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default)'
    )
    
    args = parser.parse_args()
    
    try:
        result = validate_presentation(filepath=args.file)
        
        print(json.dumps(result, indent=2))
        
        # Exit with error code if critical issues
        if result["status"] == "critical":
            sys.exit(1)
        else:
            sys.exit(0)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__
        }
        
        print(json.dumps(error_result, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()

```

