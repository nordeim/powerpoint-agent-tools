#!/usr/bin/env python3
"""
PowerPoint Add Shape Tool v3.1.0
Add shapes (rectangle, circle, arrow, etc.) to slides with comprehensive styling options.

Fully aligned with PowerPoint Agent Core v3.0 and System Prompt v3.0.
Now includes opacity/transparency support for overlay workflows.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_add_shape.py --file presentation.pptx --slide 0 \\
        --shape rectangle --position '{"left":"20%","top":"30%"}' \\
        --size '{"width":"60%","height":"40%"}' --fill-color "#0070C0" --json

    # Overlay with transparency
    uv run tools/ppt_add_shape.py --file presentation.pptx --slide 0 \\
        --shape rectangle --position '{"left":"0%","top":"0%"}' \\
        --size '{"width":"100%","height":"100%"}' \\
        --fill-color "#000000" --fill-opacity 0.15 --json

    # Quick overlay preset
    uv run tools/ppt_add_shape.py --file presentation.pptx --slide 0 \\
        --shape rectangle --overlay --fill-color "#FFFFFF" --json

Exit Codes:
    0: Success
    1: Error occurred

Changelog v3.1.0:
- NEW: --fill-opacity parameter (0.0-1.0) for transparent fills
- NEW: --line-opacity parameter (0.0-1.0) for transparent borders
- NEW: --overlay preset flag for quick overlay creation
- NEW: Opacity validation with helpful warnings
- NEW: Contrast validation accounts for transparency
- IMPROVED: Overlay workflow documentation and examples
- IMPROVED: Error messages for opacity-related issues

Changelog v3.0.0:
- Aligned with PowerPoint Agent Core v3.0
- Added text inside shapes support
- Enhanced validation with contrast warnings
- Captures and reports shape_index from core
- Added all shape types from v3.0 core
- Improved error handling and messages
- Added z-order awareness documentation
- Consistent JSON output structure
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any, List, Optional, Tuple

# Add parent directory to path for core import
sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent,
    PowerPointAgentError,
    SlideNotFoundError,
    ShapeNotFoundError,
    ColorHelper,
    Position,
    Size,
    SLIDE_WIDTH_INCHES,
    SLIDE_HEIGHT_INCHES,
    CORPORATE_COLORS,
    __version__ as CORE_VERSION
)

# ============================================================================
# CONSTANTS
# ============================================================================

__version__ = "3.1.0"

# Available shape types (aligned with v3.0 core)
AVAILABLE_SHAPES = [
    "rectangle",
    "rounded_rectangle",
    "ellipse",
    "oval",
    "triangle",
    "arrow_right",
    "arrow_left",
    "arrow_up",
    "arrow_down",
    "diamond",
    "pentagon",
    "hexagon",
    "star",
    "heart",
    "lightning",
    "sun",
    "moon",
    "cloud",
]

# Shape type aliases for user convenience
SHAPE_ALIASES = {
    "rect": "rectangle",
    "round_rect": "rounded_rectangle",
    "circle": "ellipse",
    "arrow": "arrow_right",
    "right_arrow": "arrow_right",
    "left_arrow": "arrow_left",
    "up_arrow": "arrow_up",
    "down_arrow": "arrow_down",
    "5_point_star": "star",
}

# Default corporate colors for quick reference
COLOR_PRESETS = {
    "primary": "#0070C0",
    "secondary": "#595959",
    "accent": "#ED7D31",
    "success": "#70AD47",
    "warning": "#FFC000",
    "danger": "#C00000",
    "white": "#FFFFFF",
    "black": "#000000",
    "light_gray": "#D9D9D9",
    "dark_gray": "#404040",
}

# Overlay preset defaults (aligned with System Prompt v3.0)
OVERLAY_DEFAULTS = {
    "position": {"left": "0%", "top": "0%"},
    "size": {"width": "100%", "height": "100%"},
    "fill_opacity": 0.15,  # 15% opaque = subtle, non-competing
    "z_order_action": "send_to_back",
}


# Overlay preset defaults (aligned with System Prompt v3.0)
OVERLAY_DEFAULTS = {
    "position": {"left": "0%", "top": "0%"},
    "size": {"width": "100%", "height": "100%"},
    "fill_opacity": 0.15,  # 15% opaque = subtle, non-competing
    "z_order_action": "send_to_back",
}


# ============================================================================
# VALIDATION FUNCTIONS
# ============================================================================

def resolve_shape_type(shape_type: str) -> str:
    """
    Resolve shape type, handling aliases.
    
    Args:
        shape_type: User-provided shape type
        
    Returns:
        Canonical shape type name
    """
    shape_lower = shape_type.lower().strip()
    
    # Check aliases first
    if shape_lower in SHAPE_ALIASES:
        return SHAPE_ALIASES[shape_lower]
    
    # Check direct match
    if shape_lower in AVAILABLE_SHAPES:
        return shape_lower
    
    # Try to find partial match
    for available in AVAILABLE_SHAPES:
        if shape_lower in available or available in shape_lower:
            return available
    
    return shape_lower  # Let core handle unknown types


def resolve_color(color: Optional[str]) -> Optional[str]:
    """
    Resolve color, handling presets and validation.
    
    Args:
        color: User-provided color (hex or preset name)
        
    Returns:
        Resolved hex color or None
    """
    if color is None:
        return None
    
    color_lower = color.lower().strip()
    
    # Check presets
    if color_lower in COLOR_PRESETS:
        return COLOR_PRESETS[color_lower]
    
    # Ensure # prefix for hex colors
    if not color.startswith('#') and len(color) == 6:
        try:
            int(color, 16)
            return f"#{color}"
        except ValueError:
            pass
    
    return color


def validate_opacity(
    fill_opacity: float,
    line_opacity: float
) -> Tuple[List[str], List[str]]:
    """
    Validate opacity values and return warnings/recommendations.
    
    Args:
        fill_opacity: Fill opacity (0.0-1.0)
        line_opacity: Line opacity (0.0-1.0)
        
    Returns:
        Tuple of (warnings, recommendations)
        
    Raises:
        ValueError: If opacity values are out of range
    """
    warnings: List[str] = []
    recommendations: List[str] = []
    
    # Validate ranges
    if not 0.0 <= fill_opacity <= 1.0:
        raise ValueError(
            f"fill_opacity must be between 0.0 and 1.0, got {fill_opacity}"
        )
    
    if not 0.0 <= line_opacity <= 1.0:
        raise ValueError(
            f"line_opacity must be between 0.0 and 1.0, got {line_opacity}"
        )
    
    # Warn about extreme values
    if fill_opacity == 0.0:
        warnings.append(
            "Fill opacity is 0.0 (fully transparent). Shape fill will be invisible. "
            "Use no --fill-color if you want no fill."
        )
    elif fill_opacity < 0.05:
        warnings.append(
            f"Fill opacity {fill_opacity} is extremely low (<5%). "
            "Shape may be nearly invisible."
        )
    
    if line_opacity == 0.0 and fill_opacity == 0.0:
        warnings.append(
            "Both fill and line opacity are 0.0. Shape will be completely invisible."
        )
    
    # Overlay-specific recommendations
    if 0.1 <= fill_opacity <= 0.3:
        recommendations.append(
            f"Opacity {fill_opacity} is appropriate for overlay backgrounds. "
            "Remember to use ppt_set_z_order.py --action send_to_back after adding."
        )
    elif 0.3 < fill_opacity < 0.7:
        recommendations.append(
            f"Opacity {fill_opacity} creates a semi-transparent effect. "
            "Content behind may still be partially visible."
        )
    
    return warnings, recommendations


def validate_shape_params(
    position: Dict[str, Any],
    size: Dict[str, Any],
    fill_color: Optional[str] = None,
    fill_opacity: float = 1.0,
    line_color: Optional[str] = None,
    line_opacity: float = 1.0,
    text: Optional[str] = None,
    allow_offslide: bool = False,
    is_overlay: bool = False
) -> Dict[str, Any]:
    """
    Validate shape parameters and return warnings/recommendations.
    
    Args:
        position: Position specification
        size: Size specification
        fill_color: Fill color hex
        fill_opacity: Fill opacity (0.0-1.0)
        line_color: Line color hex
        line_opacity: Line opacity (0.0-1.0)
        text: Text to add inside shape
        allow_offslide: Allow off-slide positioning
        is_overlay: Whether this is an overlay shape
        
    Returns:
        Dict with warnings, recommendations, and validation results
    """
    warnings: List[str] = []
    recommendations: List[str] = []
    validation_results: Dict[str, Any] = {}
    
    # Opacity validation
    opacity_warnings, opacity_recommendations = validate_opacity(
        fill_opacity, line_opacity
    )
    warnings.extend(opacity_warnings)
    recommendations.extend(opacity_recommendations)
    
    # Record opacity in validation results
    validation_results["fill_opacity"] = fill_opacity
    validation_results["line_opacity"] = line_opacity
    validation_results["effective_fill_transparency"] = round(1.0 - fill_opacity, 2)
    
    # Position validation
    if position:
        _validate_position(position, warnings, allow_offslide)
    
    # Size validation
    if size:
        _validate_size(size, warnings)
    
    # Color contrast validation (accounting for opacity)
    if fill_color:
        _validate_color_contrast_with_opacity(
            fill_color, fill_opacity, line_color, text,
            warnings, recommendations, validation_results
        )
    
    # Text validation
    if text:
        _validate_text(text, warnings, recommendations)
        
        # Warn about text on transparent shapes
        if fill_opacity < 0.5:
            warnings.append(
                f"Shape has text but fill opacity is only {fill_opacity}. "
                "Text may be hard to read against varied backgrounds."
            )
    
    # Overlay-specific validation
    if is_overlay:
        _validate_overlay(
            position, size, fill_opacity, 
            warnings, recommendations, validation_results
        )
    
    return {
        "warnings": warnings,
        "recommendations": recommendations,
        "validation_results": validation_results,
        "has_warnings": len(warnings) > 0
    }


def _validate_position(
    position: Dict[str, Any],
    warnings: List[str],
    allow_offslide: bool
) -> None:
    """Validate position values."""
    try:
        for key in ["left", "top"]:
            if key in position:
                value_str = str(position[key])
                if value_str.endswith('%'):
                    pct = float(value_str.rstrip('%'))
                    if not allow_offslide and (pct < 0 or pct > 100):
                        warnings.append(
                            f"Position '{key}' is {pct}% which is outside slide bounds (0-100%). "
                            f"Shape may not be visible. Use --allow-offslide if intentional."
                        )
    except (ValueError, TypeError):
        pass


def _validate_size(size: Dict[str, Any], warnings: List[str]) -> None:
    """Validate size values."""
    try:
        for key in ["width", "height"]:
            if key in size:
                value_str = str(size[key])
                if value_str.endswith('%'):
                    pct = float(value_str.rstrip('%'))
                    if pct <= 0:
                        warnings.append(
                            f"Size '{key}' is {pct}% which is invalid (must be > 0%)."
                        )
                    elif pct < 1:
                        warnings.append(
                            f"Size '{key}' is {pct}% which is extremely small (<1%). "
                            f"Shape may be invisible."
                        )
                    elif pct > 100:
                        warnings.append(
                            f"Size '{key}' is {pct}% which exceeds slide dimensions."
                        )
    except (ValueError, TypeError):
        pass


def _validate_color_contrast_with_opacity(
    fill_color: str,
    fill_opacity: float,
    line_color: Optional[str],
    text: Optional[str],
    warnings: List[str],
    recommendations: List[str],
    validation_results: Dict[str, Any]
) -> None:
    """
    Validate color contrast for visibility and accessibility,
    accounting for transparency effects.
    """
    try:
        from pptx.dml.color import RGBColor
        
        # Parse fill color
        shape_rgb = ColorHelper.from_hex(fill_color)
        
        # Record base color info
        validation_results["fill_color_hex"] = fill_color
        validation_results["fill_color_rgb"] = {
            "r": shape_rgb.red,
            "g": shape_rgb.green,
            "b": shape_rgb.blue
        }
        
        # Calculate effective color when blended with white background
        # (common slide background)
        if fill_opacity < 1.0:
            # Alpha blending: result = alpha * foreground + (1-alpha) * background
            effective_r = int(fill_opacity * shape_rgb.red + (1 - fill_opacity) * 255)
            effective_g = int(fill_opacity * shape_rgb.green + (1 - fill_opacity) * 255)
            effective_b = int(fill_opacity * shape_rgb.blue + (1 - fill_opacity) * 255)
            effective_rgb = RGBColor(effective_r, effective_g, effective_b)
            
            validation_results["effective_color_on_white"] = {
                "r": effective_r,
                "g": effective_g,
                "b": effective_b,
                "hex": f"#{effective_r:02X}{effective_g:02X}{effective_b:02X}"
            }
        else:
            effective_rgb = shape_rgb
        
        # Check contrast against white background
        white_bg = RGBColor(255, 255, 255)
        bg_contrast = ColorHelper.contrast_ratio(effective_rgb, white_bg)
        validation_results["effective_contrast_vs_white"] = round(bg_contrast, 2)
        
        # For semi-transparent shapes, contrast is inherently lower
        if fill_opacity < 1.0 and bg_contrast < 1.5:
            # This is expected for subtle overlays
            if fill_opacity <= 0.3:
                recommendations.append(
                    f"Overlay has low contrast ({bg_contrast:.2f}:1) against white, "
                    f"which is expected for opacity {fill_opacity}. "
                    "This creates a subtle tinting effect."
                )
            else:
                warnings.append(
                    f"Shape with opacity {fill_opacity} has low effective contrast "
                    f"({bg_contrast:.2f}:1) against white backgrounds."
                )
        
        # Check contrast against black background
        black_bg = RGBColor(0, 0, 0)
        
        # Calculate effective color on black background
        if fill_opacity < 1.0:
            effective_r_black = int(fill_opacity * shape_rgb.red)
            effective_g_black = int(fill_opacity * shape_rgb.green)
            effective_b_black = int(fill_opacity * shape_rgb.blue)
            effective_rgb_black = RGBColor(
                effective_r_black, effective_g_black, effective_b_black
            )
            validation_results["effective_color_on_black"] = {
                "r": effective_r_black,
                "g": effective_g_black,
                "b": effective_b_black,
                "hex": f"#{effective_r_black:02X}{effective_g_black:02X}{effective_b_black:02X}"
            }
        else:
            effective_rgb_black = shape_rgb
        
        dark_contrast = ColorHelper.contrast_ratio(effective_rgb_black, black_bg)
        validation_results["effective_contrast_vs_black"] = round(dark_contrast, 2)
        
        # If shape has text, check text readability
        if text and fill_opacity >= 0.5:
            # Only check text contrast if shape is reasonably opaque
            text_rgb_white = RGBColor(255, 255, 255)
            text_contrast = ColorHelper.contrast_ratio(text_rgb_white, effective_rgb)
            validation_results["text_contrast_white"] = round(text_contrast, 2)
            
            text_rgb_black = RGBColor(0, 0, 0)
            text_contrast_black = ColorHelper.contrast_ratio(text_rgb_black, effective_rgb)
            validation_results["text_contrast_black"] = round(text_contrast_black, 2)
            
            if text_contrast < 4.5 and text_contrast_black >= 4.5:
                recommendations.append(
                    f"Consider using dark text on this fill for better readability "
                    f"(black text contrast: {text_contrast_black:.2f}:1)."
                )
            elif text_contrast_black < 4.5 and text_contrast >= 4.5:
                recommendations.append(
                    f"White text provides good contrast on this fill "
                    f"({text_contrast:.2f}:1 meets WCAG AA)."
                )
            elif text_contrast < 4.5 and text_contrast_black < 4.5:
                warnings.append(
                    f"Neither white nor black text has sufficient contrast on this "
                    f"semi-transparent fill. Text may be hard to read."
                )
        
        # Line color validation
        if line_color:
            line_rgb = ColorHelper.from_hex(line_color)
            line_fill_contrast = ColorHelper.contrast_ratio(line_rgb, effective_rgb)
            validation_results["line_fill_contrast"] = round(line_fill_contrast, 2)
            
            if line_fill_contrast < 1.5:
                warnings.append(
                    f"Line color {line_color} has low contrast ({line_fill_contrast:.2f}:1) "
                    f"against the effective fill color. Border may not be visible."
                )
    
    except Exception as e:
        validation_results["color_validation_error"] = str(e)


def _validate_text(
    text: str,
    warnings: List[str],
    recommendations: List[str]
) -> None:
    """Validate text content."""
    if len(text) > 500:
        warnings.append(
            f"Text content is very long ({len(text)} characters). "
            f"Consider using a text box instead."
        )
    
    if '\n' in text and text.count('\n') > 5:
        recommendations.append(
            "Text has multiple lines. Consider using bullet points or a text box "
            "for better formatting control."
        )


def _validate_overlay(
    position: Dict[str, Any],
    size: Dict[str, Any],
    fill_opacity: float,
    warnings: List[str],
    recommendations: List[str],
    validation_results: Dict[str, Any]
) -> None:
    """Validate overlay-specific requirements."""
    validation_results["is_overlay"] = True
    
    # Check if full-slide
    is_full_slide = False
    try:
        left = str(position.get("left", ""))
        top = str(position.get("top", ""))
        width = str(size.get("width", ""))
        height = str(size.get("height", ""))
        
        if (left in ["0%", "0"] and top in ["0%", "0"] and 
            width in ["100%", "100"] and height in ["100%", "100"]):
            is_full_slide = True
            validation_results["is_full_slide_overlay"] = True
    except (ValueError, TypeError):
        pass
    
    # Warn if not full-slide overlay
    if not is_full_slide:
        recommendations.append(
            "Overlay mode is set but shape is not full-slide. "
            "Consider using position={'left':'0%','top':'0%'} and "
            "size={'width':'100%','height':'100%'} for full coverage."
        )
    
    # Opacity check for overlay
    if fill_opacity > 0.3:
        warnings.append(
            f"Overlay opacity {fill_opacity} is relatively high (>30%). "
            "Content behind may be significantly obscured. "
            "System prompt recommends 0.15 for subtle overlays."
        )
    elif fill_opacity < 0.05:
        warnings.append(
            f"Overlay opacity {fill_opacity} is very low (<5%). "
            "The overlay effect may be imperceptible."
        )
    
    # Always remind about z-order
    recommendations.append(
        "IMPORTANT: After adding this overlay, run:\n"
        "  ppt_set_z_order.py --file FILE --slide N --shape INDEX --action send_to_back\n"
        "This ensures the overlay appears behind content, not on top of it."
    )


def validate_opacity(
    fill_opacity: float,
    line_opacity: float
) -> Tuple[List[str], List[str]]:
    """
    Validate opacity values and return warnings/recommendations.
    
    Args:
        fill_opacity: Fill opacity (0.0-1.0)
        line_opacity: Line opacity (0.0-1.0)
        
    Returns:
        Tuple of (warnings, recommendations)
        
    Raises:
        ValueError: If opacity values are out of range
    """
    warnings: List[str] = []
    recommendations: List[str] = []
    
    # Validate ranges
    if not 0.0 <= fill_opacity <= 1.0:
        raise ValueError(
            f"fill_opacity must be between 0.0 and 1.0, got {fill_opacity}"
        )
    
    if not 0.0 <= line_opacity <= 1.0:
        raise ValueError(
            f"line_opacity must be between 0.0 and 1.0, got {line_opacity}"
        )
    
    # Warn about extreme values
    if fill_opacity == 0.0:
        warnings.append(
            "Fill opacity is 0.0 (fully transparent). Shape fill will be invisible. "
            "Use no --fill-color if you want no fill."
        )
    elif fill_opacity < 0.05:
        warnings.append(
            f"Fill opacity {fill_opacity} is extremely low (<5%). "
            "Shape may be nearly invisible."
        )
    
    if line_opacity == 0.0 and fill_opacity == 0.0:
        warnings.append(
            "Both fill and line opacity are 0.0. Shape will be completely invisible."
        )
    
    # Overlay-specific recommendations
    if 0.1 <= fill_opacity <= 0.3:
        recommendations.append(
            f"Opacity {fill_opacity} is appropriate for overlay backgrounds. "
            "Remember to use ppt_set_z_order.py --action send_to_back after adding."
        )
    elif 0.3 < fill_opacity < 0.7:
        recommendations.append(
            f"Opacity {fill_opacity} creates a semi-transparent effect. "
            "Content behind may still be partially visible."
        )
    
    return warnings, recommendations


def validate_shape_params(
    position: Dict[str, Any],
    size: Dict[str, Any],
    fill_color: Optional[str] = None,
    fill_opacity: float = 1.0,
    line_color: Optional[str] = None,
    line_opacity: float = 1.0,
    text: Optional[str] = None,
    allow_offslide: bool = False,
    is_overlay: bool = False
) -> Dict[str, Any]:
    """
    Validate shape parameters and return warnings/recommendations.
    
    Args:
        position: Position specification
        size: Size specification
        fill_color: Fill color hex
        fill_opacity: Fill opacity (0.0-1.0)
        line_color: Line color hex
        line_opacity: Line opacity (0.0-1.0)
        text: Text to add inside shape
        allow_offslide: Allow off-slide positioning
        is_overlay: Whether this is an overlay shape
        
    Returns:
        Dict with warnings, recommendations, and validation results
    """
    warnings: List[str] = []
    recommendations: List[str] = []
    validation_results: Dict[str, Any] = {}
    
    # Opacity validation
    opacity_warnings, opacity_recommendations = validate_opacity(
        fill_opacity, line_opacity
    )
    warnings.extend(opacity_warnings)
    recommendations.extend(opacity_recommendations)
    
    # Record opacity in validation results
    validation_results["fill_opacity"] = fill_opacity
    validation_results["line_opacity"] = line_opacity
    validation_results["effective_fill_transparency"] = round(1.0 - fill_opacity, 2)
    
    # Position validation
    if position:
        _validate_position(position, warnings, allow_offslide)
    
    # Size validation
    if size:
        _validate_size(size, warnings)
    
    # Color contrast validation (accounting for opacity)
    if fill_color:
        _validate_color_contrast_with_opacity(
            fill_color, fill_opacity, line_color, text,
            warnings, recommendations, validation_results
        )
    
    # Text validation
    if text:
        _validate_text(text, warnings, recommendations)
        
        # Warn about text on transparent shapes
        if fill_opacity < 0.5:
            warnings.append(
                f"Shape has text but fill opacity is only {fill_opacity}. "
                "Text may be hard to read against varied backgrounds."
            )
    
    # Overlay-specific validation
    if is_overlay:
        _validate_overlay(
            position, size, fill_opacity, 
            warnings, recommendations, validation_results
        )
    
    return {
        "warnings": warnings,
        "recommendations": recommendations,
        "validation_results": validation_results,
        "has_warnings": len(warnings) > 0
    }


def _validate_position(
    position: Dict[str, Any],
    warnings: List[str],
    allow_offslide: bool
) -> None:
    """Validate position values."""
    try:
        for key in ["left", "top"]:
            if key in position:
                value_str = str(position[key])
                if value_str.endswith('%'):
                    pct = float(value_str.rstrip('%'))
                    if not allow_offslide and (pct < 0 or pct > 100):
                        warnings.append(
                            f"Position '{key}' is {pct}% which is outside slide bounds (0-100%). "
                            f"Shape may not be visible. Use --allow-offslide if intentional."
                        )
    except (ValueError, TypeError):
        pass


def _validate_size(size: Dict[str, Any], warnings: List[str]) -> None:
    """Validate size values."""
    try:
        for key in ["width", "height"]:
            if key in size:
                value_str = str(size[key])
                if value_str.endswith('%'):
                    pct = float(value_str.rstrip('%'))
                    if pct <= 0:
                        warnings.append(
                            f"Size '{key}' is {pct}% which is invalid (must be > 0%)."
                        )
                    elif pct < 1:
                        warnings.append(
                            f"Size '{key}' is {pct}% which is extremely small (<1%). "
                            f"Shape may be invisible."
                        )
                    elif pct > 100:
                        warnings.append(
                            f"Size '{key}' is {pct}% which exceeds slide dimensions."
                        )
    except (ValueError, TypeError):
        pass


def _validate_color_contrast_with_opacity(
    fill_color: str,
    fill_opacity: float,
    line_color: Optional[str],
    text: Optional[str],
    warnings: List[str],
    recommendations: List[str],
    validation_results: Dict[str, Any]
) -> None:
    """
    Validate color contrast for visibility and accessibility,
    accounting for transparency effects.
    """
    try:
        from pptx.dml.color import RGBColor
        
        # Parse fill color
        shape_rgb = ColorHelper.from_hex(fill_color)
        
        # Record base color info
        validation_results["fill_color_hex"] = fill_color
        validation_results["fill_color_rgb"] = {
            "r": shape_rgb.red,
            "g": shape_rgb.green,
            "b": shape_rgb.blue
        }
        
        # Calculate effective color when blended with white background
        # (common slide background)
        if fill_opacity < 1.0:
            # Alpha blending: result = alpha * foreground + (1-alpha) * background
            effective_r = int(fill_opacity * shape_rgb.red + (1 - fill_opacity) * 255)
            effective_g = int(fill_opacity * shape_rgb.green + (1 - fill_opacity) * 255)
            effective_b = int(fill_opacity * shape_rgb.blue + (1 - fill_opacity) * 255)
            effective_rgb = RGBColor(effective_r, effective_g, effective_b)
            
            validation_results["effective_color_on_white"] = {
                "r": effective_r,
                "g": effective_g,
                "b": effective_b,
                "hex": f"#{effective_r:02X}{effective_g:02X}{effective_b:02X}"
            }
        else:
            effective_rgb = shape_rgb
        
        # Check contrast against white background
        white_bg = RGBColor(255, 255, 255)
        bg_contrast = ColorHelper.contrast_ratio(effective_rgb, white_bg)
        validation_results["effective_contrast_vs_white"] = round(bg_contrast, 2)
        
        # For semi-transparent shapes, contrast is inherently lower
        if fill_opacity < 1.0 and bg_contrast < 1.5:
            # This is expected for subtle overlays
            if fill_opacity <= 0.3:
                recommendations.append(
                    f"Overlay has low contrast ({bg_contrast:.2f}:1) against white, "
                    f"which is expected for opacity {fill_opacity}. "
                    "This creates a subtle tinting effect."
                )
            else:
                warnings.append(
                    f"Shape with opacity {fill_opacity} has low effective contrast "
                    f"({bg_contrast:.2f}:1) against white backgrounds."
                )
        
        # Check contrast against black background
        black_bg = RGBColor(0, 0, 0)
        
        # Calculate effective color on black background
        if fill_opacity < 1.0:
            effective_r_black = int(fill_opacity * shape_rgb.red)
            effective_g_black = int(fill_opacity * shape_rgb.green)
            effective_b_black = int(fill_opacity * shape_rgb.blue)
            effective_rgb_black = RGBColor(
                effective_r_black, effective_g_black, effective_b_black
            )
            validation_results["effective_color_on_black"] = {
                "r": effective_r_black,
                "g": effective_g_black,
                "b": effective_b_black,
                "hex": f"#{effective_r_black:02X}{effective_g_black:02X}{effective_b_black:02X}"
            }
        else:
            effective_rgb_black = shape_rgb
        
        dark_contrast = ColorHelper.contrast_ratio(effective_rgb_black, black_bg)
        validation_results["effective_contrast_vs_black"] = round(dark_contrast, 2)
        
        # If shape has text, check text readability
        if text and fill_opacity >= 0.5:
            # Only check text contrast if shape is reasonably opaque
            text_rgb_white = RGBColor(255, 255, 255)
            text_contrast = ColorHelper.contrast_ratio(text_rgb_white, effective_rgb)
            validation_results["text_contrast_white"] = round(text_contrast, 2)
            
            text_rgb_black = RGBColor(0, 0, 0)
            text_contrast_black = ColorHelper.contrast_ratio(text_rgb_black, effective_rgb)
            validation_results["text_contrast_black"] = round(text_contrast_black, 2)
            
            if text_contrast < 4.5 and text_contrast_black >= 4.5:
                recommendations.append(
                    f"Consider using dark text on this fill for better readability "
                    f"(black text contrast: {text_contrast_black:.2f}:1)."
                )
            elif text_contrast_black < 4.5 and text_contrast >= 4.5:
                recommendations.append(
                    f"White text provides good contrast on this fill "
                    f"({text_contrast:.2f}:1 meets WCAG AA)."
                )
            elif text_contrast < 4.5 and text_contrast_black < 4.5:
                warnings.append(
                    f"Neither white nor black text has sufficient contrast on this "
                    f"semi-transparent fill. Text may be hard to read."
                )
        
        # Line color validation
        if line_color:
            line_rgb = ColorHelper.from_hex(line_color)
            line_fill_contrast = ColorHelper.contrast_ratio(line_rgb, effective_rgb)
            validation_results["line_fill_contrast"] = round(line_fill_contrast, 2)
            
            if line_fill_contrast < 1.5:
                warnings.append(
                    f"Line color {line_color} has low contrast ({line_fill_contrast:.2f}:1) "
                    f"against the effective fill color. Border may not be visible."
                )
    
    except Exception as e:
        validation_results["color_validation_error"] = str(e)


def _validate_text(
    text: str,
    warnings: List[str],
    recommendations: List[str]
) -> None:
    """Validate text content."""
    if len(text) > 500:
        warnings.append(
            f"Text content is very long ({len(text)} characters). "
            f"Consider using a text box instead."
        )
    
    if '\n' in text and text.count('\n') > 5:
        recommendations.append(
            "Text has multiple lines. Consider using bullet points or a text box "
            "for better formatting control."
        )


def _validate_overlay(
    position: Dict[str, Any],
    size: Dict[str, Any],
    fill_opacity: float,
    warnings: List[str],
    recommendations: List[str],
    validation_results: Dict[str, Any]
) -> None:
    """Validate overlay-specific requirements."""
    validation_results["is_overlay"] = True
    
    # Check if full-slide
    is_full_slide = False
    try:
        left = str(position.get("left", ""))
        top = str(position.get("top", ""))
        width = str(size.get("width", ""))
        height = str(size.get("height", ""))
        
        if (left in ["0%", "0"] and top in ["0%", "0"] and 
            width in ["100%", "100"] and height in ["100%", "100"]):
            is_full_slide = True
            validation_results["is_full_slide_overlay"] = True
    except (ValueError, TypeError):
        pass
    
    # Warn if not full-slide overlay
    if not is_full_slide:
        recommendations.append(
            "Overlay mode is set but shape is not full-slide. "
            "Consider using position={'left':'0%','top':'0%'} and "
            "size={'width':'100%','height':'100%'} for full coverage."
        )
    
    # Opacity check for overlay
    if fill_opacity > 0.3:
        warnings.append(
            f"Overlay opacity {fill_opacity} is relatively high (>30%). "
            "Content behind may be significantly obscured. "
            "System prompt recommends 0.15 for subtle overlays."
        )
    elif fill_opacity < 0.05:
        warnings.append(
            f"Overlay opacity {fill_opacity} is very low (<5%). "
            "The overlay effect may be imperceptible."
        )
    
    # Always remind about z-order
    recommendations.append(
        "IMPORTANT: After adding this overlay, run:\n"
        "  ppt_set_z_order.py --file FILE --slide N --shape INDEX --action send_to_back\n"
        "This ensures the overlay appears behind content, not on top of it."
    )


# ============================================================================
# MAIN FUNCTION
# ============================================================================

def add_shape(
    filepath: Path,
    slide_index: int,
    shape_type: str,
    position: Dict[str, Any],
    size: Dict[str, Any],
    fill_color: Optional[str] = None,
    fill_opacity: float = 1.0,
    line_color: Optional[str] = None,
    line_opacity: float = 1.0,
    line_width: float = 1.0,
    text: Optional[str] = None,
    allow_offslide: bool = False,
    is_overlay: bool = False
) -> Dict[str, Any]:
    """
    Add shape to slide with comprehensive validation and opacity support.
    
    Args:
        filepath: Path to PowerPoint file
        slide_index: Target slide index (0-based)
        shape_type: Type of shape to add
        position: Position specification dict
        size: Size specification dict
        fill_color: Fill color (hex or preset name)
        fill_opacity: Fill opacity (0.0=transparent to 1.0=opaque, default: 1.0)
        line_color: Line/border color (hex or preset name)
        line_opacity: Line/border opacity (0.0=transparent to 1.0=opaque, default: 1.0)
        line_width: Line width in points
        text: Optional text to add inside shape
        allow_offslide: Allow positioning outside slide bounds
        is_overlay: Whether this is an overlay shape (applies overlay defaults)
        
    Returns:
        Result dict with shape details and validation info
        
    Raises:
        FileNotFoundError: If file doesn't exist
        SlideNotFoundError: If slide index is invalid
        PowerPointAgentError: If shape creation fails
        ValueError: If opacity values are out of range
    """
    # Validate file exists
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    # Resolve shape type
    resolved_shape = resolve_shape_type(shape_type)
    
    # Resolve colors
    resolved_fill = resolve_color(fill_color)
    resolved_line = resolve_color(line_color)
    
    # Apply overlay defaults if --overlay flag is set
    if is_overlay:
        # Use overlay default position/size if not explicitly provided
        if position == {} or position is None:
            position = OVERLAY_DEFAULTS["position"].copy()
        if size == {} or size is None:
            size = OVERLAY_DEFAULTS["size"].copy()
        # Apply default overlay opacity if still at 1.0
        if fill_opacity == 1.0:
            fill_opacity = OVERLAY_DEFAULTS["fill_opacity"]
    
    # Validate parameters (including opacity)
    validation = validate_shape_params(
        position=position,
        size=size,
        fill_color=resolved_fill,
        fill_opacity=fill_opacity,
        line_color=resolved_line,
        line_opacity=line_opacity,
        text=text,
        allow_offslide=allow_offslide,
        is_overlay=is_overlay
    )
    
    # Open presentation and add shape
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Validate slide index
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range",
                details={
                    "slide_index": slide_index,
                    "total_slides": total_slides,
                    "valid_range": f"0-{total_slides-1}"
                }
            )
        
        # Get presentation version before change
        version_before = agent.get_presentation_version()
        
        # Add shape using core with opacity support
        add_result = agent.add_shape(
            slide_index=slide_index,
            shape_type=resolved_shape,
            position=position,
            size=size,
            fill_color=resolved_fill,
            fill_opacity=fill_opacity,      # NEW: Pass fill opacity
            line_color=resolved_line,
            line_opacity=line_opacity,      # NEW: Pass line opacity
            line_width=line_width,
            text=text
        )
        
        # Save changes
        agent.save()
        
        # Get presentation version after change
        version_after = agent.get_presentation_version()
    
    # Build result
    result = {
        "status": "success" if not validation["has_warnings"] else "warning",
        "file": str(filepath),
        "slide_index": slide_index,
        "shape_type": resolved_shape,
        "shape_type_requested": shape_type,
        "shape_index": add_result.get("shape_index"),
        "position": add_result.get("position", position),
        "size": add_result.get("size", size),
        "styling": {
            "fill_color": resolved_fill,
            "fill_opacity": fill_opacity,
            "fill_transparency": round(1.0 - fill_opacity, 2),
            "line_color": resolved_line,
            "line_opacity": line_opacity,
            "line_width": line_width
        },
        "text": text,
        "is_overlay": is_overlay,
        "presentation_version": {
            "before": version_before,
            "after": version_after
        },
        "core_version": CORE_VERSION,
        "tool_version": __version__
    }
    
    # Add validation results
    if validation["validation_results"]:
        result["validation"] = validation["validation_results"]
    
    if validation["warnings"]:
        result["warnings"] = validation["warnings"]
    
    if validation["recommendations"]:
        result["recommendations"] = validation["recommendations"]
    
    # Add z-order note
    notes = [
        "Shape added to top of z-order (in front of existing shapes).",
        f"Shape index {add_result.get('shape_index')} may change if other shapes are added/removed."
    ]
    
    if is_overlay or fill_opacity < 1.0:
        notes.insert(1, 
            "Use ppt_set_z_order.py --action send_to_back to move overlay behind content."
        )
    else:
        notes.append("Use ppt_set_z_order.py to change layering if needed.")
    
    result["notes"] = notes
    
    # Add next step for overlay
    if is_overlay:
        result["next_step"] = {
            "command": "ppt_set_z_order.py",
            "args": {
                "--file": str(filepath),
                "--slide": slide_index,
                "--shape": add_result.get("shape_index"),
                "--action": "send_to_back"
            },
            "description": "Send overlay to back so it appears behind content"
        }
    
    return result


# ============================================================================
# CLI INTERFACE
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Add shape to PowerPoint slide (v3.1 with opacity support)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

AVAILABLE SHAPES:
  Basic:        rectangle, rounded_rectangle, ellipse/oval, triangle, diamond
  Arrows:       arrow_right, arrow_left, arrow_up, arrow_down
  Polygons:     pentagon, hexagon
  Decorative:   star, heart, lightning, sun, moon, cloud

SHAPE ALIASES:
  rect → rectangle, circle → ellipse, arrow → arrow_right

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

OPACITY/TRANSPARENCY:
  --fill-opacity 1.0    Fully opaque (default)
  --fill-opacity 0.5    50% transparent (half see-through)
  --fill-opacity 0.15   85% transparent (subtle overlay, recommended)
  --fill-opacity 0.0    Fully transparent (invisible)

  Note: Opacity is how SOLID the shape is (1.0 = solid, 0.0 = invisible)
        Transparency is the opposite (0.0 = solid, 1.0 = invisible)

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

OVERLAY MODE (--overlay):
  Quick preset for creating background overlays:
  - Full-slide position and size
  - 15% opacity (subtle, non-competing)
  - Reminder to use ppt_set_z_order.py after

  Usage:
    uv run tools/ppt_add_shape.py --file deck.pptx --slide 0 \\
      --shape rectangle --overlay --fill-color white --json

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

POSITION FORMATS:
  Percentage:   {"left": "20%", "top": "30%"}
  Inches:       {"left": 2.0, "top": 3.0}
  Anchor:       {"anchor": "center", "offset_x": 0, "offset_y": -1.0}
  Grid:         {"grid_row": 2, "grid_col": 3, "grid_size": 12}

ANCHOR POINTS:
  top_left, top_center, top_right
  center_left, center, center_right
  bottom_left, bottom_center, bottom_right

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

COLOR PRESETS:
  primary (#0070C0)    secondary (#595959)    accent (#ED7D31)
  success (#70AD47)    warning (#FFC000)      danger (#C00000)
  white (#FFFFFF)      black (#000000)
  light_gray (#D9D9D9) dark_gray (#404040)

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

EXAMPLES:

  # Semi-transparent callout box
  uv run tools/ppt_add_shape.py \\
    --file presentation.pptx --slide 0 --shape rounded_rectangle \\
    --position '{"left":"10%","top":"15%"}' \\
    --size '{"width":"30%","height":"15%"}' \\
    --fill-color primary --fill-opacity 0.8 --text "Key Point" --json

  # Subtle white overlay for text readability (System Prompt default)
  uv run tools/ppt_add_shape.py \\
    --file presentation.pptx --slide 2 --shape rectangle \\
    --position '{"left":"0%","top":"0%"}' \\
    --size '{"width":"100%","height":"100%"}' \\
    --fill-color "#FFFFFF" --fill-opacity 0.15 --json
  # Then run: ppt_set_z_order.py ... --action send_to_back

  # Quick overlay using --overlay preset
  uv run tools/ppt_add_shape.py \\
    --file presentation.pptx --slide 3 --shape rectangle \\
    --overlay --fill-color black --json
  # Automatically uses opacity 0.15 and full-slide positioning

  # Dark overlay for light text on busy background
  uv run tools/ppt_add_shape.py \\
    --file presentation.pptx --slide 4 --shape rectangle \\
    --overlay --fill-color "#000000" --fill-opacity 0.3 --json

  # Fully opaque shape (default behavior)
  uv run tools/ppt_add_shape.py \\
    --file presentation.pptx --slide 0 --shape ellipse \\
    --position '{"anchor":"center"}' \\
    --size '{"width":"20%","height":"20%"}' \\
    --fill-color "#FFC000" --json

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Z-ORDER (LAYERING):
  Shapes are added on TOP of existing shapes by default.
  
  For overlays, you MUST send them to back:
    1. Add the overlay shape (with --overlay or --fill-opacity)
    2. Note the shape_index from the output
    3. Run: ppt_set_z_order.py --file FILE --slide N --shape INDEX --action send_to_back
  
  Shape indices change after z-order operations - always re-query with
  ppt_get_slide_info.py before referencing shapes by index.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
        """
    )
    
    # Required arguments
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file path (must exist)'
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
        help=f'Shape type: {", ".join(AVAILABLE_SHAPES[:8])}... (see --help for full list)'
    )
    
    parser.add_argument(
        '--position',
        type=json.loads,
        default={},
        help='Position dict as JSON string (e.g., \'{"left":"20%%","top":"30%%"}\'). '
             'With --overlay, defaults to full-slide positioning.'
    )
    
    # Optional arguments
    parser.add_argument(
        '--size',
        type=json.loads,
        help='Size dict as JSON string (e.g., \'{"width":"40%%","height":"30%%"}\'). '
             'Defaults to 20%% x 20%% (or 100%% x 100%% with --overlay).'
    )
    
    parser.add_argument(
        '--fill-color',
        help='Fill color: hex (#0070C0) or preset name (primary, success, etc.)'
    )
    
    parser.add_argument(
        '--fill-opacity',
        type=float,
        default=1.0,
        help='Fill opacity: 0.0 (transparent) to 1.0 (opaque). '
             'Default: 1.0. For overlays, use 0.15 for subtle effect.'
    )
    
    parser.add_argument(
        '--line-color',
        help='Line/border color: hex or preset name'
    )
    
    parser.add_argument(
        '--line-opacity',
        type=float,
        default=1.0,
        help='Line/border opacity: 0.0 (transparent) to 1.0 (opaque). Default: 1.0'
    )
    
    parser.add_argument(
        '--line-width',
        type=float,
        default=1.0,
        help='Line width in points (default: 1.0)'
    )
    
    parser.add_argument(
        '--text',
        help='Text to add inside the shape'
    )
    
    parser.add_argument(
        '--overlay',
        action='store_true',
        help='Overlay preset: full-slide, 15%% opacity, z-order reminder. '
             'Equivalent to: --position \'{"left":"0%%","top":"0%%"}\' '
             '--size \'{"width":"100%%","height":"100%%"}\' --fill-opacity 0.15'
    )
    
    parser.add_argument(
        '--allow-offslide',
        action='store_true',
        help='Allow positioning outside slide bounds (suppresses warnings)'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default: true)'
    )
    
    parser.add_argument(
        '--version',
        action='version',
        version=f'%(prog)s {__version__} (core: {CORE_VERSION})'
    )
    
    args = parser.parse_args()
    
    try:
        # Process size - use defaults or merge from position
        size = args.size if args.size else {}
        position = args.position if args.position else {}
        
        # Allow size in position dict for convenience
        if "width" in position and "width" not in size:
            size["width"] = position.pop("width")
        if "height" in position and "height" not in size:
            size["height"] = position.pop("height")
        
        # Apply defaults based on overlay mode
        if args.overlay:
            # Overlay mode: default to full-slide
            if "left" not in position:
                position["left"] = "0%"
            if "top" not in position:
                position["top"] = "0%"
            if "width" not in size:
                size["width"] = "100%"
            if "height" not in size:
                size["height"] = "100%"
        else:
            # Normal mode: default to 20% x 20%
            if "width" not in size:
                size["width"] = "20%"
            if "height" not in size:
                size["height"] = "20%"
        
        # Execute
        result = add_shape(
            filepath=args.file,
            slide_index=args.slide,
            shape_type=args.shape,
            position=position,
            size=size,
            fill_color=args.fill_color,
            fill_opacity=args.fill_opacity,
            line_color=args.line_color,
            line_opacity=args.line_opacity,
            line_width=args.line_width,
            text=args.text,
            allow_offslide=args.allow_offslide,
            is_overlay=args.overlay
        )
        
        # Output
        if args.json:
            print(json.dumps(result, indent=2))
        else:
            status_icon = "✅" if result["status"] == "success" else "⚠️"
            print(f"{status_icon} Added {result['shape_type']} to slide {result['slide_index']}")
            print(f"   Shape index: {result['shape_index']}")
            if args.fill_color:
                opacity_str = f" (opacity: {args.fill_opacity})" if args.fill_opacity < 1.0 else ""
                print(f"   Fill: {result['styling']['fill_color']}{opacity_str}")
            if args.line_color:
                print(f"   Line: {result['styling']['line_color']} ({args.line_width}pt)")
            if args.text:
                print(f"   Text: {args.text[:50]}{'...' if len(args.text) > 50 else ''}")
            if args.overlay:
                print(f"\n   ⚠️  OVERLAY MODE: Remember to run ppt_set_z_order.py --action send_to_back")
            if result.get("warnings"):
                print("\n   Warnings:")
                for warning in result["warnings"]:
                    print(f"   ⚠️  {warning}")
        
        sys.exit(0)
        
    except ValueError as e:
        # Validation errors (including opacity out of range)
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Check that opacity values are between 0.0 and 1.0"
        }
        if args.json:
            print(json.dumps(error_result, indent=2))
        else:
            print(f"❌ Error: {e}", file=sys.stderr)
        sys.exit(1)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify the file path exists and is accessible."
        }
        if args.json:
            print(json.dumps(error_result, indent=2))
        else:
            print(f"❌ Error: {e}", file=sys.stderr)
        sys.exit(1)
        
    except SlideNotFoundError as e:
        error_result = {
            "status": "error",
            "error": e.message,
            "error_type": "SlideNotFoundError",
            "details": e.details,
            "suggestion": "Use ppt_get_info.py to check available slide indices."
        }
        if args.json:
            print(json.dumps(error_result, indent=2))
        else:
            print(f"❌ Error: {e.message}", file=sys.stderr)
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": e.message,
            "error_type": type(e).__name__,
            "details": e.details
        }
        if args.json:
            print(json.dumps(error_result, indent=2))
        else:
            print(f"❌ Error: {e.message}", file=sys.stderr)
        sys.exit(1)
        
    except json.JSONDecodeError as e:
        error_result = {
            "status": "error",
            "error": f"Invalid JSON in argument: {e}",
            "error_type": "JSONDecodeError",
            "suggestion": "Ensure --position and --size are valid JSON strings."
        }
        if args.json:
            print(json.dumps(error_result, indent=2))
        else:
            print(f"❌ Error: Invalid JSON - {e}", file=sys.stderr)
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
