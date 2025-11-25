# Deep Analysis & Enhancement: Core Library Opacity Support

## Current Implementation Analysis

### What Exists

```python
# In add_shape():
if fill_color:
    shape.fill.solid()
    shape.fill.fore_color.rgb = ColorHelper.from_hex(fill_color)
# ❌ No opacity handling

# In format_shape():
if transparency is not None:
    try:
        shape.fill.solid()
        # Note: python-pptx doesn't directly support transparency
        # This is a best-effort implementation
        changes.append("transparency_attempted")  # ❌ Does nothing!
    except Exception:
        pass
```

### The Problem

The current `format_shape` transparency implementation is essentially a **no-op**. It doesn't actually set any transparency. The comment acknowledges this limitation.

### The Solution

`python-pptx` doesn't expose transparency as a direct API, but we can **manipulate the underlying OOXML** to achieve it. The transparency is controlled via an `<a:alpha>` child element within the color specification.

---

## OOXML Structure for Transparency

### Fill Transparency

```xml
<a:solidFill>
  <a:srgbClr val="0070C0">
    <a:alpha val="15000"/>  <!-- 15% opacity (85% transparent) -->
  </a:srgbClr>
</a:solidFill>
```

### Line Transparency

```xml
<a:ln w="12700">
  <a:solidFill>
    <a:srgbClr val="000000">
      <a:alpha val="50000"/>  <!-- 50% opacity -->
    </a:srgbClr>
  </a:solidFill>
</a:ln>
```

### Alpha Value Scale

| Opacity | Alpha Value | Effect |
|---------|-------------|--------|
| 1.0 | 100000 (or omit) | Fully opaque |
| 0.5 | 50000 | 50% see-through |
| 0.15 | 15000 | Subtle overlay |
| 0.0 | 0 | Fully transparent (invisible) |

---

## Enhanced Core Implementation

### Complete Enhanced `powerpoint_agent_core.py` Shape Operations Section

```python
    # ========================================================================
    # SHAPE OPERATIONS
    # ========================================================================
    
    def _set_fill_opacity(self, shape, opacity: float) -> bool:
        """
        Set the fill opacity of a shape by manipulating the underlying XML.
        
        Args:
            shape: The shape object with a fill
            opacity: Opacity value (0.0 = fully transparent, 1.0 = fully opaque)
            
        Returns:
            True if opacity was set, False if not applicable
            
        Note:
            python-pptx doesn't directly expose fill transparency, so we
            manipulate the OOXML directly. The alpha value uses a scale
            where 100000 = 100% opaque.
        """
        if opacity >= 1.0:
            # No need to set alpha for fully opaque - it's the default
            return True
        
        if opacity < 0.0:
            opacity = 0.0
        
        try:
            # Access the shape's spPr (shape properties) element
            spPr = shape._sp.spPr
            if spPr is None:
                return False
            
            # Find the solidFill element
            solidFill = spPr.find(qn('a:solidFill'))
            if solidFill is None:
                return False
            
            # Find the color element (could be srgbClr or schemeClr)
            color_elem = solidFill.find(qn('a:srgbClr'))
            if color_elem is None:
                color_elem = solidFill.find(qn('a:schemeClr'))
            if color_elem is None:
                return False
            
            # Calculate alpha value (Office uses 0-100000 scale, where 100000 = 100%)
            alpha_value = int(opacity * 100000)
            
            # Remove existing alpha element if present
            existing_alpha = color_elem.find(qn('a:alpha'))
            if existing_alpha is not None:
                color_elem.remove(existing_alpha)
            
            # Create and add new alpha element
            # Using SubElement to create properly namespaced element
            nsmap = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
            alpha_elem = etree.SubElement(color_elem, qn('a:alpha'))
            alpha_elem.set('val', str(alpha_value))
            
            return True
            
        except Exception as e:
            # Log but don't fail - opacity is enhancement, not critical
            self._log_warning(f"Could not set fill opacity: {e}")
            return False
    
    def _set_line_opacity(self, shape, opacity: float) -> bool:
        """
        Set the line/border opacity of a shape by manipulating the underlying XML.
        
        Args:
            shape: The shape object with a line
            opacity: Opacity value (0.0 = fully transparent, 1.0 = fully opaque)
            
        Returns:
            True if opacity was set, False if not applicable
            
        Note:
            Line opacity requires the line to have a solid fill. We manipulate
            the OOXML <a:ln><a:solidFill><a:srgbClr><a:alpha> structure.
        """
        if opacity >= 1.0:
            return True
        
        if opacity < 0.0:
            opacity = 0.0
        
        try:
            # Access the shape's spPr element
            spPr = shape._sp.spPr
            if spPr is None:
                return False
            
            # Find the line element
            ln = spPr.find(qn('a:ln'))
            if ln is None:
                return False
            
            # Find solidFill within line
            solidFill = ln.find(qn('a:solidFill'))
            if solidFill is None:
                # Line might not have a fill yet - try to find/create one
                return False
            
            # Find color element
            color_elem = solidFill.find(qn('a:srgbClr'))
            if color_elem is None:
                color_elem = solidFill.find(qn('a:schemeClr'))
            if color_elem is None:
                return False
            
            # Calculate and set alpha
            alpha_value = int(opacity * 100000)
            
            existing_alpha = color_elem.find(qn('a:alpha'))
            if existing_alpha is not None:
                color_elem.remove(existing_alpha)
            
            alpha_elem = etree.SubElement(color_elem, qn('a:alpha'))
            alpha_elem.set('val', str(alpha_value))
            
            return True
            
        except Exception as e:
            self._log_warning(f"Could not set line opacity: {e}")
            return False
    
    def _ensure_line_solid_fill(self, shape, color_hex: str) -> bool:
        """
        Ensure the shape's line has a solid fill with the specified color.
        This is necessary before setting line opacity.
        
        Args:
            shape: The shape object
            color_hex: Hex color string for the line
            
        Returns:
            True if successful
        """
        try:
            # Set line color through python-pptx first
            shape.line.color.rgb = ColorHelper.from_hex(color_hex)
            
            # Now ensure the XML structure is correct for opacity
            spPr = shape._sp.spPr
            ln = spPr.find(qn('a:ln'))
            
            if ln is None:
                return False
            
            # Check if solidFill exists
            solidFill = ln.find(qn('a:solidFill'))
            if solidFill is None:
                # Create solidFill structure
                solidFill = etree.SubElement(ln, qn('a:solidFill'))
                color_elem = etree.SubElement(solidFill, qn('a:srgbClr'))
                # Remove # from hex color
                color_val = color_hex.lstrip('#').upper()
                color_elem.set('val', color_val)
            
            return True
            
        except Exception as e:
            self._log_warning(f"Could not ensure line solid fill: {e}")
            return False
    
    def _log_warning(self, message: str) -> None:
        """
        Log a warning message. Override in subclasses for custom logging.
        
        Args:
            message: Warning message to log
        """
        # Default implementation - can be enhanced with proper logging
        import sys
        print(f"WARNING: {message}", file=sys.stderr)
    
    def add_shape(
        self,
        slide_index: int,
        shape_type: str,
        position: Dict[str, Any],
        size: Dict[str, Any],
        fill_color: Optional[str] = None,
        fill_opacity: float = 1.0,
        line_color: Optional[str] = None,
        line_opacity: float = 1.0,
        line_width: float = 1.0,
        text: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Add shape to slide with optional transparency/opacity support.
        
        Args:
            slide_index: Target slide index
            shape_type: Shape type name (rectangle, ellipse, arrow_right, etc.)
            position: Position dict (percentage, inches, anchor, or grid)
            size: Size dict (percentage or inches)
            fill_color: Fill color hex (e.g., "#0070C0") or None for no fill
            fill_opacity: Fill opacity from 0.0 (transparent) to 1.0 (opaque).
                         Default is 1.0 (fully opaque). Use 0.15 for subtle overlays.
            line_color: Line/border color hex or None for no line
            line_opacity: Line opacity from 0.0 (transparent) to 1.0 (opaque).
                         Default is 1.0 (fully opaque).
            line_width: Line width in points (default: 1.0)
            text: Optional text to add inside shape
            
        Returns:
            Dict with shape_index, position, size, and applied styling details
            
        Raises:
            SlideNotFoundError: If slide index is invalid
            ValueError: If size is not specified or opacity is out of range
            
        Example:
            # Subtle white overlay for improved text readability
            agent.add_shape(
                slide_index=0,
                shape_type="rectangle",
                position={"left": "0%", "top": "0%"},
                size={"width": "100%", "height": "100%"},
                fill_color="#FFFFFF",
                fill_opacity=0.15  # 15% opaque = 85% transparent
            )
        """
        # Validate opacity ranges
        if not 0.0 <= fill_opacity <= 1.0:
            raise ValueError(
                f"fill_opacity must be between 0.0 and 1.0, got {fill_opacity}"
            )
        if not 0.0 <= line_opacity <= 1.0:
            raise ValueError(
                f"line_opacity must be between 0.0 and 1.0, got {line_opacity}"
            )
        
        slide = self._get_slide(slide_index)
        
        left, top = Position.from_dict(position)
        width, height = Size.from_dict(size)
        
        if width is None or height is None:
            raise ValueError("Shape must have explicit width and height")
        
        # Map shape type string to MSO constant
        shape_type_map = {
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
        
        mso_shape = shape_type_map.get(
            shape_type.lower(),
            MSO_AUTO_SHAPE_TYPE.RECTANGLE
        )
        
        # Add shape
        shape = slide.shapes.add_shape(
            mso_shape,
            Inches(left), Inches(top),
            Inches(width), Inches(height)
        )
        
        # Track what was actually applied
        styling_applied = {
            "fill_color": None,
            "fill_opacity": 1.0,
            "fill_opacity_applied": False,
            "line_color": None,
            "line_opacity": 1.0,
            "line_opacity_applied": False,
            "line_width": line_width
        }
        
        # Apply fill color and opacity
        if fill_color:
            shape.fill.solid()
            shape.fill.fore_color.rgb = ColorHelper.from_hex(fill_color)
            styling_applied["fill_color"] = fill_color
            styling_applied["fill_opacity"] = fill_opacity
            
            # Apply fill opacity if not fully opaque
            if fill_opacity < 1.0:
                opacity_set = self._set_fill_opacity(shape, fill_opacity)
                styling_applied["fill_opacity_applied"] = opacity_set
        else:
            # No fill - make background transparent
            shape.fill.background()
        
        # Apply line color and opacity
        if line_color:
            # Ensure line has solid fill for opacity support
            self._ensure_line_solid_fill(shape, line_color)
            shape.line.width = Pt(line_width)
            styling_applied["line_color"] = line_color
            styling_applied["line_opacity"] = line_opacity
            
            # Apply line opacity if not fully opaque
            if line_opacity < 1.0:
                opacity_set = self._set_line_opacity(shape, line_opacity)
                styling_applied["line_opacity_applied"] = opacity_set
        else:
            # No line
            shape.line.fill.background()
        
        # Add text if provided
        if text and shape.has_text_frame:
            shape.text_frame.text = text
        
        shape_index = len(slide.shapes) - 1
        
        return {
            "slide_index": slide_index,
            "shape_index": shape_index,
            "shape_type": shape_type,
            "position": {"left": left, "top": top},
            "size": {"width": width, "height": height},
            "styling": styling_applied,
            "has_text": text is not None,
            "text_preview": text[:50] + "..." if text and len(text) > 50 else text
        }
    
    def format_shape(
        self,
        slide_index: int,
        shape_index: int,
        fill_color: Optional[str] = None,
        fill_opacity: Optional[float] = None,
        line_color: Optional[str] = None,
        line_opacity: Optional[float] = None,
        line_width: Optional[float] = None,
        # Deprecated parameter - kept for backward compatibility
        transparency: Optional[float] = None
    ) -> Dict[str, Any]:
        """
        Format existing shape with optional transparency/opacity support.
        
        Args:
            slide_index: Target slide index
            shape_index: Shape index on slide
            fill_color: Fill color hex (e.g., "#0070C0")
            fill_opacity: Fill opacity from 0.0 (transparent) to 1.0 (opaque)
            line_color: Line/border color hex
            line_opacity: Line opacity from 0.0 (transparent) to 1.0 (opaque)
            line_width: Line width in points
            transparency: DEPRECATED - Use fill_opacity instead.
                         If provided, converted to fill_opacity (transparency = 1 - opacity)
            
        Returns:
            Dict with formatting changes applied and their status
            
        Raises:
            SlideNotFoundError: If slide index is invalid
            ShapeNotFoundError: If shape index is invalid
            ValueError: If opacity values are out of range
            
        Example:
            # Make an existing shape semi-transparent
            agent.format_shape(
                slide_index=0,
                shape_index=3,
                fill_opacity=0.5  # 50% opaque
            )
        """
        shape = self._get_shape(slide_index, shape_index)
        
        changes = []
        changes_detail = {}
        
        # Handle deprecated transparency parameter
        if transparency is not None and fill_opacity is None:
            # Convert transparency to opacity (they're inverses)
            fill_opacity = 1.0 - transparency
            changes.append("transparency_converted_to_opacity")
            changes_detail["transparency_deprecated"] = True
            changes_detail["converted_opacity"] = fill_opacity
        
        # Validate opacity ranges
        if fill_opacity is not None and not 0.0 <= fill_opacity <= 1.0:
            raise ValueError(
                f"fill_opacity must be between 0.0 and 1.0, got {fill_opacity}"
            )
        if line_opacity is not None and not 0.0 <= line_opacity <= 1.0:
            raise ValueError(
                f"line_opacity must be between 0.0 and 1.0, got {line_opacity}"
            )
        
        # Apply fill color
        if fill_color is not None:
            shape.fill.solid()
            shape.fill.fore_color.rgb = ColorHelper.from_hex(fill_color)
            changes.append("fill_color")
            changes_detail["fill_color"] = fill_color
        
        # Apply fill opacity
        if fill_opacity is not None:
            # Ensure shape has solid fill before applying opacity
            if fill_color is None:
                # Need to solidify fill first if not already done
                try:
                    shape.fill.solid()
                except Exception:
                    pass
            
            if fill_opacity < 1.0:
                success = self._set_fill_opacity(shape, fill_opacity)
                if success:
                    changes.append("fill_opacity")
                    changes_detail["fill_opacity"] = fill_opacity
                    changes_detail["fill_opacity_applied"] = True
                else:
                    changes.append("fill_opacity_failed")
                    changes_detail["fill_opacity_applied"] = False
            else:
                changes.append("fill_opacity_reset")
                changes_detail["fill_opacity"] = 1.0
        
        # Apply line color
        if line_color is not None:
            self._ensure_line_solid_fill(shape, line_color)
            changes.append("line_color")
            changes_detail["line_color"] = line_color
        
        # Apply line opacity
        if line_opacity is not None:
            if line_opacity < 1.0:
                success = self._set_line_opacity(shape, line_opacity)
                if success:
                    changes.append("line_opacity")
                    changes_detail["line_opacity"] = line_opacity
                    changes_detail["line_opacity_applied"] = True
                else:
                    changes.append("line_opacity_failed")
                    changes_detail["line_opacity_applied"] = False
            else:
                changes.append("line_opacity_reset")
                changes_detail["line_opacity"] = 1.0
        
        # Apply line width
        if line_width is not None:
            shape.line.width = Pt(line_width)
            changes.append("line_width")
            changes_detail["line_width"] = line_width
        
        return {
            "slide_index": slide_index,
            "shape_index": shape_index,
            "changes_applied": changes,
            "changes_detail": changes_detail,
            "success": "failed" not in "".join(changes)
        }
    
    def remove_shape(self, slide_index: int, shape_index: int) -> Dict[str, Any]:
        """
        Remove shape from slide.
        
        Args:
            slide_index: Target slide index
            shape_index: Shape index to remove
            
        Returns:
            Dict with removal details
            
        Raises:
            SlideNotFoundError: If slide index is invalid
            ShapeNotFoundError: If shape index is invalid
        """
        slide = self._get_slide(slide_index)
        shape = self._get_shape(slide_index, shape_index)
        
        # Get shape info before removal
        shape_name = shape.name
        shape_type = str(shape.shape_type)
        
        # Remove shape from slide
        sp = shape.element
        sp.getparent().remove(sp)
        
        return {
            "slide_index": slide_index,
            "removed_shape_index": shape_index,
            "removed_shape_name": shape_name,
            "removed_shape_type": shape_type,
            "new_shape_count": len(slide.shapes)
        }
    
    def set_z_order(
        self,
        slide_index: int,
        shape_index: int,
        action: str
    ) -> Dict[str, Any]:
        """
        Change the z-order (stacking order) of a shape.
        
        Args:
            slide_index: Target slide index
            shape_index: Shape index to modify
            action: One of "bring_to_front", "send_to_back", 
                   "bring_forward", "send_backward"
            
        Returns:
            Dict with z-order change details including old and new positions
            
        Raises:
            SlideNotFoundError: If slide index is invalid
            ShapeNotFoundError: If shape index is invalid
            ValueError: If action is invalid
        """
        valid_actions = {"bring_to_front", "send_to_back", "bring_forward", "send_backward"}
        if action not in valid_actions:
            raise ValueError(f"Invalid action: {action}. Must be one of {valid_actions}")
        
        slide = self._get_slide(slide_index)
        shape = self._get_shape(slide_index, shape_index)
        
        # Access the shape tree XML element
        sp_tree = slide.shapes._spTree
        element = shape.element
        
        # Find current position in XML tree
        current_index = -1
        shape_elements = [child for child in sp_tree if child.tag.endswith('}sp') or 
                         child.tag.endswith('}pic') or child.tag.endswith('}graphicFrame')]
        
        for i, child in enumerate(sp_tree):
            if child == element:
                current_index = i
                break
        
        if current_index == -1:
            raise PowerPointAgentError(
                "Could not locate shape in XML tree",
                details={"slide_index": slide_index, "shape_index": shape_index}
            )
        
        new_index = current_index
        max_index = len(sp_tree) - 1
        
        # Execute the z-order action
        if action == "bring_to_front":
            sp_tree.remove(element)
            sp_tree.append(element)
            new_index = len(sp_tree) - 1
            
        elif action == "send_to_back":
            sp_tree.remove(element)
            # Insert after nvGrpSpPr and grpSpPr (indices 0 and 1 typically)
            sp_tree.insert(2, element)
            new_index = 2
            
        elif action == "bring_forward":
            if current_index < max_index:
                sp_tree.remove(element)
                sp_tree.insert(current_index + 1, element)
                new_index = current_index + 1
                
        elif action == "send_backward":
            if current_index > 2:  # Don't go before required elements
                sp_tree.remove(element)
                sp_tree.insert(current_index - 1, element)
                new_index = current_index - 1
        
        return {
            "slide_index": slide_index,
            "shape_index": shape_index,
            "action": action,
            "z_order_change": {
                "from": current_index,
                "to": new_index
            },
            "warning": "Shape indices may have changed after z-order operation. Re-query slide info."
        }
    
    def add_table(
        self,
        slide_index: int,
        rows: int,
        cols: int,
        position: Dict[str, Any],
        size: Dict[str, Any],
        data: Optional[List[List[Any]]] = None,
        header_row: bool = True
    ) -> Dict[str, Any]:
        """
        Add table to slide.
        
        Args:
            slide_index: Target slide index
            rows: Number of rows
            cols: Number of columns
            position: Position dict
            size: Size dict
            data: Optional 2D list of cell values
            header_row: Whether first row is header (styling hint)
            
        Returns:
            Dict with shape_index and table details
        """
        slide = self._get_slide(slide_index)
        
        left, top = Position.from_dict(position)
        width, height = Size.from_dict(size)
        
        if width is None or height is None:
            raise ValueError("Table must have explicit width and height")
        
        # Create table
        table_shape = slide.shapes.add_table(
            rows, cols,
            Inches(left), Inches(top),
            Inches(width), Inches(height)
        )
        
        table = table_shape.table
        
        # Populate with data if provided
        cells_filled = 0
        if data:
            for row_idx, row_data in enumerate(data):
                if row_idx >= rows:
                    break
                for col_idx, cell_value in enumerate(row_data):
                    if col_idx >= cols:
                        break
                    table.cell(row_idx, col_idx).text = str(cell_value)
                    cells_filled += 1
        
        shape_index = len(slide.shapes) - 1
        
        return {
            "slide_index": slide_index,
            "shape_index": shape_index,
            "rows": rows,
            "cols": cols,
            "cells_filled": cells_filled,
            "position": {"left": left, "top": top},
            "size": {"width": width, "height": height}
        }
    
    def add_connector(
        self,
        slide_index: int,
        from_shape_index: int,
        to_shape_index: int,
        connector_type: str = "straight"
    ) -> Dict[str, Any]:
        """
        Add connector line between two shapes.
        
        Args:
            slide_index: Target slide index
            from_shape_index: Starting shape index
            to_shape_index: Ending shape index
            connector_type: "straight", "elbow", or "curved"
            
        Returns:
            Dict with connector details
        """
        slide = self._get_slide(slide_index)
        
        shape1 = self._get_shape(slide_index, from_shape_index)
        shape2 = self._get_shape(slide_index, to_shape_index)
        
        # Calculate center points
        x1 = shape1.left + shape1.width // 2
        y1 = shape1.top + shape1.height // 2
        x2 = shape2.left + shape2.width // 2
        y2 = shape2.top + shape2.height // 2
        
        # Map connector type
        connector_map = {
            "straight": MSO_CONNECTOR.STRAIGHT,
            "elbow": MSO_CONNECTOR.ELBOW,
            "curved": MSO_CONNECTOR.CURVE
        }
        mso_connector = connector_map.get(connector_type.lower(), MSO_CONNECTOR.STRAIGHT)
        
        # Add connector
        connector = slide.shapes.add_connector(
            mso_connector,
            x1, y1, x2, y2
        )
        
        shape_index = len(slide.shapes) - 1
        
        return {
            "slide_index": slide_index,
            "shape_index": shape_index,
            "from_shape": from_shape_index,
            "to_shape": to_shape_index,
            "connector_type": connector_type
        }
```

---

## Required Imports

Add these imports at the top of `powerpoint_agent_core.py`:

```python
# Add to existing imports section
from lxml import etree
from pptx.oxml.ns import qn

# qn() creates qualified names for XML namespace handling
# Example: qn('a:solidFill') -> '{http://schemas.openxmlformats.org/drawingml/2006/main}solidFill'
```

---

## Summary of Changes

### New Private Helper Methods

| Method | Purpose |
|--------|---------|
| `_set_fill_opacity()` | Manipulates OOXML to set fill alpha value |
| `_set_line_opacity()` | Manipulates OOXML to set line alpha value |
| `_ensure_line_solid_fill()` | Ensures line has proper XML structure for opacity |
| `_log_warning()` | Standardized warning logging |

### Enhanced `add_shape()` Method

| Change | Description |
|--------|-------------|
| New `fill_opacity` param | Float 0.0-1.0, default 1.0 |
| New `line_opacity` param | Float 0.0-1.0, default 1.0 |
| Opacity validation | ValueError if out of range |
| Enhanced return dict | Includes `styling` with opacity details |
| XML manipulation | Uses `_set_fill_opacity()` for transparency |

### Enhanced `format_shape()` Method

| Change | Description |
|--------|-------------|
| New `fill_opacity` param | Now actually works! |
| New `line_opacity` param | Supports line transparency |
| Deprecated `transparency` | Converted to `fill_opacity` for backward compatibility |
| Enhanced return dict | Includes `changes_detail` with specifics |
| Success indicator | Returns `success` boolean |

### Backward Compatibility

| Aspect | Status |
|--------|--------|
| Default behavior | ✅ Unchanged (opacity 1.0) |
| Existing code | ✅ Works without changes |
| Old `transparency` param | ✅ Auto-converted to opacity |
| Return dict | ✅ Superset of old format |

---

## Verification Test

After implementing, verify with:

```bash
# Test that help shows new parameters
uv run tools/ppt_add_shape.py --help | grep -E "opacity|overlay"

# Test creating a transparent overlay
uv run tools/ppt_add_shape.py \
  --file test.pptx --slide 0 --shape rectangle \
  --position '{"left":"0%","top":"0%"}' \
  --size '{"width":"100%","height":"100%"}' \
  --fill-color "#FFFFFF" --fill-opacity 0.15 --json

# Verify the output includes opacity info
# Expected in response:
# "styling": {
#   "fill_color": "#FFFFFF",
#   "fill_opacity": 0.15,
#   "fill_opacity_applied": true,
#   ...
# }
```

---

## Complete Integration Checklist

- [ ] Add imports (`lxml.etree`, `pptx.oxml.ns.qn`)
- [ ] Add `_set_fill_opacity()` method
- [ ] Add `_set_line_opacity()` method
- [ ] Add `_ensure_line_solid_fill()` method
- [ ] Add `_log_warning()` method
- [ ] Update `add_shape()` signature and implementation
- [ ] Update `format_shape()` signature and implementation
- [ ] Update `ppt_add_shape.py` to v3.1.0 (already done)
- [ ] Update `ppt_format_shape.py` to support new parameters
- [ ] Add unit tests for opacity functionality
- [ ] Update system prompt to remove the opacity gap note
