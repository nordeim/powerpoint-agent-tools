#!/usr/bin/env python3
"""
Comprehensive Test Suite for Shape Opacity Features
Tests add_shape() and format_shape() opacity functionality.

Author: PowerPoint Agent Team
Version: 1.0.0

Usage:
    # Run all tests
    pytest tests/test_shape_opacity.py -v
    
    # Run specific test
    pytest tests/test_shape_opacity.py::TestAddShapeOpacity::test_fill_opacity_overlay -v
    
    # Run with coverage
    pytest tests/test_shape_opacity.py --cov=core.powerpoint_agent_core -v
"""

import sys
import json
import tempfile
import shutil
from pathlib import Path
from typing import Dict, Any

import pytest

# Add parent directory to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent,
    PowerPointAgentError,
    SlideNotFoundError,
    ShapeNotFoundError,
)


# ============================================================================
# FIXTURES
# ============================================================================

@pytest.fixture
def temp_dir():
    """Create a temporary directory for test files."""
    dirpath = tempfile.mkdtemp()
    yield Path(dirpath)
    shutil.rmtree(dirpath, ignore_errors=True)


@pytest.fixture
def test_presentation(temp_dir) -> Path:
    """Create a test presentation with a slide."""
    pptx_path = temp_dir / "test_opacity.pptx"
    
    with PowerPointAgent() as agent:
        agent.create_new()
        # Add a slide with content
        agent.add_slide(layout_name="Blank")
        agent.save(pptx_path)
    
    return pptx_path


@pytest.fixture
def presentation_with_shape(temp_dir) -> tuple:
    """Create a test presentation with an existing shape."""
    pptx_path = temp_dir / "test_with_shape.pptx"
    
    with PowerPointAgent() as agent:
        agent.create_new()
        agent.add_slide(layout_name="Blank")
        
        # Add a shape to format later
        result = agent.add_shape(
            slide_index=0,
            shape_type="rectangle",
            position={"left": "20%", "top": "20%"},
            size={"width": "30%", "height": "30%"},
            fill_color="#0070C0"
        )
        shape_index = result["shape_index"]
        
        agent.save(pptx_path)
    
    return pptx_path, shape_index


# ============================================================================
# TEST: add_shape() OPACITY FUNCTIONALITY
# ============================================================================

class TestAddShapeOpacity:
    """Tests for add_shape() opacity features."""
    
    def test_add_shape_default_opacity(self, test_presentation):
        """Test that default opacity is 1.0 (fully opaque)."""
        with PowerPointAgent(test_presentation) as agent:
            agent.open(test_presentation)
            
            result = agent.add_shape(
                slide_index=0,
                shape_type="rectangle",
                position={"left": "10%", "top": "10%"},
                size={"width": "20%", "height": "20%"},
                fill_color="#FF0000"
            )
            
            agent.save()
        
        assert result["styling"]["fill_opacity"] == 1.0
        assert result["styling"]["fill_opacity_applied"] == False
        assert result["styling"]["line_opacity"] == 1.0
    
    def test_add_shape_fill_opacity_overlay(self, test_presentation):
        """Test creating an overlay with 15% opacity."""
        with PowerPointAgent(test_presentation) as agent:
            agent.open(test_presentation)
            
            result = agent.add_shape(
                slide_index=0,
                shape_type="rectangle",
                position={"left": "0%", "top": "0%"},
                size={"width": "100%", "height": "100%"},
                fill_color="#FFFFFF",
                fill_opacity=0.15
            )
            
            agent.save()
        
        assert result["styling"]["fill_color"] == "#FFFFFF"
        assert result["styling"]["fill_opacity"] == 0.15
        assert result["styling"]["fill_opacity_applied"] == True
        assert "shape_index" in result
    
    def test_add_shape_fill_opacity_half(self, test_presentation):
        """Test 50% opacity."""
        with PowerPointAgent(test_presentation) as agent:
            agent.open(test_presentation)
            
            result = agent.add_shape(
                slide_index=0,
                shape_type="ellipse",
                position={"left": "30%", "top": "30%"},
                size={"width": "40%", "height": "40%"},
                fill_color="#00FF00",
                fill_opacity=0.5
            )
            
            agent.save()
        
        assert result["styling"]["fill_opacity"] == 0.5
        assert result["styling"]["fill_opacity_applied"] == True
    
    def test_add_shape_fill_opacity_zero(self, test_presentation):
        """Test fully transparent (0.0 opacity)."""
        with PowerPointAgent(test_presentation) as agent:
            agent.open(test_presentation)
            
            result = agent.add_shape(
                slide_index=0,
                shape_type="rectangle",
                position={"left": "10%", "top": "10%"},
                size={"width": "20%", "height": "20%"},
                fill_color="#0000FF",
                fill_opacity=0.0
            )
            
            agent.save()
        
        assert result["styling"]["fill_opacity"] == 0.0
        assert result["styling"]["fill_opacity_applied"] == True
    
    def test_add_shape_line_opacity(self, test_presentation):
        """Test line/border opacity."""
        with PowerPointAgent(test_presentation) as agent:
            agent.open(test_presentation)
            
            result = agent.add_shape(
                slide_index=0,
                shape_type="rectangle",
                position={"left": "10%", "top": "10%"},
                size={"width": "20%", "height": "20%"},
                fill_color="#FFFFFF",
                line_color="#000000",
                line_opacity=0.5,
                line_width=2.0
            )
            
            agent.save()
        
        assert result["styling"]["line_color"] == "#000000"
        assert result["styling"]["line_opacity"] == 0.5
        assert result["styling"]["line_opacity_applied"] == True
    
    def test_add_shape_both_opacities(self, test_presentation):
        """Test both fill and line opacity together."""
        with PowerPointAgent(test_presentation) as agent:
            agent.open(test_presentation)
            
            result = agent.add_shape(
                slide_index=0,
                shape_type="rounded_rectangle",
                position={"left": "20%", "top": "20%"},
                size={"width": "60%", "height": "60%"},
                fill_color="#0070C0",
                fill_opacity=0.3,
                line_color="#ED7D31",
                line_opacity=0.7,
                line_width=3.0
            )
            
            agent.save()
        
        assert result["styling"]["fill_opacity"] == 0.3
        assert result["styling"]["fill_opacity_applied"] == True
        assert result["styling"]["line_opacity"] == 0.7
        assert result["styling"]["line_opacity_applied"] == True
    
    def test_add_shape_opacity_validation_fill_too_high(self, test_presentation):
        """Test that fill_opacity > 1.0 raises ValueError."""
        with PowerPointAgent(test_presentation) as agent:
            agent.open(test_presentation)
            
            with pytest.raises(ValueError) as excinfo:
                agent.add_shape(
                    slide_index=0,
                    shape_type="rectangle",
                    position={"left": "10%", "top": "10%"},
                    size={"width": "20%", "height": "20%"},
                    fill_color="#FF0000",
                    fill_opacity=1.5
                )
            
            assert "fill_opacity must be between 0.0 and 1.0" in str(excinfo.value)
    
    def test_add_shape_opacity_validation_fill_negative(self, test_presentation):
        """Test that fill_opacity < 0.0 raises ValueError."""
        with PowerPointAgent(test_presentation) as agent:
            agent.open(test_presentation)
            
            with pytest.raises(ValueError) as excinfo:
                agent.add_shape(
                    slide_index=0,
                    shape_type="rectangle",
                    position={"left": "10%", "top": "10%"},
                    size={"width": "20%", "height": "20%"},
                    fill_color="#FF0000",
                    fill_opacity=-0.5
                )
            
            assert "fill_opacity must be between 0.0 and 1.0" in str(excinfo.value)
    
    def test_add_shape_opacity_validation_line_too_high(self, test_presentation):
        """Test that line_opacity > 1.0 raises ValueError."""
        with PowerPointAgent(test_presentation) as agent:
            agent.open(test_presentation)
            
            with pytest.raises(ValueError) as excinfo:
                agent.add_shape(
                    slide_index=0,
                    shape_type="rectangle",
                    position={"left": "10%", "top": "10%"},
                    size={"width": "20%", "height": "20%"},
                    line_color="#000000",
                    line_opacity=2.0
                )
            
            assert "line_opacity must be between 0.0 and 1.0" in str(excinfo.value)
    
    def test_add_shape_no_color_no_opacity_applied(self, test_presentation):
        """Test that opacity is not applied when no color is specified."""
        with PowerPointAgent(test_presentation) as agent:
            agent.open(test_presentation)
            
            result = agent.add_shape(
                slide_index=0,
                shape_type="rectangle",
                position={"left": "10%", "top": "10%"},
                size={"width": "20%", "height": "20%"},
                fill_opacity=0.5  # Opacity specified but no color
            )
            
            agent.save()
        
        # fill_opacity_applied should be False because no fill_color
        assert result["styling"]["fill_opacity_applied"] == False


# ============================================================================
# TEST: format_shape() OPACITY FUNCTIONALITY
# ============================================================================

class TestFormatShapeOpacity:
    """Tests for format_shape() opacity features."""
    
    def test_format_shape_fill_opacity(self, presentation_with_shape):
        """Test applying fill opacity to existing shape."""
        pptx_path, shape_index = presentation_with_shape
        
        with PowerPointAgent(pptx_path) as agent:
            agent.open(pptx_path)
            
            result = agent.format_shape(
                slide_index=0,
                shape_index=shape_index,
                fill_opacity=0.5
            )
            
            agent.save()
        
        assert "fill_opacity" in result["changes_applied"]
        assert result["changes_detail"]["fill_opacity"] == 0.5
        assert result["changes_detail"]["fill_opacity_applied"] == True
        assert result["success"] == True
    
    def test_format_shape_line_opacity(self, presentation_with_shape):
        """Test applying line opacity to existing shape."""
        pptx_path, shape_index = presentation_with_shape
        
        with PowerPointAgent(pptx_path) as agent:
            agent.open(pptx_path)
            
            result = agent.format_shape(
                slide_index=0,
                shape_index=shape_index,
                line_color="#000000",
                line_opacity=0.3
            )
            
            agent.save()
        
        assert "line_opacity" in result["changes_applied"]
        assert result["changes_detail"]["line_opacity"] == 0.3
        assert result["changes_detail"]["line_opacity_applied"] == True
    
    def test_format_shape_transparency_deprecated(self, presentation_with_shape):
        """Test that deprecated transparency parameter still works."""
        pptx_path, shape_index = presentation_with_shape
        
        with PowerPointAgent(pptx_path) as agent:
            agent.open(pptx_path)
            
            # Using old transparency parameter (0.8 transparency = 0.2 opacity)
            result = agent.format_shape(
                slide_index=0,
                shape_index=shape_index,
                transparency=0.8
            )
            
            agent.save()
        
        assert "transparency_converted_to_opacity" in result["changes_applied"]
        assert result["changes_detail"]["transparency_deprecated"] == True
        assert result["changes_detail"]["transparency_value"] == 0.8
        assert result["changes_detail"]["converted_opacity"] == pytest.approx(0.2)
        assert result["changes_detail"]["fill_opacity"] == pytest.approx(0.2)
    
    def test_format_shape_fill_opacity_overrides_transparency(self, presentation_with_shape):
        """Test that fill_opacity takes precedence over transparency."""
        pptx_path, shape_index = presentation_with_shape
        
        with PowerPointAgent(pptx_path) as agent:
            agent.open(pptx_path)
            
            # Both provided - fill_opacity should win
            result = agent.format_shape(
                slide_index=0,
                shape_index=shape_index,
                fill_opacity=0.7,
                transparency=0.5
            )
            
            agent.save()
        
        assert "transparency_ignored" in result["changes_applied"]
        assert result["changes_detail"]["transparency_ignored"] == True
        assert result["changes_detail"]["fill_opacity"] == 0.7
    
    def test_format_shape_color_and_opacity(self, presentation_with_shape):
        """Test applying color and opacity together."""
        pptx_path, shape_index = presentation_with_shape
        
        with PowerPointAgent(pptx_path) as agent:
            agent.open(pptx_path)
            
            result = agent.format_shape(
                slide_index=0,
                shape_index=shape_index,
                fill_color="#FF0000",
                fill_opacity=0.4,
                line_color="#00FF00",
                line_opacity=0.6,
                line_width=2.5
            )
            
            agent.save()
        
        assert "fill_color" in result["changes_applied"]
        assert "fill_opacity" in result["changes_applied"]
        assert "line_color" in result["changes_applied"]
        assert "line_opacity" in result["changes_applied"]
        assert "line_width" in result["changes_applied"]
        assert result["success"] == True
    
    def test_format_shape_opacity_validation(self, presentation_with_shape):
        """Test that invalid opacity raises ValueError."""
        pptx_path, shape_index = presentation_with_shape
        
        with PowerPointAgent(pptx_path) as agent:
            agent.open(pptx_path)
            
            with pytest.raises(ValueError):
                agent.format_shape(
                    slide_index=0,
                    shape_index=shape_index,
                    fill_opacity=1.5
                )
    
    def test_format_shape_full_opacity_reset(self, presentation_with_shape):
        """Test that opacity 1.0 resets to fully opaque."""
        pptx_path, shape_index = presentation_with_shape
        
        with PowerPointAgent(pptx_path) as agent:
            agent.open(pptx_path)
            
            result = agent.format_shape(
                slide_index=0,
                shape_index=shape_index,
                fill_opacity=1.0
            )
            
            agent.save()
        
        assert "fill_opacity_reset" in result["changes_applied"]
        assert result["changes_detail"]["fill_opacity"] == 1.0


# ============================================================================
# TEST: XML VERIFICATION (Advanced)
# ============================================================================

class TestOpacityXMLVerification:
    """Advanced tests that verify the actual XML structure."""
    
    def test_alpha_element_created(self, test_presentation):
        """Verify that alpha element is created in XML."""
        with PowerPointAgent(test_presentation) as agent:
            agent.open(test_presentation)
            
            result = agent.add_shape(
                slide_index=0,
                shape_type="rectangle",
                position={"left": "10%", "top": "10%"},
                size={"width": "20%", "height": "20%"},
                fill_color="#0070C0",
                fill_opacity=0.15
            )
            
            agent.save()
        
        # Re-open and check XML
        with PowerPointAgent(test_presentation) as agent:
            agent.open(test_presentation)
            
            shape = agent._get_shape(0, result["shape_index"])
            spPr = shape._sp.spPr
            
            # Navigate to alpha element
            from pptx.oxml.ns import qn
            solidFill = spPr.find(qn('a:solidFill'))
            assert solidFill is not None, "solidFill element not found"
            
            color_elem = solidFill.find(qn('a:srgbClr'))
            assert color_elem is not None, "srgbClr element not found"
            
            alpha_elem = color_elem.find(qn('a:alpha'))
            assert alpha_elem is not None, "alpha element not found"
            
            alpha_val = int(alpha_elem.get('val'))
            # 0.15 * 100000 = 15000
            assert alpha_val == 15000, f"Expected 15000, got {alpha_val}"


# ============================================================================
# TEST: INTEGRATION WITH CLI TOOL
# ============================================================================

class TestCLIToolIntegration:
    """Tests for ppt_add_shape.py CLI integration."""
    
    def test_cli_add_shape_with_overlay(self, test_presentation):
        """Test CLI tool with overlay preset."""
        import subprocess
        
        result = subprocess.run(
            [
                "uv", "run", "tools/ppt_add_shape.py",
                "--file", str(test_presentation),
                "--slide", "0",
                "--shape", "rectangle",
                "--overlay",
                "--fill-color", "#FFFFFF",
                "--json"
            ],
            capture_output=True,
            text=True
        )
        
        assert result.returncode == 0, f"CLI failed: {result.stderr}"
        
        output = json.loads(result.stdout)
        assert output["status"] in ["success", "warning"]
        assert output["styling"]["fill_opacity"] == 0.15
        assert output["is_overlay"] == True
    
    def test_cli_add_shape_with_explicit_opacity(self, test_presentation):
        """Test CLI tool with explicit opacity values."""
        import subprocess
        
        result = subprocess.run(
            [
                "uv", "run", "tools/ppt_add_shape.py",
                "--file", str(test_presentation),
                "--slide", "0",
                "--shape", "ellipse",
                "--position", '{"left":"20%","top":"20%"}',
                "--size", '{"width":"40%","height":"40%"}',
                "--fill-color", "#0070C0",
                "--fill-opacity", "0.5",
                "--line-color", "#000000",
                "--line-opacity", "0.8",
                "--json"
            ],
            capture_output=True,
            text=True
        )
        
        assert result.returncode == 0, f"CLI failed: {result.stderr}"
        
        output = json.loads(result.stdout)
        assert output["styling"]["fill_opacity"] == 0.5
        assert output["styling"]["line_opacity"] == 0.8


# ============================================================================
# TEST: EDGE CASES
# ============================================================================

class TestEdgeCases:
    """Edge case and boundary tests."""
    
    def test_opacity_boundary_values(self, test_presentation):
        """Test opacity at exact boundary values."""
        with PowerPointAgent(test_presentation) as agent:
            agent.open(test_presentation)
            
            # Test 0.0 (minimum valid)
            result_min = agent.add_shape(
                slide_index=0,
                shape_type="rectangle",
                position={"left": "0%", "top": "0%"},
                size={"width": "10%", "height": "10%"},
                fill_color="#FF0000",
                fill_opacity=0.0
            )
            assert result_min["styling"]["fill_opacity"] == 0.0
            
            # Test 1.0 (maximum valid)
            result_max = agent.add_shape(
                slide_index=0,
                shape_type="rectangle",
                position={"left": "10%", "top": "0%"},
                size={"width": "10%", "height": "10%"},
                fill_color="#00FF00",
                fill_opacity=1.0
            )
            assert result_max["styling"]["fill_opacity"] == 1.0
            
            agent.save()
    
    def test_opacity_very_small_value(self, test_presentation):
        """Test very small opacity value."""
        with PowerPointAgent(test_presentation) as agent:
            agent.open(test_presentation)
            
            result = agent.add_shape(
                slide_index=0,
                shape_type="rectangle",
                position={"left": "10%", "top": "10%"},
                size={"width": "20%", "height": "20%"},
                fill_color="#0000FF",
                fill_opacity=0.001
            )
            
            agent.save()
        
        assert result["styling"]["fill_opacity"] == 0.001
        assert result["styling"]["fill_opacity_applied"] == True
    
    def test_multiple_shapes_different_opacities(self, test_presentation):
        """Test adding multiple shapes with different opacities."""
        with PowerPointAgent(test_presentation) as agent:
            agent.open(test_presentation)
            
            shapes = []
            opacities = [0.1, 0.25, 0.5, 0.75, 0.9]
            
            for i, opacity in enumerate(opacities):
                result = agent.add_shape(
                    slide_index=0,
                    shape_type="rectangle",
                    position={"left": f"{i * 20}%", "top": "40%"},
                    size={"width": "15%", "height": "20%"},
                    fill_color="#0070C0",
                    fill_opacity=opacity
                )
                shapes.append(result)
            
            agent.save()
        
        for i, shape in enumerate(shapes):
            assert shape["styling"]["fill_opacity"] == opacities[i]
            assert shape["styling"]["fill_opacity_applied"] == True


# ============================================================================
# RUN TESTS
# ============================================================================

if __name__ == "__main__":
    pytest.main([__file__, "-v", "--tb=short"])
