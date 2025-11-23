#!/usr/bin/env python3
import sys
import unittest
from unittest.mock import MagicMock
from pathlib import Path

# Add project root to path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from tools.ppt_capability_probe import extract_theme_colors

class TestThemeExtraction(unittest.TestCase):
    def test_missing_theme_warning(self):
        """Test that a warning is issued when the theme object is missing."""
        # Mock a slide master with no theme attribute
        mock_master = MagicMock()
        del mock_master.theme # Ensure attribute raises AttributeError or is missing
        # Alternatively, just don't set it if it's a MagicMock, but safer to be explicit about what we test.
        # The code uses getattr(obj, 'theme', None), so we just need to ensure it returns None.
        # MagicMock by default creates attributes on access.
        # Let's use a plain class or configure the mock.
        
        class MockMaster:
            pass
            
        master = MockMaster()
        warnings = []
        
        colors = extract_theme_colors(master, warnings)
        
        self.assertEqual(colors, {}, "Colors should be empty")
        
        # This assertion is expected to FAIL before the fix
        self.assertTrue(len(warnings) > 0, "Should have warnings for missing theme")
        self.assertIn("Theme object unavailable", warnings[0] if warnings else "", "Warning message mismatch")

    def test_missing_color_scheme_warning(self):
        """Test that a warning is issued when the color scheme is missing."""
        class MockTheme:
            pass
            
        class MockMaster:
            theme = MockTheme()
            
        master = MockMaster()
        warnings = []
        
        colors = extract_theme_colors(master, warnings)
        
        self.assertEqual(colors, {}, "Colors should be empty")
        
        # This assertion is expected to FAIL before the fix
        self.assertTrue(len(warnings) > 0, "Should have warnings for missing color scheme")
        self.assertIn("Theme color scheme unavailable", warnings[0] if warnings else "", "Warning message mismatch")

if __name__ == "__main__":
    unittest.main()
