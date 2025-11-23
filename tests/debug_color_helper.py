#!/usr/bin/env python3
import sys
from pathlib import Path

# Add project root to path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

try:
    from core.powerpoint_agent_core import ColorHelper, RGBColor
    
    print("Testing ColorHelper...")
    
    # Test 1: from_hex
    try:
        c1 = ColorHelper.from_hex("#FFFFFF")
        print(f"ColorHelper.from_hex('#FFFFFF') -> {c1} (type: {type(c1)})")
        print(f"Attributes: r={getattr(c1, 'r', 'N/A')}, g={getattr(c1, 'g', 'N/A')}, b={getattr(c1, 'b', 'N/A')}")
        print(f"String representation: '{str(c1)}'")
        print(f"Dir: {dir(c1)}")
    except Exception as e:
        print(f"Error in from_hex: {e}")

    # Test 2: luminance
    try:
        lum = ColorHelper.luminance(c1)
        print(f"Luminance: {lum}")
    except Exception as e:
        print(f"Error in luminance: {e}")

except ImportError as e:
    print(f"ImportError: {e}")
except Exception as e:
    print(f"Unexpected error: {e}")
