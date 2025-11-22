import sys
import os
import json
import subprocess
from pptx.enum.shapes import PP_PLACEHOLDER

def test_enum_coverage():
    print("Testing Enum Coverage...")
    from tools.ppt_capability_probe import PLACEHOLDER_TYPE_MAP
    
    missing = []
    for name in dir(PP_PLACEHOLDER):
        if name.isupper():
            # Get the value to check if it's in the map
            val = getattr(PP_PLACEHOLDER, name)
            code = val.value if hasattr(val, 'value') else val
            
            if isinstance(code, int) and code not in PLACEHOLDER_TYPE_MAP:
                missing.append(name)
    
    if missing:
        print(f"❌ Tool implementation misses these enum members: {missing}")
    else:
        print("✅ Tool implementation covers all enum members.")

def test_mutual_exclusivity():
    print("\nTesting Mutual Exclusivity...")
    # Create a dummy pptx
    subprocess.run(["uv", "run", "tools/ppt_create_new.py", "--output", "test_probe.pptx", "--json"], check=True, capture_output=True)
    
    # Run with both flags
    result = subprocess.run(
        ["uv", "run", "tools/ppt_capability_probe.py", "--file", "test_probe.pptx", "--json", "--summary"],
        capture_output=True,
        text=True
    )
    
    # If it succeeds (exit code 0) and prints JSON, the check failed to block it.
    if result.returncode == 0:
        try:
            json.loads(result.stdout)
            print("❌ Mutual exclusivity check FAILED (Tool ran and output JSON).")
        except:
            print("❓ Tool ran but output was not JSON. Check output.")
    else:
        print("✅ Mutual exclusivity check PASSED (Tool failed as expected).")

if __name__ == "__main__":
    test_enum_coverage()
    test_mutual_exclusivity()
