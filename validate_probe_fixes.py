import sys
import os
import json
import subprocess
from pptx.enum.shapes import PP_PLACEHOLDER

def test_enum_coverage():
    print("Testing Enum Coverage...")
    # Simulate the current implementation's known_types
    known_types = [
        'TITLE', 'BODY', 'CENTER_TITLE', 'SUBTITLE', 'DATE', 
        'SLIDE_NUMBER', 'FOOTER', 'HEADER', 'OBJECT', 'CHART',
        'TABLE', 'CLIP_ART', 'PICTURE', 'MEDIA_CLIP', 'ORG_CHART'
    ]
    
    missing = []
    for name in dir(PP_PLACEHOLDER):
        if name.isupper() and name not in known_types:
            missing.append(name)
    
    if missing:
        print(f"❌ Current implementation misses these enum members: {missing}")
    else:
        print("✅ Current implementation covers all enum members.")

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
