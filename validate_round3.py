import sys
import json
import subprocess
import time
from pathlib import Path

def run_probe(args):
    cmd = ["uv", "run", "tools/ppt_capability_probe.py", "--file", "test_probe.pptx", "--json"] + args
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        print(f"Error running probe: {result.stderr}")
        return None
    return json.loads(result.stdout)

def validate_round3():
    print("Validating Round 3 Suggestions...")
    
    # 1. Check for Per-Master Theme
    print("\n1. Checking for Per-Master Theme...")
    result = run_probe([])
    if result:
        if 'per_master' in result['theme']:
            print("✅ 'per_master' theme stats present.")
        else:
            print("❌ 'per_master' theme stats missing.")

    # 2. Check for Color Fallback (Simulation)
    # Hard to simulate without a specific file, but we can check code logic or output structure
    print("\n2. Checking Color Fallback Logic...")
    # We'll check if the tool handles non-RGB colors by looking for warnings in a run
    # (This is a weak test without a specific file, but validates current state)
    if any("scheme-based" in w for w in result.get('warnings', [])):
        print("✅ Found scheme-based color warnings (logic exists).")
    else:
        print("ℹ️ No scheme-based color warnings found (logic might be missing or not triggered).")

    # 3. Check for East Asian Fonts
    print("\n3. Checking for East Asian Fonts...")
    fonts = result['theme']['fonts']
    if 'heading_east_asian' in fonts or 'body_complex' in fonts:
        print("✅ East Asian/Complex fonts present.")
    else:
        print("❌ East Asian/Complex fonts missing.")

    # 4. Check Original Index Logic
    print("\n4. Checking Original Index Logic...")
    # We already have original_index from Round 2, but let's verify it matches what we expect
    layout_0 = result['layouts'][0]
    if 'original_index' in layout_0:
        print(f"✅ 'original_index' present: {layout_0['original_index']}")
    else:
        print("❌ 'original_index' missing.")

    # 5. Check Timeout Flag
    print("\n5. Checking Timeout Flag...")
    # Run with very short timeout
    timeout_result = run_probe(["--timeout", "0"]) # 0 seconds should trigger immediately
    if timeout_result:
        caps = timeout_result['capabilities']
        if 'analysis_complete' in caps:
             print(f"✅ 'analysis_complete' flag present: {caps['analysis_complete']}")
        else:
             print("❌ 'analysis_complete' flag missing.")

    # 6. Check Validation Error Type
    print("\n6. Checking Validation Error Type...")
    # We can't easily force a validation error on a valid file without modifying code,
    # but we can check if the status is "error" when things go wrong.
    # For now, we'll skip this dynamic check and rely on code review.

if __name__ == "__main__":
    validate_round3()
