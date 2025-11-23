import sys
import json
import subprocess
from pathlib import Path

def run_probe(args):
    cmd = ["uv", "run", "tools/ppt_capability_probe.py"] + args
    result = subprocess.run(cmd, capture_output=True, text=True)
    return result

def validate_round4():
    print("Validating Round 4 Suggestions...")
    
    # Run probe on Presentation.pptx (Essential Mode)
    result = run_probe(["--file", "Presentation.pptx", "--json"])
    if result.returncode != 0:
        print(f"❌ Error running probe: {result.stderr}")
        return

    try:
        data = json.loads(result.stdout)
        
        # 1. Check for warnings_count in metadata
        print("\n1. Checking for warnings_count in metadata...")
        meta = data['metadata']
        if 'warnings_count' in meta:
            print(f"✅ warnings_count present: {meta['warnings_count']}")
        else:
            print("❌ warnings_count MISSING in metadata!")

        # 2. Check for Scheme Color Warning Logic
        # We can't easily trigger the "schemeColor" warning with Presentation.pptx (it has no theme)
        # But we can check that we don't have duplicate warnings about it.
        print("\n2. Checking for Duplicate Color Warnings...")
        warnings = data.get('warnings', [])
        print(f"Current Warnings: {warnings}")
        
        color_warnings = [w for w in warnings if "color" in w.lower()]
        if len(color_warnings) > 1:
            # If we have "Theme color scheme unavailable" AND "Theme colors include scheme references", that might be okay depending on logic
            # But the suggestion is to consolidate.
            # In Presentation.pptx, we expect "Theme color scheme unavailable or empty"
            print(f"ℹ️ Color warnings found: {color_warnings}")
        elif len(color_warnings) == 1:
             print(f"✅ Single color warning found: {color_warnings[0]}")
        else:
            print("ℹ️ No color warnings found.")

    except json.JSONDecodeError:
        print(f"❌ Could not decode JSON: {result.stdout}")

if __name__ == "__main__":
    validate_round4()
