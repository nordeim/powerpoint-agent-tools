import sys
import json
import subprocess
from pathlib import Path

def run_probe(args):
    cmd = ["uv", "run", "tools/ppt_capability_probe.py"] + args
    result = subprocess.run(cmd, capture_output=True, text=True)
    return result

def validate_round3():
    print("Validating Round 3 Suggestions...")
    
    # Run probe on Presentation.pptx (Essential Mode)
    result = run_probe(["--file", "Presentation.pptx", "--json"])
    if result.returncode != 0:
        print(f"❌ Error running probe: {result.stderr}")
        return

    try:
        data = json.loads(result.stdout)
        
        # 1. Check for analysis_mode in metadata
        print("\n1. Checking for analysis_mode in metadata...")
        meta = data['metadata']
        if 'analysis_mode' in meta:
            print(f"✅ analysis_mode present: {meta['analysis_mode']}")
        else:
            print("❌ analysis_mode MISSING in metadata!")

        # 2. Check for Consolidated Font Warnings
        print("\n2. Checking for Consolidated Font Warnings...")
        warnings = data.get('warnings', [])
        print(f"Current Warnings: {warnings}")
        
        font_warnings = [w for w in warnings if "font" in w.lower()]
        if len(font_warnings) > 1:
            print(f"❌ Multiple font warnings found: {len(font_warnings)}. Should be consolidated.")
        elif len(font_warnings) == 1:
            print(f"✅ Single font warning found: {font_warnings[0]}")
        else:
            print("ℹ️ No font warnings found (unexpected for this test file).")

        # 3. Check for Scheme Color Warning Logic (Code Inspection required for full validation, but we can check output)
        # Since Presentation.pptx has no theme, we expect "Theme color scheme unavailable or empty"
        # We can't easily test the "schemeColor" warning without a specific file, but we can verify the code structure later.

    except json.JSONDecodeError:
        print(f"❌ Could not decode JSON: {result.stdout}")

if __name__ == "__main__":
    validate_round3()
