import sys
import json
import subprocess
from pathlib import Path

def run_probe(args):
    cmd = ["uv", "run", "tools/ppt_capability_probe.py"] + args
    result = subprocess.run(cmd, capture_output=True, text=True)
    return result

def validate_round5():
    print("Validating Round 5 Suggestions...")
    
    result = run_probe(["--file", "test_probe.pptx", "--json", "--deep"])
    if result.returncode != 0:
        print(f"❌ Error running probe: {result.stderr}")
        return

    try:
        data = json.loads(result.stdout)
        
        # 1. Check for EMU positions
        print("\n1. Checking for EMU positions...")
        layout_0 = data['layouts'][0]
        if 'placeholders' in layout_0 and len(layout_0['placeholders']) > 0:
            ph = layout_0['placeholders'][0]
            if 'position_emu' in ph:
                print("✅ 'position_emu' present.")
            else:
                print("❌ 'position_emu' missing.")
        else:
            print("ℹ️ No placeholders found in first layout to check.")

        # 2. Check for Instantiation Complete
        print("\n2. Checking for Instantiation Complete flag...")
        if 'instantiation_complete' in layout_0:
            print("✅ 'instantiation_complete' present.")
        else:
            print("❌ 'instantiation_complete' missing.")

        # 3. Check for Capability Strategy Hints
        print("\n3. Checking for Capability Strategy Hints...")
        caps = data['capabilities']
        if 'footer_support_mode' in caps:
            print("✅ 'footer_support_mode' present.")
        else:
            print("❌ 'footer_support_mode' missing.")

        # 4. Check for Per-Master Theme (should be there from Round 3)
        print("\n4. Checking for Per-Master Theme...")
        if 'per_master' in data['theme']:
            print("✅ 'theme.per_master' present.")
        else:
            print("❌ 'theme.per_master' missing.")

    except json.JSONDecodeError:
        print(f"❌ Could not decode JSON: {result.stdout}")

if __name__ == "__main__":
    validate_round5()
