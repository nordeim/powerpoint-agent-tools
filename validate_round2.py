import sys
import json
import subprocess
from pathlib import Path

def run_probe(args):
    cmd = ["uv", "run", "tools/ppt_capability_probe.py"] + args
    result = subprocess.run(cmd, capture_output=True, text=True)
    return result

def validate_round2():
    print("Validating Round 2 Suggestions...")
    
    # Run probe on test file (Essential Mode)
    result = run_probe(["--file", "Presentation.pptx", "--json"])
    if result.returncode != 0:
        print(f"❌ Error running probe: {result.stderr}")
        return

    try:
        data = json.loads(result.stdout)
        
        # 1. Check for Placeholder Map & Types in Essential Mode
        print("\n1. Checking for Placeholder Map & Types (Essential Mode)...")
        layout_0 = data['layouts'][0]
        
        if 'placeholder_types' in layout_0:
            print(f"✅ placeholder_types present: {len(layout_0['placeholder_types'])} types")
        else:
            print("❌ placeholder_types MISSING in essential mode!")
            
        if 'placeholder_map' in layout_0:
            print(f"✅ placeholder_map present: {layout_0['placeholder_map']}")
        else:
            print("❌ placeholder_map MISSING in essential mode!")

        # 2. Check for Standardized Capability Layout References
        print("\n2. Checking for Standardized Capability Layout References...")
        caps = data['capabilities']
        # We need a layout that actually has capabilities to check this
        # In test_probe.pptx (from previous logs), it seemed to have footer/slide# support?
        # Wait, the reviewer said "Capability arrays empty" for the log they reviewed?
        # Let's check what we get.
        
        checked_any = False
        for cap_list_name in ['layouts_with_footer', 'layouts_with_slide_number', 'layouts_with_date']:
            if cap_list_name in caps and len(caps[cap_list_name]) > 0:
                ref = caps[cap_list_name][0]
                if 'original_index' in ref and 'master_index' in ref:
                    print(f"✅ {cap_list_name} references standardized.")
                else:
                    print(f"❌ {cap_list_name} references missing fields: {list(ref.keys())}")
                checked_any = True
        
        if not checked_any:
            print("ℹ️ No capabilities detected to check references.")

        # 3. Check for Theme Warnings
        print("\n3. Checking for Theme Warnings...")
        warnings = data.get('warnings', [])
        print(f"Warnings found: {warnings}")
        # We can't strictly enforce the specific warning unless we know the file has scheme colors but no RGB
        # But we can check if the logic seems to be running (e.g. no crash)

    except json.JSONDecodeError:
        print(f"❌ Could not decode JSON: {result.stdout}")

if __name__ == "__main__":
    validate_round2()
