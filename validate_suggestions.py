import sys
import json
import subprocess
from pathlib import Path

def run_probe(args):
    cmd = ["uv", "run", "tools/ppt_capability_probe.py"] + args
    result = subprocess.run(cmd, capture_output=True, text=True)
    return result

def validate_suggestions():
    print("Validating Improvement Suggestions...")
    
    # Run probe on test file
    result = run_probe(["--file", "test_probe.pptx", "--json"])
    if result.returncode != 0:
        print(f"❌ Error running probe: {result.stderr}")
        return

    try:
        data = json.loads(result.stdout)
        
        # 1. Check for Standardized Capability Layout References
        print("\n1. Checking for Standardized Capability Layout References...")
        caps = data['capabilities']
        if 'layouts_with_footer' in caps and len(caps['layouts_with_footer']) > 0:
            ref = caps['layouts_with_footer'][0]
            if 'original_index' in ref and 'master_index' in ref:
                print("✅ Capability layout references standardized.")
            else:
                print(f"❌ Capability layout references missing fields: {list(ref.keys())}")
        else:
            print("ℹ️ No footer layouts to check.")

        # 2. Check for Symmetric Recommendations
        print("\n2. Checking for Symmetric Recommendations...")
        recs = caps.get('recommendations', [])
        has_slide_num_rec = any("Slide number placeholders available" in r for r in recs)
        has_date_rec = any("Date placeholders available" in r for r in recs)
        
        if caps['has_slide_number_placeholders'] and not has_slide_num_rec:
            print("❌ Missing slide number recommendation enumeration.")
        elif caps['has_slide_number_placeholders']:
             print("✅ Slide number recommendation present.")
             
        if caps['has_date_placeholders'] and not has_date_rec:
            print("❌ Missing date recommendation enumeration.")
        elif caps['has_date_placeholders']:
            print("✅ Date recommendation present.")

        # 3. Check for Metadata Audit Fields
        print("\n3. Checking for Metadata Audit Fields...")
        meta = data['metadata']
        if 'layout_count_total' in meta and 'layout_count_analyzed' in meta:
            print("✅ Metadata audit fields present.")
        else:
            print("❌ Metadata audit fields missing.")

        # 4. Check for Placeholder Map (Essential Mode)
        print("\n4. Checking for Placeholder Map (Essential Mode)...")
        layout_0 = data['layouts'][0]
        if 'placeholder_map' in layout_0:
            print("✅ Placeholder map present.")
        else:
            print("❌ Placeholder map missing.")

    except json.JSONDecodeError:
        print(f"❌ Could not decode JSON: {result.stdout}")

if __name__ == "__main__":
    validate_suggestions()
