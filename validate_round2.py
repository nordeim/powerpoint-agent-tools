import sys
import json
import subprocess
from pathlib import Path

def run_probe(args):
    cmd = ["uv", "run", "tools/ppt_capability_probe.py", "--file", "test_probe.pptx", "--json"] + args
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        print(f"Error running probe: {result.stderr}")
        return None
    return json.loads(result.stdout)

def validate_round2():
    print("Validating Round 2 Suggestions...")
    
    # 1. Test Max Layouts Index Shifting
    print("\n1. Testing Max Layouts Index Shifting...")
    full_result = run_probe([])
    limited_result = run_probe(["--max-layouts", "2"])
    
    if full_result and limited_result:
        full_layout_names = [l['name'] for l in full_result['layouts']]
        limited_layout_names = [l['name'] for l in limited_result['layouts']]
        
        print(f"Full Layouts (First 3): {full_layout_names[:3]}")
        print(f"Limited Layouts: {limited_layout_names}")
        
        # Check indices
        full_indices = [l['index'] for l in full_result['layouts'][:2]]
        limited_indices = [l['index'] for l in limited_result['layouts']]
        
        print(f"Full Indices: {full_indices}")
        print(f"Limited Indices: {limited_indices}")
        
        if full_indices == limited_indices:
            print("Indices match (expected for simple slicing).")
        else:
            print("Indices shifted!")

    # 2. Check for Per-Master Stats
    print("\n2. Checking for Per-Master Stats...")
    if 'per_master' not in full_result['capabilities']:
        print("❌ 'per_master' stats missing from capabilities.")
    else:
        print("✅ 'per_master' stats present.")

    # 3. Check for Aspect Ratio Float
    print("\n3. Checking for Aspect Ratio Float...")
    if 'aspect_ratio_float' not in full_result['slide_dimensions']:
        print("❌ 'aspect_ratio_float' missing from slide_dimensions.")
    else:
        print("✅ 'aspect_ratio_float' present.")

    # 4. Check for Schema Version
    print("\n4. Checking for Schema Version...")
    if 'schema_version' not in full_result['metadata']:
        print("❌ 'schema_version' missing from metadata.")
    else:
        print("✅ 'schema_version' present.")

if __name__ == "__main__":
    validate_round2()
