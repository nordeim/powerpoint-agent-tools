import sys
import json
import subprocess
from pathlib import Path

def run_probe(args):
    cmd = ["uv", "run", "tools/ppt_capability_probe.py"] + args
    result = subprocess.run(cmd, capture_output=True, text=True)
    return result

def validate_round6():
    print("Validating Round 6 Suggestions...")
    
    # 1. Check for Duplicate per_master (Output Check)
    print("\n1. Checking for Duplicate per_master (Output Check)...")
    # This is hard to check via JSON output because json.dump will deduplicate keys.
    # But we can check if the code has it twice. We'll rely on code review for this.
    print("ℹ️ Duplicate per_master check requires code inspection (will fix).")

    # 2. Check for original_index in Summary
    print("\n2. Checking for original_index in Summary...")
    # We need to force max_layouts to see if it uses index vs original_index
    result = run_probe(["--file", "test_probe.pptx", "--summary", "--max-layouts", "1"])
    if result.returncode == 0:
        # If original_index is used, it should match the index in the file.
        # Since we are just taking the first one, index and original_index are likely 0.
        # But we can check if the code uses the key.
        print("ℹ️ Summary output check is heuristic. Will implement fix.")
    else:
        print(f"❌ Error running summary: {result.stderr}")

    # 3. Check for metadata.timeout_seconds
    print("\n3. Checking for metadata.timeout_seconds...")
    result = run_probe(["--file", "test_probe.pptx", "--json", "--timeout", "10"])
    if result.returncode == 0:
        try:
            data = json.loads(result.stdout)
            if 'timeout_seconds' in data['metadata']:
                print(f"✅ 'metadata.timeout_seconds' present: {data['metadata']['timeout_seconds']}")
            else:
                print("❌ 'metadata.timeout_seconds' missing.")
        except:
            print("❌ JSON decode failed")
    else:
        print(f"❌ Error running probe: {result.stderr}")

    # 4. Check for Consolidated Non-RGB Warning
    print("\n4. Checking for Consolidated Non-RGB Warning...")
    # We don't have a file with scheme colors easily available to force this, 
    # but we can check if the warning is present in the list of potential warnings if we mock it.
    # For now, we'll mark as missing.
    print("❌ Consolidated warning logic missing (expected).")

if __name__ == "__main__":
    validate_round6()
