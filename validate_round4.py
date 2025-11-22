import sys
import json
import subprocess
from pathlib import Path

def run_probe(args):
    cmd = ["uv", "run", "tools/ppt_capability_probe.py", "--file", "test_probe.pptx"] + args
    result = subprocess.run(cmd, capture_output=True, text=True)
    return result

def validate_round4():
    print("Validating Round 4 Suggestions...")
    
    # 1. Check Error JSON Structure
    print("\n1. Checking Error JSON Structure...")
    # Run with a non-existent file to trigger error
    result = run_probe(["--file", "non_existent.pptx", "--json"])
    if result.returncode != 0:
        try:
            error_json = json.loads(result.stdout)
            if "warnings" in error_json:
                print("✅ 'warnings' field present in error JSON.")
            else:
                print("❌ 'warnings' field missing from error JSON.")
        except json.JSONDecodeError:
            print(f"❌ Could not decode error JSON: {result.stdout}")
    else:
        print("❌ Expected error for non-existent file, got success.")

    # 2. Check Summary Output for Masters Section
    print("\n2. Checking Summary Output for Masters Section...")
    result = run_probe(["--summary"])
    if result.returncode == 0:
        if "Master Slides:" in result.stdout and "layouts" in result.stdout: 
            # Heuristic check for master section details
            print("✅ 'Master Slides' section found in summary.")
        else:
            print("❌ 'Master Slides' detailed section missing from summary (only total count might be there).")
    else:
        print(f"❌ Error running summary: {result.stderr}")

    # 3. Check for Duplicate Warnings
    print("\n3. Checking for Duplicate Warnings...")
    # We can't easily force duplicates without a specific file, but we can check if the code has dedup logic
    # For now, we'll just note this for implementation.
    print("ℹ️ Duplicate warning check requires code inspection/implementation.")

if __name__ == "__main__":
    validate_round4()
