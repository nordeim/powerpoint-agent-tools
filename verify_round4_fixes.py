import sys
import json
import subprocess
from pathlib import Path

def run_probe(args):
    # Don't hardcode file here, let args decide
    cmd = ["uv", "run", "tools/ppt_capability_probe.py"] + args
    result = subprocess.run(cmd, capture_output=True, text=True)
    return result

def verify_round4_fixes():
    print("Verifying Round 4 Fixes...")
    
    # 1. Check Error JSON Structure (Warnings)
    print("\n1. Checking Error JSON Structure (Warnings)...")
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
    result = run_probe(["--file", "test_probe.pptx", "--summary"])
    if result.returncode == 0:
        if "Master Slides:" in result.stdout and "Master 0:" in result.stdout: 
            print("✅ 'Master Slides' detailed section found in summary.")
        else:
            print("❌ 'Master Slides' detailed section missing from summary.")
    else:
        print(f"❌ Error running summary: {result.stderr}")

if __name__ == "__main__":
    verify_round4_fixes()
