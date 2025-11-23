#!/usr/bin/env python3
import sys
import json
import subprocess
from pathlib import Path

# Add project root to path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

def run_tool(args):
    cmd = [sys.executable, str(project_root / "tools" / "ppt_add_table.py")] + args + ["--json"]
    result = subprocess.run(cmd, capture_output=True, text=True)
    try:
        return json.loads(result.stdout)
    except json.JSONDecodeError:
        print(f"Failed to decode JSON: {result.stdout}")
        return {"status": "error", "error": "Invalid JSON output"}

def verify_table_validation():
    sample_pptx = project_root / "samples" / "sample.pptx"
    if not sample_pptx.exists():
        print("Sample file not found, skipping test")
        return

    print("Testing Table Validation...")

    # 1. Valid Table
    print("\n1. Testing Valid Table...")
    res = run_tool([
        "--file", str(sample_pptx),
        "--slide", "0",
        "--rows", "3",
        "--cols", "3",
        "--position", '{"left":"10%","top":"10%"}',
        "--size", '{"width":"50%","height":"30%"}'
    ])
    if res.get("status") == "success" and not res.get("warnings"):
        print("✅ Valid table passed")
    else:
        print(f"❌ Valid table failed: {res}")

    # 2. Off-slide Table (Expect Warning)
    print("\n2. Testing Off-slide Table...")
    res = run_tool([
        "--file", str(sample_pptx),
        "--slide", "0",
        "--rows", "3",
        "--cols", "3",
        "--position", '{"left":"110%","top":"10%"}',
        "--size", '{"width":"50%","height":"30%"}'
    ])
    warnings = res.get("warnings", [])
    if any("outside slide bounds" in w for w in warnings):
        print("✅ Off-slide warning detected")
    else:
        print(f"❌ Off-slide warning missing: {res}")

    # 3. Tiny Table (Expect Warning)
    print("\n3. Testing Tiny Table...")
    res = run_tool([
        "--file", str(sample_pptx),
        "--slide", "0",
        "--rows", "5",
        "--cols", "5",
        "--position", '{"left":"10%","top":"10%"}',
        "--size", '{"width":"5%","height":"5%"}'
    ])
    warnings = res.get("warnings", [])
    if any("very small" in w for w in warnings):
        print("✅ Tiny table warning detected")
    else:
        print(f"❌ Tiny table warning missing: {res}")

if __name__ == "__main__":
    verify_table_validation()
