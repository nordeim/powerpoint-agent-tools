#!/usr/bin/env python3
import sys
import json
import subprocess
from pathlib import Path

# Add project root to path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

def run_tool(args):
    cmd = [sys.executable, str(project_root / "tools" / "ppt_add_shape.py")] + args + ["--json"]
    result = subprocess.run(cmd, capture_output=True, text=True)
    try:
        return json.loads(result.stdout)
    except json.JSONDecodeError:
        print(f"Failed to decode JSON: {result.stdout}")
        return {"status": "error", "error": "Invalid JSON output"}

def verify_shape_validation():
    sample_pptx = project_root / "samples" / "sample.pptx"
    if not sample_pptx.exists():
        print("Sample file not found, skipping test")
        return

    print("Testing Shape Validation...")

    # 1. Valid Shape
    print("\n1. Testing Valid Shape...")
    res = run_tool([
        "--file", str(sample_pptx),
        "--slide", "0",
        "--shape", "rectangle",
        "--position", '{"left":"10%","top":"10%"}',
        "--size", '{"width":"20%","height":"20%"}',
        "--fill-color", "#0070C0"
    ])
    if res.get("status") == "success" and not res.get("warnings"):
        print("✅ Valid shape passed")
    else:
        print(f"❌ Valid shape failed: {res}")

    # 2. Off-slide Shape (Expect Warning)
    print("\n2. Testing Off-slide Shape...")
    res = run_tool([
        "--file", str(sample_pptx),
        "--slide", "0",
        "--shape", "rectangle",
        "--position", '{"left":"110%","top":"10%"}',
        "--size", '{"width":"20%","height":"20%"}'
    ])
    warnings = res.get("warnings", [])
    if any("outside slide bounds" in w for w in warnings):
        print("✅ Off-slide warning detected")
    else:
        print(f"❌ Off-slide warning missing: {res}")

    # 3. Tiny Shape (Expect Warning)
    print("\n3. Testing Tiny Shape...")
    res = run_tool([
        "--file", str(sample_pptx),
        "--slide", "0",
        "--shape", "rectangle",
        "--position", '{"left":"10%","top":"10%"}',
        "--size", '{"width":"0.5%","height":"0.5%"}'
    ])
    warnings = res.get("warnings", [])
    if any("extremely small" in w for w in warnings):
        print("✅ Tiny shape warning detected")
    else:
        print(f"❌ Tiny shape warning missing: {res}")

    # 4. Low Contrast (Expect Warning)
    print("\n4. Testing Low Contrast Shape...")
    res = run_tool([
        "--file", str(sample_pptx),
        "--slide", "0",
        "--shape", "rectangle",
        "--position", '{"left":"10%","top":"10%"}',
        "--size", '{"width":"20%","height":"20%"}',
        "--fill-color", "#FFFFFF" # White on White (assumed)
    ])
    warnings = res.get("warnings", [])
    if any("contrast" in w.lower() for w in warnings):
        print("✅ Low contrast warning detected")
    else:
        print(f"❌ Low contrast warning missing: {res}")

if __name__ == "__main__":
    verify_shape_validation()
