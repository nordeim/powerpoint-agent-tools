#!/usr/bin/env python3
import sys
import json
import subprocess
from pathlib import Path

# Add project root to path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

def verify_version():
    tool_path = project_root / "tools" / "ppt_capability_probe.py"
    sample_path = project_root / "samples" / "sample.pptx"
    
    print(f"Running probe on {sample_path} to check versions...")
    
    try:
        result = subprocess.run(
            [sys.executable, str(tool_path), "--file", str(sample_path), "--json"],
            capture_output=True,
            text=True,
            check=True
        )
        output = json.loads(result.stdout)
        
        metadata = output.get("metadata", {})
        tool_version = metadata.get("tool_version")
        schema_version = metadata.get("schema_version")
        
        print(f"Detected Tool Version: {tool_version}")
        print(f"Detected Schema Version: {schema_version}")
        
        expected_tool_version = "1.1.1"
        expected_schema_version = "capability_probe.v1.1.1"
        
        errors = []
        if tool_version != expected_tool_version:
            errors.append(f"Tool version mismatch: expected {expected_tool_version}, got {tool_version}")
            
        if schema_version != expected_schema_version:
            errors.append(f"Schema version mismatch: expected {expected_schema_version}, got {schema_version}")
            
        if errors:
            print("❌ Version verification failed:")
            for error in errors:
                print(f"  - {error}")
            sys.exit(1)
            
        print("✅ Version verification passed!")
        
    except subprocess.CalledProcessError as e:
        print(f"❌ Probe failed: {e.stderr}")
        sys.exit(1)
    except json.JSONDecodeError:
        print(f"❌ Invalid JSON output: {result.stdout}")
        sys.exit(1)
    except Exception as e:
        print(f"❌ Unexpected error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    verify_version()
