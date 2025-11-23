#!/usr/bin/env python3
import sys
import json
import subprocess
from pathlib import Path

# Add project root to path
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))

from core.strict_validator import validate_against_schema

def verify_probe():
    tool_path = project_root / "tools" / "ppt_capability_probe.py"
    schema_path = project_root / "schemas" / "capability_probe.v1.1.1.schema.json"
    sample_path = project_root / "samples" / "sample.pptx"
    
    print(f"Running probe on {sample_path}...")
    
    try:
        result = subprocess.run(
            [sys.executable, str(tool_path), "--file", str(sample_path), "--json"],
            capture_output=True,
            text=True,
            check=True
        )
        output = json.loads(result.stdout)
        
        print("Probe successful. Validating against schema...")
        validate_against_schema(output, str(schema_path))
        print("✅ Schema validation passed!")
        
    except subprocess.CalledProcessError as e:
        print(f"❌ Probe failed: {e.stderr}")
        sys.exit(1)
    except json.JSONDecodeError:
        print(f"❌ Invalid JSON output: {result.stdout}")
        sys.exit(1)
    except ValueError as e:
        print(f"❌ Schema validation failed:\n{e}")
        sys.exit(1)
    except Exception as e:
        print(f"❌ Unexpected error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    verify_probe()
