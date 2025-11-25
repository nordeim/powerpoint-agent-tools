#!/usr/bin/env bash
set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ROOT_DIR="$(dirname "$SCRIPT_DIR")"
WRAPPER="$ROOT_DIR/scripts/probe_wrapper.sh"

# Mock ppt_capability_probe.py
function create_mock_probe_success {
    cat <<EOF > "$ROOT_DIR/ppt_capability_probe.py"
#!/usr/bin/env python3
import json
print(json.dumps({
    "tool_name": "ppt_capability_probe",
    "tool_version": "1.0.0",
    "schema_version": "1.0.0",
    "file": "test.pptx",
    "probe_timestamp": "2025-01-01T00:00:00Z",
    "capabilities": {
        "can_read": True,
        "can_write": True,
        "layouts": ["Title"],
        "slide_dimensions": {"width_pt": 720, "height_pt": 540}
    }
}))
EOF
    chmod +x "$ROOT_DIR/ppt_capability_probe.py"
}

function create_mock_probe_fail {
    cat <<EOF > "$ROOT_DIR/ppt_capability_probe.py"
#!/usr/bin/env python3
import sys
sys.exit(1)
EOF
    chmod +x "$ROOT_DIR/ppt_capability_probe.py"
}

# Mock fallback tools
function create_mock_fallbacks {
    cat <<EOF > "$ROOT_DIR/ppt_get_info.py"
#!/usr/bin/env python3
import json
print(json.dumps({"tool_name": "ppt_get_info", "slide_count": 1}))
EOF
    chmod +x "$ROOT_DIR/ppt_get_info.py"

    cat <<EOF > "$ROOT_DIR/ppt_get_slide_info.py"
#!/usr/bin/env python3
import json
print(json.dumps({"tool_name": "ppt_get_slide_info", "index": 0}))
EOF
    chmod +x "$ROOT_DIR/ppt_get_slide_info.py"
}

# Setup
export PATH="$ROOT_DIR:$PATH"
TEST_FILE="/tmp/test_deck.pptx"
touch "$TEST_FILE"

echo "Running Test 1: Probe Success"
create_mock_probe_success
"$WRAPPER" "$TEST_FILE" > /dev/null
echo "PASS"

echo "Running Test 2: Probe Failure -> Fallback Success"
create_mock_probe_fail
create_mock_fallbacks
"$WRAPPER" "$TEST_FILE" > /dev/null
echo "PASS"

# Cleanup
rm "$TEST_FILE"
rm "$ROOT_DIR/ppt_capability_probe.py"
rm "$ROOT_DIR/ppt_get_info.py"
rm "$ROOT_DIR/ppt_get_slide_info.py"
