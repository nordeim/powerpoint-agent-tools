#!/usr/bin/env bash
# ppt_capability_probe_v1.1.1_tests.sh
set -euo pipefail

TOOL="tools/ppt_capability_probe.py"
FILE="${1:-sample.pptx}"

echo "═══════════════════════════════════════════════════════════════"
echo "ppt_capability_probe.py v1.1.1 - Minimal Regression Test Suite"
echo "File: $FILE"
echo "═══════════════════════════════════════════════════════════════"

# Test 1: JSON Contract & version checks
echo "Test 1: JSON Contract & version"
uv run "$TOOL" --file "$FILE" --json \
  | jq '{status, meta: .metadata | {tool_version, schema_version, operation_id, duration_ms, timeout_seconds, layout_count_total, layout_count_analyzed}, keys_present: [.slide_dimensions != null, .layouts != null, .theme != null, .capabilities != null, .warnings != null, .info != null]}'

# Test 2: Atomic verification (MD5 before/after)
echo "Test 2: Atomic verification"
md5sum "$FILE" > /tmp/before.md5
uv run "$TOOL" --file "$FILE" --json > /tmp/probe.json
md5sum "$FILE" > /tmp/after.md5
if diff /tmp/before.md5 /tmp/after.md5 > /dev/null; then
  echo "✅ PASS: File unchanged (atomic read verified)"
else
  echo "❌ FAIL: File was modified!"
  exit 1
fi

# Test 3: Summary mode displays original_index for layouts
echo "Test 3: Summary layout index (original_index)"
uv run "$TOOL" --file "$FILE" --max-layouts 2 --summary | sed -n '/Available Layouts:/,/Theme/p' | head -n 10

# Test 4: Timeout wiring and analysis_complete flag
echo "Test 4: Timeout wiring & analysis_complete"
uv run "$TOOL" --file "$FILE" --json --deep --timeout 1 \
  | jq '{analysis_complete: .capabilities.analysis_complete, timeout_seconds: .metadata.timeout_seconds}'

# Test 5: Theme warnings for non-RGB scheme references (if present)
echo "Test 5: Theme colors warning"
uv run "$TOOL" --file "$FILE" --json \
  | jq '.warnings | map(select(test("Theme colors include scheme references"))) | length'

echo "═══════════════════════════════════════════════════════════════"
echo "Tests complete"
