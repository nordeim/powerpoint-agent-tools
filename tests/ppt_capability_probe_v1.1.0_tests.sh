#!/bin/bash
# ppt_capability_probe_v1.1.0_tests.sh
# Comprehensive test suite for capability probe tool

echo "═══════════════════════════════════════════════════════════════"
echo "ppt_capability_probe.py v1.1.0 - Comprehensive Test Suite"
echo "═══════════════════════════════════════════════════════════════"
echo ""

# Test 1: JSON Contract Validation
echo "Test 1: JSON Contract Validation"
echo "─────────────────────────────────────────────────────────────"
uv run tools/ppt_capability_probe.py \
  --file bitcoin_market_report_nov2024_v2.pptx \
  --json | jq '{status, operation_id: .metadata.operation_id, duration_ms: .metadata.duration_ms, library_versions: .metadata.library_versions}'

echo ""

# Test 2: Atomic Read Verification
echo "Test 2: Atomic Read Verification"
echo "─────────────────────────────────────────────────────────────"
echo "Calculating checksum before probe..."
md5sum bitcoin_market_report_nov2024_v2.pptx > /tmp/before.md5

uv run tools/ppt_capability_probe.py \
  --file bitcoin_market_report_nov2024_v2.pptx \
  --verify-atomic \
  --json > /tmp/probe_result.json

echo "Calculating checksum after probe..."
md5sum bitcoin_market_report_nov2024_v2.pptx > /tmp/after.md5

if diff /tmp/before.md5 /tmp/after.md5 > /dev/null; then
  echo "✅ PASS: File unchanged (atomic read verified)"
else
  echo "❌ FAIL: File was modified!"
fi

echo ""

# Test 3: Placeholder Type Enum Accuracy
echo "Test 3: Placeholder Type Detection (Enum-Based)"
echo "─────────────────────────────────────────────────────────────"
uv run tools/ppt_capability_probe.py \
  --file bitcoin_market_report_nov2024_v2.pptx \
  --json | jq '.layouts[0].placeholder_types'

echo ""

# Test 4: Deep Mode Position Accuracy
echo "Test 4: Deep Mode (Transient Instantiation)"
echo "─────────────────────────────────────────────────────────────"
uv run tools/ppt_capability_probe.py \
  --file bitcoin_market_report_nov2024_v2.pptx \
  --deep \
  --json | jq '.layouts[0].placeholders[0] | {type, position_source, position_percent}'

echo ""

# Test 5: Capability Detection
echo "Test 5: Capability Detection"
echo "─────────────────────────────────────────────────────────────"
uv run tools/ppt_capability_probe.py \
  --file bitcoin_market_report_nov2024_v2.pptx \
  --json | jq '.capabilities | {has_footer_placeholders, has_slide_number_placeholders, total_layouts, total_master_slides}'

echo ""

# Test 6: Theme Extraction
echo "Test 6: Theme Extraction (Font Scheme API)"
echo "─────────────────────────────────────────────────────────────"
uv run tools/ppt_capability_probe.py \
  --file bitcoin_market_report_nov2024_v2.pptx \
  --json | jq '.theme.fonts'

echo ""

# Test 7: Warnings Array
echo "Test 7: Warnings and Info Arrays"
echo "─────────────────────────────────────────────────────────────"
uv run tools/ppt_capability_probe.py \
  --file bitcoin_market_report_nov2024_v2.pptx \
  --json | jq '{warnings, info}'

echo ""

# Test 8: Summary Mode
echo "Test 8: Human-Friendly Summary Mode"
echo "─────────────────────────────────────────────────────────────"
uv run tools/ppt_capability_probe.py \
  --file bitcoin_market_report_nov2024_v2.pptx \
  --summary | head -20

echo ""
echo "═══════════════════════════════════════════════════════════════"
echo "Test Suite Complete"
echo "═══════════════════════════════════════════════════════════════"
