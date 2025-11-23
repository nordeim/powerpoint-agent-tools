#!/bin/bash
set -e

# Activate venv
source .venv/bin/activate

# Use the generated file from the previous run (or create a dummy one if needed)
OUTPUT_FILE="bitcoin_market_report_nov2024_v2.pptx"

if [ ! -f "$OUTPUT_FILE" ]; then
    echo "Error: $OUTPUT_FILE not found. Please run the demo script first (until failure)."
    exit 1
fi

echo "Running ppt_validate_presentation.py..."
VALIDATION_OUTPUT=$(python3 tools/ppt_validate_presentation.py --file "$OUTPUT_FILE" --json)

echo "Attempting to parse output with ORIGINAL jq command (should fail)..."
if echo "$VALIDATION_OUTPUT" | jq -r '.issues | to_entries[] | "     - \(.key): \(.value | length) items"' 2>/dev/null; then
    echo "Original command SUCCEEDED (Unexpected)"
else
    echo "Original command FAILED (Expected)"
fi

echo "Attempting to parse output with NEW jq command (should pass)..."
if echo "$VALIDATION_OUTPUT" | jq -r '.summary | to_entries[] | "     - \(.key): \(.value) items"'; then
    echo "New command SUCCEEDED (Expected)"
else
    echo "New command FAILED (Unexpected)"
    exit 1
fi
