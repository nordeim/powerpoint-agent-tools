#!/usr/bin/env bash
set -euo pipefail

TOOL="tools/ppt_capability_probe.py"

echo "═══════════════════════════════════════════════════════════════"
echo "ppt_capability_probe.py v1.1.1 - Comprehensive Regression Suite"
echo "═══════════════════════════════════════════════════════════════"

# 1. Success path (essential mode)
uv run "$TOOL" --file samples/sample.pptx --json | jq '.status, .metadata.tool_version, .slide_dimensions.aspect_ratio'

# 2. Success path (deep mode)
uv run "$TOOL" --file samples/sample.pptx --json --deep | jq '.layouts[0].placeholders[0].position_source, .layouts[0].instantiation_complete'

# 3. Error path (missing file)
if uv run "$TOOL" --file samples/missing.pptx --json; then
  echo "❌ FAIL: Missing file did not error"
else
  echo "✅ PASS: Missing file error path"
fi

# 4. Error path (locked file simulation)
chmod 000 samples/locked.pptx
if uv run "$TOOL" --file samples/locked.pptx --json; then
  echo "❌ FAIL: Locked file did not error"
else
  echo "✅ PASS: Locked file error path"
fi
chmod 644 samples/locked.pptx

# 5. Multi-master deck
uv run "$TOOL" --file samples/multi_master.pptx --json | jq '.capabilities.per_master, .theme.per_master'

# 6. Theme scheme-only colors
uv run "$TOOL" --file samples/scheme_colors.pptx --json | jq '.warnings'

# 7. Missing font scheme
uv run "$TOOL" --file samples/no_font_scheme.pptx --json | jq '.theme.fonts, .warnings'

# 8. Large deck with max-layouts
uv run "$TOOL" --file samples/large.pptx --json --max-layouts 5 | jq '.metadata.layout_count_total, .metadata.layout_count_analyzed, .info'

# 9. Timeout partial analysis
uv run "$TOOL" --file samples/large.pptx --json --deep --timeout 1 | jq '.capabilities.analysis_complete, .warnings'

echo "═══════════════════════════════════════════════════════════════"
echo "Comprehensive tests complete"

