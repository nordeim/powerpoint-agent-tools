# Review of Regression Test Output

Your minimal regression suite ran cleanly:

- **Test 1 (JSON Contract & version):**  
  ✅ Contract fields present, schema_version matches `capability_probe.v1.1.1`, all required keys flagged as present.  
  Duration and layout counts look correct.

- **Test 2 (Atomic verification):**  
  ✅ File unchanged before/after probe. Atomic read verified.

- **Test 3 (Summary layout index):**  
  ✅ Summary shows `original_index` values, not truncated indices. Layout names and placeholder types are correctly surfaced.

- **Test 4 (Timeout wiring & analysis_complete):**  
  ✅ With `--timeout 1`, analysis still completed quickly on sample deck. Flagging works (`analysis_complete: true`, `timeout_seconds: 1`).

- **Test 5 (Theme colors warning):**  
  ✅ No warnings for non-RGB scheme references in sample deck. Correctly returns `0`.

Overall: the minimal suite validates core contract, atomicity, summary index mapping, timeout wiring, and theme warnings. It’s a solid smoke test.

---

# Gaps in Current Test Coverage

The current suite does **not** exercise:

1. **Error paths**  
   - Locked file / permission denied.  
   - Non-existent file.  
   - Schema validation failure.

2. **Deep mode instantiation**  
   - Placeholder positions marked as `instantiated`.  
   - `instantiation_complete` flags.  
   - EMU values in placeholder payload.

3. **Multi-master decks**  
   - Per-master stats in capabilities.  
   - Per-master theme colors/fonts.

4. **Theme edge cases**  
   - Decks with scheme-only colors (non-RGB).  
   - Decks with missing font schemes.

5. **Large deck truncation**  
   - `--max-layouts` limiting analysis.  
   - Metadata layout_count_total vs layout_count_analyzed.

6. **Timeout partial analysis**  
   - Deck large enough to trigger timeout.  
   - `analysis_complete: false` flagged.

---

# Comprehensive Test Plan

## Functional Scenarios

- **Success path (essential mode):** Validate JSON contract, summary output, atomic verification.
- **Success path (deep mode):** Validate placeholder payloads include EMU, instantiation flags.
- **Error path (locked file):** Expect `status:error`, `error_type:PermissionError`.
- **Error path (missing file):** Expect `status:error`, `error_type:FileNotFoundError`.
- **Error path (schema validation failure):** Inject missing field, expect `status:error`, `error_type:SchemaValidationError`.

## Edge Cases

- **Multi-master deck:** Validate `capabilities.per_master` and `theme.per_master`.
- **Theme scheme-only colors:** Validate consolidated warning “Theme colors include scheme references without explicit RGB”.
- **Missing font scheme:** Validate fallback to Calibri with warning.
- **Large deck with max-layouts:** Validate info message and layout_count_analyzed < layout_count_total.
- **Timeout partial analysis:** Validate `analysis_complete:false` and timeout warning.

---

# Extended Test Harness (bash + jq)

```bash
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
```

---

# Next Steps

- **Prepare sample decks:**  
  - `sample.pptx` (basic single-master).  
  - `multi_master.pptx` (two masters).  
  - `scheme_colors.pptx` (theme with scheme-only colors).  
  - `no_font_scheme.pptx` (theme missing font scheme).  
  - `large.pptx` (≥50 layouts).  
  - `locked.pptx` (permission denied simulation).

- **Integrate into CI:**  
  - Add these tests to GitHub Actions workflow.  
  - Validate outputs against schema JSON.  
  - Block merges on failures.

---

✅ With this extended suite, you’ll cover success, error, deep mode, multi-master, theme edge cases, large deck truncation, and timeout behavior — ensuring the probe tool is robust across real-world scenarios.
