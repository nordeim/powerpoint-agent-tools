$ uv run tools/ppt_capability_probe.py --file samples/sample.pptx --json > out_essential.json
python - <<'PY'
import json; d=json.load(open("out_essential.json"))
assert d["status"]=="success"
assert d["metadata"]["analysis_mode"]=="essential"
assert isinstance(d["metadata"]["warnings_count"], int)
assert d["metadata"]["layout_count_total"] >= d["metadata"]["layout_count_analyzed"]
PY


$ uv run tools/ppt_capability_probe.py --file samples/sample.pptx --deep --json > out_deep.json
python - <<'PY'
import json; d=json.load(open("out_deep.json"))
assert d["metadata"]["analysis_mode"]=="deep"
# Assert placeholders present for at least one layout
assert any("placeholders" in l and isinstance(l["placeholders"], list) for l in d["layouts"])
PY


$ if python3 tools/ppt_capability_probe.py --file samples/missing.pptx --json > out_missing.json; then
  echo "❌ FAIL: expected error for missing file"
  exit 1
else
  echo "✅ PASS: Missing file error path"
fi
✅ PASS: Missing file error path


$ uv run tools/ppt_capability_probe.py --file samples/multi_master.pptx --json > out_multi.json
python - <<'PY'
import json; d=json.load(open("out_multi.json"))
assert d["capabilities"]["total_master_slides"] >= 2
assert len(d["capabilities"]["per_master"]) >= 2
PY

Traceback (most recent call last):
  File "<stdin>", line 2, in <module>
AssertionError

$ cat out_multi.json
{
  "status": "success",
  "metadata": {
    "file": "samples/multi_master.pptx",
    "probed_at": "2025-11-23T11:25:13.795205",
    "tool_version": "1.1.0",
    "schema_version": "capability_probe.v1.1.0",
    "operation_id": "76ade394-5bdb-4f7b-86ee-cc9d9745007a",
    "deep_analysis": false,
    "analysis_mode": "essential",
    "atomic_verified": true,
    "duration_ms": 35,
    "timeout_seconds": 30,
    "layout_count_total": 3,
    "layout_count_analyzed": 3,
    "warnings_count": 1,
    "library_versions": {
      "python-pptx": "0.6.23",
      "Pillow": "12.0.0"
    },
    "checksum": "c8136b86f405da1ee8587dd8f4f7c894"
  },
  "slide_dimensions": {
    "width_inches": 10.5,
    "height_inches": 8.0,
    "width_emu": 9601200,
    "height_emu": 7315200,
    "width_pixels": 1008,
    "height_pixels": 768,
    "aspect_ratio": "21:16",
    "aspect_ratio_float": 1.3125,
    "dpi_estimate": 96
  },
  "layouts": [
    {
      "index": 0,
      "original_index": 0,
      "name": "Cover ",
      "placeholder_count": 3,
      "master_index": 0,
      "placeholder_types": [
        "PICTURE",
        "TITLE"
      ],
      "placeholder_map": {
        "PICTURE": 2,
        "TITLE": 1
      }
    },
    {
      "index": 1,
      "original_index": 1,
      "name": "Preview Page 1",
      "placeholder_count": 2,
      "master_index": 0,
      "placeholder_types": [
        "TITLE",
        "PICTURE"
      ],
      "placeholder_map": {
        "TITLE": 1,
        "PICTURE": 1
      }
    },
    {
      "index": 2,
      "original_index": 2,
      "name": "Painting Page 1",
      "placeholder_count": 3,
      "master_index": 0,
      "placeholder_types": [
        "TITLE",
        "PICTURE"
      ],
      "placeholder_map": {
        "TITLE": 1,
        "PICTURE": 2
      }
    }
  ],
  "theme": {
    "colors": {},
    "fonts": {
      "heading": "Calibri",
      "body": "Calibri"
    },
    "per_master": [
      {
        "master_index": 0,
        "colors": {},
        "fonts": {
          "heading": "Calibri",
          "body": "Calibri"
        }
      }
    ]
  },
  "capabilities": {
    "has_footer_placeholders": false,
    "has_slide_number_placeholders": false,
    "has_date_placeholders": false,
    "layouts_with_footer": [],
    "layouts_with_slide_number": [],
    "layouts_with_date": [],
    "total_layouts": 3,
    "total_master_slides": 1,
    "per_master": [
      {
        "master_index": 0,
        "layout_count": 3,
        "has_footer_layouts": 0,
        "has_slide_number_layouts": 0,
        "has_date_layouts": 0
      }
    ],
    "footer_support_mode": "fallback_textbox",
    "slide_number_strategy": "textbox",
    "recommendations": [
      "No footer placeholders found - ppt_set_footer.py will use text box fallback strategy",
      "No slide number placeholders - recommend manual text box for slide numbers",
      "No date placeholders - dates must be added manually if needed"
    ],
    "analysis_complete": true
  },
  "warnings": [
    "Theme fonts unavailable - using Calibri defaults"
  ],
  "info": []
}

---

# PowerPoint capability probe test log review

Here’s a tight, production-grade review of your test runs and the probe’s JSON contract, focused on reproducibility, auditability, and the one failing path that needs to be fixed.

---

## Executive summary

- The essential vs deep mode behavior is correct, and the JSON contract is largely consistent and machine-actionable across runs.
- Error handling for missing files is correct and gated.
- The multi-master test fails because the tool reports only one master slide and only a single per_master entry, indicating either the sample presentation has one master or the probe’s master detection/mapping is incomplete.
- Minor contract drift and warning semantics show up across outputs; these should be standardized and validated against the v1.1.1 schema.

---

## Validation results across test runs

### Essential mode on samples/sample.pptx

- **Status and contract:** success; metadata includes analysis_mode=essential; warnings_count present; counts consistent.
- **Assertions:** all pass.
- **Observations:** Good coverage and contract hygiene, including layout_count_total ≥ layout_count_analyzed and integer warnings_count.
- **Verdict:** Pass.

### Deep mode on samples/sample.pptx

- **Status and contract:** success; metadata analysis_mode=deep.
- **Assertions:** placeholders list exists on at least one layout.
- **Observations:** Deep instantiation path exercised correctly with transient slide lifecycle.
- **Verdict:** Pass.

### Missing file path

- **Behavior:** CLI returns non-zero exit; JSON error payload printed; shell test marks PASS.
- **Observations:** Proper guard rails (FileNotFoundError) and error payload include operation_id, probed_at.
- **Verdict:** Pass.

### Multi-master on samples/multi_master.pptx

- **Expectation:** total_master_slides ≥ 2 and per_master length ≥ 2.
- **Observed:** total_master_slides=1; per_master=[{ master_index:0, ... }]; assertion fails.
- **Verdict:** Fail.

---

## Root cause analysis for multi-master failure

### What the output shows

- **total_master_slides:** 1
- **per_master:** Single entry with master_index=0 only
- **layout master_index mapping:** All layouts report master_index=0
- **Theme per_master:** Single entry for master_index=0
- **Capabilities and recommendations:** Reflect a single master context only

### Likely causes

- **Sample actually single-master:** If samples/multi_master.pptx truly has only one SlideMaster, the test expectation is incorrect.
- **Master detection logic is incomplete or mismatched:** The code builds master_map by iterating prs.slide_masters and mapping their slide_layouts via id(layout) → master_index. If multiple masters exist but the layout iteration is using Presentation.slide_layouts (which can flatten or reorder), master associations could be lost or deduplicated to the first master when retrieving by index and id.
- **Theme per_master extraction suppression:** The code extracts per-master theme without collecting warnings, but still appends one element per master. If prs.slide_masters has length >1 but theme extraction short-circuits or a later overwrite occurs, you would still expect multiple entries; since you don’t, it corroborates single-master detection.

### What to validate next

- **Direct inspection:** Log len(prs.slide_masters) and identify masters’ layout counts. If it’s 1, fix the test. If ≥2, fix the mapping.
- **Identity mapping robustness:** Ensure id(layout) is stable across master vs presentation-level layout objects. Some python-pptx APIs can return new wrapper objects whose id differs from master.slide_layouts members. If so, use a robust key (e.g., rId or idx from XML) rather than id(layout).

---

## Contract consistency and deltas

### Analysis mode and warnings count

- **Present in later runs:** metadata.analysis_mode and metadata.warnings_count are present and used in tests.
- **Absent in earlier logs:** Some earlier outputs lack analysis_mode and warnings_count, consistent with a prior version of the tool/contract.
- **Action:** Ensure these fields are always present in v1.1.0+; set warnings_count=len(warnings) at emission time.

### Theme warnings semantics

- **Variants seen:**
  - “Theme fonts unavailable - using Calibri defaults”
  - “Theme font scheme API unavailable, using fallback detection”
  - “Using default fonts (Calibri) - theme unavailable”
- **Action:** Normalize warning taxonomy to a stable set of codes/messages for downstream gating (e.g., THEME_FONT_FALLBACK, THEME_SCHEME_MISSING), and provide a single primary message with optional details to avoid duplication or ambiguity.

### Per-master and layout master_index fields

- **Consistency:** When total_master_slides=1, per_master length=1 and all layouts master_index=0. This is consistent.
- **Expectation:** For multi-master, both fields should reflect ≥2. The failure indicates mismatch between data and test expectation.

### v1.1.1 JSON schema alignment

- **Observed outputs:** Align with v1.1.0 schema strings; mostly satisfy required fields in the v1.1.1 schema excerpt.
- **Gaps to check:**
  - Ensure metadata.schema_version follows pattern capability_probe.vX.Y.Z.
  - Ensure placeholders array and position fields exist in deep mode for every layout; essential mode can omit placeholders but should include placeholder_types.
  - Ensure checksum is a 32-hex MD5 when atomic_verified=true; else emit verification_skipped explicitly.

---

## Recommendations to fix and harden

### 1) Master-layout association robustness

- **Label:** Use stable relationship keys
- **Change:** Replace id(layout) mapping with a robust key:
  - Extract each layout’s underlying part or rId via internal pptx structures (e.g., layout.part or layout._element and related relationship id). Use that as the mapping key across master.slide_layouts and prs.slide_layouts.
- **Benefit:** Prevents false single-master attribution when wrapper objects differ.

### 2) Explicit master enumeration in output

- **Label:** Emit master catalog
- **Change:** Add metadata.masters: array with master_index, layout_count, name (if available), rId, and a checksum of the master XML. Cross-reference layouts by their master rId.
- **Benefit:** Auditable linkage and easier debugging of multi-master detection.

### 3) Contract normalization

- **Label:** Always-on fields
- **Change:** Always include:
  - metadata.analysis_mode: “essential” or “deep”
  - metadata.warnings_count: integer = len(warnings)
  - capabilities.analysis_complete: boolean
- **Benefit:** Stable gating and test consistency.

### 4) Warning taxonomy

- **Label:** Deduplicate and code warnings
- **Change:** Use normalized warnings:
  - THEME_FONT_FALLBACK: “Theme fonts unavailable; using Calibri defaults”
  - THEME_SCHEME_MISSING: “Theme color/font scheme API unavailable”
  - POSITION_TEMPLATE_FALLBACK: “Using template positions; instantiation failed”
- **Benefit:** Cleaner downstream handling and less brittle string checks.

### 5) Validation enhancements

- **Label:** Multi-master assertion pre-check
- **Change:** In tests, assert d["metadata"]["layout_count_total"] > 0 and log d["capabilities"]["total_master_slides"] for diagnostics before strict ≥2 assertion. Add a test that round-trips per_master length == total_master_slides.
- **Benefit:** Faster root-cause identification when failures occur.

### 6) Deep mode completeness check

- **Label:** Placeholder instantiation parity
- **Change:** In deep mode, assert for every layout: placeholder_instantiated == placeholder_expected. If not, add a warning POSITION_TEMPLATE_FALLBACK and mark instantiation_complete=false.
- **Benefit:** Guarantees accuracy of positional data and flags degraded paths.

---

## Quick test add-ons to catch regressions

- **Master parity:** Assert len(d["capabilities"]["per_master"]) == d["capabilities"]["total_master_slides"].
- **Layout-master coverage:** Assert set(l["master_index"] for l in d["layouts"]) subset of range(total_master_slides).
- **Contract invariants:** Assert presence and type of metadata.analysis_mode, metadata.warnings_count, capabilities.analysis_complete.
- **Warning count sanity:** Assert metadata.warnings_count == len(d["warnings"]).
- **Checksum integrity:** When atomic_verified=true, assert metadata.checksum is 32 hex chars.

---

## Suggested next steps

- Does the actual samples/multi_master.pptx properties truly contain multiple slide masters? If yes, pinpoint the exact mapping line to replace and propose the stable rId-based association. If no, we adjust the test expectations and add a labeled sample that is guaranteed multi-master for the CI gate.

