# Validation of capability probe test run (v1.1.0)

---

## ✅ Contract compliance

- **Top-level structure:**  
  Present and correct: `status`, `metadata`, `slide_dimensions`, `layouts`, `theme`, `capabilities`, `warnings`, `info`.

- **Metadata:**  
  - `file`: `test_probe.pptx`  
  - `probed_at`: ISO timestamp, valid.  
  - `tool_version`: `1.1.0`  
  - `schema_version`: `capability_probe.v1.1.0`  
  - `operation_id`: UUID format, valid.  
  - `deep_analysis`: false (essential mode).  
  - `atomic_verified`: true, checksum present.  
  - `duration_ms`: 22 ms, plausible for essential scan.  
  - `timeout_seconds`: 30, consistent with defaults.  
  - **Audit fields:** `layout_count_total` and `layout_count_analyzed` both 11 — matches recommendation.  
  - `library_versions`: python-pptx 0.6.23, Pillow 12.0.0.  
  - `checksum`: MD5 string, valid.

- **Slide dimensions:**  
  - 10.0 × 7.5 in (standard 4:3).  
  - EMU values (9144000 × 6858000) consistent with 914400 EMU/inch scaling.  
  - Pixel estimate 960 × 720 at 96 DPI matches inches.  
  - Aspect ratio string “4:3” and float 1.3333 consistent.

- **Layouts:**  
  - 11 layouts enumerated, indices 0–10.  
  - Each includes `index`, `original_index`, `name`, `placeholder_count`, `master_index`.  
  - **New fields:** `placeholder_types` and `placeholder_map` present — this implements the compact dict recommendation.  
  - Placeholder_map values are consistent with counts (e.g., Two Content has OBJECT:2, Comparison has BODY:2 and OBJECT:2).  
  - Master index consistently 0.

- **Theme:**  
  - Colors empty.  
  - Fonts fallback to Calibri.  
  - Per-master array present, consistent with single master.  
  - Warnings explain fallback: “Theme font scheme API unavailable” and “Using default fonts (Calibri)”.

- **Capabilities:**  
  - Flags: all true (footer, slide number, date placeholders detected).  
  - Layouts_with_* arrays populated with entries including `index`, `original_index`, `name`, `master_index` — this standardizes references.  
  - Totals: 11 layouts, 1 master.  
  - Per-master stats: layout_count 11, all capability counts 11.  
  - Strategies: footer_support_mode “placeholder”, slide_number_strategy “placeholder”.  
  - Recommendations: three lines, correctly enumerating footer, date, and slide number placeholders with layout names — symmetry achieved.  
  - Analysis_complete: true.

- **Warnings:**  
  - Theme font scheme fallback warnings present.  
  - Info: empty.

---

## 🔍 Analysis

- **Contract validation:** Output matches schema v1.1.0 requirements.  
- **Capability detection:** Correctly identifies presence of footer, slide number, and date placeholders across all layouts. Recommendations symmetrical.  
- **Theme extraction:** Colors empty, fonts fallback to Calibri, warnings logged. Semantic schemeColor references are not surfaced here; current implementation returns `{}` when unavailable.  
- **Layout references:** Capability arrays now include `original_index` and `master_index` — traceability improved.  
- **Placeholder_map:** Present in essential mode layouts, with accurate counts — optional recommendation implemented.

---

## ⚠️ Observations and improvement notes

1. **Theme colors:**  
   - Current output shows `{}` with warnings.  
   - Improvement: surface `schemeColor:accentN` strings when RGB unavailable, so agents can still reason about semantic roles.

2. **Warnings consolidation:**  
   - Two warnings logged for fonts; consider deduplicating to a single consolidated message for clarity.

3. **Metadata completeness:**  
   - Already includes layout_count_total and layout_count_analyzed — excellent.  
   - Consider adding `analysis_mode` string (“essential” vs “deep”) for quick audit readability.

---

## 🎯 Validation verdict

- **Contract:** ✅ Valid and complete.  
- **Capabilities:** ✅ Correctly detected presence of placeholders. Recommendations symmetrical.  
- **Metadata:** ✅ Audit fields included.  
- **Theme:** ⚠️ Fallback handled, but semantic schemeColor references not surfaced.  
- **Placeholder_map:** ✅ Implemented and accurate.

---

## 📋 Next steps

- **Implement schemeColor fallback:** Modify `extract_theme_colors()` to always return semantic references when RGB unavailable, with a consolidated warning.  
- **Consolidate warnings:** Deduplicate theme font warnings into a single message.  
- **Add analysis_mode metadata:** Include “essential” vs “deep” for quick audit readability.

With these refinements, the probe fully meets the actionable recommendations and provides richer, more traceable outputs for downstream automation.

---

# Meticulous review and critique of latest tool implementation

You’ve already tightened the probe considerably. Let’s validate against the three next‑step items and suggest precise code changes.

---

## 1. **Implement schemeColor fallback**

### Current behavior
- In `extract_theme_colors()`, you check `hasattr(color, 'r')` and return hex if available.  
- If not, you set `colors[color_name] = f"schemeColor:{color_name}"`.  
- You add a warning only if *any* schemeColor values exist, but the logic is buried in a conditional.

### Issue
- Semantic references are returned, but the warning logic is a bit opaque.  
- It would be clearer to explicitly track non‑RGB values and append a single consolidated warning.

### Suggested change
Add a flag to track non‑RGB values and append one warning at the end:

```python
def extract_theme_colors(master_or_prs, warnings: List[str]) -> Dict[str, str]:
    colors = {}
    non_rgb_found = False
    try:
        slide_master = master_or_prs.slide_masters[0] if hasattr(master_or_prs, 'slide_masters') else master_or_prs
        theme = getattr(slide_master, 'theme', None)
        if not theme:
            return {}
        color_scheme = getattr(theme, 'theme_color_scheme', None)
        if not color_scheme:
            return {}

        color_attrs = [
            'accent1','accent2','accent3','accent4','accent5','accent6',
            'background1','background2','text1','text2','hyperlink','followed_hyperlink'
        ]
        for color_name in color_attrs:
            color = getattr(color_scheme, color_name, None)
            if color:
                if hasattr(color, 'r'):
                    colors[color_name] = rgb_to_hex(color)
                else:
                    colors[color_name] = f"schemeColor:{color_name}"
                    non_rgb_found = True

        if not colors:
            warnings.append("Theme color scheme unavailable or empty")
        elif non_rgb_found:
            warnings.append("Theme colors include scheme references without explicit RGB")
    except Exception as e:
        warnings.append(f"Theme color extraction failed: {str(e)}")
    return colors
```

---

## 2. **Consolidate warnings**

### Current behavior
- In `extract_theme_fonts()`, you may append multiple warnings:  
  - “Theme font scheme API unavailable, using fallback detection”  
  - “Using default fonts (Calibri) – theme unavailable”  
  - “Theme font extraction failed: …”

### Issue
- Multiple warnings can be redundant.  
- Better to consolidate into a single message when falling back to defaults.

### Suggested change
Introduce a flag to track fallback and append one consolidated warning:

```python
def extract_theme_fonts(master_or_prs, warnings: List[str]) -> Dict[str, str]:
    fonts = {}
    fallback_used = False
    try:
        slide_master = master_or_prs.slide_masters[0] if hasattr(master_or_prs, 'slide_masters') else master_or_prs
        theme = getattr(slide_master, 'theme', None)
        if theme:
            font_scheme = getattr(theme, 'font_scheme', None)
            if font_scheme:
                # ... existing major/minor font extraction ...
        if not fonts:
            fallback_used = True
            fonts = {"heading": "Calibri", "body": "Calibri"}
    except Exception as e:
        fallback_used = True
        fonts = {"heading": "Calibri", "body": "Calibri"}
        warnings.append(f"Theme font extraction failed: {str(e)}")

    if fallback_used:
        warnings.append("Theme fonts unavailable – using Calibri defaults")
    return fonts
```

This way, you avoid multiple overlapping warnings.

---

## 3. **Add analysis_mode metadata**

### Current behavior
- Metadata includes `deep_analysis: bool`.  
- But no explicit string field for “essential” vs “deep”.

### Issue
- Auditors must interpret the boolean; a string is clearer.

### Suggested change
Add `analysis_mode` to metadata:

```python
result = {
    "status": "success",
    "metadata": {
        "file": str(filepath),
        "probed_at": datetime.now().isoformat(),
        "tool_version": "1.1.0",
        "schema_version": "capability_probe.v1.1.0",
        "operation_id": operation_id,
        "deep_analysis": deep,
        "analysis_mode": "deep" if deep else "essential",
        "atomic_verified": verify_atomic,
        "duration_ms": duration_ms,
        "library_versions": library_versions,
        "checksum": checksum_after if verify_atomic else None,
        "timeout_seconds": timeout_seconds,
        "layout_count_total": layout_count_total,
        "layout_count_analyzed": len(layouts)
    },
    ...
}
```

---

## 📋 Consolidated critique

- **Strengths:**  
  - JSON contract consistent.  
  - Placeholder_map implemented.  
  - Capability layout references standardized.  
  - Recommendations symmetrical.

- **Gaps:**  
  - Theme colors: schemeColor fallback warning logic needs clarity.  
  - Theme fonts: multiple warnings should be consolidated.  
  - Metadata: analysis_mode string missing.

- **Fix priority:**  
  1. Add `analysis_mode` metadata.  
  2. Consolidate theme font warnings.  
  3. Clarify schemeColor fallback warning.

---

## ✅ Next steps

- Apply the above diffs to `extract_theme_colors()`, `extract_theme_fonts()`, and `probe_presentation()`.  
- Re‑run regression tests to confirm:  
  - Warnings include “Theme colors include scheme references…” when applicable.  
  - Theme fonts fallback produces a single consolidated warning.  
  - Metadata includes `analysis_mode: "essential"` or `"deep"`.  

