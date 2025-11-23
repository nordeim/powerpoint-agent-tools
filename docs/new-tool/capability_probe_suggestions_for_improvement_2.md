# Review and validation of capability probe test run (v1.1.0)

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
  - `duration_ms`: 15 ms, plausible for essential scan.  
  - `timeout_seconds`: 30, consistent with defaults.  
  - **Audit fields:** `layout_count_total` and `layout_count_analyzed` both 11 — excellent, matches actionable recommendation.  
  - `library_versions`: python-pptx 0.6.23, Pillow 12.0.0.  
  - `checksum`: MD5 string, valid.

- **Slide dimensions:**  
  - 10.0 × 7.5 in (standard 4:3).  
  - EMU values (9144000 × 6858000) consistent with 914400 EMU/inch.  
  - Pixel estimate 960 × 720 at 96 DPI matches inches.  
  - Aspect ratio string “4:3” and float 1.3333 consistent.

- **Layouts:**  
  - 11 layouts enumerated, indices 0–10.  
  - Each includes `index`, `original_index`, `name`, `placeholder_count`, `master_index`.  
  - Placeholder types omitted in essential mode here, but counts are consistent.  
  - Master index consistently 0.

- **Theme:**  
  - Colors empty.  
  - Fonts fallback to Calibri.  
  - Per-master array present, consistent with single master.  
  - Warnings explain fallback: “Theme font scheme API unavailable” and “Using default fonts (Calibri)”.

- **Capabilities:**  
  - Flags: all false (no footer, slide number, date placeholders detected).  
  - Layouts_with_* arrays empty.  
  - Totals: 11 layouts, 1 master.  
  - Per-master stats: layout_count 11, all capability counts 0.  
  - Strategies: footer_support_mode “fallback_textbox”, slide_number_strategy “textbox”.  
  - Recommendations: three lines, correctly reflecting absence of footer, slide number, and date placeholders.  
  - Analysis_complete: true.

- **Warnings:**  
  - Theme font scheme fallback warnings present.  
  - Info: empty.

---

## 🔍 Analysis

- **Contract validation:** Output matches schema v1.1.0 requirements. Audit metadata fields (`layout_count_total`, `layout_count_analyzed`) are included — this was a key recommendation and is now consistently applied.  
- **Capability detection:** Correctly identifies absence of footer, slide number, and date placeholders. Recommendations are symmetrical across all three roles — another improvement applied.  
- **Theme extraction:** Colors empty, fonts fallback to Calibri, warnings logged. Semantic schemeColor references are not surfaced here; current implementation returns `{}` when unavailable.  
- **Layout references:** Capability arrays are empty, so no layout_ref objects to inspect. In runs where placeholders exist, ensure `original_index` and `master_index` are included — this log doesn’t exercise that path.  
- **Placeholder_map:** Not present in essential mode layouts. Counts are provided, but no compact map of type → count. Optional recommendation not yet implemented.

---

## ⚠️ Observations and improvement notes

1. **Theme colors:**  
   - Current output shows `{}` with warnings.  
   - Improvement: surface `schemeColor:accentN` strings when RGB unavailable, so agents can still reason about semantic roles.

2. **Capability layout references:**  
   - Not exercised here (arrays empty).  
   - Ensure when populated, entries include `original_index` and `master_index` for traceability.

3. **Placeholder_map:**  
   - Optional recommendation not yet implemented.  
   - Adding `placeholder_map` in essential mode would make downstream decisions lighter.

---

## 🎯 Validation verdict

- **Contract:** ✅ Valid and complete.  
- **Capabilities:** ✅ Correctly detected absence of placeholders. Recommendations symmetrical.  
- **Metadata:** ✅ Audit fields included.  
- **Theme:** ⚠️ Fallback handled, but semantic schemeColor references not surfaced.  
- **Placeholder_map:** ⚠️ Optional improvement not yet applied.

---

## 📋 Next steps

- **Implement schemeColor fallback:** Modify `extract_theme_colors()` to always return semantic references when RGB unavailable, with a consolidated warning.  
- **Standardize capability layout references:** Ensure `original_index` and `master_index` are included in capability arrays when populated.  
- **Add placeholder_map:** In essential mode, include a compact dict of type → count per layout.  

With these refinements, the probe will fully meet the actionable recommendations and provide richer, more traceable outputs for downstream automation.

---

# Meticulous review and critique of latest tool implementation

You’ve already hardened the probe substantially. Let’s validate against the three next‑step items and suggest precise code changes.

---

## 1. **Implement schemeColor fallback**

### Current behavior
- In `extract_theme_colors()`, you check `hasattr(color, 'r')` and return hex if available.  
- If not, you set `colors[color_name] = f"schemeColor:{color_name}"`.  
- You suppress warnings to avoid spam.

### Issue
- Semantic references are returned, but no consolidated warning is logged.  
- Agents may not realize they’re consuming scheme references instead of RGB.

### Suggested change
Add a flag to track non‑RGB values and append a single warning at the end.

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

## 2. **Standardize capability layout references**

### Current behavior
- In `analyze_capabilities()`, you build `layout_ref = {"index": layout['index'], "name": layout['name']}`.  
- In the latest patch you added `original_index` and `master_index` in some places, but not consistently.

### Issue
- Capability arrays may omit `original_index` and `master_index`, reducing traceability.

### Suggested change
Always include both fields:

```python
layout_ref = {
    "index": layout['index'],
    "original_index": layout.get('original_index', layout['index']),
    "name": layout['name'],
    "master_index": layout.get('master_index')
}
```

Apply this consistently for `layouts_with_footer`, `layouts_with_slide_number`, and `layouts_with_date`.

---

## 3. **Add placeholder_map in essential mode**

### Current behavior
- In essential mode (`deep=False`), you only collect `placeholder_types` list.  
- No compact map of type → count is provided.

### Issue
- Downstream consumers must parse arrays to count roles.  
- A dict would make decisions lighter and more auditable.

### Suggested change
Add a `placeholder_map` dict in essential mode:

```python
else:
    placeholder_map = {}
    placeholder_types = []
    for shape in layout.placeholders:
        try:
            ph_type = shape.placeholder_format.type
            ph_type_name = get_placeholder_type_name(ph_type)
            placeholder_map[ph_type_name] = placeholder_map.get(ph_type_name, 0) + 1
            if ph_type_name not in placeholder_types:
                placeholder_types.append(ph_type_name)
        except Exception:
            pass
    layout_info["placeholder_types"] = placeholder_types
    layout_info["placeholder_map"] = placeholder_map
```

---

## 📋 Consolidated critique

- **Strengths:**  
  - JSON contract is consistent.  
  - Deep‑mode instantiation is safe.  
  - Theme/font fallbacks are robust.  
  - Atomic verification enforced.

- **Gaps:**  
  - Theme colors: schemeColor fallback present but warning suppressed.  
  - Capability layout references: not standardized.  
  - Essential mode: lacks placeholder_map.

- **Fix priority:**  
  1. Standardize capability layout references (critical for traceability).  
  2. Implement consolidated schemeColor warning (important for usability).  
  3. Add placeholder_map (optional but valuable for downstream efficiency).

---

## ✅ Next steps

- Apply the above diffs to `extract_theme_colors()`, `analyze_capabilities()`, and `detect_layouts_with_instantiation()`.  
- Re‑run regression tests to confirm:  
  - Warnings include “Theme colors include scheme references…” when applicable.  
  - Capability arrays include `original_index` and `master_index`.  
  - Essential mode layouts include `placeholder_map`.  

