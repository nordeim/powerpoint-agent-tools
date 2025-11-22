# Meticulous review of ppt_capability_probe.py v1.1.0

You’ve built a strong, production-grade probe. It’s cleanly structured, JSON-first, atomic, and aligned with machine-actionable introspection. Below is a rigorous validation and critique with concrete, low-risk fixes to harden reliability, auditability, and maintainability.

---

## Contract validation and output integrity

- **JSON contract compliance:** 
  - Top-level keys are present: status, metadata (file, probed_at, tool_version, operation_id, duration_ms), slide_dimensions, layouts, theme, capabilities, warnings, info. 
  - Contract is consistent on both success and error paths. 
  - Validation runs before print and marks `status=error` on schema gaps. Good.

- **Durability signals:**
  - **Atomic read verification:** Checksum before/after prevents silent mutation. Good.
  - **Operation metadata:** UUID `operation_id`, duration, library versions, and checksum logged. Good.
  - **Timeout behavior:** `analysis_complete` flag is set based on runtime. Good.

- **Usability:**
  - **Deep vs essential modes:** Clear behavior and flags. Good.
  - **Multiple masters:** Per-master mapping, master index association, and per-master stats. Good.
  - **Theme extraction:** Robust font scheme and color scheme fallbacks with warnings. Good.

Direct verdict: Contract is solid and production-ready with minor issues to correct.

---

## Functional review and edge-case handling

- **Placeholder detection:** 
  - Dynamic enum mapping from `PP_PLACEHOLDER` with graceful unknown handling. Good.
  - Deep mode instantiation uses transient slide lifecycle with safe cleanup. Good.
  - Position precision includes inches, EMUs, percentages, and source marking. Good.

- **Layout analysis:**
  - `index` reflects sliced subset index when `--max-layouts` truncates, while `original_index` preserves actual file layout position. Good.
  - Per-master association map is correctly inferred and surfaced. Good.

- **Theme extraction:**
  - Fonts: `major_font` and `minor_font` with East Asian/complex script precedence, fallback to on-shape detection and Calibri defaults with warnings. Good.
  - Colors: Proper detection via theme color scheme; non-RGB scheme references labeled as `schemeColor:*`. Good.

- **Capabilities:**
  - Accurate detection for footer, slide number, and date placeholders via type codes, mapped layout references, strategy flags, and per-master tallies. Good.

- **Error and safety:**
  - Locked file and path validation behaviors are correct.
  - Internal pptx structures used for deletion (`_sldIdLst` and `drop_rel`) handled inside try blocks. Good.

---

## Bugs, risks, and inconsistencies to fix

- **Duplicate key assignment (capabilities.per_master):**
  - `analyze_capabilities()` sets `"per_master"` twice, causing redundant assignment. Remove duplicate.
  - Impact: Cosmetic redundancy; potential confusion in code review or downstream processing.

- **Unused imports and symbols:**
  - `Inches` and `BytesIO` are imported but unused.
  - `PowerPointAgent` is imported but only `PowerPointAgentError` is used.
  - Impact: Minor lint failures; noise for maintainers.

- **Layout index semantics in summary:**
  - Summary prints `layout['index']` (subset index when `--max-layouts` is used), not `original_index`. 
  - Impact: UI may mislead when partial analysis truncates layouts; displaying `original_index` is clearer.

- **Theme color handling for non-RGB:**
  - Colors default to `schemeColor:*` when actual RGB isn’t available. This is fine, but consider surfacing a single warning summarizing “non-RGB theme references found” once to aid audits.
  - Impact: Better operator situational awareness.

- **Timeout signaling:**
  - `detect_layouts_with_instantiation()` adds a timeout warning mid-loop and `analysis_complete` is computed afterward. Good overall; consider adding `metadata.timeout_seconds` so the contract shows configured ceiling for audits.
  - Impact: Improved traceability.

- **DPI and aspect ratio outputs:**
  - DPI is a static estimate. Already labeled as “estimate,” which is appropriate. Good to keep, but mention in `info` when aspect ratio is nonstandard (already derived via rational approximation; consider adding a single info line).

---

## Recommended code changes (surgical diffs)

### 1) Remove duplicate per_master assignment in capabilities

```diff
def analyze_capabilities(layouts: List[Dict[str, Any]], prs) -> Dict[str, Any]:
    ...
-    return {
+    result = {
         "has_footer_placeholders": has_footer,
         "has_slide_number_placeholders": has_slide_number,
         "has_date_placeholders": has_date,
         "layouts_with_footer": layouts_with_footer,
         "layouts_with_slide_number": layouts_with_slide_number,
         "layouts_with_date": layouts_with_date,
         "total_layouts": len(layouts),
         "total_master_slides": len(prs.slide_masters),
-        "per_master": list(per_master_stats.values()),
-        "per_master": list(per_master_stats.values()),
+        "per_master": list(per_master_stats.values()),
         "footer_support_mode": "placeholder" if has_footer else "fallback_textbox",
         "slide_number_strategy": "placeholder" if has_slide_number else "textbox",
         "recommendations": recommendations
     }
+    return result
```

### 2) Remove unused imports

```diff
- from pptx.util import Inches
- from io import BytesIO
- from core.powerpoint_agent_core import PowerPointAgent, PowerPointAgentError
+ from core.powerpoint_agent_core import PowerPointAgentError
```

### 3) Prefer original_index in summary for clarity

```diff
-        lines.append(f"  [{layout['index']}] {layout['name']} ({ph_count} placeholder{'s' if ph_count != 1 else ''})")
+        display_idx = layout.get('original_index', layout['index'])
+        lines.append(f"  [{display_idx}] {layout['name']} ({ph_count} placeholder{'s' if ph_count != 1 else ''})")
```

### 4) Add metadata for timeout ceiling (traceability)

```diff
result = {
    "status": "success",
    "metadata": {
        ...
+       "timeout_seconds": timeout_seconds,
    },
    ...
}
```

### 5) Add consolidated info/warning for non-RGB theme references

```diff
colors = {}
...
for color_name in color_attrs:
    ...
        else:
            colors[color_name] = f"schemeColor:{color_name}"
+           non_rgb_found = True
...
if not colors:
    warnings.append("Theme color scheme unavailable or empty")
+elif 'non_rgb_found' in locals() and non_rgb_found:
+    warnings.append("Theme colors include scheme references without explicit RGB")
```

---

## Performance and safety notes

- **Transient instantiation cleanup:** Your `_add_transient_slide` generator and defensive deletion are correct for in-memory probing and not saving back to disk. The checksum guard closes the loop on atomicity. Good.

- **Max layouts behavior:** Slicing handled in analysis function. For clarity and predictability, including `metadata.layout_count_total` vs `metadata.layout_count_analyzed` would make partial probes explicit.

- **Warnings hygiene:** You already avoid warning spam for per-master theme extraction. Good. The consolidated non-RGB color warning above complements this approach.

---

## QA checklist to run post-fix

- **Schema validation:**
  - Run success path and error path; ensure duplicate `per_master` is gone.
  - Confirm new `metadata.timeout_seconds` appears when set.

- **Index clarity test:**
  - Use `--max-layouts 3` on a template with >3 layouts; confirm summary shows `original_index`.

- **Theme extraction tests:**
  - On a theme with non-RGB scheme references, confirm single warning appears.
  - On a theme with RGB colors, confirm warning is absent.

- **Atomic verification test:**
  - Check md5 before/after across deep and essential modes; confirm match.

- **Timeout test:**
  - Use a low `--timeout` with a large template; confirm `analysis_complete=false` and timeout warning is present.

---

## Final recommendations

- Implement the 5 diffs above to remove redundancy, tighten clarity, and enhance auditability.
- Consider adding:
  - `metadata.layout_count_total` and `metadata.layout_count_analyzed` for large templates.
  - A small “capability summary” block in `info` (e.g., “footer=placeholder, slide_numbers=textbox”) to aid quick scan.
- Keep the current defensive approach to internal pptx structures; it's pragmatic and well-contained.

