# PowerPoint capability probe v1.1.0 review

You’ve leveled this tool up. It’s approaching “production artifact” quality. Below is a meticulous validation, pinpoint critiques, and concrete refactors to take it all the way.

---

## Contract and CLI behavior

- **JSON contract completeness:** Required fields are present and validated; error-path JSON mirrors success shape with status/error/error_type/metadata — solid.
- **Mutual exclusivity:** Summary vs JSON is enforced correctly in the latest file. One variant earlier allowed both; your final variant fixes it. Keep the strict parser.error guard only.
- **Operation metadata:** UUID, duration_ms, library versions, checksum, deep_analysis, atomic_verified — all there and consistent.
- **Timeout signaling:** You append warnings and set capabilities.analysis_complete. Good. Consider surfacing analysis_complete in metadata too for top-level status visibility.
- **Exit codes:** 0 on success, 1 on error (with JSON error payload). Good.

---

## Atomicity, transient instantiation, and internal APIs

- **Atomic integrity:** MD5 pre/post; error on mismatch. Good.
- **Transient slide lifecycle:** The final variant uses a safe add→map→drop_rel→delete pattern. One earlier helper _add_transient_slide yielded a slide and cleaned up in finally — more robust if reused across loops. Prefer the generator helper for cleanup guarantees and reduce duplicate internal calls.
- **Internal XML pokes:** You safely manipulate `prs.slides._sldIdLst` and `prs.part.drop_rel(rId)`. That’s necessary because python-pptx lacks delete, but it’s tightly coupled to internal APIs. Keep it, but guard with explicit try/except and consistent warnings (you do).
- **Layout limiting:** Final code avoids mutating `prs.slide_layouts._sldLayoutLst` (good). Earlier variants mutated it; that risks undefined behavior. The current approach of slicing a local list is correct.

Recommendations:
- **Refactor transient add/remove into a single utility** to guarantee cleanup even if analysis throws mid-loop.
- **Guard relation removal** with contextual info in warnings (layout name, index, rId) for postmortem triage.

---

## Placeholder detection, positions, and master associations

- **Enum mapping:** Dynamic discovery of PP_PLACEHOLDER enum values via dir() and .value fallback is correct. You return UNKNOWN_<code> for non-mapped — good.
- **Type codes vs names:** You store both type_code and resolved name; capability logic matches on type_code. Perfect for precision.
- **Position accuracy:** Deep mode instantiates slides and aligns instantiated placeholders back to layout placeholders via idx — this is the right strategy to get runtime positions.
- **Master association:** You build a map of layout object ids to master_index. Good for multi-master statistics.

Refinements:
- **Placeholder completeness flag:** Add a per-layout boolean `instantiation_complete` and count of placeholders matched vs expected (layout.placeholders length). Helps detect partial instantiation (e.g., master-level placeholders not realized).
- **Coordinate fidelity:** You present both inches and percent; add EMU values in the placeholder payload for full fidelity and audit consistency with top-level slide_dimensions.
  - Example addition:
    - position_emu: { left: int, top: int }
    - size_emu: { width: int, height: int }
- **Anchor baseline:** Include slide dimensions in each placeholder payload only when needed? Alternatively, keep as-is but consider a single per-layout `slide_reference` block with width/height to reduce repeated fields.

---

## Theme extraction (colors and fonts)

- **Color scheme:** You access `theme.theme_color_scheme` and convert to hex when RGB present; else note `schemeColor:<name>`. This is pragmatic.
- **Font scheme:** You correctly use major/minor with latin/east_asian/complex_script, capturing typeface safely. Fallback to scanning master shapes, then default Calibri with warnings — great resilience.

Recommendations:
- **Expose raw scheme references:** When color has no explicit RGB (theme references), include both:
  - color_value: "#RRGGBB" or null
  - scheme_ref: "accent1" etc.
- **Per-master theme info:** You collect per-master arrays in some variants; the final file retains per_master in theme only in earlier version. Reintroduce `theme.per_master` consistently for multi-master decks:
  - theme: { colors, fonts, per_master: [{ master_index, colors, fonts }] }

---

## Capabilities and recommendations

- **Capability flags:** Footer/slide_number/date detection by type_code across instantiated or template placeholders — correct.
- **Per-master stats:** You compute counts for layouts per master and per capability — valuable for multi-template decks.
- **Recommendations:** Clear and actionable. Consider including the original_index for layouts and a compact list of indices for quick targeting.

Enhancements:
- **Footer support mode:** Add `footer_support_mode: "placeholder" | "fallback_textbox" | "none"` based on detection; this aligns with downstream tool decisions.
- **Slide-number plan:** If no slide-number placeholders, add a `slide_number_strategy: "textbox"` hint to downstream automation.

---

## Robustness, performance, and edge cases

- **Timeout:** You append warning and stop processing; you also set analysis_complete false. Good.
- **Huge templates:** `--max-layouts` slices analyzed list; maintain info message. Consider also sampling masters evenly or prioritizing layouts with common names (Title and Content, Title Slide).
- **Locked/permission:** You probe read access up front. Good.

Performance notes:
- **Repeated theme extraction:** In per-master loop variants you suppressed warnings for noise — keep that approach but include per_master data.
- **Deep mode memory:** Transient slide creation is per-layout; it’s fine. Consider short-circuit when timeout threshold is near to avoid partial cleanup issues.

---

## QA nits and small correctness fixes

- **Consistent field names:**
  - Ensure the final summary and JSON use the same casing and field presence: `schema_version` appears in some variants; keep it in the final for traceability.
- **Function signatures:** Some earlier files used detect_layouts_with_instantiation with differing parameters. Your final file standardizes the signature with timeout and max_layouts; keep only that.
- **Exception uniformity:** When schema validation fails, you set status=error and error_type=SchemaValidationError but still return the result for printing. That’s fine; consider also adding `metadata.operation_id` unchanged for continuity and adding `metadata.validation_failures` array with missing fields.

---

## Suggested refactors (diff-style snippets)

#### 1) Transient slide lifecycle helper (reuse + cleanup guarantees)
```python
def add_and_cleanup_transient_slide(prs, layout):
    slide = None
    idx = None
    try:
        slide = prs.slides.add_slide(layout)
        idx = len(prs.slides) - 1
        return slide, idx
    finally:
        if idx is not None and idx < len(prs.slides):
            try:
                rId = prs.slides._sldIdLst[idx].rId
                prs.part.drop_rel(rId)
                del prs.slides._sldIdLst[idx]
            except Exception:
                pass
```

Usage:
```python
if deep:
    try:
        temp_slide, added_idx = add_and_cleanup_transient_slide(prs, layout)
        # map placeholders...
    except Exception as e:
        warnings.append(f"... {e}")
        # fallback...
```

#### 2) Placeholder payload: add EMU and instantiation completeness
```python
# After computing inches
left_emu = int(shape.left) if hasattr(shape, 'left') else 0
top_emu = int(shape.top) if hasattr(shape, 'top') else 0
width_emu = int(shape.width) if hasattr(shape, 'width') else 0
height_emu = int(shape.height) if hasattr(shape, 'height') else 0

payload.update({
  "position_emu": {"left": left_emu, "top": top_emu},
  "size_emu": {"width": width_emu, "height": height_emu}
})
```

Per layout:
```python
layout_info["instantiation_complete"] = len(placeholders) == len(layout.placeholders)
layout_info["placeholder_expected"] = len(layout.placeholders)
layout_info["placeholder_instantiated"] = len(placeholders)
```

#### 3) Theme per-master consistency
```python
theme_per_master = []
for m_idx, master in enumerate(prs.slide_masters):
    theme_per_master.append({
        "master_index": m_idx,
        "colors": extract_theme_colors(master, []),
        "fonts": extract_theme_fonts(master, [])
    })

result["theme"]["per_master"] = theme_per_master
```

#### 4) Capability strategy hints
```python
capabilities["footer_support_mode"] = "placeholder" if capabilities["has_footer_placeholders"] else "fallback_textbox"
capabilities["slide_number_strategy"] = "placeholder" if capabilities["has_slide_number_placeholders"] else "textbox"
```

---

## Final verdict

- **Production-ready core:** Yes — consistent JSON contract, atomicity verified, deep-mode accuracy via transient slides, robust theme extraction, multi-master support.
- **Gaps to close:** Per-master theme consistently included, placeholder payloads enriched with EMU and instantiation completeness, consolidated transient lifecycle helper, explicit strategy hints for downstream tools, and metadata.analysis_complete mirrored at top-level for quick pipeline checks.

If you want, I’ll produce a minimal PR patch set with the four refactors above, plus a quick test that asserts instantiation completeness and per-master theme presence.
