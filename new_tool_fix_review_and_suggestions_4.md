# PowerPoint capability probe v1.1.0 review

You’ve sharpened this tool into something agents can actually rely on. The design choices show care for atomicity, schema contracts, and real template introspection. Below is a meticulous review: what’s solid, what’s risky, and what I would change with concrete, code-level fixes.

---

## Strengths that are production-grade

- **JSON contract consistency:** Top-level status, metadata.operation_id, metadata.duration_ms, and library versions are present. Error path mirrors structure well enough for automated pipelines.
- **Placeholder detection:** Uses PP_PLACEHOLDER from python-pptx with a dynamic map. No numeric guessing. Good fallback via UNKNOWN_X.
- **Deep-mode instantiation:** Transient slide lifecycle is handled defensively with cleanup. Accurate runtime positions via instantiated placeholders mapped by idx is exactly the right approach.
- **Atomicity & verification:** Pre/post MD5 checksums with violation raising is an excellent operational gate.
- **Theme extraction:** Font scheme and color scheme extraction implemented with graceful fallbacks, plus per-master theme reporting without warning spam.
- **Capability analysis:** Enumerates footer/slide number/date placeholders with layout references and master associations (per_master stats). Recommendations are pragmatic.
- **Timeouts & limits:** Timeout protection and max-layouts constraints with info/warnings to callers. Good pattern for large templates.

---

## Critical issues to fix before broader rollout

### 1) Hidden mutation risk on “max layouts”
- **Issue:** In some versions, you slice internal pptx structures:
  - `prs.slide_layouts._sldLayoutLst = prs.slide_layouts._sldLayoutLst[:max_layouts]`
- **Impact:** This mutates in-memory relationships and can risk output state on save (even though you don’t save in probe). It violates the spirit of read-only inspection and risks version-specific python-pptx internals.
- **Fix:** Never reassign internal lists. Instead, only slice the list you iterate over, not the underlying XML structures.

  Example:
  ```python
  all_layouts = list(prs.slide_layouts)
  if max_layouts and len(all_layouts) > max_layouts:
      info.append(f"Limited analysis to first {max_layouts} of {len(all_layouts)} layouts")
      layouts_to_process = all_layouts[:max_layouts]
  else:
      layouts_to_process = all_layouts
  # pass layouts_to_process to detect_layouts_with_instantiation(...)
  ```

### 2) Deep-mode cleanup and slide index reliability
- **Issue:** Some versions compute `added_idx = len(prs.slides) - 1` on deletion, which is safe when only one slide is added. But the safer approach is to retain the exact sldId index captured during addition.
- **Fix:** Your generator `_add_transient_slide` is the right model. Prefer that everywhere you instantiate.

  Example (already present in best version):
  ```python
  for temp_slide in _add_transient_slide(prs, layout):
      # ...work...
  # generator ensures exact cleanup by stored index
  ```

### 3) Placeholder enumeration: layout vs master
- **Issue:** `temp_slide.shapes` vs `temp_slide.placeholders`. You correctly map instantiated placeholders by idx, but relying on `temp_slide.shapes` alone can miss master placeholders not instantiated. You handled this by iterating layout.placeholders and stitching from instantiated_map — keep that pattern consistently.
- **Fix:** Ensure every deep-mode path builds from `layout.placeholders` and then sources geometry from `instantiated_map` where available. You’ve done this well in the latest file; make that canonical.

### 4) Theme color conversion robustness
- **Issue:** Not all scheme colors are direct RGB; some can be theme references (tints, shades). You guard with `hasattr(color, 'r')` and fallback to “schemeColor:name”. Good. But callers may expect hex consistently.
- **Fix:** Add an optional normalization flag or document that non-RGB scheme colors are returned as a scheme reference string. If you can resolve theme variants (tints/shades), map them, otherwise explicitly note in warnings once: “Some theme colors are scheme references, not resolved to hex.”

### 5) Font extraction: latin/east_asian/complex_script handling
- **Issue:** In some paths, font values may be objects not strings. You already use `getattr(latin, 'typeface', str(latin))`. Keep this across major/minor fonts consistently, and set an explicit precedence: latin > east_asian > complex_script.
- **Fix:** Normalize:

  ```python
  def _font_name(font_obj):
      return getattr(font_obj, 'typeface', str(font_obj)) if font_obj else None

  fonts['heading'] = _font_name(getattr(major, 'latin', None)) or \
                     _font_name(getattr(major, 'east_asian', None)) or \
                     _font_name(getattr(major, 'complex_script', None))
  ```

### 6) Error-path schema parity
- **Issue:** In the final main() error branch, you return:
  ```
  {
    "status": "error",
    "error": "...",
    "error_type": "...",
    "metadata": {...}
  }
  ```
  But you omit slide_dimensions/layouts/theme/capabilities/warnings/info, so downstream code expecting top-level arrays must branch. It’s acceptable, but document this contract and guarantee metadata presence.
- **Fix:** Add a single `warnings` array at top level in error path to maintain operational signals, even if empty.

---

## High-value improvements (clarity, resilience, and UX)

### CLI and output behavior
- **Mutual exclusivity enforcement:** Your best version enforces `--summary` XOR `--json` via parser.error — keep this. In earlier draft, you had a “pass” branch; remove that entirely.
- **Default behavior:** Good default to JSON when neither flag is set. Keep it.
- **Include analysis completeness flag:** You added `capabilities["analysis_complete"]` when timeout occurs. Keep this; it’s operational gold for agents.

### Aspect ratio reporting
- **Current:** You estimate DPI and produce pixel dims. Solid. You also include `aspect_ratio_float` in the best version.
- **Enhancement:** Always include both a canonical ratio string and float value:
  - `aspect_ratio_string` (e.g., “16:9” or reduced fraction)
  - `aspect_ratio_float` (e.g., 1.7778)

### Multiple masters reporting
- **Current:** You compute per-master stats in the best version and attach `capabilities.per_master`. Good.
- **Enhancement:** In summary mode, add a short “Masters” section showing layout counts and placeholder presence per master so humans can pick the right layout.

### Warnings and info hygiene
- **Current:** You avoid flooding warnings for per-master theme extraction; smart.
- **Enhancement:** Add deduplication on warnings to avoid repeating the same message.

---

## Security, performance, and maintenance notes

- **Atomicity:** Checksums provide strong assurance. For very large files, MD5 costs are acceptable for probe operations; document the performance impact at scale.
- **Performance:** Timeout protection inside `detect_layouts_with_instantiation` is good. Consider exposing a `--stop-on-timeout` flag vs partial results (you already return partial with a warning).
- **Internal APIs:** Avoid touching private attributes (`_sldIdLst`, `_sldLayoutLst`) in any future refactor unless unavoidable. Your generator approach minimizes this risk.

---

## Tests and validation to add

- **Locked file handling:** Ensure PermissionError path prints consistent error JSON with metadata and “warnings: []”.
- **Multi-master template:** Assert per_master stats match layout associations.
- **Theme-less decks:** Validate fonts/colors fallbacks and single warning.
- **Timeout scenario:** Simulate slow deep instantiation, verify `analysis_complete` false and partial layouts returned.
- **Schema validation:** You already run validate_output and flip status to “error” if missing fields — add a test that asserts the warning for schema errors includes the specific missing paths.

---

## Concrete diffs I recommend

#### Remove internal layout list mutation
```diff
- prs.slide_layouts._sldLayoutLst = prs.slide_layouts._sldLayoutLst[:max_layouts]
+ # Do not mutate internal pptx structures; only slice iteration list
+ layouts_to_process = list(prs.slide_layouts)[:max_layouts]
+ layouts = detect_layouts_with_instantiation(
+     prs, slide_width, slide_height, deep, warnings,
+     timeout_start=start_time, timeout_seconds=timeout_seconds,
+     max_layouts=max_layouts
+ )
```

#### Error-path: include warnings
```diff
error_result = {
    "status": "error",
    "error": str(e),
    "error_type": type(e).__name__,
    "metadata": {
        "file": str(args.file) if args.file else None,
        "tool_version": "1.1.0",
        "operation_id": str(uuid.uuid4()),
        "probed_at": datetime.now().isoformat()
    },
+   "warnings": []
}
```

#### Font normalization helper
```python
def _font_name(font_obj):
    return getattr(font_obj, 'typeface', str(font_obj)) if font_obj else None
```

Apply in extract_theme_fonts as noted above.

---

## Bottom line

Direct answer: This v1.1.0 is near production-ready. Fix the internal layout list mutation in older variants, standardize deep-mode cleanup via the generator everywhere, and tighten font/color normalization and error-path signals. With those changes, you’ll have a robust, auditable probe that agents can depend on for template capability discovery at scale.
