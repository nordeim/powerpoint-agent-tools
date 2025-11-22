# Review of ppt_capability_probe.py v1.1.0 implementation

You did the hard, essential work: closed the JSON contract, removed placeholder guessing, added deep instantiation, hardened theme extraction, and surfaced operational signals. The architecture aligns with your stateless, inspection-first philosophy and is close to production-grade. Below is a meticulous critique with precise, actionable fixes and micro-patches to tighten robustness, correctness, and orchestration hygiene.

---

## Contract, CLI, and orchestration hygiene

- **Mutual exclusivity validation missing:** You intend summary and JSON to be exclusive, but the guard `if args.summary and args.output_json and len(sys.argv) > 3: pass` is a no-op. This can break pipelines.

  - Patch:
    ```python
    if args.summary and args.output_json:
        print(json.dumps({
            "status": "error",
            "error": "Flags --summary and --json are mutually exclusive",
            "error_type": "ArgumentError",
            "metadata": {
                "tool_version": "1.1.0",
                "probed_at": datetime.now().isoformat()
            }
        }, indent=2))
        sys.exit(1)
    ```

- **Stdout discipline:** You print JSON or summary on stdout; good. Keep all logging off stdout. Avoid printing anything before JSON (including import errors)—you’re already compliant.

- **Operation metadata completeness:** You added `library_versions`, `operation_id`, and `duration_ms`. Consider adding an optional `hostname` (non-sensitive) to help correlate distributed runs.

  - Optional:
    ```python
    import socket
    host = socket.gethostname()
    # ...
    "hostname": host
    ```

---

## Placeholder type mapping and correctness

- **Enum use is right, but coverage is partial:** Your `known_types` list may miss valid enum members on some Office versions (e.g., `CONTENT`, `CENTERED_TITLE`, `VERTICAL_TITLE`, `MEDIA_CLIP` variations). Building the map from the enum itself is safer.

  - Patch:
    ```python
    def build_placeholder_type_map() -> Dict[int, str]:
        type_map = {}
        for name in dir(PP_PLACEHOLDER):
            if name.isupper():
                enum_val = getattr(PP_PLACEHOLDER, name, None)
                try:
                    type_map[int(enum_val)] = name
                except Exception:
                    pass
        return type_map
    ```

- **Capability detection depends on codes existing:** You derive `footer_type_code`, `slide_number_type_code`, and `date_type_code` by scanning the map. If a member is missing, you silently degrade. Add an explicit warning when any canonical type is absent.

  - Patch:
    ```python
    if footer_type_code is None:
        warnings.append("FOOTER placeholder enum not found; capability detection may be incomplete")
    ```

---

## Deep instantiation and mutation risks

- **Transient slide approach inside the original Presentation is risky:** You add and then remove slides by manipulating private XML structures (`_sldIdLst`, `drop_rel`). If an exception occurs between add/drop, the in-memory model is mutated and could affect subsequent probes or future operations—even if the file checksum is unchanged. Also, you’re probing the same object you loaded from disk.

  - Safer approach: instantiate in a separate in-memory clone. python-pptx doesn’t provide a formal clone API, but you can open from bytes to avoid touching the original object.

  - Patch (conceptual):
    ```python
    from io import BytesIO

    def load_prs_bytes(path: Path):
        with open(path, 'rb') as f:
            return BytesIO(f.read())

    prs_bytes = load_prs_bytes(filepath)
    prs_for_instantiation = Presentation(prs_bytes)  # isolated copy

    # Use prs_for_instantiation for deep instantiation
    # Keep original prs for non-mutating reads
    ```

  - If you prefer your current method, add a guard to ensure cleanup is robust:
    ```python
    try:
        temp_slide = prs.slides.add_slide(layout)
        # ... collect data ...
    finally:
        # Clean even on exceptions
        try:
            rId = prs.slides._sldIdLst[-1].rId
            prs.part.drop_rel(rId)
            del prs.slides._sldIdLst[-1]
        except Exception as cleanup_err:
            warnings.append(f"Cleanup after instantiation failed: {cleanup_err}")
    ```

- **Layout limiting by slicing internal lists:** You truncate `prs.slide_layouts._sldLayoutLst`. That’s a mutation of the in-memory model which could affect subsequent logic. Prefer limiting via indices in your loop rather than altering internal lists.

  - Patch:
    ```python
    def detect_layouts_with_instantiation(prs, slide_width, slide_height, deep, warnings, max_layouts=None):
        layouts = []
        for idx, layout in enumerate(prs.slide_layouts):
            if max_layouts is not None and idx >= max_layouts:
                warnings.append(f"Stopped at max_layouts={max_layouts}")
                break
            # existing logic...
    ```

  - Update the call site accordingly and remove `_sldLayoutLst` slicing.

---

## Multiple masters accounting and associations

- **You report `total_master_slides` but don’t associate layouts with masters:** In multi-master decks, layout names can collide; operational planning benefits from per-master grouping.

  - Improvement:
    - Add master index for each layout:
      ```python
      master_idx = None
      for m_idx, master in enumerate(prs.slide_masters):
          if layout in master.slide_layouts:
              master_idx = m_idx
              break
      layout_info["master_index"] = master_idx
      ```
    - Provide per-master stats:
      ```python
      "masters": [
        {"index": 0, "layout_count": 7},
        {"index": 1, "layout_count": 5}
      ]
      ```

---

## Theme extraction robustness

- **Colors:** Accessing `theme_color_scheme` is fine for common themes. Some templates have custom schemes or missing members. You already emit warnings if empty; enhance with explicit mapping of expected keys and a “missing” list.

  - Patch:
    ```python
    missing = []
    for color_name in color_attrs:
        color = getattr(color_scheme, color_name, None)
        if color:
            colors[color_name] = rgb_to_hex(color)
        else:
            missing.append(color_name)
    if missing:
        warnings.append(f"Theme colors missing keys: {', '.join(missing)}")
    ```

- **Fonts:** Using `font_scheme` is correct. Ensure you normalize to plain strings; some environments return objects whose repr looks like strings.

  - Patch:
    ```python
    def font_name_safe(val):
        return str(val) if val else None
    fonts['heading'] = font_name_safe(getattr(font_scheme.major_font, 'latin', None))
    fonts['body'] = font_name_safe(getattr(font_scheme.minor_font, 'latin', None))
    ```

---

## Dimensions and aspect ratio fidelity

- **Aspect ratio calculation bounds:** Using a 0.01 tolerance in inches is fine, but converting via pixels and then fraction can produce simplified ratios that look precise but aren’t (e.g., custom page sizes). Return both the raw float and the canonical string for clarity.

  - Patch:
    ```python
    "aspect_ratio_value": round(ratio, 6),
    "aspect_ratio": aspect_ratio
    ```

- **EMU exposure:** You expose EMUs as integers; good. Consider explicitly labeling units in keys (you do) and include centimeters for international teams.

  - Optional:
    ```python
    "width_cm": round(width_inches * 2.54, 2),
    "height_cm": round(height_inches * 2.54, 2),
    ```

---

## Output validation and warnings

- **Validator reports missing fields via warnings but doesn’t re-validate after append:** You append a validation warning and return success. That’s okay for non-critical gaps, but it may hide schema regressions.

  - Suggestion: If required fields are missing, return `"status": "error"` with a clear message. If they’re optional, keep as warnings.

  - Patch:
    ```python
    is_valid, missing_fields = validate_output(result)
    if not is_valid:
        result["status"] = "error"
        result["error"] = "Output schema validation failed"
        result["error_type"] = "SchemaValidationError"
        result["warnings"].append(f"Missing: {', '.join(missing_fields)}")
        return result
    ```

---

## Error handling and edge cases

- **Permission errors:** You read one byte to test permission, good. Also catch `OSError` for network shares and locked files with non-PermissionError variants.

  - Patch:
    ```python
    try:
        with open(filepath, 'rb') as f:
            f.read(1)
    except (PermissionError, OSError) as e:
        raise PermissionError(f"Unable to read file: {filepath} ({e})")
    ```

- **Timeout for deep analysis:** You document timeout protection but don’t enforce one. Add an optional `--timeout-ms` and abort deep loop gracefully with partial results and a warning.

  - Patch (simplified):
    ```python
    def probe_presentation(..., timeout_ms: Optional[int] = None):
        start = time.perf_counter()
        # inside detect loop:
        if timeout_ms and (time.perf_counter() - start) * 1000 > timeout_ms:
            warnings.append("Deep analysis timed out; results are partial")
            break
    ```

- **Summary mode dimensions:** You print `DPI Estimate` hardcoded as 96, not an issue, but if future probes change DPI logic, reflect that in summary mode from the computed value.

---

## Micro-nits and consistency

- **Imports:** You import `Inches` but don’t use it in this file; remove unused imports.
- **String casing:** In summary, you mix ✓/✗ and textual labels—consistent and fine. Keep stable for monitoring parsing.
- **Top-level info/warnings:** Good. Consider adding `warnings_count` and `info_count` to metadata to avoid scanning arrays in orchestrators.

  - Patch:
    ```python
    "metadata": {
      # ...
      "warnings_count": len(warnings),
      "info_count": len(info)
    }
    ```

---

## Testing recommendations

- **Placeholder enum coverage:** Validate detection across default Office themes and custom templates with unusual placeholder mixes (e.g., vertical title layouts).
- **Multi-master decks:** Confirm per-master counts; ensure layout indices refer to global list, and master_index disambiguates.
- **Deep vs essential parity:** Verify that capability flags are identical between modes; deep only adds positional fidelity.
- **Instantiation cleanup:** Simulate exceptions during deep instantiation to test cleanup warnings and no residual mutation.
- **Large deck performance:** Measure with `--max-layouts` and `--timeout-ms` to ensure graceful partial results.

---

## Verdict

Strong release that fixes critical correctness and contract issues. Address deep instantiation mutation risks (prefer in-memory copy), remove internal list slicing, tighten mutual exclusivity validation, expand enum mapping coverage, and add per-master associations. With these refinements, you’ll have a robust, auditable probe that reliably informs agent decisions under real-world templates and operational constraints.

If you want, I can provide a compact PR with:
- Enum map via reflection
- Clean deep instantiation using a BytesIO clone
- Max-layouts without internal slicing
- Mutual exclusivity enforcement
- Per-master layout association
- Enhanced theme warnings and aspect ratio metadata
- Optional timeout parameter and counts in metadata
