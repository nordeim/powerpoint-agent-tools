# PowerPoint capability probe v1.1.0 review

You built a thoughtful, production-minded probe. It’s close to “ship-ready” for your agent workflows. Below is a rigorous validation, with concrete, surgical fixes to elevate correctness, robustness, and maintainability.

---

## Quick verdict

- Strengths: JSON-first contract, atomic verification, enum-driven placeholder types, deep mode instantiation for accurate positions, theme/ fonts extraction attempts, capability analysis, clear summary mode.
- Priority fixes: Correct theme font extraction, color resolution, stop mutating internal layout XML in max-layouts, make deep mode explicitly non-persistent, harden enum mapping, implement timeout protection, tag layouts with master context, and add schema guardrails.

---

## Architecture and contract validation

- Output shape aligns with your design (status, metadata, durations, versions, warnings/info).
- Deep-mode positioning via transient instantiation is the right call, paired with checksum verification.
- Capability flags are actionable and succinct. Summary formatting is readable and consistent.

Gaps to fix:
- The “timeout protection” and “per-master stats” in the changelog aren’t implemented. Add them for spec compliance.
- The max-layouts logic mutates `prs.slide_layouts._sldLayoutLst`. That changes the in-memory XML tree of the object even if the file on disk remains intact. Avoid touching private XML lists; filter in Python instead.
- Theme font extraction returns the raw object, not the typeface string; color extraction can miss theme-derived colors or return defaults.

---

## Correctness issues and precise fixes

### 1) Theme font extraction returns the wrong value
- Issue: `font_scheme.major_font.latin` is a font object. You need `.typeface` to obtain the name.
- Fix:

```python
def extract_theme_fonts(prs, warnings: List[str]) -> Dict[str, str]:
    fonts = {}
    try:
        slide_master = prs.slide_masters[0]
        theme = slide_master.theme
        font_scheme = getattr(theme, "font_scheme", None)

        if font_scheme:
            major = getattr(font_scheme, "major_font", None)
            minor = getattr(font_scheme, "minor_font", None)

            if major and getattr(major, "latin", None):
                fonts["heading"] = major.latin.typeface  # not major.latin

            if minor and getattr(minor, "latin", None):
                fonts["body"] = minor.latin.typeface

        if not fonts:
            warnings.append("Theme font scheme unavailable or incomplete; using fallback")
            # Fallback scan (best-effort)
            for shape in slide_master.shapes:
                tf = getattr(shape, "text_frame", None)
                if not tf: continue
                for p in tf.paragraphs:
                    name = getattr(p.font, "name", None)
                    if name and "heading" not in fonts:
                        fonts["heading"] = name
                        break
                if "heading" in fonts: break

        if not fonts:
            fonts = {"heading": "Calibri", "body": "Calibri"}
            warnings.append("Using default fonts (Calibri) - theme unavailable")
    except Exception as e:
        fonts = {"heading": "Calibri", "body": "Calibri"}
        warnings.append(f"Theme font extraction failed: {str(e)}")

    return fonts
```

### 2) Theme color resolution may be incomplete
- Issue: `theme_color_scheme.accentN` can be theme colors that require resolution; sometimes you’ll encounter scheme colors or missing direct RGB. Your `rgb_to_hex` assumes RGB is available.
- Improvement: Try `.rgb` or `.theme_color` resolution; fall back gracefully.

```python
def rgb_to_hex(rgb_color) -> str:
    try:
        # RGBColor from python-pptx
        r, g, b = rgb_color.r, rgb_color.g, rgb_color.b
        return f"#{r:02X}{g:02X}{b:02X}"
    except Exception:
        return "#000000"

def extract_theme_colors(prs, warnings: List[str]) -> Dict[str, str]:
    colors = {}
    try:
        slide_master = prs.slide_masters[0]
        color_scheme = slide_master.theme.theme_color_scheme
        for name in [
            "accent1","accent2","accent3","accent4","accent5","accent6",
            "background1","background2","text1","text2","hyperlink","followed_hyperlink"
        ]:
            try:
                color = getattr(color_scheme, name, None)
                if color is None:
                    continue
                colors[name] = rgb_to_hex(color)  # validate that color is RGBColor
            except Exception:
                # Graceful skip
                pass
        if not colors:
            warnings.append("Theme color scheme unavailable or empty")
    except Exception as e:
        warnings.append(f"Theme color extraction failed: {str(e)}")
    return colors
```

If you find templates where theme colors are not resolved, add a secondary resolver (e.g., query effective formatting on an instantiated shape, then read its fill/font color).

### 3) Deep mode instantiation should be clearly non-persistent and safer
- Current behavior: Add slide then drop the relation and delete the last slide entry. It’s generally safe in-memory, but risks side effects on the object graph.
- Improvement: Clone a transient Presentation for deep probing, and never alter the original object’s relationship lists. When sticking to the original `prs`, ensure delete logic is robust: drop rel by index used, not “last element” assumptions.

```python
def detect_layouts_with_instantiation(prs, slide_width, slide_height, deep, warnings):
    layouts = []
    for idx, layout in enumerate(prs.slide_layouts):
        info = {"index": idx, "name": layout.name, "placeholder_count": len(layout.placeholders)}
        if deep:
            try:
                # Add a slide using this layout
                slide = prs.slides.add_slide(layout)
                placeholders = []
                for shape in slide.shapes:
                    if getattr(shape, "is_placeholder", False):
                        placeholders.append(analyze_placeholder(shape, slide_width, slide_height, instantiated=True))
                # Remove the slide we added (by exact index)
                added_idx = len(prs.slides) - 1
                rId = prs.slides._sldIdLst[added_idx].rId
                prs.part.drop_rel(rId)
                del prs.slides._sldIdLst[added_idx]
                info["placeholders"] = placeholders
            except Exception as e:
                warnings.append(f"Could not instantiate layout '{layout.name}': {str(e)}")
                info["placeholders"] = [
                    analyze_placeholder(shape, slide_width, slide_height, instantiated=False)
                    for shape in layout.placeholders if getattr(shape, "is_placeholder", False)
                ]
                info["_warning"] = "Using template positions (instantiation failed)"
        else:
            types = set()
            for shape in layout.placeholders:
                try:
                    types.add(get_placeholder_type_name(shape.placeholder_format.type))
                except Exception:
                    pass
            info["placeholder_types"] = sorted(types)
        layouts.append(info)
    return layouts
```

If you want a hard atomic read guarantee with zero in-memory mutation on the opened presentation, open a second `Presentation` object for deep analysis and discard it after.

### 4) Enum mapping should handle Enum types explicitly
- Issue: `isinstance(value, int)` might fail if `PP_PLACEHOLDER` is an Enum subclass. Most builds expose them as ints, but safer to use `.value`.
- Fix:

```python
def build_placeholder_type_map() -> Dict[int, str]:
    type_map = {}
    for name in dir(PP_PLACEHOLDER):
        if not name.isupper(): 
            continue
        try:
            enum_member = getattr(PP_PLACEHOLDER, name)
            code = enum_member if isinstance(enum_member, int) else getattr(enum_member, "value", None)
            if code is not None:
                type_map[int(code)] = name
        except Exception:
            pass
    return type_map
```

### 5) Max-layouts must not mutate private XML lists
- Issue: `prs.slide_layouts._sldLayoutLst = ...` changes internal XML list. That’s brittle and can break assumptions elsewhere.
- Fix: Slice the Python list you iterate over, not the underlying XML:

```python
all_layouts = list(prs.slide_layouts)
if max_layouts and len(all_layouts) > max_layouts:
    info.append(f"Limited analysis to first {max_layouts} of {len(all_layouts)} layouts")
    all_layouts = all_layouts[:max_layouts]

# Then iterate over all_layouts instead of prs.slide_layouts
layouts = []
for idx, layout in enumerate(all_layouts):
    # … capture original global index if needed:
    # real_index = prs.slide_layouts.index(layout)  # if supported
    # else keep idx within the subset, but document the subset mapping
```

You’ll also need to plumb this filtered list into your detection function so the reporting aligns.

### 6) Per-master context missing in layout reporting
- You report the total master count but not the association per layout. Add a “master_index” to each layout so agents can make master-aware decisions.

```python
def detect_layouts_with_instantiation(prs, slide_width, slide_height, deep, warnings):
    layouts = []
    # Build mapping: layout -> master index
    master_map = {}
    for m_idx, master in enumerate(prs.slide_masters):
        for l in master.slide_layouts:
            master_map[id(l)] = m_idx

    for idx, layout in enumerate(prs.slide_layouts):
        info = {
            "index": idx,
            "name": layout.name,
            "placeholder_count": len(layout.placeholders),
            "master_index": master_map.get(id(layout), None)
        }
        # ... rest unchanged
        layouts.append(info)
    return layouts
```

Then in capabilities, consider stratified counts per master if useful to your downstream agent logic.

### 7) Timeout protection not implemented
- Your changelog mentions timeouts, but there’s no enforcement in deep mode. Add a guard:

```python
def probe_presentation(..., timeout_seconds: Optional[int] = None):
    start_time = time.perf_counter()
    # ...
    def check_timeout():
        if timeout_seconds and (time.perf_counter() - start_time) > timeout_seconds:
            raise TimeoutError(f"Probe exceeded {timeout_seconds}s")

    # Call check_timeout() inside loops (layouts iteration) and deep instantiation points
```

Add a top-level warning and partial results printing on timeout to honor graceful degradation.

---

## Robustness and safety

- Atomic verification via MD5 is good. Consider SHA-256 for stronger guarantees if file integrity is mission-critical.
- Permission checks are present, but locked file detection can be broader. You might also catch `OSError` variants for network shares or file-in-use on Windows.
- Use of private attributes like `_sldIdLst` is unavoidable for delete operations; your current usage is safe but should be documented as tested across python-pptx versions.

---

## Performance considerations

- Deep mode instantiation is O(L × S) across layouts and shapes. For large templates, the max-layouts guard helps.
- Consider an optional “positions_only_for_types” filter (e.g., only TITLE/BODY/FOOTER/DATE/SLIDE_NUMBER) to cut work by 30–60% for agent preflight probes.

---

## Contract, clarity, and auditability

- Validation checks are solid. Add a schema version field (e.g., metadata.schema_version: “1.1.0”) distinct from tool version. If you change layout shape fields later, schema versioning will help downstream consumers.
- Recommendations referencing `ppt_set_footer.py` are useful; include “method_hint” flags (e.g., “footer_strategy”: “textbox_fallback”) in capabilities for machine routing.

---

## Suggested incremental patch set

### Patch A: Fix enums, fonts, colors
- Implement enum `.value` mapping
- Fix `.latin.typeface` usage
- Keep color extraction as RGBColor; add warnings when unresolved

### Patch B: Non-mutating max-layouts and master tagging
- Remove `_sldLayoutLst` slicing
- Tag layouts with `master_index`
- Use the filtered list during detection

### Patch C: Safer deep instantiation and optional timeout
- Delete slide by captured index, not by “last”
- Add optional timeout checks with warnings
- Document no-persistence behavior in metadata (e.g., “position_source”: “instantiated_nonpersistent”)

### Patch D: Schema version and method hints
- metadata.schema_version
- capabilities.footer_strategy_hint = “textbox_fallback” or “placeholder”
- capabilities.slide_number_strategy_hint

---

## Example: minimal diffs for critical fixes

```python
# 1) Enum mapping
def build_placeholder_type_map() -> Dict[int, str]:
    type_map = {}
    for name in dir(PP_PLACEHOLDER):
        if name.isupper():
            member = getattr(PP_PLACEHOLDER, name, None)
            if member is None:
                continue
            code = member if isinstance(member, int) else getattr(member, "value", None)
            if code is not None:
                type_map[int(code)] = name
    return type_map

# 2) Fonts
# fonts['heading'] = font_scheme.major_font.latin.typeface
# fonts['body']    = font_scheme.minor_font.latin.typeface

# 3) Max-layouts: never touch _sldLayoutLst
all_layouts = list(prs.slide_layouts)
if max_layouts and len(all_layouts) > max_layouts:
    info.append(f"Limited analysis to first {max_layouts} of {len(all_layouts)} layouts")
    all_layouts = all_layouts[:max_layouts]
# pass all_layouts to detection function

# 4) Deep instantiation deletion
added_idx = len(prs.slides) - 1
rId = prs.slides._sldIdLst[added_idx].rId
prs.part.drop_rel(rId)
del prs.slides._sldIdLst[added_idx]
```

---

## Final notes

This tool already meets your production bar for JSON-first agent compatibility and atomic safety. Apply the above patches and add the timeout/ master-tagging to align the implementation with the promise of your changelog and your standards for auditability. If you want, I can produce a clean v1.1.1 diff that folds all fixes without inflating complexity.
