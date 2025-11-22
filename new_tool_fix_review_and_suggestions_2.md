# PowerPoint capability probe v1.1.0 review

You’ve built a solid, production-grade probe with clear contracts, defensive checks, enum-driven correctness, and deep-mode instantiation for positional accuracy. Below is a meticulous validation and critique with precise fixes, risk notes, and test vectors to harden it further.

---

## Contract and CLI correctness

- **JSON contract (good):** Top-level status, metadata.operation_id, metadata.duration_ms, library versions, warnings/info arrays are present and validated. Output validator enforces presence without overreach. 
- **CLI mode exclusivity (good):** You default to JSON when neither flag is set and explicitly error if both are used. 
- **Timeout and layout limit (good):** You expose both and short-circuit deep analysis if time exceeds the allotted budget, appending warnings.

- **Minor polish:**
  - **Labeling consistency:** Summary shows “Atomic Verified: ✓/✗,” but error output uses “operation_id” and “probed_at” while summary doesn’t explicitly show them. It’s fine, but consider aligning summary content to reduce variance across modes.
  - **Exit code clarity:** Summary and JSON paths both correctly exit 0; error path exits 1. Add a one-line comment near sys.exit calls to keep future contributors disciplined.

---

## Deep-mode instantiation and atomicity

- **Instantiation deletion (good):** You remove the transient slide by retrieving rId at added index and dropping the relationship plus list node. This is the correct python-pptx pattern for safe deletion.
- **Atomic read verification (good):** MD5 pre/post on file and raising PowerPointAgentError on mismatch is the right fail-fast gate.

- **Risk & improvement:**
  - **Internal XML assumptions:** Accessing prs.slides._sldIdLst and prs.part.drop_rel relies on python-pptx internals; future versions could change. Mitigate by:
    - **Improvement:** Encapsulate transient-slide lifecycle in a helper function with explicit try/finally to guarantee cleanup even if downstream analysis raises.
    - **Improvement:** Add a defensive check that length increased by one before deletion; if not, warn and skip deletion to avoid index mismatch in highly unusual edge cases.
  - **Master placeholders not instantiating:** You correctly fall back to template-level shapes and mark position_source. Good. Add a small flag at layout level like “position_mode: mixed” when both instantiated and template placeholders are present to help downstream consumers reason about composite accuracy.

---

## Placeholder and theme fidelity

- **Placeholder type mapping (good):** Dynamic mapping from PP_PLACEHOLDER members, robust against enum value drift, with UNKNOWN fallback labeling.
- **Accurate type usage (good):** Capability detection uses type_code, not guessed numbers.
- **Positions (good):** EMU-to-inches conversion and percentages tied to slide dimensions are correct. You add position_source and size percent; solid.

- **Theme extraction (good):** You correctly navigate theme.theme_color_scheme and theme.font_scheme with safe getattr and fallbacks. Colors are hex-converted via r/g/b.

- **Improvements:**
  - **Slide masters (colors/fonts):** You only read from the first master. In multi-master decks, theme variants may differ. 
    - **Add:** theme.masters array with per-master colors and fonts, plus chosen “primary_master_index” denoting the master that governs most layouts (tie to your master_map and count of associated layouts).
  - **Color conversion robustness:** OOXML theme colors can be scheme-based (tints/shades) or non-RGB (like “schemeColor” references).
    - **Add:** A small resolver that returns a best-effort final RGB, falling back to hex of nearest accent if direct RGB is unavailable, and adds a warning like “Non-RGB theme color references encountered; normalized to scheme base”.
  - **Font scheme richness:** Capture East Asian/Complex script fonts if available:
    - **Add:** fonts.east_asian and fonts.complex if font_scheme.major/minor expose them, else omit keys.

---

## Capabilities and multi-master support

- **Capability analysis (good):** Flags for footer, slide numbers, date; layout refs contain index/name; recommendations are pragmatic.
- **Per-master association (partial):** You compute master_map for layouts, but capabilities don’t show per-master stats.

- **Improvements:**
  - **Per-master stats:** 
    - **Add:** capabilities.per_master = [{master_index, layout_count, has_footer_layouts, has_slide_number_layouts, has_date_layouts}] for targeted agent decisions (e.g., pick the master with slide numbers).
  - **Template coverage metric:**
    - **Add:** coverage ratios: percent of layouts with title, content, two-content, footer, etc., to guide agents toward safer layout choices.

---

## Performance, safety, and edge cases

- **Performance controls (good):** max_layouts and timeout guard deep mode. You append info when limiting layouts.
- **File validation (good):** Existence, is_file, and permission read test before load.

- **Edge cases to cover:**
  - **Very large templates:** 
    - **Add:** If deep mode is enabled and max_layouts is None with layout_count > N (e.g., 30), append a warning and auto-limit to N unless user overrides. This guards runaway time.
  - **Locked/Read-only files on certain OSes:** The initial 1-byte read tests permission, but some OSes allow read while preventing new relations; your workflow is read-only though. Fine, but:
    - **Add:** A warning if Presentation() load raises but you catch earlier permission. Not strictly necessary, just polish.
  - **Non-standard slide sizes:** You compute aspect ratio with Fraction(…, limit_denominator=20). Good. 
    - **Add:** dims.aspect_ratio_float for exact numeric ratio to support precise layout math in agents.

---

## API ergonomics and JSON schema enhancements

- **Schema validation (good):** validate_output ensures required fields exist and appends warnings if missing.
- **Human summary (good):** Clean, focused, with layout enumerations and capability flags.

- **Improvements:**
  - **Schema versioning:** 
    - **Add:** metadata.schema_version: "capability_probe.v1.1.0" so downstream parsers can evolve safely.
  - **Slide layout taxonomy:**
    - **Add:** For each layout, include a normalized role label (e.g., role: "TITLE_AND_CONTENT", "BLANK") inferred from name heuristic or placeholder pattern, with warning if ambiguous. This helps agents pick layouts deterministically.
  - **Indices consistency:** Ensure layout indices in layouts match original order even when max_layouts applied; you already slice layouts_to_process but keep index from enumerate after slicing. Good; note that indices then map to the sliced set, not absolute. 
    - **Add:** original_index to preserve absolute indices when max_layouts is used.

---

## Concrete code-level recommendations

- **Encapsulate transient slide lifecycle:**
  - Create helper add_and_cleanup_slide(prs, layout) that:
    - Adds slide, returns slide and added index
    - Yields control to analysis
    - In finally: drop_rel and del _sldIdLst with defensive checks and warning if mismatch
- **Per-master reporting:**
  - Build master_stats during master_map construction:
    - master_stats[m_idx] = {layout_count, has_footer, has_slide_number, has_date}
  - Emit capabilities.per_master and reference in recommendations, e.g., “Master 0: 8 layouts, footer on 3; Master 1: 4 layouts, no footer.”
- **Theme robustness:**
  - Attempt to resolve scheme colors with fallbacks; when unresolved, annotate colors["_warnings"] and include count.
- **Indices preservation:**
  - In detect_layouts_with_instantiation, add original_index captured before slicing and keep index as positional in analysis order.

---

## Test vectors to harden behavior

- **Multi-master deck:** 
  - Expect per-master stats populated, master_index present in each layout, recommendations reference masters.
- **Theme with scheme colors only:** 
  - Ensure colors extracted or warnings indicate normalization fallback.
- **Locked file (Windows):**
  - Simulate PermissionError; ensure error JSON includes error_type and metadata.probed_at.
- **Large deck (50+ layouts):**
  - Use --max-layouts 20; info includes “Limited analysis…”; indices preserved via original_index.
- **Timeout test:**
  - Set --timeout 1 with deep mode; expect warning “Probe exceeded 1s timeout…” and partial layouts present.
- **Non-standard slide size:**
  - Expect dims.aspect_ratio numeric plus string fraction, and DPI estimate present.
- **Placeholders missing/unknown:**
  - Ensure UNKNOWN_x entries labeled; recommendations still accurate.

---

## Quick wins to land now

- **Add per-master stats to capabilities and a concise summary line.**
- **Expose original_index when max_layouts is used to preserve mapping.**
- **Wrap transient slide add/remove in a helper with try/finally and defensive checks.**
- **Add schema_version to metadata for safer downstream evolution.**
- **Include aspect_ratio_float in slide_dimensions for numerical consumers.**

You’re close to “set-and-forget” production quality. These refinements push reliability under multi-master themes, large/odd templates, and future library changes—exactly where automation trips without extra rigor.
