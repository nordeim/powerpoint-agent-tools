# Strategic Enhancement Plan: Integrating the Capability Probe

## Executive Summary
The `ppt_capability_probe.py` tool represents a paradigm shift for the AI Presentation Architect. By moving from **reactive inspection** (checking slides as needed) to **proactive deep analysis** (building a complete mental model upfront), the agent can achieve higher precision, better brand alignment, and fewer runtime errors.

## 1. The Paradigm Shift

| Feature | Current Workflow | Enhanced Workflow (with Probe) |
| :--- | :--- | :--- |
| **Inspection** | Reactive: `ppt_get_info.py` (Basic metadata) | Proactive: `ppt_capability_probe.py --deep` (Full mental model) |
| **Layouts** | Guesswork: "Title and Content" (Hope it exists) | Precision: Select by capability (e.g., "Layout index 4 has Title + 2 Bodies") |
| **Positioning** | Heuristic: "Left 10%, Top 20%" | Semantic: "Use exact placeholder bounds from Master" |
| **Branding** | Manual: User must supply hex codes | Automatic: Extract and use Theme Colors/Fonts |
| **Integrity** | Post-check: `ppt_validate_presentation.py` | Continuous: Atomic verification + Checksums |

## 2. Core Integration Strategies

### Strategy A: The "Mental Model" Initialization
**Protocol**: At the start of *any* session involving an existing file or template, the agent MUST run:
```bash
uv run tools/ppt_capability_probe.py --file template.pptx --deep --json
```
**Benefit**: This provides a "God View" of the presentation:
-   **Layout Map**: Exact names and indices of every available layout.
-   **Placeholder Geometry**: Precise `(left, top, width, height)` for every placeholder on every layout.
-   **Theme DNA**: The exact color palette and font scheme used by the master.

### Strategy B: Theme-Aware Content Generation
**Problem**: Agents often generate content that looks "off" because it uses generic colors (e.g., standard blue) instead of the corporate brand.
**Solution**:
1.  Read `theme.colors` from the probe output.
2.  Map semantic roles to hex codes:
    -   `Primary` -> `accent1`
    -   `Secondary` -> `accent2`
    -   `Background` -> `background1`
3.  Apply these colors dynamically to charts, shapes, and text without user prompts.

### Strategy C: Precision Layout Targeting
**Problem**: Agents struggle to find the right layout for complex content (e.g., "3 columns with headers").
**Solution**:
1.  Analyze `layouts` array in probe output.
2.  Filter for layouts with specific `placeholder_types` (e.g., count of `BODY` or `OBJECT` placeholders).
3.  Select the layout by **Name** or **Index** with 100% confidence.

## 3. Proposed System Prompt Updates

### Update 1: Mandatory Inspection Protocol
*Replace the current "Inspect Before Edit" rule with:*

> **Deep Inspection Protocol**:
> - **Initialization**: Upon receiving a file, IMMEDIATELY run `ppt_capability_probe.py --deep --json`.
> - **Context Loading**: Ingest the `layouts`, `theme`, and `slide_dimensions` into your working context.
> - **Layout Selection**: Never guess layout names. Use the exact names returned by the probe.

### Update 2: Visual Design Framework
*Enhance the "Canonical Color Palettes" section:*

> **Theme-Derived Palette**:
> - **Priority**: ALWAYS prefer extracted theme colors over generic palettes.
> - **Mapping**:
>   - Use `theme.colors.accent1` for primary elements (titles, main bars).
>   - Use `theme.colors.accent2` for secondary data.
>   - Use `theme.fonts.heading` for all title text.

### Update 3: Tool Catalog
*Add the Probe to Domain 7 (Inspection & Analysis):*

> | Tool | Critical Arguments | Purpose |
> | :--- | :--- | :--- |
> | `ppt_capability_probe.py` | `--file PATH` (req), `--deep` | **Primary Inspection**: Deep analysis of layouts, theme, and capabilities. |

## 4. Example Workflow: "Smart" Slide Creation

**User Request**: "Add a slide with our branding showing 3 key metrics."

**Agent Action**:
1.  **Probe**: `ppt_capability_probe.py --file deck.pptx --deep --json`
2.  **Analyze**:
    -   Find layout with 3 `OBJECT` or `BODY` placeholders. -> Found "Three Column Content" (Index 4).
    -   Extract Theme Colors: `accent1` = `#0070C0`.
3.  **Execute**:
    -   `ppt_add_slide.py --layout "Three Column Content"`
    -   `ppt_add_chart.py ... --color "#0070C0"` (Matches brand automatically)

## 5. Conclusion
Integrating the Capability Probe transforms the agent from a "blind tool user" to a "context-aware architect." It enables the agent to see the presentation as the user sees it—structurally and visually—before making a single edit.
