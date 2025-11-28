# 🎯 Consolidated Review & Deep-Dive Analysis
## PowerPoint Agent Tools v3.1.0

---

## Executive Summary

I have completed a comprehensive analysis of all three documents. PowerPoint Agent Tools is a **production-grade, governance-first toolkit** enabling AI agents to programmatically engineer PowerPoint presentations with precision, safety, and visual intelligence.

| Dimension | Key Insight |
|-----------|-------------|
| **Maturity** | v3.1.0 - Production-ready with opacity support, version tracking, and approval token framework |
| **Complexity** | 39 CLI tools + 1 core library + JSON schemas + comprehensive test infrastructure |
| **Architecture** | Hub-and-Spoke: Stateless CLI tools wrapping a stateful-while-open core |
| **Governance** | Enterprise-grade: Clone-before-edit, approval tokens, version hashing, audit trails |
| **Target Users** | AI agents (primary), automation engineers, accessibility specialists |

---

## Part 1: WHAT — Project Identity & Components

### 1.1 Core Mission Statement

> **"Enabling AI agents to engineer presentations with precision, safety, and visual intelligence"**

This is not a simple PowerPoint manipulation library. It's an **AI-agent-optimized orchestration layer** with:
- Machine-parseable I/O (JSON-first)
- Governance controls (approval tokens, version tracking)
- Design intelligence (typography, color theory, WCAG compliance)
- Operational resilience (probe fallbacks, timeout handling, graceful degradation)

### 1.2 Component Inventory

```
┌─────────────────────────────────────────────────────────────────────────┐
│                    POWERPOINT AGENT TOOLS v3.1.0                        │
├─────────────────────────────────────────────────────────────────────────┤
│                                                                         │
│  ┌─────────────────────────────────────────────────────────────────┐   │
│  │                    ORCHESTRATION LAYER                          │   │
│  │        (AI Agents / Human Developers / CI/CD Pipelines)         │   │
│  └────────────────────────────┬────────────────────────────────────┘   │
│                               │                                         │
│  ┌────────────────────────────▼────────────────────────────────────┐   │
│  │                      CLI TOOLS (39 Spokes)                      │   │
│  │                                                                  │   │
│  │  ┌──────────┐ ┌──────────┐ ┌──────────┐ ┌──────────┐            │   │
│  │  │ Creation │ │  Slides  │ │ Content  │ │  Images  │            │   │
│  │  │ (4 tools)│ │(6 tools) │ │(6 tools) │ │(4 tools) │            │   │
│  │  └──────────┘ └──────────┘ └──────────┘ └──────────┘            │   │
│  │  ┌──────────┐ ┌──────────┐ ┌──────────┐ ┌──────────┐            │   │
│  │  │  Shapes  │ │ Data Viz │ │Inspection│ │Validation│            │   │
│  │  │(6 tools) │ │(4 tools) │ │(4 tools) │ │(4 tools) │            │   │
│  │  └──────────┘ └──────────┘ └──────────┘ └──────────┘            │   │
│  └────────────────────────────┬────────────────────────────────────┘   │
│                               │                                         │
│  ┌────────────────────────────▼────────────────────────────────────┐   │
│  │                   CORE LIBRARY (The Hub)                         │   │
│  │                                                                  │   │
│  │  ┌─────────────────────────────────────────────────────────┐    │   │
│  │  │              powerpoint_agent_core.py                    │    │   │
│  │  │                                                          │    │   │
│  │  │  • PowerPointAgent class (context manager)               │    │   │
│  │  │  • XML manipulation (opacity, z-order)                   │    │   │
│  │  │  • Position/Size resolution (%, inches, anchor, grid)    │    │   │
│  │  │  • Version hashing (geometry-aware in v3.1.0)            │    │   │
│  │  │  • Color helpers & contrast calculation                  │    │   │
│  │  │  • Path validation & security                            │    │   │
│  │  └─────────────────────────────────────────────────────────┘    │   │
│  │                                                                  │   │
│  │  ┌─────────────────────────────────────────────────────────┐    │   │
│  │  │              strict_validator.py                         │    │   │
│  │  │  • JSON Schema validation (Draft-07/2019-09/2020-12)     │    │   │
│  │  │  • Custom format checkers (hex-color, percentage, etc.)  │    │   │
│  │  └─────────────────────────────────────────────────────────┘    │   │
│  └────────────────────────────┬────────────────────────────────────┘   │
│                               │                                         │
│  ┌────────────────────────────▼────────────────────────────────────┐   │
│  │                      python-pptx + lxml                         │   │
│  │                   (Underlying Libraries)                        │   │
│  └─────────────────────────────────────────────────────────────────┘   │
│                                                                         │
└─────────────────────────────────────────────────────────────────────────┘
```

### 1.3 Tool Catalog (39 Tools by Domain)

| Domain | Tools | Key Capabilities |
|--------|-------|------------------|
| **Creation** (4) | `ppt_create_new`, `ppt_create_from_template`, `ppt_create_from_structure`, `ppt_clone_presentation` | Initialize, template-based creation, structure-driven generation |
| **Slides** (6) | `ppt_add_slide`, `ppt_delete_slide`⚠️, `ppt_duplicate_slide`, `ppt_reorder_slides`, `ppt_set_slide_layout`, `ppt_set_footer` | Full slide lifecycle management |
| **Content** (6) | `ppt_set_title`, `ppt_add_text_box`, `ppt_add_bullet_list`, `ppt_format_text`, `ppt_replace_text`, `ppt_add_notes` | Text content creation and formatting |
| **Images** (4) | `ppt_insert_image`, `ppt_replace_image`, `ppt_crop_image`, `ppt_set_image_properties` | Image lifecycle with alt-text support |
| **Shapes** (6) | `ppt_add_shape`, `ppt_format_shape`, `ppt_add_connector`, `ppt_set_background`, `ppt_set_z_order`, `ppt_remove_shape`⚠️ | Visual elements with opacity support |
| **Data Viz** (4) | `ppt_add_chart`, `ppt_update_chart_data`, `ppt_format_chart`, `ppt_add_table` | Charts and tables |
| **Inspection** (4) | `ppt_get_info`, `ppt_get_slide_info`, `ppt_extract_notes`, `ppt_capability_probe` | Non-destructive analysis |
| **Validation** (4) | `ppt_validate_presentation`, `ppt_check_accessibility`, `ppt_export_images`, `ppt_export_pdf` | Quality assurance and export |

> ⚠️ = Requires approval token (destructive operation)

### 1.4 Version 3.1.0 Innovations

| Feature | Description | Impact |
|---------|-------------|--------|
| **Opacity Support** | `fill_opacity` and `line_opacity` parameters (0.0-1.0) | Enables overlays, watermarks, subtle effects |
| **Geometry-Aware Versioning** | Hash includes shape positions, not just text | Detects visual changes, not just content |
| **Z-Order Control** | `ppt_set_z_order.py` with XML manipulation | Layer management for overlays |
| **Speaker Notes** | `ppt_add_notes.py` with append/prepend/overwrite modes | Presentation scripting, accessibility alternatives |
| **Deprecation Path** | `transparency` → `fill_opacity` conversion with warnings | Backward compatibility |

---

## Part 2: WHY — Design Philosophy & Value Proposition

### 2.1 The Four Pillars (Design Contract)

```
┌─────────────────────────────────────────────────────────────────────────┐
│                         THE FOUR PILLARS                                │
├────────────────┬────────────────┬────────────────┬─────────────────────┤
│   STATELESS    │    ATOMIC      │   COMPOSABLE   │    ACCESSIBLE       │
├────────────────┼────────────────┼────────────────┼─────────────────────┤
│ Each call is   │ Open→Modify→   │ Tools chain    │ WCAG 2.1 AA         │
│ independent    │ Save→Close     │ via JSON       │ compliance          │
│                │                │                │                     │
│ No memory of   │ One action     │ Output feeds   │ Alt text, contrast, │
│ previous calls │ per invocation │ next input     │ reading order       │
├────────────────┴────────────────┴────────────────┴─────────────────────┤
│                                                                         │
│                         + VISUAL-AWARE                                  │
│            Typography scales, color theory, content density             │
│                                                                         │
└─────────────────────────────────────────────────────────────────────────┘
```

### 2.2 Why Statelessness Matters

AI agents face unique challenges that stateless design directly addresses:

| Challenge | Stateless Solution |
|-----------|-------------------|
| Context window limits | No need to remember previous operations |
| Parallel execution | No shared state = no race conditions |
| Error recovery | Each operation is self-contained and restartable |
| Pipeline composition | Tools chain cleanly via stdout/stdin |
| Debugging | Each operation is independently reproducible |

### 2.3 Safety Hierarchy (Immutable)

```
PRECEDENCE ORDER:
1. Never perform destructive operations without approval  ← HIGHEST
2. Always work on cloned copies, never source files
3. Validate before delivery, always
4. Fail safely—incomplete is better than corrupted
5. Document everything for audit and rollback
6. Refresh indices after structural changes
7. Dry-run before actual execution for replacements       ← LOWEST
```

### 2.4 The Governance Value Proposition

This is not just a "nice to have" — it's production-critical:

| Governance Feature | Business Value |
|-------------------|----------------|
| **Clone-Before-Edit** | Zero risk to source materials; legal/compliance safe |
| **Approval Tokens** | Audit trail for destructive operations; SOC2 compatible |
| **Version Tracking** | Race condition detection; concurrent edit protection |
| **JSON-First I/O** | Machine-parseable; integrates with any orchestrator |
| **Accessibility Validation** | Legal compliance (ADA, Section 508, WCAG) |

---

## Part 3: HOW — Architecture & Implementation Patterns

### 3.1 Hub-and-Spoke Architecture (Detailed)

```
┌─────────────────────────────────────────────────────────────────────────┐
│                           SPOKE LAYER                                   │
│                                                                         │
│   ┌─────────────────┐                                                   │
│   │ CLI Tool        │  Responsibilities:                                │
│   │ (ppt_*.py)      │  • Argparse argument handling                     │
│   │                 │  • Input validation                               │
│   │                 │  • JSON output formatting                         │
│   │                 │  • Exit code management                           │
│   │                 │  • Error-to-JSON conversion                       │
│   │                 │  • HYGIENE BLOCK (stderr suppression)             │
│   └────────┬────────┘                                                   │
│            │                                                            │
│            │ Calls methods on                                           │
│            ▼                                                            │
├─────────────────────────────────────────────────────────────────────────┤
│                            HUB LAYER                                    │
│                                                                         │
│   ┌─────────────────────────────────────────────────────────────────┐  │
│   │                    PowerPointAgent                               │  │
│   │                                                                  │  │
│   │  LIFECYCLE:                                                      │  │
│   │  ┌─────────┐  ┌──────────┐  ┌────────────┐  ┌─────────────────┐ │  │
│   │  │ __init__│→ │  open()  │→ │  methods() │→ │ save() + close()│ │  │
│   │  │         │  │          │  │            │  │ (via __exit__)  │ │  │
│   │  └─────────┘  └──────────┘  └────────────┘  └─────────────────┘ │  │
│   │                                                                  │  │
│   │  RESPONSIBILITIES:                                               │  │
│   │  • File locking                                                  │  │
│   │  • python-pptx Presentation object management                   │  │
│   │  • Position/Size resolution (%, inches, anchor, grid)           │  │
│   │  • XML manipulation (opacity via lxml, z-order)                 │  │
│   │  • Version hash calculation (geometry-aware)                     │  │
│   │  • Color parsing and contrast calculation                        │  │
│   │  • Path traversal security                                       │  │
│   │                                                                  │  │
│   │  RETURN CONTRACT (v3.1.0+):                                      │  │
│   │  • All mutation methods return Dict, not primitives              │  │
│   │  • Dicts include: result data, styling details, version info    │  │
│   └─────────────────────────────────────────────────────────────────┘  │
│                                                                         │
└─────────────────────────────────────────────────────────────────────────┘
```

### 3.2 The Five Golden Rules (Inviolable)

| Rule | Implementation | Failure Mode |
|------|----------------|--------------|
| **1. Output Hygiene** | Hygiene Block: `sys.stderr = open(os.devnull, 'w')` at top of every tool | `jq: parse error` from library warnings |
| **2. Clone-Before-Edit** | Check file path starts with `/work/` or use `ppt_clone_presentation.py` first | Source file corruption, data loss |
| **3. Fail Safely with JSON** | Catch-all exception handler outputs `{"status": "error", ...}` | Orchestrator cannot parse failure reason |
| **4. Version Tracking** | Capture `presentation_version_before` and `presentation_version_after` | Undetected concurrent modifications |
| **5. Approval Tokens** | Validate token presence and scope before destructive operations | Unauthorized data destruction |

### 3.3 The Five-Phase Workflow

```
┌─────────────────────────────────────────────────────────────────────────┐
│                        FIVE-PHASE WORKFLOW                              │
├─────────────────────────────────────────────────────────────────────────┤
│                                                                         │
│  ┌─────────────┐                                                        │
│  │  PHASE 0    │  REQUEST INTAKE & CLASSIFICATION                       │
│  │             │  • Classify: SIMPLE | STANDARD | COMPLEX | DESTRUCTIVE │
│  │             │  • Determine approval requirements                     │
│  │             │  • Assess risk level                                   │
│  └──────┬──────┘                                                        │
│         ▼                                                               │
│  ┌─────────────┐                                                        │
│  │  PHASE 1    │  DISCOVER (Deep Inspection)                            │
│  │             │  • ppt_capability_probe.py --deep                      │
│  │             │  • Capture presentation_version                        │
│  │             │  • Extract theme colors/fonts                          │
│  │             │  • Map existing elements                               │
│  │             │  • Fallback probes if timeout                          │
│  └──────┬──────┘                                                        │
│         ▼                                                               │
│  ┌─────────────┐                                                        │
│  │  PHASE 2    │  PLAN (Manifest-Driven Design)                         │
│  │             │  • Create Change Manifest (JSON schema v3.0)           │
│  │             │  • Document design decisions with rationale            │
│  │             │  • Define preflight checklist                          │
│  │             │  • Specify rollback commands                           │
│  │             │  • Obtain approval tokens if needed                    │
│  └──────┬──────┘                                                        │
│         ▼                                                               │
│  ┌─────────────┐                                                        │
│  │  PHASE 3    │  CREATE (Design-Intelligent Execution)                 │
│  │             │  • Execute manifest operations sequentially            │
│  │             │  • Verify version before each mutation                 │
│  │             │  • Refresh indices after structural changes            │
│  │             │  • Record version after each mutation                  │
│  │             │  • Retry transient errors with backoff                 │
│  └──────┬──────┘                                                        │
│         ▼                                                               │
│  ┌─────────────┐                                                        │
│  │  PHASE 4    │  VALIDATE (Quality Assurance Gates)                    │
│  │             │  • ppt_validate_presentation.py                        │
│  │             │  • ppt_check_accessibility.py                          │
│  │             │  • Visual coherence check                              │
│  │             │  • Remediate issues if found                           │
│  └──────┬──────┘                                                        │
│         ▼                                                               │
│  ┌─────────────┐                                                        │
│  │  PHASE 5    │  DELIVER (Production Handoff)                          │
│  │             │  • Final validation confirmation                       │
│  │             │  • Generate delivery package                           │
│  │             │  • Document lessons learned                            │
│  │             │  • Provide next-step recommendations                   │
│  └─────────────┘                                                        │
│                                                                         │
└─────────────────────────────────────────────────────────────────────────┘
```

### 3.4 Critical Pattern: Shape Index Management

This is the **#1 source of bugs** in PowerPoint automation. The documents emphasize this repeatedly:

```
⚠️ SHAPE INDICES ARE POSITIONAL AND SHIFT AFTER STRUCTURAL OPERATIONS

Operations That Invalidate Indices:
┌────────────────────┬─────────────────────────────────────────────────┐
│ Operation          │ Effect                                          │
├────────────────────┼─────────────────────────────────────────────────┤
│ add_shape()        │ Adds new index at end                           │
│ remove_shape()     │ Shifts ALL subsequent indices down by 1         │
│ set_z_order()      │ Reorders indices (XML element moves in tree)    │
│ delete_slide()     │ Invalidates ALL indices on that slide           │
│ add_slide()        │ May affect slide indices (not shape indices)    │
└────────────────────┴─────────────────────────────────────────────────┘

MANDATORY PROTOCOL:
1. Before targeting shapes: Run ppt_get_slide_info.py
2. After index-invalidating operations: IMMEDIATELY refresh via ppt_get_slide_info.py
3. Never cache shape indices across operations
4. Use shape names/identifiers when available, not just indices
5. Document index refresh in manifest operation notes
```

**Correct Pattern:**
```python
# ✅ CORRECT - re-query after structural changes
result1 = agent.add_shape(slide_index=0, ...)  # Returns shape_index: 5
agent.remove_shape(slide_index=0, shape_index=result1["shape_index"])

# IMMEDIATELY refresh indices
slide_info = agent.get_slide_info(slide_index=0)

# Find target shape by characteristics, not cached index
for shape in slide_info["shapes"]:
    if shape["name"] == "target_shape":
        agent.format_shape(slide_index=0, shape_index=shape["index"], ...)
```

### 3.5 Critical Pattern: Overlay Workflow

The **overlay pattern** for text readability is a key use case:

```bash
# Complete Overlay Workflow (CLI)

# 1. Clone for safety
uv run tools/ppt_clone_presentation.py \
  --source original.pptx --output work.pptx --json

# 2. Add overlay shape with opacity
uv run tools/ppt_add_shape.py --file work.pptx --slide 2 --shape rectangle \
  --position '{"left":"0%","top":"0%"}' \
  --size '{"width":"100%","height":"100%"}' \
  --fill-color "#FFFFFF" --fill-opacity 0.15 --json

# 3. MANDATORY: Refresh shape indices
uv run tools/ppt_get_slide_info.py --file work.pptx --slide 2 --json
# → Note the new shape index (e.g., 7)

# 4. Send overlay to back
uv run tools/ppt_set_z_order.py --file work.pptx --slide 2 \
  --shape 7 --action send_to_back --json

# 5. MANDATORY: Refresh indices again
uv run tools/ppt_get_slide_info.py --file work.pptx --slide 2 --json

# 6. Validate contrast after overlay
uv run tools/ppt_check_accessibility.py --file work.pptx --json
```

### 3.6 XML Manipulation Details

`python-pptx` doesn't expose transparency or z-order natively. The core uses `lxml` for direct XML manipulation:

**Opacity (OOXML Scale):**
```python
# OOXML uses 0-100,000 scale (not 0-1)
# opacity=0.15 becomes val="15000"

from lxml import etree
from pptx.oxml.ns import qn

spPr = shape._sp.spPr
solidFill = spPr.find(qn('a:solidFill'))
color_elem = solidFill.find(qn('a:srgbClr'))
alpha_elem = etree.SubElement(color_elem, qn('a:alpha'))
alpha_elem.set('val', str(int(0.15 * 100000)))  # 15% opacity → 15000
```

**Z-Order:**
```python
# Z-order is determined by element position in XML tree
# Moving element = physically relocating in <p:spTree>
# This changes the shape's index in the shapes collection!
```

### 3.7 Version Tracking Protocol

```
VERSION TRACKING PROTOCOL (v3.1.0)

GEOMETRY-AWARE HASHING:
• v2.0: Only text content was hashed
• v3.1.0: Hash includes {left}:{top}:{width}:{height} for every shape
• Moving a shape by 1 pixel changes the version hash

PROTOCOL:
1. After clone: Capture initial presentation_version from ppt_get_info.py
2. Before each mutation: Verify current version matches expected
3. With each mutation: Record expected version in manifest
4. After each mutation: Capture new version, update manifest
5. On version mismatch: ABORT → Re-probe → Update manifest → Seek guidance

RESPONSE STRUCTURE:
{
    "status": "success",
    "file": "/path/to/file.pptx",
    "presentation_version_before": "a1b2c3d4",
    "presentation_version_after": "e5f6g7h8",
    "changes_made": { ... }
}
```

### 3.8 Error Handling Matrix

| Exit Code | Category | Meaning | Retryable | Action |
|-----------|----------|---------|-----------|--------|
| **0** | Success | Operation completed | N/A | Proceed |
| **1** | Usage Error | Invalid arguments, bad input | No | Fix arguments |
| **2** | Validation Error | Schema/content invalid | No | Fix input data |
| **3** | Transient Error | Timeout, I/O, network | Yes | Retry with backoff |
| **4** | Permission Error | Approval token missing/invalid | No | Obtain token |
| **5** | Internal Error | Unexpected failure | Maybe | Investigate |

### 3.9 Probe Resilience Framework

```
PROBE DECISION TREE (with timeout protection)

1. Validate absolute path
2. Check file readability
3. Verify disk space ≥ 100MB
4. Attempt deep probe with 15s timeout
   ├── Success → Return full probe JSON
   └── Failure → Retry with exponential backoff (2s, 4s, 8s)
5. If all retries fail:
   ├── Attempt fallback probes (ppt_get_info + ppt_get_slide_info)
   │   ├── Success → Return merged minimal JSON with probe_fallback: true
   │   └── Failure → Return structured error JSON
   └── Exit with appropriate code

TRANSIENT SLIDE PATTERN (for deep analysis):
• Add slide temporarily for placeholder analysis
• Use generator pattern with try/finally for guaranteed cleanup
• File is not saved, so transient slide disappears anyway
```

---

## Part 4: Design Intelligence System

### 4.1 Visual Hierarchy Framework

```
VISUAL HIERARCHY PYRAMID

           ▲ PRIMARY (Title, Key Message)
          ╱ ╲   Largest, Boldest, Top Position
         ╱   ╲
        ╱─────╲ SECONDARY (Subtitles, Headers)
       ╱       ╲   Medium Size, Supporting Position
      ╱─────────╲ TERTIARY (Body, Details, Data)
     ╱           ╲   Smallest, Dense Information
    ╱─────────────╲ AMBIENT (Backgrounds, Overlays)
   ╱_______________╲   Subtle, Non-Competing
```

### 4.2 Typography Scale

| Element | Minimum | Recommended | Maximum |
|---------|---------|-------------|---------|
| Main Title | 36pt | 44pt | 60pt |
| Slide Title | 28pt | 32pt | 40pt |
| Subtitle | 20pt | 24pt | 28pt |
| Body Text | 16pt | 18pt | 24pt |
| Bullet Points | 14pt | 16pt | 20pt |
| Captions | 12pt | 14pt | 16pt |
| Footer/Legal | 10pt | 12pt | 14pt |
| **NEVER BELOW** | **10pt** | — | — |

### 4.3 Color System Priority

```
PRIORITY ORDER:
1. Theme-extracted colors (from probe)
2. Semantic color mapping (accent1 → primary, etc.)
3. Canonical fallback palettes (only if theme extraction fails)

CANONICAL PALETTES:
• Corporate: #0070C0 primary, #595959 secondary, #ED7D31 accent
• Modern: #2E75B6 primary, #404040 secondary, #FFC000 accent
• Minimal: #000000 primary, #808080 secondary, #C00000 accent
• Data-Rich: #2A9D8F primary with 5-color chart palette
```

### 4.4 Content Density (6×6 Rule)

```
STANDARD LIMITS:
├── Maximum 6 bullet points per slide
├── Maximum 6 words per bullet point
└── One key message per slide

EXTENDED (requires explicit documentation):
├── Data-dense slides: Up to 8 bullets, 10 words
├── Reference slides: Dense text acceptable
└── Must document exception in manifest design_decisions
```

### 4.5 Accessibility Requirements (WCAG 2.1 AA)

| Check | Requirement | Tool | Remediation |
|-------|-------------|------|-------------|
| Alt text | All images must have descriptive alt text | `ppt_check_accessibility` | `ppt_set_image_properties --alt-text` |
| Color contrast | Text ≥4.5:1 (body), ≥3:1 (large) | `ppt_check_accessibility` | `ppt_format_text --color` |
| Reading order | Logical tab order for screen readers | `ppt_check_accessibility` | Manual reordering |
| Font size | No text below 10pt, prefer ≥12pt | Manual verification | `ppt_format_text --font-size` |
| Color independence | Information not conveyed by color alone | Manual verification | Add patterns/labels |

---

## Part 5: Testing & Contribution Standards

### 5.1 Required Test Coverage

| Category | What to Test |
|----------|--------------|
| **Happy Path** | Normal usage succeeds |
| **Edge Cases** | Boundary values (0, 1, max, empty) |
| **Error Cases** | Invalid inputs raise correct exceptions |
| **Validation** | Invalid ranges/formats rejected |
| **Backward Compat** | Deprecated features still work |
| **CLI Integration** | Tool produces valid JSON |
| **Governance** | Clone-before-edit enforced, tokens validated |
| **Version Tracking** | Presentation versions captured correctly |
| **Index Freshness** | Shape indices refreshed after structural changes |

### 5.2 Code Standards

| Aspect | Requirement |
|--------|-------------|
| Python Version | 3.8+ |
| Type Hints | Mandatory for all function signatures |
| Docstrings | Required for modules, classes, functions |
| Line Length | 100 characters (soft limit) |
| Formatting | `black` with default settings |
| Linting | `ruff` with no errors |
| Imports | Grouped: stdlib → third-party → local |

### 5.3 Tool Naming Convention

```
ppt_<verb>_<noun>.py

Examples:
• ppt_add_shape.py
• ppt_get_slide_info.py
• ppt_validate_presentation.py
• ppt_check_accessibility.py
```

---

## Part 6: Key Insights & Gotchas

### 6.1 Common Failure Modes

| Error | Cause | Fix |
|-------|-------|-----|
| `jq: parse error: Invalid numeric literal` | Non-JSON output to stdout (library warnings, print statements) | Apply Hygiene Block |
| `TypeError: '<=' not supported between 'int' and 'dict'` | Treating v3.1.0 return value as primitive | Use `result["key"]` pattern |
| `SlideNotFoundError` | Index out of range (0-indexed) | Verify with `ppt_get_info.py` first |
| `ShapeNotFoundError` | Stale index after structural change | Refresh with `ppt_get_slide_info.py` |
| `LayoutNotFoundError` | Guessing layout names | Use probe to discover actual names |
| `VersionConflictError` | External modification detected | Re-probe and update manifest |

### 6.2 The "Footer Trap"

```
THE TRAP:
• agent.prs.slide_masters[0].placeholders might contain a footer
• This makes you think the footer works

THE REALITY:
• Individual slides might have "Hide Background Graphics" on
• Or simply haven't instantiated that placeholder

THE FIX:
• ppt_set_footer.py uses Dual Strategy:
  1. Try setting the placeholder
  2. Check if any slides were actually updated
  3. If 0 slides updated, fall back to Text Box Overlay
```

### 6.3 Chart Update Limitations

```
⚠️ python-pptx has LIMITED chart update support

RISKY:
agent.update_chart_data(slide_index=0, chart_index=0, data=new_data)
# May fail if schema doesn't match exactly

PREFERRED:
agent.remove_shape(slide_index=0, shape_index=chart_index)  # Delete old
agent.add_chart(slide_index=0, chart_type="column", data=new_data, ...)  # Create new
```

---

## Part 7: Summary & Readiness Assessment

### 7.1 What I Now Understand

| Dimension | Confidence | Notes |
|-----------|------------|-------|
| Project Identity & Mission | ✅ High | AI-agent-optimized, governance-first PowerPoint automation |
| Architecture | ✅ High | Hub-and-Spoke, stateless tools, stateful-while-open core |
| Tool Catalog | ✅ High | 39 tools across 8 domains |
| Design Philosophy | ✅ High | Four Pillars + Visual-Aware + Safety Hierarchy |
| Workflow Phases | ✅ High | Discover → Plan → Create → Validate → Deliver |
| Critical Patterns | ✅ High | Shape index management, overlays, version tracking |
| Gotchas | ✅ High | Footer trap, chart limitations, hygiene block |
| Testing Standards | ✅ High | Coverage requirements, test patterns |
| Contribution Requirements | ✅ High | PR checklist, naming conventions |

### 7.2 My Operating Posture Going Forward

When you bring me coding tasks for this project, I will:

1. **Classify the request** (SIMPLE/STANDARD/COMPLEX/DESTRUCTIVE)
2. **Identify affected components** (core library, specific tools, tests)
3. **Apply the Five Golden Rules** rigorously
4. **Follow the Five-Phase Workflow** for non-trivial tasks
5. **Ensure shape index freshness** after structural operations
6. **Track presentation versions** for mutations
7. **Validate accessibility** before delivery
8. **Produce production-ready, well-documented code** with tests

---

## PowerPointAgent Core Handbook v3.1.4 — Deep Analysis

---

## Executive Summary

I have integrated the **PowerPointAgent Core Programming Handbook v3.1.4** with my existing understanding. This document provides the **implementation-level details** that complete my knowledge of the system.

| Dimension | Previous Understanding | New Depth Added |
|-----------|----------------------|-----------------|
| File Locking | "Uses file locking" | `os.open` with `O_CREAT\|O_EXCL`, 10s timeout, `FileLockError` |
| Version Hashing | "Geometry-aware" | SHA-256 of: slide count + layouts + `{left}:{top}:{width}:{height}` + text runs |
| Approval Tokens | "Required for destructive ops" | HMAC-SHA256 structure, scope strings, generation pattern |
| Z-Order | "XML manipulation" | Move `<p:sp>` in `<p:spTree>`: end=front, index 2=back |
| Opacity | "OOXML injection" | `<a:alpha val="50000"/>` inside `<a:srgbClr>`, 0-100000 scale |
| Performance | Not detailed | O(N) complexity table with concrete timings |
| Backward Compat | "Dict returns in v3.1" | Silent clamping removed (breaking), migration patterns |

---

## Part 1: Core Architecture Deep-Dive

### 1.1 Context Manager Lifecycle (Precise)

```
┌─────────────────────────────────────────────────────────────────────────┐
│                    PowerPointAgent LIFECYCLE                            │
├─────────────────────────────────────────────────────────────────────────┤
│                                                                         │
│  with PowerPointAgent(filepath) as agent:                               │
│       │                                                                 │
│       ▼                                                                 │
│  ┌─────────────────────────────────────────────────────────────────┐   │
│  │  __init__(filepath)                                              │   │
│  │  • Store filepath                                                │   │
│  │  • Initialize internal state                                     │   │
│  │  • NO file operations yet                                        │   │
│  └──────────────────────────┬──────────────────────────────────────┘   │
│                             │                                           │
│                             ▼                                           │
│  ┌─────────────────────────────────────────────────────────────────┐   │
│  │  open(filepath, acquire_lock=True)                               │   │
│  │                                                                  │   │
│  │  LOCKING MECHANISM (if acquire_lock=True):                       │   │
│  │  ┌────────────────────────────────────────────────────────────┐ │   │
│  │  │  os.open(lockfile, O_CREAT | O_EXCL)                       │ │   │
│  │  │  ├── Success → Lock acquired                               │ │   │
│  │  │  └── errno.EEXIST → Retry with 10s timeout                 │ │   │
│  │  │      └── Timeout → Raise FileLockError                     │ │   │
│  │  └────────────────────────────────────────────────────────────┘ │   │
│  │                                                                  │   │
│  │  PATH VALIDATION:                                                │   │
│  │  • Traversal protection (if allowed_base_dirs set)              │   │
│  │  • Extension check: .pptx, .pptm, .potx                         │   │
│  │                                                                  │   │
│  │  LOAD:                                                           │   │
│  │  • self.prs = Presentation(filepath)                            │   │
│  └──────────────────────────┬──────────────────────────────────────┘   │
│                             │                                           │
│                             ▼                                           │
│  ┌─────────────────────────────────────────────────────────────────┐   │
│  │  MUTATION OPERATIONS (user code)                                 │   │
│  │                                                                  │   │
│  │  • add_shape(), format_shape(), add_text_box(), etc.            │   │
│  │  • Each returns Dict with results + version tracking            │   │
│  │  • Destructive ops require approval_token                       │   │
│  └──────────────────────────┬──────────────────────────────────────┘   │
│                             │                                           │
│                             ▼                                           │
│  ┌─────────────────────────────────────────────────────────────────┐   │
│  │  save(filepath=None)                                             │   │
│  │                                                                  │   │
│  │  • If filepath=None → Overwrite source                          │   │
│  │  • Ensures parent directories exist                             │   │
│  │  • Atomic write operation                                       │   │
│  └──────────────────────────┬──────────────────────────────────────┘   │
│                             │                                           │
│                             ▼                                           │
│  ┌─────────────────────────────────────────────────────────────────┐   │
│  │  __exit__() (automatic)                                          │   │
│  │                                                                  │   │
│  │  • Release file lock (remove lockfile)                          │   │
│  │  • Close Presentation object                                    │   │
│  │  • Clean up resources                                           │   │
│  └─────────────────────────────────────────────────────────────────┘   │
│                                                                         │
└─────────────────────────────────────────────────────────────────────────┘
```

### 1.2 Version Hashing Algorithm (Geometry-Aware)

```
VERSION HASH COMPUTATION (SHA-256, prefix 16 chars)

INPUT COMPONENTS:
┌─────────────────────────────────────────────────────────────────────────┐
│  1. SLIDE COUNT                                                         │
│     └── int: len(prs.slides)                                           │
│                                                                         │
│  2. LAYOUT NAMES (per slide)                                            │
│     └── List[str]: [slide.slide_layout.name for slide in prs.slides]   │
│                                                                         │
│  3. SHAPE GEOMETRY (per shape, per slide)                               │
│     └── Format: "{left}:{top}:{width}:{height}"                        │
│     └── Units: EMUs (English Metric Units)                             │
│     └── Example: "914400:914400:4572000:2743200"                       │
│                                                                         │
│  4. TEXT CONTENT (per text frame)                                       │
│     └── SHA-256 hash of concatenated text runs                         │
└─────────────────────────────────────────────────────────────────────────┘

ALGORITHM:
1. Collect all components into deterministic string
2. Compute SHA-256 hash
3. Return first 16 characters as version identifier

IMPLICATIONS:
• Moving a shape by 1 pixel → Different version
• Changing any text → Different version  
• Adding/removing shapes → Different version
• Reordering slides → Different version
• Formatting changes WITHOUT geometry/text → SAME version (!)
```

### 1.3 Position Resolution System

```python
# Position.from_dict() Resolution Logic

RESOLUTION PRIORITY:
1. Check for "anchor" key → Anchor-based positioning
2. Check for "grid_row" key → Grid-based positioning  
3. Check for percentage strings → Percentage-based
4. Assume numeric values → Absolute (inches)

ANCHOR SYSTEM:
┌─────────────────────────────────────────────────────────────────────────┐
│                                                                         │
│    top_left ─────── top_center ─────── top_right                       │
│        │                │                  │                            │
│        │                │                  │                            │
│   center_left ────── center ─────── center_right                       │
│        │                │                  │                            │
│        │                │                  │                            │
│   bottom_left ──── bottom_center ──── bottom_right                     │
│                                                                         │
└─────────────────────────────────────────────────────────────────────────┘

GRID SYSTEM (12-column default):
┌───┬───┬───┬───┬───┬───┬───┬───┬───┬───┬───┬───┐
│ 1 │ 2 │ 3 │ 4 │ 5 │ 6 │ 7 │ 8 │ 9 │10 │11 │12 │
└───┴───┴───┴───┴───┴───┴───┴───┴───┴───┴───┴───┘
{"grid_row": 2, "grid_col": 3, "grid_span": 6} 
→ Starts at column 3, spans 6 columns (50% width)
```

---

## Part 2: API Implementation Details

### 2.1 Destructive Operations — Token Enforcement

```python
# INTERNAL ENFORCEMENT PATTERN

def delete_slide(self, index: int, approval_token: str = None) -> Dict:
    """
    Delete slide at index.
    
    SECURITY: Requires valid approval_token with scope 'delete:slide'
    """
    # 1. Token validation FIRST
    if not approval_token:
        raise ApprovalTokenError(
            "Approval token required for slide deletion",
            details={"operation": "delete_slide", "slide_index": index}
        )
    
    # 2. Validate token structure and scope
    self._validate_approval_token(approval_token, required_scope="delete:slide")
    
    # 3. Capture version before
    version_before = self.get_presentation_version()
    
    # 4. Perform deletion
    rId = self.prs.slides._sldIdLst[index].rId
    self.prs.part.drop_rel(rId)
    del self.prs.slides._sldIdLst[index]
    
    # 5. Capture version after
    version_after = self.get_presentation_version()
    
    return {
        "deleted_index": index,
        "total_slides": len(self.prs.slides),
        "presentation_version_before": version_before,
        "presentation_version_after": version_after
    }
```

**Token Scopes:**
| Scope | Required For |
|-------|-------------|
| `delete:slide` | `delete_slide()` |
| `remove:shape` | `remove_shape()` |
| `replace:all` | Mass text replacement (future) |

### 2.2 Shape Operations — Return Structures (v3.1.4)

```python
# add_shape() Return Structure
{
    "shape_index": 5,
    "shape_type": "rectangle",
    "position": {"left": "10%", "top": "20%"},
    "size": {"width": "30%", "height": "40%"},
    "styling": {
        "fill_color": "#0070C0",
        "fill_opacity": 0.5,
        "fill_opacity_applied": True,  # Confirms XML injection succeeded
        "line_color": None,
        "line_opacity": 1.0
    },
    "presentation_version_before": "a1b2c3d4e5f6g7h8",
    "presentation_version_after": "i9j0k1l2m3n4o5p6"
}

# format_shape() Return Structure  
{
    "shape_index": 5,
    "changes_applied": ["fill_color", "fill_opacity", "transparency_converted_to_opacity"],
    "changes_detail": {
        "fill_color": {"from": "#FFFFFF", "to": "#0070C0"},
        "fill_opacity": {"from": 1.0, "to": 0.5},
        "converted_opacity": 0.5  # If transparency param was used
    },
    "presentation_version_before": "...",
    "presentation_version_after": "..."
}
```

### 2.3 Z-Order Implementation Details

```python
# set_z_order() Internal Implementation

def set_z_order(self, slide_index: int, shape_index: int, action: str) -> Dict:
    """
    Modify shape stacking order via XML tree manipulation.
    
    CRITICAL: Invalidates shape indices! Caller must refresh.
    """
    slide = self.prs.slides[slide_index]
    spTree = slide.shapes._spTree  # XML tree: <p:spTree>
    
    # Get shape's XML element
    shape = slide.shapes[shape_index]
    sp_element = shape._sp  # <p:sp> element
    
    # Remove from current position
    spTree.remove(sp_element)
    
    # Reinsert at new position
    if action == "bring_to_front":
        spTree.append(sp_element)  # End of list = top layer
        
    elif action == "send_to_back":
        # Index 0-1 are typically background/master refs
        # Insert at index 2 to be behind all content shapes
        spTree.insert(2, sp_element)
        
    elif action == "bring_forward":
        current_idx = list(spTree).index(sp_element)
        spTree.insert(current_idx + 1, sp_element)
        
    elif action == "send_backward":
        current_idx = list(spTree).index(sp_element)
        new_idx = max(2, current_idx - 1)  # Don't go before index 2
        spTree.insert(new_idx, sp_element)
    
    return {
        "shape_index_original": shape_index,
        "action": action,
        "warning": "Shape indices have changed. Refresh via get_slide_info()."
    }
```

### 2.4 Opacity Injection (XML Level)

```xml
<!-- BEFORE: Shape with solid fill, no transparency -->
<p:sp>
  <p:spPr>
    <a:solidFill>
      <a:srgbClr val="0070C0"/>
    </a:solidFill>
  </p:spPr>
</p:sp>

<!-- AFTER: Shape with 50% opacity (fill_opacity=0.5) -->
<p:sp>
  <p:spPr>
    <a:solidFill>
      <a:srgbClr val="0070C0">
        <a:alpha val="50000"/>  <!-- 50000 = 50% on OOXML scale -->
      </a:srgbClr>
    </a:solidFill>
  </p:spPr>
</p:sp>
```

```python
# _set_fill_opacity() Implementation Pattern

def _set_fill_opacity(self, shape, opacity: float):
    """
    Inject <a:alpha> element into shape's fill color.
    
    Args:
        opacity: 0.0 (invisible) to 1.0 (opaque)
    """
    from lxml import etree
    from pptx.oxml.ns import qn
    
    # Convert 0.0-1.0 to OOXML scale (0-100000)
    ooxml_alpha = int(opacity * 100000)
    
    spPr = shape._sp.spPr
    
    # Find or create <a:solidFill>
    solidFill = spPr.find(qn('a:solidFill'))
    if solidFill is None:
        # Shape has no fill - need to create one
        solidFill = etree.SubElement(spPr, qn('a:solidFill'))
        srgbClr = etree.SubElement(solidFill, qn('a:srgbClr'))
        srgbClr.set('val', 'FFFFFF')  # Default white
    else:
        srgbClr = solidFill.find(qn('a:srgbClr'))
        if srgbClr is None:
            # May have theme color - need to handle
            srgbClr = etree.SubElement(solidFill, qn('a:srgbClr'))
            srgbClr.set('val', 'FFFFFF')
    
    # Find or create <a:alpha>
    alpha = srgbClr.find(qn('a:alpha'))
    if alpha is None:
        alpha = etree.SubElement(srgbClr, qn('a:alpha'))
    
    # Set value
    alpha.set('val', str(ooxml_alpha))
```

### 2.5 Chart Update Strategy (3-Tier Fallback)

```python
def update_chart_data(self, slide_index: int, chart_index: int, data: Dict) -> Dict:
    """
    Update chart with new data using 3-tier fallback strategy.
    """
    chart_shape = self._get_chart_shape(slide_index, chart_index)
    chart = chart_shape.chart
    
    # TIER 1: Best - Use replace_data() if available
    try:
        chart_data = self._build_chart_data(data)
        chart.replace_data(chart_data)
        return {"strategy": "replace_data", "success": True}
        
    except AttributeError:
        # Older python-pptx version
        pass
    
    # TIER 2: Fallback - Recreate chart in-place
    try:
        # Capture position and size
        left, top = chart_shape.left, chart_shape.top
        width, height = chart_shape.width, chart_shape.height
        chart_type = self._detect_chart_type(chart)
        
        # Remove old chart
        sp = chart_shape._element
        sp.getparent().remove(sp)
        
        # Create new chart with same geometry
        self.add_chart(
            slide_index=slide_index,
            chart_type=chart_type,
            data=data,
            position={"left": left, "top": top},  # EMUs
            size={"width": width, "height": height}  # EMUs
        )
        
        return {
            "strategy": "recreate",
            "success": True,
            "warning": "Chart was recreated. Some custom styling may be lost."
        }
        
    except Exception as e:
        return {
            "strategy": "failed",
            "success": False,
            "error": str(e)
        }
```

---

## Part 3: Performance Characteristics

### 3.1 Operation Complexity Table

| Operation | Complexity | 10-Slide Deck | 50-Slide Deck | Scaling Factor |
|-----------|------------|---------------|---------------|----------------|
| `get_presentation_version()` | O(N) shapes | ~15ms | ~75ms | Linear with total shapes |
| `capability_probe(deep=True)` | O(M) layouts | ~120ms | ~600ms+ | Creates/destroys slides |
| `add_shape()` | O(1) | ~8ms | ~8ms | Constant (XML injection) |
| `format_shape()` | O(1) | ~5ms | ~5ms | Constant |
| `replace_text(global)` | O(N) text runs | ~25ms | ~125ms | Linear with text volume |
| `save()` | I/O bound | ~50ms | ~200ms+ | Dominated by file size |
| `get_slide_info()` | O(K) shapes/slide | ~10ms | ~10ms | Per-slide constant |

### 3.2 Performance Guidelines

```
OPTIMIZATION STRATEGIES:

1. BATCHING (within context manager)
   ┌────────────────────────────────────────────────────────────────┐
   │  with PowerPointAgent(filepath) as agent:                      │
   │      agent.open(filepath)                                      │
   │      # Multiple operations...                                  │
   │      agent.add_shape(...)                                      │
   │      agent.add_shape(...)                                      │
   │      agent.format_shape(...)                                   │
   │      agent.save()  # Single save for all operations            │
   │  # More efficient than separate tool calls                     │
   └────────────────────────────────────────────────────────────────┘

2. SHALLOW PROBES
   • Use deep=False unless layout geometry is strictly required
   • Deep probe creates/destroys transient slides (expensive)

3. LIMITS
   • Avoid decks >100 slides for interactive sessions
   • Avoid files >50MB (images dominate size)
   • Implement 15s timeout for probes

4. VERSION CACHING (within single operation)
   • get_presentation_version() is O(N) - called twice per mutation
   • For multi-step operations, consider caching version_before
```

### 3.3 Memory Considerations

```
MEMORY FOOTPRINT:

Presentation Object:
• Base overhead: ~10-20MB for python-pptx internals
• Per slide: ~0.5-2MB depending on content
• Images: Loaded into memory (can be large!)

Guidelines:
• Process large decks in sections if possible
• Use image compression when inserting
• Close agent promptly after use (context manager handles this)
```

---

## Part 4: Security Implementation

### 4.1 Approval Token Generation (Reference)

```python
import hmac
import hashlib
import base64
import json
import time

def generate_approval_token(
    scope: str, 
    user: str, 
    secret_key: bytes,
    expiry_seconds: int = 3600
) -> str:
    """
    Generate HMAC-SHA256 approval token for destructive operations.
    
    Args:
        scope: Required scope (e.g., "delete:slide", "remove:shape")
        user: User identifier for audit trail
        secret_key: Shared secret for HMAC signing
        expiry_seconds: Token validity period (default 1 hour)
        
    Returns:
        Token string in format: "HMAC-SHA256:{base64_payload}.{signature}"
    """
    payload = {
        "scope": scope,
        "user": user,
        "issued": time.time(),
        "expiry": time.time() + expiry_seconds,
        "single_use": True
    }
    
    # Serialize and encode payload
    json_payload = json.dumps(payload, sort_keys=True)
    b64_payload = base64.urlsafe_b64encode(json_payload.encode()).decode()
    
    # Generate HMAC-SHA256 signature
    signature = hmac.new(
        secret_key, 
        b64_payload.encode(), 
        hashlib.sha256
    ).hexdigest()
    
    return f"HMAC-SHA256:{b64_payload}.{signature}"


def validate_approval_token(
    token: str, 
    required_scope: str, 
    secret_key: bytes
) -> Dict:
    """
    Validate approval token.
    
    Raises:
        ApprovalTokenError: If token is invalid, expired, or lacks scope
    """
    if not token or not token.startswith("HMAC-SHA256:"):
        raise ApprovalTokenError("Invalid token format")
    
    try:
        _, payload_sig = token.split(":", 1)
        b64_payload, signature = payload_sig.rsplit(".", 1)
        
        # Verify signature
        expected_sig = hmac.new(
            secret_key,
            b64_payload.encode(),
            hashlib.sha256
        ).hexdigest()
        
        if not hmac.compare_digest(signature, expected_sig):
            raise ApprovalTokenError("Invalid token signature")
        
        # Decode payload
        payload = json.loads(base64.urlsafe_b64decode(b64_payload))
        
        # Check expiry
        if time.time() > payload["expiry"]:
            raise ApprovalTokenError("Token expired")
        
        # Check scope
        if payload["scope"] != required_scope:
            raise ApprovalTokenError(
                f"Token scope '{payload['scope']}' does not match required '{required_scope}'"
            )
        
        return payload
        
    except (ValueError, KeyError, json.JSONDecodeError) as e:
        raise ApprovalTokenError(f"Token parsing failed: {e}")
```

### 4.2 Path Validation

```python
def _validate_path(self, filepath: Path) -> Path:
    """
    Security-hardened path validation.
    
    Checks:
    1. Path traversal protection (if allowed_base_dirs configured)
    2. Extension whitelist enforcement
    3. Existence verification
    """
    filepath = Path(filepath).resolve()  # Resolve symlinks
    
    # 1. Traversal protection
    if self.allowed_base_dirs:
        if not any(filepath.is_relative_to(base) for base in self.allowed_base_dirs):
            raise PathValidationError(
                f"Path '{filepath}' is outside allowed directories",
                details={"allowed": [str(d) for d in self.allowed_base_dirs]}
            )
    
    # 2. Extension whitelist
    ALLOWED_EXTENSIONS = {'.pptx', '.pptm', '.potx'}
    if filepath.suffix.lower() not in ALLOWED_EXTENSIONS:
        raise PathValidationError(
            f"Extension '{filepath.suffix}' not allowed",
            details={"allowed": list(ALLOWED_EXTENSIONS)}
        )
    
    return filepath
```

---

## Part 5: Backward Compatibility Details

### 5.1 Breaking Changes in v3.1.x

| Change | v3.0.x Behavior | v3.1.x Behavior | Migration |
|--------|-----------------|-----------------|-----------|
| **add_slide() return** | Returns `int` (index) | Returns `Dict` | Use `result["slide_index"]` |
| **Index clamping** | Silent clamping to valid range | Raises `SlideNotFoundError` | Add explicit validation |
| **Approval tokens** | Not required | Required for destructive ops | Implement token generation |
| **transparency param** | Primary parameter | Deprecated (warning logged) | Use `fill_opacity` |

### 5.2 Migration Patterns

```python
# MIGRATION: add_slide() return value

# v3.0.x pattern (DEPRECATED but still works via duck typing)
idx = agent.add_slide("Blank")
if isinstance(idx, dict):
    idx = idx["slide_index"]  # Handle both versions

# v3.1.x pattern (RECOMMENDED)
result = agent.add_slide("Blank")
idx = result["slide_index"]
version = result["presentation_version_after"]


# MIGRATION: Index validation

# v3.0.x pattern (relied on silent clamping)
agent.add_slide("Blank", index=999)  # Would clamp to end

# v3.1.x pattern (explicit validation required)
try:
    agent.add_slide("Blank", index=999)
except SlideNotFoundError as e:
    # Handle invalid index
    valid_range = e.details["available"]
    

# MIGRATION: Transparency to Opacity

# v3.0.x pattern (DEPRECATED)
agent.format_shape(slide_index=0, shape_index=1, transparency=0.5)
# ⚠️ Logs warning, converts to fill_opacity=0.5 internally

# v3.1.x pattern (RECOMMENDED)  
agent.format_shape(slide_index=0, shape_index=1, fill_opacity=0.5)
```

---

## Part 6: Troubleshooting Reference

### 6.1 Common Error Scenarios

| Symptom | Root Cause | Diagnosis | Solution |
|---------|------------|-----------|----------|
| `ShapeNotFoundError: index 10 out of range (0-8)` | Stale indices after structural op | Check last `remove_shape` or `set_z_order` | Call `get_slide_info()` immediately after |
| `FileLockError: timeout after 10s` | Stale `.pptx.lock` file or concurrent access | Check for orphaned lock files | Delete `.pptx.lock` or wait for other process |
| `ApprovalTokenError: token required` | Destructive op without token | Check operation type | Generate token with correct scope |
| `Chart formatting lost` | `replace_data()` failed, fell back to recreate | Check python-pptx version | Accept or manually re-apply via `format_chart()` |
| `Opacity not applied` | Shape has no `<a:solidFill>` | Check if shape has gradient/pattern fill | Set explicit fill color first |

### 6.2 Debugging OOXML

```python
# Inspect shape's raw XML for debugging

from lxml import etree

def debug_shape_xml(agent, slide_index: int, shape_index: int):
    """Print shape's XML for debugging visual features."""
    slide = agent.prs.slides[slide_index]
    shape = slide.shapes[shape_index]
    
    xml_str = etree.tostring(
        shape._sp, 
        pretty_print=True, 
        encoding='unicode'
    )
    print(xml_str)
    
    # Check for specific elements
    spPr = shape._sp.spPr
    from pptx.oxml.ns import qn
    
    solidFill = spPr.find(qn('a:solidFill'))
    if solidFill is not None:
        print("✅ Has <a:solidFill>")
        alpha = solidFill.find('.//' + qn('a:alpha'))
        if alpha is not None:
            print(f"✅ Has <a:alpha val='{alpha.get('val')}'>")
        else:
            print("❌ No <a:alpha> - opacity not set")
    else:
        print("❌ No <a:solidFill> - cannot inject opacity")
```

### 6.3 Lock File Recovery

```bash
# Find and remove stale lock files (Unix/Mac)
find /path/to/presentations -name "*.pptx.lock" -mmin +60 -delete

# Windows PowerShell
Get-ChildItem -Path "C:\presentations" -Filter "*.pptx.lock" | 
    Where-Object { $_.LastWriteTime -lt (Get-Date).AddHours(-1) } |
    Remove-Item
```

---

## Part 7: Integrated Mental Model

### 7.1 Complete Understanding Summary

```
┌─────────────────────────────────────────────────────────────────────────┐
│              POWERPOINT AGENT CORE — COMPLETE MENTAL MODEL              │
├─────────────────────────────────────────────────────────────────────────┤
│                                                                         │
│  IDENTITY:                                                              │
│  • Stateful-while-open core wrapped by stateless CLI tools              │
│  • Security-hardened with locking, tokens, path validation              │
│  • Geometry-aware versioning for change detection                       │
│  • XML manipulation for features missing in python-pptx                 │
│                                                                         │
│  LIFECYCLE:                                                             │
│  __init__ → open(lock) → [mutations] → save() → __exit__(unlock)       │
│                                                                         │
│  CRITICAL PROTOCOLS:                                                    │
│  1. Version tracking: before/after on every mutation                    │
│  2. Index freshness: refresh after add/remove/z-order                   │
│  3. Token enforcement: destructive ops require approval                 │
│  4. Path safety: traversal protection + extension whitelist             │
│                                                                         │
│  RETURN CONTRACT (v3.1.4):                                              │
│  • All mutations return Dict with result + versions                     │
│  • Never return bare primitives from mutation methods                   │
│  • Include styling details for visual operations                        │
│                                                                         │
│  XML "MAGIC":                                                           │
│  • Opacity: Inject <a:alpha val="N"/> into <a:srgbClr>                 │
│  • Z-Order: Move <p:sp> within <p:spTree> (index 2 = back, end = front)│
│  • Scale: OOXML uses 0-100000, core converts from 0.0-1.0              │
│                                                                         │
│  PERFORMANCE:                                                           │
│  • Version hashing: O(N) shapes (~15ms/10 slides)                      │
│  • Deep probe: O(M) layouts with 15s timeout                           │
│  • Shape ops: O(1) constant time                                       │
│  • Save: I/O bound, scales with file size                              │
│                                                                         │
│  BACKWARD COMPAT:                                                       │
│  • Dict returns work with legacy code via duck typing                   │
│  • transparency → fill_opacity conversion with warning                  │
│  • Silent clamping REMOVED (intentional breaking change)                │
│                                                                         │
└─────────────────────────────────────────────────────────────────────────┘
```

---

## ✅ Core Handbook Integration Complete

I have now integrated the **PowerPointAgent Core Programming Handbook v3.1.4** with my understanding of the overall ecosystem. My knowledge now spans:

| Layer | Documents Integrated | Confidence |
|-------|---------------------|------------|
| **Project Overview** | README, Quick Start | ✅ High |
| **Agent Operations** | AGENT_SYSTEM_PROMPT.md | ✅ High |
| **Programming Guide** | Comprehensive Programming & Troubleshooting Guide | ✅ High |
| **Core Implementation** | PowerPoint Agent Core Handbook v3.1.4 | ✅ High |

**Ready for development tasks** with complete understanding of:
- Architecture and design philosophy
- CLI tool implementation patterns
- Core library internals and APIs
- XML manipulation techniques
- Security and governance enforcement
- Performance characteristics
- Troubleshooting procedures
- Backward compatibility requirements
