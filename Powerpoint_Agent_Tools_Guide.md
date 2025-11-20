# 📘 PowerPoint Agent Tools: Technical Reference Guide

This comprehensive guide covers the usage, architecture, and best practices for the **PowerPoint Agent Tool Suite**. It is designed for developers and AI agents interacting with the 30-tool CLI ecosystem.

## 🧠 Core Concepts

### 1. Stateless Architecture
The tools are designed to be **stateless**. The agent does not hold the file open between commands.
*   **Atomic Operations:** Each CLI command Opens -> Modifies -> Saves -> Closes.
*   **Concurrency:** A file locking mechanism prevents race conditions.
*   **Implication:** You must provide the `--file` path for every command.

### 2. The Inspection Pattern
To edit an element, you must first know its ID (Index).
1.  **Inspect:** Run `ppt_get_slide_info.py` to get a JSON list of shapes on a slide.
2.  **Target:** Identify the `shape_index` of the object you want to modify.
3.  **Act:** Use tools like `ppt_format_shape.py` or `ppt_format_text.py` using that index.

### 3. The Coordinate System
PowerPoint uses absolute positioning (Inches), but these tools abstract this via the **Position Dictionary**.
*   **Canvas:** Standard Wide slide is **10.0" wide x 7.5" high**.
*   **Responsive:** Use **Percentages** (`"left": "10%"`) for layouts that survive aspect ratio changes.

---

## 📚 Tool Catalog (30 Tools)

### 🏗️ Domain 1: Creation & Architecture

| Tool | Description | Key Arguments |
|------|-------------|---------------|
| `ppt_create_new.py` | Initialize a blank deck. | `--output`, `--slides`, `--layout` |
| `ppt_create_from_template.py` | Initialize from a corporate `.pptx`. | `--template`, `--output` |
| `ppt_create_from_structure.py` | **(Recommended)** Build full deck from JSON. | `--structure`, `--output` |
| `ppt_clone_presentation.py` | Create a "Save As" copy. | `--source`, `--output` |

### 🎞️ Domain 2: Slide Management

| Tool | Description | Key Arguments |
|------|-------------|---------------|
| `ppt_add_slide.py` | Insert new slide. | `--layout`, `--index` |
| `ppt_delete_slide.py` | Remove slide. | `--index` |
| `ppt_duplicate_slide.py` | Clone slide (keep layout/content). | `--index` |
| `ppt_reorder_slides.py` | Move slide position. | `--from-index`, `--to-index` |
| `ppt_set_slide_layout.py` | Apply new master layout. | `--layout` |

### 📝 Domain 3: Text & Content

| Tool | Description | Key Arguments |
|------|-------------|---------------|
| `ppt_set_title.py` | Set Title/Subtitle placeholders. | `--title`, `--subtitle` |
| `ppt_add_text_box.py` | Add free-floating text. | `--text`, `--position`, `--size` |
| `ppt_add_bullet_list.py` | Add lists (bullet/numbered). | `--items`, `--bullet-style` |
| `ppt_format_text.py` | Style existing text (req. shape idx). | `--shape`, `--font-size`, `--color` |
| `ppt_replace_text.py` | Global find/replace. | `--find`, `--replace`, `--match-case` |

### 🖼️ Domain 4: Images & Media

| Tool | Description | Key Arguments |
|------|-------------|---------------|
| `ppt_insert_image.py` | Add image file. | `--image`, `--size "auto"` |
| `ppt_replace_image.py` | Swap image (preserve layout). | `--old-image`, `--new-image` |
| `ppt_set_image_properties.py` | Set Alt-Text/Transparency. | `--shape`, `--alt-text` |

### 🎨 Domain 5: Vector Graphics

| Tool | Description | Key Arguments |
|------|-------------|---------------|
| `ppt_add_shape.py` | Add geometry (Arrow, Rect, Star). | `--shape`, `--fill-color` |
| `ppt_format_shape.py` | Style fill/border. | `--shape`, `--line-width` |
| `ppt_add_connector.py` | Draw line between shapes. | `--from-shape`, `--to-shape` |

### 📊 Domain 6: Data Visualization

| Tool | Description | Key Arguments |
|------|-------------|---------------|
| `ppt_add_chart.py` | Add Column/Pie/Line charts. | `--chart-type`, `--data` |
| `ppt_format_chart.py` | Update title/legend. | `--chart`, `--legend` |
| `ppt_add_table.py` | Add data grid. | `--rows`, `--cols`, `--data` |

### 🔍 Domain 7: Inspection & Analysis

| Tool | Description | Output Data |
|------|-------------|-------------|
| `ppt_get_info.py` | Deck metadata. | Slide count, dimensions, layouts. |
| `ppt_get_slide_info.py` | **Critical.** Slide contents. | Shape indices, types, text content. |
| `ppt_extract_notes.py` | Speaker notes. | Dictionary of notes per slide. |

### 🛡️ Domain 8: Validation & Output

| Tool | Description | Key Features |
|------|-------------|--------------|
| `ppt_validate_presentation.py` | Health check. | Missing assets, empty slides. |
| `ppt_check_accessibility.py` | WCAG 2.1 audit. | Contrast, Alt Text, Reading Order. |
| `ppt_export_pdf.py` | PDF Conversion. | Requires LibreOffice. |
| `ppt_export_images.py` | Slide -> IMG. | PNG/JPG support. |

---

## 📐 Reference: Positioning & Data Schemas

### 1. Position Dictionary
AI Agents should prefer **Percentage** for robustness.

```json
// Percentage (Responsive)
{"left": "10%", "top": "20%"}

// Anchor (Layout aware)
{"anchor": "bottom_right", "offset_x": -0.5, "offset_y": -0.5}

// Grid (Structured)
{"grid_row": 2, "grid_col": 6, "grid_size": 12}

// Absolute (Precise)
{"left": 1.5, "top": 2.0}
```

### 2. Chart Data Schema
Used for `ppt_add_chart.py`.
```json
{
  "categories": ["Q1", "Q2", "Q3", "Q4"],
  "series": [
    {
      "name": "Revenue",
      "values": [10.5, 12.0, 15.5, 18.0]
    },
    {
      "name": "Profit",
      "values": [2.5, 3.0, 4.5, 6.0]
    }
  ]
}
```

---

## 🧪 Workflow Recipes

### Recipe 1: The "Auditor" (Quality Assurance)
This workflow analyzes a deck and fixes common issues.

```bash
# 1. Check for accessibility violations
uv python tools/ppt_check_accessibility.py --file deck.pptx --json

# 2. (Agent Logic) If missing alt text on Slide 2, Shape 1:
uv python tools/ppt_set_image_properties.py \
  --file deck.pptx \
  --slide 2 \
  --shape 1 \
  --alt-text "Quarterly Revenue Graph showing 15% growth" \
  --json

# 3. Verify branding (Replace old company name)
uv python tools/ppt_replace_text.py \
  --file deck.pptx \
  --find "OldCorp" \
  --replace "NewCorp" \
  --match-case \
  --json

# 4. Final validation
uv python tools/ppt_validate_presentation.py --file deck.pptx --json
```

### Recipe 2: The "Architect" (Complex Diagram)
Builds a flowchart using shapes and connectors.

```bash
# 1. Add "Start" Node
uv python tools/ppt_add_shape.py \
  --file flow.pptx --slide 0 --shape rounded_rectangle \
  --position '{"grid_col":1, "grid_row":2}' --size '{"width":"15%","height":"10%"}' \
  --json
# (Assume this returns shape_index 0)

# 2. Add "Process" Node
uv python tools/ppt_add_shape.py \
  --file flow.pptx --slide 0 --shape rectangle \
  --position '{"grid_col":4, "grid_row":2}' --size '{"width":"15%","height":"10%"}' \
  --json
# (Assume this returns shape_index 1)

# 3. Connect them
uv python tools/ppt_add_connector.py \
  --file flow.pptx --slide 0 \
  --from-shape 0 --to-shape 1 \
  --type straight \
  --json
```

### Recipe 3: The "Analyst" (Data Report)
Generates a report from raw data.

```bash
# 1. Create from Corporate Template
uv python tools/ppt_create_from_template.py \
  --template master.pptx --output report.pptx --slides 1 --json

# 2. Add Table
uv python tools/ppt_add_table.py \
  --file report.pptx --slide 0 \
  --rows 4 --cols 3 \
  --headers "Metric,Value,YoY" \
  --data '[["Users","1M","+10%"],["Rev","$5M","+20%"],["Churn","2%","-1%"]]' \
  --position '{"left":"5%","top":"20%"}' --size '{"width":"40%","height":"50%"}' \
  --json

# 3. Add Chart (Visualizing the same data)
uv python tools/ppt_add_chart.py \
  --file report.pptx --slide 0 \
  --chart-type bar \
  --data '{"categories":["Users","Rev"],"series":[{"name":"YoY","values":[10,20]}]}' \
  --position '{"left":"50%","top":"20%"}' --size '{"width":"45%","height":"50%"}' \
  --json

# 4. Export for email
uv python tools/ppt_export_pdf.py --file report.pptx --output report.pdf --json
```

---

## 🚑 Troubleshooting

### Common Exit Code 1 Errors

1.  **`SlideNotFoundError`**:
    *   *Cause:* You requested index `5` but the deck only has `5` slides (indices 0-4).
    *   *Fix:* Run `ppt_get_info.py` to check `slide_count`.

2.  **`LayoutNotFoundError`**:
    *   *Cause:* You requested "Title Slide" but the template calls it "Title".
    *   *Fix:* Run `ppt_get_info.py` to see the `layouts` list.

3.  **`ImageNotFoundError`**:
    *   *Cause:* The tool cannot access the image path.
    *   *Fix:* Use absolute paths or ensure the agent's working directory is correct.

4.  **`InvalidPositionError`**:
    *   *Cause:* Malformed JSON in the `--position` argument.
    *   *Fix:* Ensure strict JSON syntax (double quotes for keys/strings).
    *   *Example:* `'{"left": "10%"}'` (Correct) vs `{'left': '10%'}` (Incorrect Python dict syntax).
