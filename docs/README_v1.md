# 🎨 PowerPoint Agent Tools

<div align="center">

![PowerPoint Agent Tools](https://img.shields.io/badge/PowerPoint-Agent_Tools-B7472A?style=for-the-badge&logo=microsoft-powerpoint&logoColor=white)
[![Python Version](https://img.shields.io/badge/python-3.8+-blue.svg?style=for-the-badge&logo=python&logoColor=white)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg?style=for-the-badge)](https://opensource.org/licenses/MIT)
[![PRs Welcome](https://img.shields.io/badge/PRs-welcome-brightgreen.svg?style=for-the-badge)](http://makeapullrequest.com)

**Production-grade PowerPoint manipulation for AI agents**

Build, edit, and enhance PowerPoint presentations programmatically through simple CLI tools designed for AI consumption.

[Features](#-features) • [Quick Start](#-quick-start) • [Installation](#-installation) • [Documentation](#-documentation) • [Examples](#-examples) • [Contributing](#-contributing)

</div>

---

## 🎯 Why PowerPoint Agent Tools?

Traditional PowerPoint libraries require complex Python code and offer limited AI-agent-friendly APIs. **PowerPoint Agent Tools** provides:

✅ **CLI-First Design** - AI agents call simple commands, no Python knowledge required  
✅ **JSON Everywhere** - All outputs machine-parsable for easy integration  
✅ **Flexible Positioning** - 5 positioning systems (%, anchor, grid, Excel-like, absolute)  
✅ **Data-Driven** - Create charts, tables from JSON data files  
✅ **Stateless Operations** - No session management, perfect for distributed systems  
✅ **Visual Design** - Shapes, colors, layouts for professional presentations  
✅ **Bulk Operations** - Find/replace, clone, template-based workflows  
✅ **Accessibility** - Built-in alt-text, WCAG-aware color helpers  

---

## 🚀 Quick Start

Create a professional presentation in 60 seconds:

```bash
# Install
pip install python-pptx Pillow

# Create presentation
uv python tools/ppt_create_new.py \
  --output quarterly_results.pptx \
  --slides 3 \
  --json

# Set title
uv python tools/ppt_set_title.py \
  --file quarterly_results.pptx \
  --slide 0 \
  --title "Q4 Results" \
  --subtitle "Record Breaking Performance" \
  --json

# Add bullet points
uv python tools/ppt_add_bullet_list.py \
  --file quarterly_results.pptx \
  --slide 1 \
  --items "Revenue up 45%,Customer growth 60%,Market share 23%" \
  --position '{"left":"10%","top":"25%"}' \
  --size '{"width":"80%","height":"60%"}' \
  --json

# Add chart
cat > data.json << 'EOF'
{
  "categories": ["Q1", "Q2", "Q3", "Q4"],
  "series": [{"name": "Revenue", "values": [100, 120, 140, 160]}]
}
EOF

uv python tools/ppt_add_chart.py \
  --file quarterly_results.pptx \
  --slide 2 \
  --chart-type column \
  --data data.json \
  --position '{"left":"10%","top":"20%"}' \
  --size '{"width":"80%","height":"65%"}' \
  --json
```

**Result:** Professional 3-slide presentation with title, bullets, and chart! 🎉

---

## 📦 Installation

### Requirements

- **Python:** 3.8 or higher
- **Dependencies:** `python-pptx`, `Pillow`
- **Optional:** `pandas` (for advanced data operations)

### Method 1: Quick Install (pip)

```bash
# Clone repository
git clone https://github.com/your-org/powerpoint-agent-tools.git
cd powerpoint-agent-tools

# Install dependencies
pip install -r requirements.txt

# Verify installation
python tools/ppt_create_new.py --help
```

### Method 2: Using uv (Recommended)

```bash
# Install uv (fast Python package manager)
curl -LsSf https://astral.sh/uv/install.sh | sh

# Clone and install
git clone https://github.com/your-org/powerpoint-agent-tools.git
cd powerpoint-agent-tools
uv pip install -r requirements.txt

# Test
uv python tools/ppt_create_new.py --help
```

### Method 3: Development Setup

```bash
# Clone and setup dev environment
git clone https://github.com/your-org/powerpoint-agent-tools.git
cd powerpoint-agent-tools

# Create virtual environment
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install with dev dependencies
pip install -r requirements.txt
pip install pytest pytest-cov  # For testing

# Run tests
pytest test_basic_tools.py -v
```

### Dependencies (`requirements.txt`)

```txt
python-pptx==0.6.23    # Core PowerPoint manipulation
Pillow>=10.0.0         # Image processing
pandas>=2.0.0          # Optional: DataFrame support
```

---

## 🛠️ Tool Catalog

### 📁 Creation Tools (1 tool)

| Tool | Purpose | Key Features |
|------|---------|--------------|
| [`ppt_create_new.py`](#ppt_create_newpy) | Create new presentation | Templates, layouts, multi-slide |

### 📊 Slide Management (2 tools)

| Tool | Purpose | Key Features |
|------|---------|--------------|
| [`ppt_add_slide.py`](#ppt_add_slidepy) | Add slide | Any position, layouts, auto-title |
| [`ppt_set_title.py`](#ppt_set_titlepy) | Set title/subtitle | Title + subtitle support |

### ✏️ Content Creation (5 tools)

| Tool | Purpose | Key Features |
|------|---------|--------------|
| [`ppt_add_text_box.py`](#ppt_add_text_boxpy) | Add text box | 5 positioning systems, formatting |
| [`ppt_insert_image.py`](#ppt_insert_imagepy) | Insert image | Auto aspect ratio, compression, alt-text |
| [`ppt_add_bullet_list.py`](#ppt_add_bullet_listpy) | Add lists | Bullets, numbered, custom formatting |
| [`ppt_add_chart.py`](#ppt_add_chartpy) | Add chart | 9 chart types, data validation |
| [`ppt_add_table.py`](#ppt_add_tablepy) | Add table | Headers, JSON data import |

### 🎨 Visual Design (1 tool)

| Tool | Purpose | Key Features |
|------|---------|--------------|
| [`ppt_add_shape.py`](#ppt_add_shapepy) | Add shapes | 10 shapes, colors, layering |

### 🔧 Utilities (1 tool)

| Tool | Purpose | Key Features |
|------|---------|--------------|
| [`ppt_replace_text.py`](#ppt_replace_textpy) | Find/replace | Bulk updates, dry-run, case-sensitive |

**Total:** 10 production-ready tools

---

## 📖 Tool Documentation

### `ppt_create_new.py`

Create new PowerPoint presentation with specified slides.

**Usage:**
```bash
uv python tools/ppt_create_new.py \
  --output presentation.pptx \
  --slides 5 \
  --layout "Title and Content" \
  --json
```

**Arguments:**
- `--output PATH` (required) - Output file path
- `--slides N` - Number of slides (default: 1)
- `--template PATH` - Template .pptx file
- `--layout NAME` - Slide layout (default: "Title and Content")
- `--json` - JSON output

**Available Layouts:**
- Title Slide
- Title and Content
- Section Header
- Two Content
- Comparison
- Title Only
- Blank
- Picture with Caption

**Example Response:**
```json
{
  "status": "success",
  "file": "presentation.pptx",
  "slides_created": 5,
  "file_size_bytes": 28432,
  "available_layouts": ["Title Slide", "Title and Content", ...]
}
```

---

### `ppt_add_slide.py`

Add new slide to existing presentation.

**Usage:**
```bash
uv python tools/ppt_add_slide.py \
  --file presentation.pptx \
  --layout "Title and Content" \
  --title "Revenue Growth" \
  --index 2 \
  --json
```

**Arguments:**
- `--file PATH` (required) - Presentation file
- `--layout NAME` (required) - Layout name
- `--index N` - Position (0-based, default: end)
- `--title TEXT` - Optional title to set
- `--json` - JSON output

---

### `ppt_set_title.py`

Set slide title and subtitle.

**Usage:**
```bash
uv python tools/ppt_set_title.py \
  --file presentation.pptx \
  --slide 0 \
  --title "Q4 Results" \
  --subtitle "Financial Review" \
  --json
```

**Arguments:**
- `--file PATH` (required)
- `--slide INDEX` (required) - 0-based slide index
- `--title TEXT` (required)
- `--subtitle TEXT` - Optional subtitle
- `--json` - JSON output

---

### `ppt_add_text_box.py`

Add text box with flexible positioning.

**Usage:**
```bash
uv python tools/ppt_add_text_box.py \
  --file presentation.pptx \
  --slide 1 \
  --text "Revenue increased 45%" \
  --position '{"left":"20%","top":"30%"}' \
  --size '{"width":"60%","height":"10%"}' \
  --font-size 24 \
  --bold \
  --color "#0070C0" \
  --json
```

**Arguments:**
- `--file PATH` (required)
- `--slide INDEX` (required)
- `--text TEXT` (required)
- `--position JSON` (required) - See [Positioning Guide](#-positioning-systems)
- `--size JSON` (required)
- `--font-name NAME` - Default: Calibri
- `--font-size N` - Default: 18
- `--bold` - Bold text
- `--italic` - Italic text
- `--color HEX` - Text color (e.g., #FF0000)
- `--alignment {left,center,right,justify}` - Default: left
- `--json` - JSON output

---

### `ppt_insert_image.py`

Insert image with automatic aspect ratio handling.

**Usage:**
```bash
uv python tools/ppt_insert_image.py \
  --file presentation.pptx \
  --slide 0 \
  --image logo.png \
  --position '{"left":"5%","top":"5%"}' \
  --size '{"width":"15%","height":"auto"}' \
  --alt-text "Company Logo" \
  --compress \
  --json
```

**Arguments:**
- `--file PATH` (required)
- `--slide INDEX` (required)
- `--image PATH` (required)
- `--position JSON` (required)
- `--size JSON` - Use "auto" for aspect ratio
- `--compress` - Compress before inserting
- `--alt-text TEXT` - Accessibility (recommended)
- `--json` - JSON output

**Supported Formats:** PNG, JPG, JPEG, GIF, BMP

---

### `ppt_add_bullet_list.py`

Add bullet or numbered list.

**Usage:**
```bash
uv python tools/ppt_add_bullet_list.py \
  --file presentation.pptx \
  --slide 1 \
  --items "Revenue up 45%,Customers +60%,Market share 23%" \
  --position '{"left":"10%","top":"25%"}' \
  --size '{"width":"80%","height":"60%"}' \
  --bullet-style bullet \
  --font-size 22 \
  --json
```

**Arguments:**
- `--file PATH` (required)
- `--slide INDEX` (required)
- `--items TEXT` (required) - Comma-separated
- `--items-file PATH` - JSON array file
- `--position JSON` (required)
- `--size JSON` (required)
- `--bullet-style {bullet,numbered,none}` - Default: bullet
- `--font-size N` - Default: 18
- `--color HEX` - Text color
- `--json` - JSON output

---

### `ppt_add_chart.py`

Add data visualization chart.

**Usage:**
```bash
# Create data file
cat > data.json << 'EOF'
{
  "categories": ["Q1", "Q2", "Q3", "Q4"],
  "series": [
    {"name": "Revenue", "values": [100, 120, 140, 160]},
    {"name": "Costs", "values": [80, 90, 100, 110]}
  ]
}
EOF

uv python tools/ppt_add_chart.py \
  --file presentation.pptx \
  --slide 2 \
  --chart-type column \
  --data data.json \
  --position '{"left":"10%","top":"20%"}' \
  --size '{"width":"80%","height":"60%"}' \
  --title "Revenue vs Costs" \
  --json
```

**Chart Types:**
- `column` - Vertical bars
- `column_stacked` - Stacked vertical
- `bar` - Horizontal bars
- `bar_stacked` - Stacked horizontal
- `line` - Line chart
- `line_markers` - Line with markers
- `pie` - Pie chart
- `area` - Area chart
- `scatter` - Scatter plot

**Arguments:**
- `--file PATH` (required)
- `--slide INDEX` (required)
- `--chart-type TYPE` (required)
- `--data PATH` (required) - JSON data file
- `--position JSON` (required)
- `--size JSON` (required)
- `--title TEXT` - Chart title
- `--json` - JSON output

---

### `ppt_add_table.py`

Add data table.

**Usage:**
```bash
# Create table data
cat > table_data.json << 'EOF'
[
  ["Q1", "10.5", "8.2", "2.3"],
  ["Q2", "12.8", "9.1", "3.7"],
  ["Q3", "15.2", "10.5", "4.7"],
  ["Q4", "18.6", "12.1", "6.5"]
]
EOF

uv python tools/ppt_add_table.py \
  --file presentation.pptx \
  --slide 3 \
  --rows 5 \
  --cols 4 \
  --headers "Quarter,Revenue,Costs,Profit" \
  --data table_data.json \
  --position '{"left":"15%","top":"25%"}' \
  --size '{"width":"70%","height":"50%"}' \
  --json
```

**Arguments:**
- `--file PATH` (required)
- `--slide INDEX` (required)
- `--rows N` (required)
- `--cols N` (required)
- `--position JSON` (required)
- `--size JSON` (required)
- `--data PATH` - JSON 2D array
- `--headers TEXT` - Comma-separated
- `--json` - JSON output

---

### `ppt_add_shape.py`

Add shape for visual design.

**Usage:**
```bash
uv python tools/ppt_add_shape.py \
  --file presentation.pptx \
  --slide 1 \
  --shape rounded_rectangle \
  --position '{"left":"10%","top":"15%"}' \
  --size '{"width":"30%","height":"15%"}' \
  --fill-color "#0070C0" \
  --line-color "#FFFFFF" \
  --line-width 2 \
  --json
```

**Available Shapes:**
- `rectangle`, `rounded_rectangle`, `ellipse`, `triangle`
- `arrow_right`, `arrow_left`, `arrow_up`, `arrow_down`
- `star`, `heart`

**Arguments:**
- `--file PATH` (required)
- `--slide INDEX` (required)
- `--shape TYPE` (required)
- `--position JSON` (required)
- `--size JSON` (required)
- `--fill-color HEX` - Fill color
- `--line-color HEX` - Border color
- `--line-width FLOAT` - Border width
- `--json` - JSON output

---

### `ppt_replace_text.py`

Find and replace text across presentation.

**Usage:**
```bash
# Dry run first (recommended)
uv python tools/ppt_replace_text.py \
  --file presentation.pptx \
  --find "2023" \
  --replace "2024" \
  --dry-run \
  --json

# Apply changes
uv python tools/ppt_replace_text.py \
  --file presentation.pptx \
  --find "2023" \
  --replace "2024" \
  --match-case \
  --json
```

**Arguments:**
- `--file PATH` (required)
- `--find TEXT` (required)
- `--replace TEXT` (required)
- `--match-case` - Case-sensitive
- `--dry-run` - Preview without changes
- `--json` - JSON output

---

## 🎯 Positioning Systems

PowerPoint Agent Tools offers **5 flexible positioning systems** for precise element placement:

### 1. Percentage (Recommended for AI)

**Best for:** Responsive layouts that work on any slide size

```json
{"left": "20%", "top": "30%"}
```

**Example:**
```bash
--position '{"left":"10%","top":"25%"}'
--size '{"width":"80%","height":"60%"}'
```

**Advantages:**
- ✅ Responsive to slide dimensions
- ✅ Intuitive (0% = edge, 100% = opposite edge)
- ✅ Works with 16:9 and 4:3 aspect ratios

---

### 2. Anchor Points

**Best for:** Standard locations (headers, footers, centered content)

```json
{"anchor": "center", "offset_x": 0, "offset_y": -1.0}
```

**Available Anchors:**
```
top_left      top_center      top_right
center_left   center          center_right
bottom_left   bottom_center   bottom_right
```

**Example:**
```bash
# Centered title
--position '{"anchor":"center","offset_y":-2.0}'

# Bottom right copyright
--position '{"anchor":"bottom_right","offset_x":-0.5,"offset_y":-0.3}'
```

**Advantages:**
- ✅ Semantic positioning
- ✅ Easy to understand intent
- ✅ Consistent placement

---

### 3. Grid System (12×12)

**Best for:** Structured layouts, aligning multiple elements

```json
{"grid_row": 2, "grid_col": 3, "grid_size": 12}
```

**Example:**
```bash
# Top-left quadrant
--position '{"grid_row":1,"grid_col":1,"grid_size":12}'

# Center area
--position '{"grid_row":5,"grid_col":4,"grid_size":12}'
```

**Advantages:**
- ✅ Precise alignment
- ✅ Like Bootstrap grid
- ✅ Consistent spacing

---

### 4. Excel-like Grid

**Best for:** Familiar reference for Excel users

```json
{"grid": "C4"}
```

**Example:**
```bash
--position '{"grid":"A1"}'   # Top-left
--position '{"grid":"C4"}'   # Center-ish
--position '{"grid":"L12"}'  # Bottom-right
```

**Advantages:**
- ✅ Familiar to spreadsheet users
- ✅ Concise notation
- ✅ Easy to communicate

---

### 5. Absolute Inches

**Best for:** Precise positioning from design specifications

```json
{"left": 2.0, "top": 3.0}
```

**Slide Dimensions (16:9):**
- Width: 10.0 inches
- Height: 7.5 inches

**Example:**
```bash
--position '{"left":1.5,"top":2.0}'
```

**Advantages:**
- ✅ Exact positioning
- ✅ Matches design specs
- ✅ No calculations needed

---

## 📚 Examples

### Example 1: Quarterly Business Review

```bash
#!/bin/bash
# Generate Q4 Business Review presentation

PPTX="q4_review.pptx"

# Create presentation
uv python tools/ppt_create_new.py \
  --output $PPTX \
  --slides 1 \
  --layout "Title Slide" \
  --json

# Title slide
uv python tools/ppt_set_title.py \
  --file $PPTX \
  --slide 0 \
  --title "Q4 2024 Business Review" \
  --subtitle "Record-Breaking Performance" \
  --json

# Add company logo
uv python tools/ppt_insert_image.py \
  --file $PPTX \
  --slide 0 \
  --image assets/logo.png \
  --position '{"anchor":"top_right","offset_x":-0.5,"offset_y":0.5}' \
  --size '{"width":"12%","height":"auto"}' \
  --alt-text "Company Logo" \
  --json

# Agenda slide
uv python tools/ppt_add_slide.py \
  --file $PPTX \
  --layout "Title and Content" \
  --title "Agenda" \
  --json

uv python tools/ppt_add_bullet_list.py \
  --file $PPTX \
  --slide 1 \
  --items "Financial highlights,Customer growth,Product launches,2025 outlook" \
  --bullet-style numbered \
  --position '{"left":"15%","top":"25%"}' \
  --size '{"width":"70%","height":"60%"}' \
  --font-size 24 \
  --json

# Revenue chart
cat > revenue_data.json << 'EOF'
{
  "categories": ["Q1", "Q2", "Q3", "Q4"],
  "series": [
    {"name": "2023", "values": [10.2, 11.5, 12.8, 14.3]},
    {"name": "2024", "values": [12.5, 15.2, 18.7, 22.1]}
  ]
}
EOF

uv python tools/ppt_add_slide.py \
  --file $PPTX \
  --layout "Title and Content" \
  --title "Revenue Growth" \
  --json

uv python tools/ppt_add_chart.py \
  --file $PPTX \
  --slide 2 \
  --chart-type column \
  --data revenue_data.json \
  --position '{"left":"10%","top":"20%"}' \
  --size '{"width":"80%","height":"65%"}' \
  --title "Quarterly Revenue ($M)" \
  --json

# Add emphasis callout
uv python tools/ppt_add_shape.py \
  --file $PPTX \
  --slide 2 \
  --shape rounded_rectangle \
  --position '{"left":"70%","top":"12%"}' \
  --size '{"width":"25%","height":"12%"}' \
  --fill-color "#00B050" \
  --json

uv python tools/ppt_add_text_box.py \
  --file $PPTX \
  --slide 2 \
  --text "54% Growth!" \
  --position '{"left":"72%","top":"14%"}' \
  --size '{"width":"21%","height":"8%"}' \
  --font-size 22 \
  --bold \
  --color "#FFFFFF" \
  --alignment center \
  --json

echo "✅ Q4 Business Review created: $PPTX"
```

---

### Example 2: Data-Driven Report from Database

```python
#!/usr/bin/env python3
"""Generate presentation from database."""
import subprocess
import json
import psycopg2

def run_tool(tool, args):
    """Execute PowerPoint tool."""
    cmd = ['uv', 'python', f'tools/{tool}', '--json'] + args
    result = subprocess.run(cmd, capture_output=True, text=True)
    return json.loads(result.stdout)

# Fetch data from database
conn = psycopg2.connect("dbname=analytics user=admin")
cur = conn.cursor()

cur.execute("""
    SELECT quarter, revenue, costs, profit 
    FROM financial_metrics 
    WHERE year = 2024 
    ORDER BY quarter
""")

data = cur.fetchall()
categories = [row[0] for row in data]
revenue_values = [float(row[1]) for row in data]
costs_values = [float(row[2]) for row in data]

# Create chart data
chart_data = {
    "categories": categories,
    "series": [
        {"name": "Revenue", "values": revenue_values},
        {"name": "Costs", "values": costs_values}
    ]
}

with open('chart_data.json', 'w') as f:
    json.dump(chart_data, f)

# Create presentation
run_tool('ppt_create_new.py', [
    '--output', 'report.pptx',
    '--slides', '1'
])

# Add chart
run_tool('ppt_add_chart.py', [
    '--file', 'report.pptx',
    '--slide', '0',
    '--chart-type', 'column',
    '--data', 'chart_data.json',
    '--position', '{"left":"10%","top":"20%"}',
    '--size', '{"width":"80%","height":"60%"}',
    '--title', '2024 Financial Performance'
])

print("✅ Report generated from database")
```

---

### Example 3: Template-Based Automation

```bash
#!/bin/bash
# Generate client-specific presentations from template

TEMPLATE="corporate_template.pptx"
CLIENT_NAME="Acme Corp"
OUTPUT="acme_presentation.pptx"

# Clone template
cp $TEMPLATE $OUTPUT

# Customize with client data
uv python tools/ppt_replace_text.py \
  --file $OUTPUT \
  --find "{{CLIENT_NAME}}" \
  --replace "$CLIENT_NAME" \
  --json

uv python tools/ppt_replace_text.py \
  --file $OUTPUT \
  --find "{{DATE}}" \
  --replace "$(date +%Y-%m-%d)" \
  --json

# Add client logo
uv python tools/ppt_insert_image.py \
  --file $OUTPUT \
  --slide 0 \
  --image "clients/${CLIENT_NAME}_logo.png" \
  --position '{"anchor":"top_right","offset_x":-0.5,"offset_y":0.5}' \
  --size '{"width":"15%","height":"auto"}' \
  --alt-text "$CLIENT_NAME Logo" \
  --json

# Add custom data slide
cat > client_metrics.json << EOF
{
  "categories": ["Jan", "Feb", "Mar", "Apr"],
  "series": [{"name": "Performance", "values": [85, 92, 88, 95]}]
}
EOF

uv python tools/ppt_add_slide.py \
  --file $OUTPUT \
  --layout "Title and Content" \
  --title "Your Performance Dashboard" \
  --json

uv python tools/ppt_add_chart.py \
  --file $OUTPUT \
  --slide 5 \
  --chart-type line_markers \
  --data client_metrics.json \
  --position '{"left":"10%","top":"20%"}' \
  --size '{"width":"80%","height":"65%"}' \
  --json

echo "✅ Client presentation created: $OUTPUT"
```

---

## 🎨 Design Best Practices

### 1. Visual Hierarchy

**Slide Structure:**
```
1. Title (largest) - 36-48pt
2. Subtitle (medium) - 24-32pt
3. Body text (standard) - 18-24pt
4. Captions (smallest) - 12-16pt
```

### 2. Color Palette

**Corporate Standard:**
```bash
PRIMARY_BLUE="#0070C0"     # Headers, important text
SECONDARY_GRAY="#595959"   # Body text
ACCENT_ORANGE="#ED7D31"    # Callouts, highlights
SUCCESS_GREEN="#70AD47"    # Positive metrics
WARNING_YELLOW="#FFC000"   # Cautions
DANGER_RED="#C00000"       # Negative metrics
```

### 3. Content Guidelines

**6×6 Rule:**
- Maximum 6 bullets per slide
- Maximum 6 words per bullet
- Exceptions: Charts, tables (visual data)

**Slide Timing:**
- 1 slide = 1-2 minutes of speaking
- 30-minute presentation = 15-20 slides
- Leave time for Q&A

### 4. Layout Patterns

**Common Layouts:**
1. **Title Slide** - Opening (title + subtitle + logo)
2. **Agenda** - Roadmap (numbered list)
3. **Section Header** - Topic breaks (large text)
4. **Title and Content** - Standard slides (title + content)
5. **Two Content** - Comparisons (side-by-side)
6. **Title Only** - Maximum content space
7. **Blank** - Custom layouts

### 5. Accessibility

**Always Include:**
- ✅ Alt text for all images
- ✅ High contrast text (4.5:1 minimum)
- ✅ Readable font sizes (18pt+)
- ✅ Descriptive slide titles
- ✅ Logical reading order

---

## 🔧 Advanced Usage

### Shell Script Automation

```bash
#!/bin/bash
# Generate monthly report automatically

MONTH=$(date +%B)
YEAR=$(date +%Y)
PPTX="monthly_report_${MONTH}_${YEAR}.pptx"

# Create base presentation
uv python tools/ppt_create_new.py \
  --output "$PPTX" \
  --slides 5 \
  --json

# Set title
uv python tools/ppt_set_title.py \
  --file "$PPTX" \
  --slide 0 \
  --title "$MONTH $YEAR Report" \
  --json

# Loop through KPIs
for kpi in revenue customers satisfaction; do
  # Fetch data (example using curl to API)
  curl -s "https://api.example.com/kpi/$kpi" > "${kpi}_data.json"
  
  # Add slide with chart
  uv python tools/ppt_add_slide.py \
    --file "$PPTX" \
    --layout "Title and Content" \
    --title "$(echo $kpi | tr '[:lower:]' '[:upper:]')" \
    --json
  
  uv python tools/ppt_add_chart.py \
    --file "$PPTX" \
    --slide $((${kpi_index} + 1)) \
    --chart-type line \
    --data "${kpi}_data.json" \
    --position '{"left":"10%","top":"20%"}' \
    --size '{"width":"80%","height":"65%"}' \
    --json
done

echo "✅ Monthly report generated: $PPTX"
```

### Python Integration

```python
#!/usr/bin/env python3
"""PowerPoint automation library."""
import subprocess
import json
from typing import Dict, Any, List

class PowerPointAgent:
    """High-level PowerPoint automation interface."""
    
    def __init__(self, filename: str):
        self.filename = filename
    
    def _run_tool(self, tool: str, args: List[str]) -> Dict[str, Any]:
        """Execute tool and return JSON response."""
        cmd = ['uv', 'python', f'tools/{tool}', '--json'] + args
        result = subprocess.run(cmd, capture_output=True, text=True)
        
        if result.returncode != 0:
            raise RuntimeError(f"Tool failed: {result.stderr}")
        
        return json.loads(result.stdout)
    
    def create(self, slides: int = 1, layout: str = "Title and Content"):
        """Create new presentation."""
        return self._run_tool('ppt_create_new.py', [
            '--output', self.filename,
            '--slides', str(slides),
            '--layout', layout
        ])
    
    def add_slide(self, layout: str, title: str = None):
        """Add slide to presentation."""
        args = ['--file', self.filename, '--layout', layout]
        if title:
            args.extend(['--title', title])
        return self._run_tool('ppt_add_slide.py', args)
    
    def add_chart(self, slide: int, chart_type: str, data: Dict,
                  position: Dict, size: Dict, title: str = None):
        """Add chart to slide."""
        # Write data to temp file
        import tempfile
        with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
            json.dump(data, f)
            data_file = f.name
        
        args = [
            '--file', self.filename,
            '--slide', str(slide),
            '--chart-type', chart_type,
            '--data', data_file,
            '--position', json.dumps(position),
            '--size', json.dumps(size)
        ]
        
        if title:
            args.extend(['--title', title])
        
        return self._run_tool('ppt_add_chart.py', args)

# Usage
agent = PowerPointAgent('output.pptx')
agent.create(slides=3)
agent.add_slide('Title and Content', 'Revenue Analysis')
agent.add_chart(
    slide=1,
    chart_type='column',
    data={
        'categories': ['Q1', 'Q2', 'Q3', 'Q4'],
        'series': [{'name': 'Revenue', 'values': [100, 120, 140, 160]}]
    },
    position={'left': '10%', 'top': '20%'},
    size={'width': '80%', 'height': '60%'},
    title='Quarterly Revenue'
)
```

---

## 🐛 Troubleshooting

### Common Issues

#### Issue 1: "python-pptx not found"

**Error:**
```
ModuleNotFoundError: No module named 'pptx'
```

**Solution:**
```bash
pip install python-pptx
# or
uv pip install python-pptx
```

---

#### Issue 2: "Invalid JSON in position argument"

**Error:**
```json
{
  "status": "error",
  "error": "Invalid JSON in position argument",
  "hint": "Use single quotes around JSON: '{\"left\":\"20%\"}'"
}
```

**Solution:**
Use single quotes around JSON, double quotes inside:
```bash
# Correct ✅
--position '{"left":"20%","top":"30%"}'

# Wrong ❌
--position {"left":"20%","top":"30%"}
--position "{"left":"20%","top":"30%"}"
```

---

#### Issue 3: "Slide index out of range"

**Error:**
```json
{
  "status": "error",
  "error": "Slide index 5 out of range (0-3)"
}
```

**Solution:**
- Slide indices are 0-based (first slide = 0)
- Check total slides first:
```bash
uv python tools/ppt_get_info.py --file presentation.pptx --json
```

---

#### Issue 4: "Image file not found"

**Error:**
```json
{
  "status": "error",
  "error": "Image file not found: logo.png"
}
```

**Solution:**
- Use absolute paths: `/full/path/to/image.png`
- Or relative from working directory: `./assets/logo.png`
- Verify file exists: `ls -la logo.png`

---

#### Issue 5: "Layout not found"

**Error:**
```json
{
  "status": "error",
  "error": "Layout 'Title Slide' not found"
}
```

**Solution:**
- Layout names are case-sensitive
- Get available layouts:
```bash
uv python tools/ppt_create_new.py --output test.pptx --slides 1 --json | jq '.available_layouts'
```

---

### Getting Help

1. **Check tool help:** `python tools/ppt_create_new.py --help`
2. **Review examples:** See [Examples](#-examples) section
3. **Read system prompt:** [`AGENT_SYSTEM_PROMPT.md`](AGENT_SYSTEM_PROMPT.md)
4. **Open issue:** [GitHub Issues](https://github.com/your-org/powerpoint-agent-tools/issues)
5. **Discussions:** [GitHub Discussions](https://github.com/your-org/powerpoint-agent-tools/discussions)

---

## 🏗️ Architecture

### Project Structure

```
powerpoint-agent-tools/
├── core/
│   ├── __init__.py
│   └── powerpoint_agent_core.py      # Core library (2200+ lines)
├── tools/                              # 10 CLI tools
│   ├── ppt_create_new.py
│   ├── ppt_add_slide.py
│   ├── ppt_set_title.py
│   ├── ppt_add_text_box.py
│   ├── ppt_insert_image.py
│   ├── ppt_add_bullet_list.py
│   ├── ppt_add_chart.py
│   ├── ppt_add_table.py
│   ├── ppt_add_shape.py
│   └── ppt_replace_text.py
├── AGENT_SYSTEM_PROMPT.md              # AI agent instructions
├── README.md                           # This file
├── requirements.txt                    # Dependencies
└── test_basic_tools.py                 # Integration tests
```

### Core Library Components

```
powerpoint_agent_core.py (2200 lines)
├── Exceptions (10 classes)
├── Constants & Enums (7 enums)
├── File Locking (FileLock)
├── Position & Size Helpers (2 classes)
├── Color Management (ColorHelper)
├── Template Preservation (TemplateProfile)
├── Accessibility Checker (AccessibilityChecker)
├── Asset Validator (AssetValidator)
└── PowerPointAgent (50+ methods)
    ├── File Operations (6 methods)
    ├── Slide Management (8 methods)
    ├── Text Operations (6 methods)
    ├── Shape Operations (5 methods)
    ├── Image Operations (5 methods)
    ├── Chart Operations (4 methods)
    ├── Layout & Theme (4 methods)
    ├── Validation (3 methods)
    ├── Export (3 methods)
    └── Utilities (5 methods)
```

### Design Patterns

**1. Stateless Tools**
- Each tool performs ONE operation
- No session state management
- Perfect for distributed systems

**2. JSON-First API**
- All outputs machine-parsable
- Structured error messages
- Easy integration with any language

**3. Composable Operations**
- Chain tools for complex workflows
- Unix philosophy: Do one thing well
- Pipe-friendly design

**4. Security-Hardened**
- Path validation
- File locking
- Input sanitization
- Resource limits

---

## 🔄 Comparison with Alternatives

| Feature | PowerPoint Agent Tools | python-pptx Direct | COM Automation | Google Slides API |
|---------|------------------------|-------------------|----------------|-------------------|
| **CLI-First** | ✅ Yes | ❌ Python only | ❌ Python only | ❌ API only |
| **AI-Friendly** | ✅ JSON outputs | ❌ Complex API | ❌ Windows only | ⚠️ Requires auth |
| **Stateless** | ✅ Yes | ⚠️ Requires coding | ❌ No | ✅ Yes |
| **Positioning** | ✅ 5 systems | ⚠️ Inches only | ⚠️ Inches only | ⚠️ Points |
| **Data Charts** | ✅ JSON input | ⚠️ Manual | ⚠️ Manual | ⚠️ Complex |
| **Templates** | ✅ Preserved | ⚠️ Limited | ✅ Full | ❌ No |
| **Accessibility** | ✅ Built-in | ⚠️ Manual | ⚠️ Manual | ✅ Good |
| **Platform** | ✅ Cross-platform | ✅ Cross-platform | ❌ Windows only | ✅ Cloud |
| **Learning Curve** | ✅ Low | ⚠️ Medium | ❌ High | ⚠️ Medium |

**Why Choose PowerPoint Agent Tools?**
- ✅ Designed specifically for AI agents
- ✅ Simple CLI interface
- ✅ Comprehensive examples and documentation
- ✅ Production-ready with error handling
- ✅ Active development and support

---

## 🗺️ Roadmap

### Version 1.1 (Q2 2024)
- [ ] Additional validation tools (`ppt_validate_presentation.py`, `ppt_check_accessibility.py`)
- [ ] Export tools (`ppt_export_pdf.py`, `ppt_export_images.py`)
- [ ] Animation support (`ppt_add_animation.py`)
- [ ] Master slide manipulation
- [ ] SmartArt diagram support

### Version 1.2 (Q3 2024)
- [ ] Real-time collaboration features
- [ ] Cloud storage integration (S3, GCS, Azure Blob)
- [ ] Batch processing tools
- [ ] Template marketplace
- [ ] Web-based preview tool

### Version 2.0 (Q4 2024)
- [ ] Natural language interface (GPT-4 integration)
- [ ] Visual diff tool for presentations
- [ ] Advanced theme management
- [ ] Video embedding support
- [ ] Presentation analytics

---

## 🤝 Contributing

We welcome contributions! Here's how to get started:

### Development Setup

```bash
# Fork and clone repository
git clone https://github.com/your-username/powerpoint-agent-tools.git
cd powerpoint-agent-tools

# Create virtual environment
python -m venv venv
source venv/bin/activate  # Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
pip install pytest pytest-cov black flake8

# Run tests
pytest test_basic_tools.py -v

# Check code style
black core/ tools/
flake8 core/ tools/
```

### Contribution Guidelines

1. **Fork** the repository
2. **Create** a feature branch (`git checkout -b feature/amazing-feature`)
3. **Write** tests for new functionality
4. **Ensure** all tests pass (`pytest -v`)
5. **Format** code (`black .`)
6. **Commit** with clear messages (`git commit -m 'Add amazing feature'`)
7. **Push** to branch (`git push origin feature/amazing-feature`)
8. **Open** a Pull Request

### Code Standards

- Follow PEP 8 style guide
- Include comprehensive docstrings
- Add type hints for all functions
- Write tests for new features
- Update documentation

### Areas for Contribution

- 🐛 Bug fixes
- ✨ New tools
- 📝 Documentation improvements
- 🧪 Additional test cases
- 🎨 Design patterns and templates
- 🌐 Internationalization
- ⚡ Performance optimizations

---

## 📄 License

MIT License

Copyright (c) 2024 PowerPoint Agent Tools Contributors

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

---

## 🙏 Acknowledgments

**Built with:**
- [python-pptx](https://python-pptx.readthedocs.io/) - Excellent PowerPoint library
- [Pillow](https://pillow.readthedocs.io/) - Image processing
- [pandas](https://pandas.pydata.org/) - Data manipulation

**Inspired by:**
- Financial modeling best practices
- Corporate presentation standards
- AI agent design patterns

**Special Thanks:**
- Open source community
- Early adopters and testers
- Contributors and maintainers

---

## 📞 Support & Contact

- **Documentation:** [Full Documentation](AGENT_SYSTEM_PROMPT.md)
- **Issues:** [GitHub Issues](https://github.com/your-org/powerpoint-agent-tools/issues)
- **Discussions:** [GitHub Discussions](https://github.com/your-org/powerpoint-agent-tools/discussions)
- **Email:** support@powerpoint-agent-tools.dev
- **Twitter:** [@ppt_agent_tools](https://twitter.com/ppt_agent_tools)

---

## ⭐ Star History

If you find this project useful, please consider starring it on GitHub!

[![Star History Chart](https://api.star-history.com/svg?repos=your-org/powerpoint-agent-tools&type=Date)](https://star-history.com/#your-org/powerpoint-agent-tools&Date)

---

<div align="center">

**Built with ❤️ for AI agents and presentation creators who demand automation**

[⬆ Back to Top](#-powerpoint-agent-tools)

</div>
