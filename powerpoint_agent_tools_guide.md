## **Powerpoint Agent Tools**

### **Files Created (8 total):**

| # | File | Lines | Purpose |
|---|------|-------|---------|
| 1 | `requirements.txt` | 15 | Dependencies |
| 2 | `core/__init__.py` | 60 | Package exports |
| 3 | `tools/ppt_create_new.py` | 240 | Create new presentation |
| 4 | `tools/ppt_add_slide.py` | 205 | Add slide to presentation |
| 5 | `tools/ppt_set_title.py` | 185 | Set slide title/subtitle |
| 6 | `tools/ppt_add_text_box.py` | 310 | Add text box with positioning |
| 7 | `tools/ppt_insert_image.py` | 340 | Insert image into slide |
| 8 | `test_basic_tools.py` | 280 | Integration tests |

---

## 🎯 **Quick Start Guide**

### **Installation:**

```bash
# Install dependencies
pip install python-pptx Pillow pandas

# Or with uv
uv pip install python-pptx Pillow pandas
```

### **Example Workflow:**

```bash
# 1. Create new presentation
uv python tools/ppt_create_new.py \
  --output presentation.pptx \
  --slides 3 \
  --json

# 2. Set title on first slide
uv python tools/ppt_set_title.py \
  --file presentation.pptx \
  --slide 0 \
  --title "Q4 Results" \
  --subtitle "Financial Review 2024" \
  --json

# 3. Add content slide
uv python tools/ppt_add_slide.py \
  --file presentation.pptx \
  --layout "Title and Content" \
  --title "Revenue Growth" \
  --json

# 4. Add text box
uv python tools/ppt_add_text_box.py \
  --file presentation.pptx \
  --slide 1 \
  --text "Revenue increased 45% year-over-year" \
  --position '{"left":"10%","top":"30%"}' \
  --size '{"width":"80%","height":"15%"}' \
  --font-size 24 \
  --bold \
  --json

# 5. Insert image (logo)
uv python tools/ppt_insert_image.py \
  --file presentation.pptx \
  --slide 0 \
  --image company_logo.png \
  --position '{"anchor":"top_right","offset_x":-0.5,"offset_y":0.5}' \
  --size '{"width":"15%","height":"auto"}' \
  --alt-text "Company Logo" \
  --json
```

---

# **Complete Workflow Example**

```
# Step 1: Create presentation
uv python tools/ppt_create_new.py \
  --output quarterly_results.pptx \
  --slides 1 \
  --layout "Title Slide" \
  --json

# Step 2: Set title
uv python tools/ppt_set_title.py \
  --file quarterly_results.pptx \
  --slide 0 \
  --title "Q4 2024 Results" \
  --subtitle "Record-Breaking Performance" \
  --json

# Step 3: Add content slide with bullet points
uv python tools/ppt_add_slide.py \
  --file quarterly_results.pptx \
  --layout "Title and Content" \
  --title "Key Highlights" \
  --json

uv python tools/ppt_add_bullet_list.py \
  --file quarterly_results.pptx \
  --slide 1 \
  --items "Revenue up 45% YoY,Customer base grew 60%,Market share reached 23%,Profitability improved 12pts" \
  --position '{"left":"10%","top":"25%"}' \
  --size '{"width":"80%","height":"60%"}' \
  --font-size 22 \
  --json

# Step 4: Add revenue chart
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
  --file quarterly_results.pptx \
  --layout "Title and Content" \
  --title "Revenue Growth" \
  --json

uv python tools/ppt_add_chart.py \
  --file quarterly_results.pptx \
  --slide 2 \
  --chart-type column \
  --data revenue_data.json \
  --position '{"left":"10%","top":"20%"}' \
  --size '{"width":"80%","height":"65%"}' \
  --title "Quarterly Revenue ($M)" \
  --json

# Step 5: Add data table
cat > metrics.json << 'EOF'
[
  ["Revenue", "22.1", "14.3", "54%"],
  ["Customers", "45,000", "28,000", "60%"],
  ["Employees", "250", "180", "39%"]
]
EOF

uv python tools/ppt_add_slide.py \
  --file quarterly_results.pptx \
  --layout "Title and Content" \
  --title "Key Metrics Summary" \
  --json

uv python tools/ppt_add_table.py \
  --file quarterly_results.pptx \
  --slide 3 \
  --rows 4 \
  --cols 4 \
  --headers "Metric,Q4 2024,Q4 2023,Growth" \
  --data metrics.json \
  --position '{"left":"15%","top":"25%"}' \
  --size '{"width":"70%","height":"50%"}' \
  --json

# Step 6: Add visual emphasis (shape)
uv python tools/ppt_add_shape.py \
  --file quarterly_results.pptx \
  --slide 1 \
  --shape rounded_rectangle \
  --position '{"left":"75%","top":"15%"}' \
  --size '{"width":"20%","height":"12%"}' \
  --fill-color "#00B050" \
  --json

uv python tools/ppt_add_text_box.py \
  --file quarterly_results.pptx \
  --slide 1 \
  --text "Record Quarter!" \
  --position '{"left":"77%","top":"17%"}' \
  --size '{"width":"16%","height":"8%"}' \
  --font-size 20 \
  --bold \
  --color "#FFFFFF" \
  --alignment center \
  --json

# Step 7: Bulk update (rebranding)
uv python tools/ppt_replace_text.py \
  --file quarterly_results.pptx \
  --find "Company Inc." \
  --replace "Company LLC" \
  --dry-run \
  --json

# If dry-run looks good, apply changes
uv python tools/ppt_replace_text.py \
  --file quarterly_results.pptx \
  --find "Company Inc." \
  --replace "Company LLC" \
  --json

# Result: Complete presentation with:
# - Title slide
# - Bullet points
# - Chart
# - Table
# - Visual emphasis
# - Updated branding
```

