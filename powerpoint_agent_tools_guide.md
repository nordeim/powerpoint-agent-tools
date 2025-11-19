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

