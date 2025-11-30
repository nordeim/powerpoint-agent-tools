awesome understanding!  Now, please think deeper and think harder to deeply and thoroughly explore systematically and meticulously how best to imagine a comprehensive system prompt for AI agent meant for PowerPoint presentation slides generation. 

Use below as your guide/reference to create a complete ideal/optimized **agent system prompt** that combines the **LLM-driven content analysis** with the **tool-driven presentation generation**, complete with **actual tool usage examples**. The actual tools listing follows after this sample agent system prompt (for reference)

---

### **AI Presentation Architect: Autonomous Presentation Generation**

#### **Identity**:

You are an **Autonomous AI Presentation Architect**, capable of autonomously generating a professional, engaging, and accessible PowerPoint presentation based on the given text content. Using the power of **LLMs** (e.g., GPT-4 or similar) for content analysis and strategic organization, combined with a set of specialized tools for slide creation, formatting, and validation, your goal is to produce high-quality presentations without direct interaction with the tool code. You handle everything from **initial content analysis** to **final presentation output**.

#### **Mission**:

Your mission is to **analyze the provided content**, structure it effectively into a presentation, and generate the resulting PowerPoint presentation autonomously. You ensure that the presentation is **visually appealing**, **accessible**, and **technically sound**, applying the necessary tools and validations along the way.

---

### **Workflow Phases**:

#### **Phase 1: Discover (Content Analysis & Design Exploration)**

**Goal**: Analyze the provided **text content** (user input) and determine how best to structure and present the data.

**Process**:

1. **Input Content**: The user provides raw text content or a brief overview.
2. **LLM Analysis**: The **LLM** (like GPT-4 or similar) processes the content and identifies key themes, topics, and essential information to present.
3. **Outline Generation**: The LLM generates an initial **outline** for the presentation, suggesting slide titles, types, and content. The LLM may also suggest the best visualizations (charts, tables, bullet points) for presenting the data.

**Example**:

* **Input**: User provides a document on **Q1 Sales Performance**.
* **LLM Output**:

  * **Slide 1**: **Title Slide** – "Q1 2024 Sales Performance"
  * **Slide 2**: **Bullet List** – Key insights (Revenue, Growth Rate, Market Share).
  * **Slide 3**: **Chart** – Q1 Revenue Growth (Bar Chart).
  * **Slide 4**: **Text Box** – Key takeaways and Q2 forecast.

---

#### **Phase 2: Plan (Design Strategy & Layout Definition)**

**Goal**: Define the **layout** and **visual structure** of the presentation.

**Tools Used**:

* **`ppt_create_from_structure.py`**: Create a PowerPoint presentation from a JSON-based structure.

  * Example Usage:

    ```bash
    uv run tools/ppt_create_from_structure.py --structure outline.json --output presentation.pptx --json
    ```
* **`ppt_set_slide_layout.py`**: Set the slide layout based on the content type (e.g., Title, Bullet List, Chart).

  * Example Usage:

    ```bash
    uv run tools/ppt_set_slide_layout.py --file presentation.pptx --slide 0 --layout "Title Slide" --json
    ```

**Process**:

1. **Slide Layouts**: Apply slide layouts as per the LLM-generated outline. For example, a **Title Slide** for the first slide, **Bullet Lists** for key insights, and **Charts** for data.
2. **Template Application**: Apply a **corporate template** to ensure branding consistency across all slides.

---

#### **Phase 3: Create (Slide Generation & Content Population)**

**Goal**: Populate the slides with the content defined in the outline. This includes text, images, charts, and other data visualizations.

**Tools Used**:

* **`ppt_add_slide.py`**: Add new slides to the presentation.

  * Example Usage:

    ```bash
    uv run tools/ppt_add_slide.py --file presentation.pptx --layout "Content" --index 1 --json
    ```
* **`ppt_add_text_box.py`**: Insert text boxes with key points or titles.

  * Example Usage:

    ```bash
    uv run tools/ppt_add_text_box.py --file presentation.pptx --slide 1 --text "Q1 Revenue: $5M" --position '{"left":"10%","top":"25%"}' --json
    ```
* **`ppt_add_bullet_list.py`**: Add bullet points for data or key insights.

  * Example Usage:

    ```bash
    uv run tools/ppt_add_bullet_list.py --file presentation.pptx --slide 2 --items "Revenue: $5M, Growth: 10%, Market Share: 15%" --position '{"left":"10%","top":"30%"}' --json
    ```
* **`ppt_add_chart.py`**: Add a chart to represent the data visually.

  * Example Usage:

    ```bash
    uv run tools/ppt_add_chart.py --file presentation.pptx --slide 3 --chart-type "bar" --data q1_revenue_data.json --position '{"left":"20%","top":"40%"}' --json
    ```
* **`ppt_insert_image.py`**: Insert images, such as logos or product images.

  * Example Usage:

    ```bash
    uv run tools/ppt_insert_image.py --file presentation.pptx --slide 4 --image "company_logo.png" --position '{"left":"15%","top":"10%"}' --size '{"width":"20%","height":"auto"}' --json
    ```

**Process**:

1. **Populate Slides**: Using the toolset, populate each slide with the appropriate content (text, images, charts, etc.) based on the design strategy.
2. **Content Formatting**: Apply necessary formatting and layout adjustments to ensure the content is clear, concise, and visually appealing.

---

#### **Phase 4: Validate (Final Quality Assurance & Accessibility Check)**

**Goal**: Ensure that the presentation is **structurally sound**, **accessible**, and adheres to presentation best practices.

**Tools Used**:

* **`ppt_validate_presentation.py`**: Validate the presentation's overall structure (e.g., missing slides, broken links).

  * Example Usage:

    ```bash
    uv run tools/ppt_validate_presentation.py --file presentation.pptx --json
    ```
* **`ppt_check_accessibility.py`**: Perform an accessibility check to ensure compliance with WCAG standards (e.g., alt text, color contrast).

  * Example Usage:

    ```bash
    uv run tools/ppt_check_accessibility.py --file presentation.pptx --json
    ```
* **`ppt_format_text.py`**: Ensure that all text is formatted according to accessibility standards, such as minimum font size and adequate contrast.

  * Example Usage:

    ```bash
    uv run tools/ppt_format_text.py --file presentation.pptx --slide 1 --shape 0 --font-name "Arial" --font-size 24 --json
    ```

**Process**:

1. **Structural Validation**: Ensure all slides are in place and content is intact. Check for any missing or malformed content.
2. **Accessibility Check**: Run accessibility checks to ensure that all images have alternative text, the font size is adequate, and the contrast ratio is sufficient for readability.
3. **Final Formatting**: Adjust any text or elements that do not meet accessibility or design standards.

---

#### **Phase 5: Deliver (Finalization & Output Generation)**

**Goal**: Finalize the presentation and export it in the required formats for delivery.

**Tools Used**:

* **`ppt_export_pdf.py`**: Export the presentation to PDF format for easy sharing and distribution.

  * Example Usage:

    ```bash
    uv run tools/ppt_export_pdf.py --file presentation.pptx --output presentation.pdf --json
    ```
* **`ppt_export_images.py`**: Export slides as images.

  * Example Usage:

    ```bash
    uv run tools/ppt_export_images.py --file presentation.pptx --output-dir images/ --format png --json
    ```

**Process**:

1. **Export to Desired Formats**: Output the presentation in PowerPoint format (.pptx) and optionally in PDF or image formats.
2. **Final Packaging**: Package the exported files for delivery, ensuring all formats are available and correctly labeled.

---

# Actual `.pptx` manipulation tools available:

+ head -25 tools/ppt_add_bullet_list.py
+ head -25 tools/ppt_add_chart.py
+ head -25 tools/ppt_add_connector.py
+ head -25 tools/ppt_add_notes.py
+ head -25 tools/ppt_add_shape.py
+ head -25 tools/ppt_add_slide.py
+ head -25 tools/ppt_add_table.py
+ head -25 tools/ppt_add_text_box.py
+ head -25 tools/ppt_capability_probe.py
+ head -25 tools/ppt_check_accessibility.py
+ head -25 tools/ppt_clone_presentation.py
+ head -25 tools/ppt_create_from_structure.py
+ head -25 tools/ppt_create_from_template.py
+ head -25 tools/ppt_create_new.py
+ head -25 tools/ppt_crop_image.py
+ head -25 tools/ppt_delete_slide.py
+ head -25 tools/ppt_duplicate_slide.py
+ head -25 tools/ppt_export_images.py
+ head -25 tools/ppt_export_pdf.py
+ head -25 tools/ppt_extract_notes.py
+ head -25 tools/ppt_format_chart.py
+ head -25 tools/ppt_format_shape.py
+ head -25 tools/ppt_format_table.py
+ head -25 tools/ppt_format_text.py
+ head -25 tools/ppt_get_info.py
+ head -25 tools/ppt_get_slide_info.py
+ head -25 tools/ppt_insert_image.py
+ head -25 tools/ppt_json_adapter.py
+ head -25 tools/ppt_merge_presentations.py
+ head -25 tools/ppt_remove_shape.py
+ head -25 tools/ppt_reorder_slides.py
+ head -25 tools/ppt_replace_image.py
+ head -25 tools/ppt_replace_text.py
+ head -25 tools/ppt_search_content.py
+ head -25 tools/ppt_set_background.py
+ head -25 tools/ppt_set_footer.py
+ head -25 tools/ppt_set_image_properties.py
+ head -25 tools/ppt_set_slide_layout.py
+ head -25 tools/ppt_set_title.py
+ head -25 tools/ppt_set_z_order.py
+ head -25 tools/ppt_update_chart_data.py
+ head -25 tools/ppt_validate_presentation.py

---

#!/usr/bin/env python3
"""
PowerPoint Add Bullet List Tool v3.1.0
Add bullet or numbered list with 6×6 rule validation and accessibility checks.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_add_bullet_list.py --file deck.pptx --slide 1 \\
        --items "Point 1,Point 2,Point 3" \\
        --position '{"left":"10%","top":"25%"}' \\
        --size '{"width":"80%","height":"60%"}' --json

Exit Codes:
    0: Success
    1: Error occurred

6×6 Rule (Best Practice):
    - Maximum 6 bullet points per slide
    - Maximum 6 words per line (~60 characters)
    - Ensures readability and audience engagement
"""

#!/usr/bin/env python3
"""
PowerPoint Add Chart Tool v3.1.0
Add data visualization chart to slide

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_add_chart.py --file presentation.pptx --slide 1 --chart-type column --data chart_data.json --position '{"left":"10%","top":"20%"}' --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

Supported Chart Types:
    column, column_stacked, bar, bar_stacked, line, line_markers,
    pie, area, scatter, doughnut
"""

import sys
import os

# --- HYGIENE BLOCK START ---
#!/usr/bin/env python3
"""
PowerPoint Add Connector Tool v3.1.0
Draw a line/connector between two shapes on a slide

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_add_connector.py --file deck.pptx --slide 0 --from-shape 0 --to-shape 1 --type straight --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

Use Cases:
    - Flowcharts and process diagrams
    - Org charts
    - Network diagrams
    - Relationship mapping
"""

import sys
import os
#!/usr/bin/env python3
"""
PowerPoint Add Speaker Notes Tool v3.1.0
Add, append, or overwrite speaker notes for a specific slide.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_add_notes.py --file deck.pptx --slide 0 --text "Key talking point" --json
    uv run tools/ppt_add_notes.py --file deck.pptx --slide 0 --text "New script" --mode overwrite --json
    uv run tools/ppt_add_notes.py --file deck.pptx --slide 0 --text "IMPORTANT:" --mode prepend --json

Exit Codes:
    0: Success
    1: Error occurred
"""

import sys
import os

sys.stderr = open(os.devnull, 'w')

import json
#!/usr/bin/env python3
"""
PowerPoint Add Shape Tool v3.1.0
Add shapes (rectangle, circle, arrow, etc.) to slides with comprehensive styling options.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_add_shape.py --file presentation.pptx --slide 0 \\
        --shape rectangle --position '{"left":"20%","top":"30%"}' \\
        --size '{"width":"60%","height":"40%"}' --fill-color "#0070C0" --json

    # Overlay with opacity
    uv run tools/ppt_add_shape.py --file presentation.pptx --slide 0 \\
        --shape rectangle --position '{"left":"0%","top":"0%"}' \\
        --size '{"width":"100%","height":"100%"}' \\
        --fill-color "#000000" --fill-opacity 0.15 --json

    # Quick overlay preset
    uv run tools/ppt_add_shape.py --file presentation.pptx --slide 0 \\
        --shape rectangle --overlay --fill-color "#FFFFFF" --json

Exit Codes:
#!/usr/bin/env python3
"""
PowerPoint Add Slide Tool v3.1.0
Add new slide to existing presentation with specific layout

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Compatible with PowerPoint Agent Core v3.1.0 (Dictionary Returns)

Usage:
    uv run tools/ppt_add_slide.py --file presentation.pptx --layout "Title and Content" --index 2 --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)
"""

import sys
import os

# --- HYGIENE BLOCK START ---
# CRITICAL: Redirect stderr to /dev/null to prevent library noise from corrupting JSON output
sys.stderr = open(os.devnull, 'w')
#!/usr/bin/env python3
"""
PowerPoint Add Table Tool v3.1.0
Add data table to slide with comprehensive validation.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_add_table.py --file presentation.pptx --slide 1 --rows 5 --cols 3 \\
        --data table_data.json --position '{"left":"10%","top":"25%"}' \\
        --size '{"width":"80%","height":"50%"}' --json

Exit Codes:
    0: Success
    1: Error occurred
"""

import sys
import os

sys.stderr = open(os.devnull, 'w')

import json
#!/usr/bin/env python3
"""
PowerPoint Add Text Box Tool v3.1.0
Add text box with flexible positioning, comprehensive validation, and accessibility checking.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_add_text_box.py --file deck.pptx --slide 0 \\
        --text "Revenue: $1.5M" --position '{"left":"20%","top":"30%"}' \\
        --size '{"width":"60%","height":"10%"}' --json

Exit Codes:
    0: Success
    1: Error occurred

Position Formats:
  1. Percentage: {"left": "20%", "top": "30%"}
  2. Inches: {"left": 2.0, "top": 3.0}
  3. Anchor: {"anchor": "center", "offset_x": 0, "offset_y": -1.0}
  4. Grid: {"grid_row": 2, "grid_col": 3, "grid_size": 12}
"""

#!/usr/bin/env python3
"""
PowerPoint Capability Probe Tool v3.1.0
Detect and report presentation template capabilities, layouts, and theme properties.

This tool provides comprehensive introspection of PowerPoint presentations to detect:
- Available layouts and their placeholders (with accurate runtime positions)
- Slide dimensions and aspect ratios
- Theme colors and fonts (using proper font scheme API)
- Template capabilities (footer support, slide numbers, dates)
- Multiple master slide support

Critical for AI agents and automation workflows to understand template capabilities
before generating content.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    # Basic probe (essential info)
    uv run tools/ppt_capability_probe.py --file template.pptx --json
    
    # Deep probe (accurate positions via transient instantiation)
    uv run tools/ppt_capability_probe.py --file template.pptx --deep --json
#!/usr/bin/env python3
"""
PowerPoint Check Accessibility Tool v3.1.0
Run WCAG 2.1 accessibility checks on presentation.

This tool performs comprehensive accessibility validation including:
- Alt text presence for images
- Color contrast ratios
- Reading order verification
- Font size compliance

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_check_accessibility.py --file presentation.pptx --json

Exit Codes:
    0: Success (check completed, see 'passed' field for result)
    1: Error occurred (file not found, crash)

Design Principles:
    - Read-only operation (acquire_lock=False)
    - JSON-first output with consistent contract
#!/usr/bin/env python3
"""
PowerPoint Clone Presentation Tool v3.1.0
Create an exact copy of a presentation for safe editing

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

⚠️ GOVERNANCE FOUNDATION - Clone-Before-Edit Principle

This tool implements the foundational safety principle: NEVER modify source
files directly. Always create a working copy first using this tool.

Usage:
    uv run tools/ppt_clone_presentation.py --source original.pptx --output work_copy.pptx --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

Safety Workflow:
    1. Clone: ppt_clone_presentation.py --source original.pptx --output work.pptx
    2. Edit: Use other tools on work.pptx
    3. Validate: ppt_validate_presentation.py --file work.pptx
#!/usr/bin/env python3
"""
PowerPoint Create From Structure Tool v3.1.0
Create a complete presentation from a JSON structure definition.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_create_from_structure.py --structure deck.json --output presentation.pptx --json

Exit Codes:
    0: Success
    1: Error occurred
"""

import sys
import os

sys.stderr = open(os.devnull, 'w')

import json
import argparse
from pathlib import Path
#!/usr/bin/env python3
"""
PowerPoint Create From Template Tool v3.1.1
Create new presentation from existing .pptx template.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.1

Usage:
    uv run tools/ppt_create_from_template.py --template corporate_template.pptx --output new_presentation.pptx --slides 10 --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

Changelog v3.1.1:
    - Added sys.stdout.flush() for pipeline safety
    - Added suggestion field to all error handlers
    - Added tool_version to all error responses
    - Added get_available_layouts() fallback for compatibility
"""

import sys
import os
#!/usr/bin/env python3
"""
PowerPoint Create New Tool v3.1.1
Create a new PowerPoint presentation with specified slides.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.1

Usage:
    uv run tools/ppt_create_new.py --output presentation.pptx --slides 5 --layout "Title and Content" --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

Note:
    For creating presentations from existing templates with branding,
    consider using ppt_create_from_template.py instead.

Changelog v3.1.1:
    - Added sys.stdout.flush() for pipeline safety
    - Added suggestion field to all error handlers
    - Added tool_version to all error responses
    - Added get_available_layouts() fallback for compatibility
#!/usr/bin/env python3
"""
PowerPoint Crop Image Tool v3.1.0
Crop an existing image on a slide by trimming edges

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_crop_image.py --file deck.pptx --slide 0 --shape 1 --left 0.1 --right 0.1 --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

Notes:
    Crop values are percentages of the original image size (0.0 to 1.0).
    For example, --left 0.1 trims 10% from the left edge.
"""

import sys
import os

# --- HYGIENE BLOCK START ---
#!/usr/bin/env python3
"""
PowerPoint Delete Slide Tool v3.1.1
Remove a slide from the presentation.

⚠️ DESTRUCTIVE OPERATION - Requires approval token with scope 'delete:slide'

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.1

Usage:
    uv run tools/ppt_delete_slide.py --file presentation.pptx --index 1 --approval-token "HMAC-SHA256:..." --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)
    4: Permission error (missing or invalid approval token)

Security:
    This tool performs a destructive operation and requires a valid approval
    token with scope 'delete:slide'. Generate tokens using the approval token
    system described in the governance documentation.

Changelog v3.1.1:
#!/usr/bin/env python3
"""
PowerPoint Duplicate Slide Tool v3.1.1
Clone an existing slide within the presentation.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.1

Usage:
    uv run tools/ppt_duplicate_slide.py --file presentation.pptx --index 0 --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

Changelog v3.1.1:
    - Added sys.stdout.flush() for pipeline safety
    - Added suggestion field to all error handlers
    - Added tool_version to all error responses
"""

import sys
import os

#!/usr/bin/env python3
"""
PowerPoint Export Images Tool v3.1.1
Export each slide as PNG or JPG image using LibreOffice.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.1

Usage:
    uv run tools/ppt_export_images.py --file presentation.pptx --output-dir output/ --format png --json

Exit Codes:
    0: Success
    1: Error occurred

Requirements:
    LibreOffice must be installed for image export:
    - Linux: sudo apt install libreoffice-impress
    - macOS: brew install --cask libreoffice
    - Windows: Download from https://www.libreoffice.org/

Changelog v3.1.1:
    - Added hygiene block for JSON pipeline safety
    - Added presentation_version tracking
#!/usr/bin/env python3
"""
PowerPoint Export PDF Tool v3.1.1
Export presentation to PDF format using LibreOffice.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.1

Usage:
    uv run tools/ppt_export_pdf.py --file presentation.pptx --output presentation.pdf --json

Exit Codes:
    0: Success
    1: Error occurred

Requirements:
    LibreOffice must be installed for PDF export:
    - Linux: sudo apt install libreoffice-impress
    - macOS: brew install --cask libreoffice
    - Windows: Download from https://www.libreoffice.org/

Changelog v3.1.1:
    - Added hygiene block for JSON pipeline safety
    - Added presentation_version tracking via PowerPointAgent
#!/usr/bin/env python3
"""
PowerPoint Extract Notes Tool v3.1.0
Extract speaker notes from all slides in a presentation.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_extract_notes.py --file presentation.pptx --json

Exit Codes:
    0: Success
    1: Error occurred
"""

import sys
import os

sys.stderr = open(os.devnull, 'w')

import json
import argparse
from pathlib import Path
#!/usr/bin/env python3
"""
PowerPoint Format Chart Tool v3.1.0
Format existing chart (title, legend position)

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_format_chart.py --file presentation.pptx --slide 1 --chart 0 --title "Revenue Growth" --legend bottom --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

Limitations:
    python-pptx has limited chart formatting support. This tool handles:
    - Chart title text
    - Legend position
    
    Not supported (requires PowerPoint):
    - Individual series colors
    - Axis formatting
    - Data labels
#!/usr/bin/env python3
"""
PowerPoint Format Shape Tool v3.1.0
Update styling of existing shapes including fill, line, opacity, and text formatting.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_format_shape.py --file presentation.pptx --slide 0 --shape 1 \\
        --fill-color "#FF0000" --fill-opacity 0.8 --json

Exit Codes:
    0: Success
    1: Error occurred

Note: The --transparency parameter is DEPRECATED. Use --fill-opacity instead.
      Opacity: 0.0 = invisible, 1.0 = opaque (opposite of transparency)
"""

import sys
import os

sys.stderr = open(os.devnull, 'w')
#!/usr/bin/env python3
"""
PowerPoint Format Table Tool v3.1.1
Style and format existing tables in presentations.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.1

Usage:
    uv run tools/ppt_format_table.py --file presentation.pptx --slide 0 --shape 2 --header-fill "#0070C0" --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

This tool formats existing tables by applying styling options including:
- Header row colors and formatting
- Data row colors with optional banding
- Font styling (name, size, color)
- Border styling (color, width)
- First column highlighting
"""

import sys
#!/usr/bin/env python3
"""
PowerPoint Format Text Tool v3.1.0
Format existing text with accessibility validation and contrast checking

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Features:
    - Font name, size, color, bold, italic formatting
    - WCAG 2.1 AA/AAA color contrast validation
    - Font size accessibility warnings (<12pt)
    - Before/after formatting comparison
    - Detailed validation results and recommendations

Usage:
    uv run tools/ppt_format_text.py --file deck.pptx --slide 0 --shape 0 --font-name "Arial" --font-size 24 --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

Accessibility:
    - Minimum font size: 12pt (14pt recommended for presentations)
#!/usr/bin/env python3
"""
PowerPoint Get Info Tool v3.1.0
Get presentation metadata (slide count, dimensions, file size, version)

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_get_info.py --file presentation.pptx --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)
"""

import sys
import os

# --- HYGIENE BLOCK START ---
# CRITICAL: Redirect stderr to /dev/null to prevent library noise from corrupting JSON output
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

#!/usr/bin/env python3
"""
PowerPoint Get Slide Info Tool v3.1.0
Get detailed information about slide content (shapes, images, text, positions)

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Features:
    - Full text content (no truncation)
    - Position information (inches and percentages)
    - Size information (inches and percentages)
    - Human-readable placeholder type names
    - Notes detection

Usage:
    uv run tools/ppt_get_slide_info.py --file presentation.pptx --slide 0 --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

Use Cases:
    - Finding shape indices for ppt_format_text.py
#!/usr/bin/env python3
"""
PowerPoint Insert Image Tool v3.1.0
Insert image into slide with automatic aspect ratio handling

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_insert_image.py --file presentation.pptx --slide 0 --image logo.png --position '{"left":"10%","top":"10%"}' --size '{"width":"20%","height":"auto"}' --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

Accessibility:
    Always use --alt-text to provide alternative text for screen readers.
    This is required for WCAG 2.1 compliance.
"""

import sys
import os

# --- HYGIENE BLOCK START ---
#!/usr/bin/env python3
"""
PowerPoint JSON Adapter Tool v3.1.1
Validates and normalizes JSON outputs from presentation CLI tools.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.1

Usage:
    uv run tools/ppt_json_adapter.py --schema ppt_get_info.schema.json --input raw.json

Behavior:
    - Validates input JSON against provided schema
    - Maps common alias keys to canonical keys
    - Emits normalized JSON to stdout
    - On validation failure, emits structured error JSON and exits non-zero

Exit Codes:
    0: Success (valid and normalized)
    2: Validation Error (schema validation failed)
    3: Input Load Error (could not read input file)
    5: Schema Load Error (could not read schema file)

Changelog v3.1.1:
#!/usr/bin/env python3
"""
PowerPoint Merge Presentations Tool v3.1.1
Combine slides from multiple presentations into one.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.1

Usage:
    uv run tools/ppt_merge_presentations.py --sources '[{"file":"a.pptx","slides":"all"},{"file":"b.pptx","slides":[0,2,4]}]' --output merged.pptx --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

This tool merges slides from multiple source presentations into a single output
presentation. You can specify which slides to include from each source.

Source Specification Format:
    [
        {"file": "path/to/first.pptx", "slides": "all"},
        {"file": "path/to/second.pptx", "slides": [0, 1, 2]},
        {"file": "path/to/third.pptx", "slides": [5, 6]}
    ]
#!/usr/bin/env python3
"""
PowerPoint Remove Shape Tool v3.1.0
Safely remove shapes from slides with comprehensive safety controls.

⚠️  DESTRUCTIVE OPERATION WARNING ⚠️
- Shape removal CANNOT be undone
- Shape indices WILL shift after removal
- Always CLONE the presentation first
- Always use --dry-run to preview

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    # Preview removal (RECOMMENDED FIRST)
    uv run tools/ppt_remove_shape.py --file deck.pptx --slide 0 --shape 2 --dry-run --json
    
    # Execute removal
    uv run tools/ppt_remove_shape.py --file deck.pptx --slide 0 --shape 2 --json

Exit Codes:
    0: Success
    1: Error occurred
#!/usr/bin/env python3
"""
PowerPoint Reorder Slides Tool v3.1.0
Move a slide from one position to another

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_reorder_slides.py --file presentation.pptx --from-index 3 --to-index 1 --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

Notes:
    - Indices are 0-based
    - Moving a slide shifts other slides accordingly
    - Original content is preserved during move
"""

import sys
import os

#!/usr/bin/env python3
"""
PowerPoint Replace Image Tool v3.1.0
Replace an existing image with a new one (preserves position and size)

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_replace_image.py --file presentation.pptx --slide 0 --old-image "logo" --new-image new_logo.png --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

Use Cases:
    - Logo updates during rebranding
    - Product photo updates
    - Chart/diagram refreshes
    - Team photo updates
"""

import sys
import os
#!/usr/bin/env python3
"""
PowerPoint Replace Text Tool v3.1.0
Find and replace text across presentation or in specific targets

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Features:
    - Global replacement (entire presentation)
    - Targeted replacement (specific slide)
    - Surgical replacement (specific shape)
    - Dry-run mode (preview without changes)
    - Case-sensitive matching option
    - Formatting-preserving replacement (run-level)
    - Location reporting

Usage:
    uv run tools/ppt_replace_text.py --file deck.pptx --find "Old" --replace "New" --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

#!/usr/bin/env python3
"""
PowerPoint Search Content Tool v3.1.1
Search for text content across all slides in a presentation.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.1

Usage:
    uv run tools/ppt_search_content.py --file presentation.pptx --query "Revenue" --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

This tool searches for text content across slides, including:
- Text in shapes and text boxes
- Slide titles and subtitles
- Speaker notes
- Table cell contents

Use this tool to locate content before using ppt_replace_text.py or to
navigate large presentations efficiently.
"""
#!/usr/bin/env python3
"""
PowerPoint Set Background Tool v3.1.0
Set slide background to a solid color or image.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_set_background.py --file deck.pptx --slide 0 --color "#FFFFFF" --json
    uv run tools/ppt_set_background.py --file deck.pptx --all-slides --color "#F5F5F5" --json
    uv run tools/ppt_set_background.py --file deck.pptx --slide 0 --image background.jpg --json

Exit Codes:
    0: Success
    1: Error occurred
"""

import sys
import os

sys.stderr = open(os.devnull, 'w')

import json
#!/usr/bin/env python3
"""
PowerPoint Set Footer Tool v3.1.0
Configure slide footer with Dual Strategy (Placeholder + Text Box Fallback).

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_set_footer.py --file deck.pptx --text "Company © 2024" --json
    uv run tools/ppt_set_footer.py --file deck.pptx --text "Confidential" --show-number --json

Exit Codes:
    0: Success
    1: Error occurred
"""

import sys
import os

sys.stderr = open(os.devnull, 'w')

import json
import argparse
#!/usr/bin/env python3
"""
PowerPoint Set Image Properties Tool v3.1.0
Set alt text and opacity for image shapes (accessibility support)

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_set_image_properties.py --file deck.pptx --slide 0 --shape 1 --alt-text "Company Logo" --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

Accessibility:
    Alt text is required for WCAG 2.1 compliance. All images should have
    descriptive alternative text that conveys the image's content and purpose.
"""

import sys
import os

# --- HYGIENE BLOCK START ---
#!/usr/bin/env python3
"""
PowerPoint Set Slide Layout Tool v3.1.0
Change the layout of an existing slide with safety warnings

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

⚠️ IMPORTANT WARNING:
    Changing slide layouts can cause CONTENT LOSS!
    - Text in removed placeholders may disappear
    - Shapes may be repositioned
    - This is a python-pptx limitation
    
    ALWAYS backup your presentation before changing layouts!

Usage:
    uv run tools/ppt_set_slide_layout.py --file presentation.pptx --slide 2 --layout "Title Only" --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

Safety:
#!/usr/bin/env python3
"""
PowerPoint Set Title Tool v3.1.0
Set slide title and optional subtitle with comprehensive validation.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_set_title.py --file presentation.pptx --slide 0 --title "Q4 Results" --json
    uv run tools/ppt_set_title.py --file deck.pptx --slide 0 --title "2024 Strategy" \\
        --subtitle "Growth & Innovation" --json

Exit Codes:
    0: Success
    1: Error occurred

Best Practices:
- Keep titles under 60 characters for readability
- Keep subtitles under 100 characters
- Use "Title Slide" layout for first slide (index 0)
- Use title case: "This Is Title Case"
- Subtitles provide context, not repetition
"""
#!/usr/bin/env python3
"""
PowerPoint Set Z-Order Tool v3.1.0
Manage shape layering (Bring to Front, Send to Back, etc.).

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_set_z_order.py --file deck.pptx --slide 0 --shape 1 --action bring_to_front --json

Exit Codes:
    0: Success
    1: Error occurred

⚠️  IMPORTANT: Shape indices change after z-order operations!
    Always refresh indices with ppt_get_slide_info.py before targeting shapes.
"""

import sys
import os

sys.stderr = open(os.devnull, 'w')

#!/usr/bin/env python3
"""
PowerPoint Update Chart Data Tool v3.1.0
Update the data of an existing chart

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_update_chart_data.py --file deck.pptx --slide 0 --chart 0 --data new_data.json --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

⚠️ LIMITATION WARNING:
    python-pptx has LIMITED chart update support. The replace_data() method
    may fail if the new data schema doesn't match the original chart exactly.
    
    If update fails, consider the alternative approach:
    1. Delete the existing chart: ppt_remove_shape.py
    2. Add a new chart with new data: ppt_add_chart.py
"""

#!/usr/bin/env python3
"""
PowerPoint Validate Presentation Tool v3.1.1
Comprehensive validation for structure, accessibility, assets, and design quality.

Fully aligned with PowerPoint Agent Core v3.1.0+ and System Prompt v3.0 validation gates.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.1

Usage:
    uv run tools/ppt_validate_presentation.py --file presentation.pptx --json
    uv run tools/ppt_validate_presentation.py --file presentation.pptx --policy strict --json

Exit Codes:
    0: Success (valid or only warnings within policy thresholds)
    1: Error occurred or critical issues exceed policy thresholds

Changelog v3.1.1:
    - Added presentation_version to output for audit trail
    - Populated fix_command for actionable remediation
    - Expanded _validate_design_rules with color and 6x6 rule checking
    - Added tool_version to output
    - Added acquire_lock documentation comments
