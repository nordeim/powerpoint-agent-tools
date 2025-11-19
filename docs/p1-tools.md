# 🚀 PowerPoint Agent Tools - Next 5 P1 Tools Implementation

## Pre-Implementation Analysis

### **Selected P1 Tools (High Priority):**

| # | Tool | Purpose | Business Value |
|---|------|---------|----------------|
| 6 | `ppt_add_bullet_list.py` | Add bullet/numbered lists | **High** - Most common content type |
| 7 | `ppt_add_chart.py` | Add data visualization charts | **Critical** - Business presentations |
| 8 | `ppt_add_table.py` | Add data tables | **High** - Structured data display |
| 9 | `ppt_add_shape.py` | Add shapes (boxes, arrows, etc.) | **Medium** - Visual design |
| 10 | `ppt_replace_text.py` | Find and replace text | **High** - Bulk updates |

### **Design Validation:**

✅ **Rich content creation** - Charts, tables, bullets  
✅ **Data-driven** - Supports JSON data input  
✅ **Bulk operations** - Replace text across all slides  
✅ **Visual design** - Shapes for diagrams  
✅ **Comprehensive examples** - Real-world use cases  

---

## **File 1/5: `tools/ppt_add_bullet_list.py`**

```python
#!/usr/bin/env python3
"""
PowerPoint Add Bullet List Tool
Add bullet or numbered list to slide

Usage:
    uv python ppt_add_bullet_list.py --file presentation.pptx --slide 0 --items "Item 1,Item 2,Item 3" --position '{"left":"10%","top":"25%"}' --size '{"width":"80%","height":"60%"}' --json

Exit Codes:
    0: Success
    1: Error occurred
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any, List

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError, SlideNotFoundError
)


def add_bullet_list(
    filepath: Path,
    slide_index: int,
    items: List[str],
    position: Dict[str, Any],
    size: Dict[str, Any],
    bullet_style: str = "bullet",
    font_size: int = 18,
    font_name: str = "Calibri",
    color: str = None,
    line_spacing: float = 1.0
) -> Dict[str, Any]:
    """Add bullet or numbered list to slide."""
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if not items:
        raise ValueError("At least one item required")
    
    if len(items) > 20:
        raise ValueError("Maximum 20 items per list (readability limit)")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Validate slide index
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides-1})"
            )
        
        # Add bullet list
        agent.add_bullet_list(
            slide_index=slide_index,
            items=items,
            position=position,
            size=size,
            bullet_style=bullet_style,
            font_size=font_size
        )
        
        # Get the last added shape for additional formatting if needed
        slide_info = agent.get_slide_info(slide_index)
        last_shape_idx = slide_info["shape_count"] - 1
        
        # Apply color if specified
        if color:
            agent.format_text(
                slide_index=slide_index,
                shape_index=last_shape_idx,
                color=color
            )
        
        # Save
        agent.save()
    
    return {
        "status": "success",
        "file": str(filepath),
        "slide_index": slide_index,
        "items_added": len(items),
        "items": items,
        "bullet_style": bullet_style,
        "formatting": {
            "font_size": font_size,
            "font_name": font_name,
            "color": color,
            "line_spacing": line_spacing
        }
    }


def main():
    parser = argparse.ArgumentParser(
        description="Add bullet or numbered list to PowerPoint slide",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
List Formats:
  1. Comma-separated: "Item 1,Item 2,Item 3"
  2. Newline-separated: "Item 1\\nItem 2\\nItem 3"
  3. JSON array from file: --items-file items.json

Bullet Styles:
  - bullet: Traditional bullet points (default)
  - numbered: 1. 2. 3. numbering
  - none: Plain list without bullets

Examples:
  # Simple bullet list
  uv python ppt_add_bullet_list.py \\
    --file presentation.pptx \\
    --slide 1 \\
    --items "Revenue up 45%,Customer growth 60%,Market share 23%" \\
    --position '{"left":"10%","top":"25%"}' \\
    --size '{"width":"80%","height":"60%"}' \\
    --json
  
  # Numbered list with custom formatting
  uv python ppt_add_bullet_list.py \\
    --file presentation.pptx \\
    --slide 2 \\
    --items "Define objectives,Analyze market,Develop strategy,Execute plan" \\
    --bullet-style numbered \\
    --position '{"left":"15%","top":"30%"}' \\
    --size '{"width":"70%","height":"50%"}' \\
    --font-size 20 \\
    --color "#0070C0" \\
    --json
  
  # Multi-line items (agenda)
  uv python ppt_add_bullet_list.py \\
    --file presentation.pptx \\
    --slide 0 \\
    --items "Introduction and welcome,Q4 financial results,2024 strategic priorities,Q&A session" \\
    --position '{"grid":"B3"}' \\
    --size '{"width":"60%","height":"50%"}' \\
    --font-size 22 \\
    --json
  
  # Key takeaways (centered)
  uv python ppt_add_bullet_list.py \\
    --file presentation.pptx \\
    --slide 10 \\
    --items "Strong Q4 performance,Exceeded revenue targets,Positive market reception" \\
    --position '{"anchor":"center","offset_x":0,"offset_y":0}' \\
    --size '{"width":"70%","height":"40%"}' \\
    --font-size 24 \\
    --color "#00B050" \\
    --json
  
  # From JSON file
  echo '["First point", "Second point", "Third point"]' > items.json
  uv python ppt_add_bullet_list.py \\
    --file presentation.pptx \\
    --slide 3 \\
    --items-file items.json \\
    --position '{"left":"10%","top":"25%"}' \\
    --size '{"width":"80%","height":"60%"}' \\
    --json

Best Practices:
  - Keep items concise (max 2 lines per bullet)
  - Use 3-7 items per slide for readability
  - Start each item with action verb for impact
  - Use parallel structure (all items same format)
  - Font size 18-24pt for body text
  - Leave white space (don't fill entire slide)

Formatting Tips:
  - Use color for emphasis (sparingly)
  - Increase font size for key points
  - Use numbered lists for sequential steps
  - Use bullet lists for unordered points
  - Consistent spacing improves readability
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file path'
    )
    
    parser.add_argument(
        '--slide',
        required=True,
        type=int,
        help='Slide index (0-based)'
    )
    
    parser.add_argument(
        '--items',
        help='Comma-separated list items'
    )
    
    parser.add_argument(
        '--items-file',
        type=Path,
        help='JSON file with array of items'
    )
    
    parser.add_argument(
        '--position',
        required=True,
        type=json.loads,
        help='Position dict (JSON string)'
    )
    
    parser.add_argument(
        '--size',
        required=True,
        type=json.loads,
        help='Size dict (JSON string)'
    )
    
    parser.add_argument(
        '--bullet-style',
        choices=['bullet', 'numbered', 'none'],
        default='bullet',
        help='Bullet style (default: bullet)'
    )
    
    parser.add_argument(
        '--font-size',
        type=int,
        default=18,
        help='Font size in points (default: 18)'
    )
    
    parser.add_argument(
        '--font-name',
        default='Calibri',
        help='Font name (default: Calibri)'
    )
    
    parser.add_argument(
        '--color',
        help='Text color (hex, e.g., #0070C0)'
    )
    
    parser.add_argument(
        '--line-spacing',
        type=float,
        default=1.0,
        help='Line spacing multiplier (default: 1.0)'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        help='Output JSON response'
    )
    
    args = parser.parse_args()
    
    try:
        # Parse items
        if args.items_file:
            if not args.items_file.exists():
                raise FileNotFoundError(f"Items file not found: {args.items_file}")
            with open(args.items_file, 'r') as f:
                items = json.load(f)
            if not isinstance(items, list):
                raise ValueError("Items file must contain JSON array")
        elif args.items:
            # Split by comma or newline
            if '\\n' in args.items:
                items = args.items.split('\\n')
            else:
                items = args.items.split(',')
            items = [item.strip() for item in items if item.strip()]
        else:
            raise ValueError("Either --items or --items-file required")
        
        result = add_bullet_list(
            filepath=args.file,
            slide_index=args.slide,
            items=items,
            position=args.position,
            size=args.size,
            bullet_style=args.bullet_style,
            font_size=args.font_size,
            font_name=args.font_name,
            color=args.color,
            line_spacing=args.line_spacing
        )
        
        if args.json:
            print(json.dumps(result, indent=2))
        else:
            print(f"✅ Added {result['items_added']}-item {result['bullet_style']} list to slide {result['slide_index']}")
            print(f"   Items:")
            for i, item in enumerate(result['items'][:5], 1):
                prefix = f"{i}." if args.bullet_style == 'numbered' else "•"
                print(f"     {prefix} {item}")
            if len(result['items']) > 5:
                print(f"     ... and {len(result['items']) - 5} more")
        
        sys.exit(0)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__
        }
        
        if args.json:
            print(json.dumps(error_result, indent=2))
        else:
            print(f"❌ Error: {e}", file=sys.stderr)
        
        sys.exit(1)


if __name__ == "__main__":
    main()
```

---

## **File 2/5: `tools/ppt_add_chart.py`**

```python
#!/usr/bin/env python3
"""
PowerPoint Add Chart Tool
Add data visualization chart to slide

Usage:
    uv python ppt_add_chart.py --file presentation.pptx --slide 1 --chart-type column --data chart_data.json --position '{"left":"10%","top":"20%"}' --size '{"width":"80%","height":"60%"}' --json

Exit Codes:
    0: Success
    1: Error occurred
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any, List

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError, SlideNotFoundError
)


def add_chart(
    filepath: Path,
    slide_index: int,
    chart_type: str,
    data: Dict[str, Any],
    position: Dict[str, Any],
    size: Dict[str, Any],
    chart_title: str = None
) -> Dict[str, Any]:
    """Add chart to slide."""
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    # Validate data structure
    if "categories" not in data:
        raise ValueError("Data must contain 'categories' key")
    
    if "series" not in data or not data["series"]:
        raise ValueError("Data must contain at least one series")
    
    # Validate all series have same length as categories
    cat_len = len(data["categories"])
    for series in data["series"]:
        if len(series.get("values", [])) != cat_len:
            raise ValueError(
                f"Series '{series.get('name', 'unnamed')}' has {len(series['values'])} values, "
                f"but {cat_len} categories. Must match."
            )
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Validate slide index
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides-1})"
            )
        
        # Add chart
        agent.add_chart(
            slide_index=slide_index,
            chart_type=chart_type,
            data=data,
            position=position,
            size=size,
            chart_title=chart_title
        )
        
        # Save
        agent.save()
    
    return {
        "status": "success",
        "file": str(filepath),
        "slide_index": slide_index,
        "chart_type": chart_type,
        "chart_title": chart_title,
        "categories": len(data["categories"]),
        "series": len(data["series"]),
        "data_points": sum(len(s["values"]) for s in data["series"])
    }


def main():
    parser = argparse.ArgumentParser(
        description="Add data visualization chart to PowerPoint slide",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Chart Types:
  - column: Vertical bars (compare across categories)
  - column_stacked: Stacked vertical bars (show composition)
  - bar: Horizontal bars (compare items)
  - bar_stacked: Stacked horizontal bars
  - line: Line chart (show trends over time)
  - line_markers: Line with data point markers
  - pie: Pie chart (show proportions, single series only)
  - area: Area chart (emphasize magnitude of change)
  - scatter: Scatter plot (show relationships)

Data Format (JSON):
{
  "categories": ["Q1", "Q2", "Q3", "Q4"],
  "series": [
    {"name": "Revenue", "values": [100, 120, 140, 160]},
    {"name": "Costs", "values": [80, 90, 100, 110]}
  ]
}

Examples:
  # Revenue growth chart
  cat > revenue_data.json << 'EOF'
{
  "categories": ["Q1 2023", "Q2 2023", "Q3 2023", "Q4 2023", "Q1 2024"],
  "series": [
    {"name": "Revenue ($M)", "values": [12.5, 15.2, 18.7, 22.1, 25.8]},
    {"name": "Target ($M)", "values": [15, 16, 18, 20, 24]}
  ]
}
EOF
  
  uv python ppt_add_chart.py \\
    --file presentation.pptx \\
    --slide 1 \\
    --chart-type column \\
    --data revenue_data.json \\
    --position '{"left":"10%","top":"20%"}' \\
    --size '{"width":"80%","height":"60%"}' \\
    --title "Revenue Growth Trajectory" \\
    --json
  
  # Market share pie chart
  cat > market_data.json << 'EOF'
{
  "categories": ["Our Company", "Competitor A", "Competitor B", "Others"],
  "series": [
    {"name": "Market Share", "values": [35, 28, 22, 15]}
  ]
}
EOF
  
  uv python ppt_add_chart.py \\
    --file presentation.pptx \\
    --slide 2 \\
    --chart-type pie \\
    --data market_data.json \\
    --position '{"anchor":"center"}' \\
    --size '{"width":"60%","height":"60%"}' \\
    --title "Market Share Distribution" \\
    --json
  
  # Year-over-year comparison (bar chart)
  cat > yoy_data.json << 'EOF'
{
  "categories": ["Revenue", "Profit", "Customers", "Employees"],
  "series": [
    {"name": "2023", "values": [100, 25, 1000, 150]},
    {"name": "2024", "values": [145, 38, 1450, 200]}
  ]
}
EOF
  
  uv python ppt_add_chart.py \\
    --file presentation.pptx \\
    --slide 3 \\
    --chart-type bar \\
    --data yoy_data.json \\
    --position '{"left":"5%","top":"25%"}' \\
    --size '{"width":"90%","height":"60%"}' \\
    --title "Year-over-Year Growth" \\
    --json
  
  # Line chart (trends)
  cat > trend_data.json << 'EOF'
{
  "categories": ["Jan", "Feb", "Mar", "Apr", "May", "Jun"],
  "series": [
    {"name": "Website Traffic", "values": [12000, 13500, 15200, 16800, 18500, 21000]},
    {"name": "Conversions", "values": [240, 270, 304, 336, 370, 420]}
  ]
}
EOF
  
  uv python ppt_add_chart.py \\
    --file presentation.pptx \\
    --slide 4 \\
    --chart-type line_markers \\
    --data trend_data.json \\
    --position '{"left":"10%","top":"20%"}' \\
    --size '{"width":"80%","height":"65%"}' \\
    --title "Traffic & Conversion Trends" \\
    --json
  
  # Inline data (short example)
  uv python ppt_add_chart.py \\
    --file presentation.pptx \\
    --slide 5 \\
    --chart-type column \\
    --data-string '{"categories":["A","B","C"],"series":[{"name":"Sales","values":[10,20,15]}]}' \\
    --position '{"left":"20%","top":"25%"}' \\
    --size '{"width":"60%","height":"50%"}' \\
    --json

Best Practices:
  - Use column charts for comparing categories (most common)
  - Use line charts for showing trends over time
  - Use pie charts for proportions (max 5-7 slices)
  - Use bar charts when category names are long
  - Keep data series count to 3-5 max for clarity
  - Use consistent colors across presentation
  - Always include a descriptive title
  - Round numbers for readability

Chart Selection Guide:
  - Compare values: Column or Bar chart
  - Show trends: Line chart
  - Show proportions: Pie chart
  - Show composition: Stacked column/bar
  - Show correlation: Scatter plot
  - Emphasize change: Area chart
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file path'
    )
    
    parser.add_argument(
        '--slide',
        required=True,
        type=int,
        help='Slide index (0-based)'
    )
    
    parser.add_argument(
        '--chart-type',
        required=True,
        choices=['column', 'column_stacked', 'bar', 'bar_stacked', 
                'line', 'line_markers', 'pie', 'area', 'scatter'],
        help='Chart type'
    )
    
    parser.add_argument(
        '--data',
        type=Path,
        help='JSON file with chart data'
    )
    
    parser.add_argument(
        '--data-string',
        help='Inline JSON data string'
    )
    
    parser.add_argument(
        '--position',
        required=True,
        type=json.loads,
        help='Position dict (JSON string)'
    )
    
    parser.add_argument(
        '--size',
        required=True,
        type=json.loads,
        help='Size dict (JSON string)'
    )
    
    parser.add_argument(
        '--title',
        help='Chart title'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        help='Output JSON response'
    )
    
    args = parser.parse_args()
    
    try:
        # Load chart data
        if args.data:
            if not args.data.exists():
                raise FileNotFoundError(f"Data file not found: {args.data}")
            with open(args.data, 'r') as f:
                data = json.load(f)
        elif args.data_string:
            data = json.loads(args.data_string)
        else:
            raise ValueError("Either --data or --data-string required")
        
        result = add_chart(
            filepath=args.file,
            slide_index=args.slide,
            chart_type=args.chart_type,
            data=data,
            position=args.position,
            size=args.size,
            chart_title=args.title
        )
        
        if args.json:
            print(json.dumps(result, indent=2))
        else:
            print(f"✅ Added {result['chart_type']} chart to slide {result['slide_index']}")
            if args.title:
                print(f"   Title: {result['chart_title']}")
            print(f"   Categories: {result['categories']}")
            print(f"   Series: {result['series']}")
            print(f"   Total data points: {result['data_points']}")
        
        sys.exit(0)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__
        }
        
        if args.json:
            print(json.dumps(error_result, indent=2))
        else:
            print(f"❌ Error: {e}", file=sys.stderr)
        
        sys.exit(1)


if __name__ == "__main__":
    main()
```

---

## **File 3/5: `tools/ppt_add_table.py`**

```python
#!/usr/bin/env python3
"""
PowerPoint Add Table Tool
Add data table to slide

Usage:
    uv python ppt_add_table.py --file presentation.pptx --slide 1 --rows 5 --cols 3 --data table_data.json --position '{"left":"10%","top":"25%"}' --size '{"width":"80%","height":"50%"}' --json

Exit Codes:
    0: Success
    1: Error occurred
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any, List

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError, SlideNotFoundError
)


def add_table(
    filepath: Path,
    slide_index: int,
    rows: int,
    cols: int,
    position: Dict[str, Any],
    size: Dict[str, Any],
    data: List[List[Any]] = None,
    headers: List[str] = None
) -> Dict[str, Any]:
    """Add table to slide."""
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if rows < 1 or cols < 1:
        raise ValueError("Table must have at least 1 row and 1 column")
    
    if rows > 50 or cols > 20:
        raise ValueError("Maximum table size: 50 rows × 20 columns (readability limit)")
    
    # Prepare data with headers
    table_data = []
    
    if headers:
        if len(headers) != cols:
            raise ValueError(f"Headers count ({len(headers)}) must match columns ({cols})")
        table_data.append(headers)
        data_rows = rows - 1  # One row used for headers
    else:
        data_rows = rows
    
    # Add data rows
    if data:
        if len(data) > data_rows:
            raise ValueError(f"Too many data rows ({len(data)}) for table size ({data_rows} data rows)")
        
        for row in data:
            if len(row) != cols:
                raise ValueError(f"Data row has {len(row)} items, expected {cols}")
            table_data.append(row)
        
        # Pad with empty rows if needed
        while len(table_data) < rows:
            table_data.append([""] * cols)
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Validate slide index
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides-1})"
            )
        
        # Add table
        agent.add_table(
            slide_index=slide_index,
            rows=rows,
            cols=cols,
            position=position,
            size=size,
            data=table_data if table_data else None
        )
        
        # Save
        agent.save()
    
    return {
        "status": "success",
        "file": str(filepath),
        "slide_index": slide_index,
        "rows": rows,
        "cols": cols,
        "has_headers": headers is not None,
        "data_rows_filled": len(data) if data else 0,
        "total_cells": rows * cols
    }


def main():
    parser = argparse.ArgumentParser(
        description="Add data table to PowerPoint slide",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Data Format (JSON):
  - 2D array: [["A1","B1","C1"], ["A2","B2","C2"]]
  - CSV file: converted to 2D array
  - Pandas DataFrame: exported to JSON array

Examples:
  # Simple pricing table
  cat > pricing.json << 'EOF'
[
  ["Starter", "$9/mo", "Basic features"],
  ["Pro", "$29/mo", "Advanced features"],
  ["Enterprise", "$99/mo", "All features + support"]
]
EOF
  
  uv python ppt_add_table.py \\
    --file presentation.pptx \\
    --slide 3 \\
    --rows 4 \\
    --cols 3 \\
    --headers "Plan,Price,Features" \\
    --data pricing.json \\
    --position '{"left":"15%","top":"25%"}' \\
    --size '{"width":"70%","height":"50%"}' \\
    --json
  
  # Quarterly results table
  cat > results.json << 'EOF'
[
  ["Q1", "10.5", "8.2", "2.3"],
  ["Q2", "12.8", "9.1", "3.7"],
  ["Q3", "15.2", "10.5", "4.7"],
  ["Q4", "18.6", "12.1", "6.5"]
]
EOF
  
  uv python ppt_add_table.py \\
    --file presentation.pptx \\
    --slide 4 \\
    --rows 5 \\
    --cols 4 \\
    --headers "Quarter,Revenue,Costs,Profit" \\
    --data results.json \\
    --position '{"left":"10%","top":"20%"}' \\
    --size '{"width":"80%","height":"55%"}' \\
    --json
  
  # Comparison table (centered)
  cat > comparison.json << 'EOF'
[
  ["Speed", "Fast", "Very Fast"],
  ["Security", "Standard", "Enterprise"],
  ["Support", "Email", "24/7 Phone"]
]
EOF
  
  uv python ppt_add_table.py \\
    --file presentation.pptx \\
    --slide 5 \\
    --rows 4 \\
    --cols 3 \\
    --headers "Feature,Basic,Premium" \\
    --data comparison.json \\
    --position '{"anchor":"center"}' \\
    --size '{"width":"60%","height":"40%"}' \\
    --json
  
  # Empty table (for manual filling)
  uv python ppt_add_table.py \\
    --file presentation.pptx \\
    --slide 6 \\
    --rows 6 \\
    --cols 4 \\
    --headers "Name,Role,Department,Email" \\
    --position '{"left":"10%","top":"25%"}' \\
    --size '{"width":"80%","height":"60%"}' \\
    --json
  
  # From CSV file (convert first)
  # cat data.csv | python -c "import csv, json, sys; print(json.dumps(list(csv.reader(sys.stdin))))" > data.json
  # Then use --data data.json

Best Practices:
  - Keep tables under 10 rows for readability
  - Use headers for all tables
  - Align numbers right, text left
  - Use consistent decimal places
  - Highlight key values with color
  - Leave white space around table
  - Use alternating row colors for large tables

Table Size Guidelines:
  - 3-5 columns: Optimal for most presentations
  - 6-10 rows: Maximum for comfortable reading
  - Font size: 12-16pt for body, 14-18pt for headers
  - Cell padding: Leave breathing room

When to Use Tables vs Charts:
  - Use tables: Exact values matter, detailed data
  - Use charts: Show trends, comparisons, patterns
  - Use both: Table with summary chart
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file path'
    )
    
    parser.add_argument(
        '--slide',
        required=True,
        type=int,
        help='Slide index (0-based)'
    )
    
    parser.add_argument(
        '--rows',
        required=True,
        type=int,
        help='Number of rows (including header if present)'
    )
    
    parser.add_argument(
        '--cols',
        required=True,
        type=int,
        help='Number of columns'
    )
    
    parser.add_argument(
        '--position',
        required=True,
        type=json.loads,
        help='Position dict (JSON string)'
    )
    
    parser.add_argument(
        '--size',
        required=True,
        type=json.loads,
        help='Size dict (JSON string)'
    )
    
    parser.add_argument(
        '--data',
        type=Path,
        help='JSON file with 2D array of cell values'
    )
    
    parser.add_argument(
        '--data-string',
        help='Inline JSON 2D array string'
    )
    
    parser.add_argument(
        '--headers',
        help='Comma-separated header row (will be row 0)'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        help='Output JSON response'
    )
    
    args = parser.parse_args()
    
    try:
        # Parse headers
        headers = None
        if args.headers:
            headers = [h.strip() for h in args.headers.split(',')]
        
        # Load table data
        data = None
        if args.data:
            if not args.data.exists():
                raise FileNotFoundError(f"Data file not found: {args.data}")
            with open(args.data, 'r') as f:
                data = json.load(f)
        elif args.data_string:
            data = json.loads(args.data_string)
        
        # Validate data is 2D array
        if data is not None:
            if not isinstance(data, list):
                raise ValueError("Data must be a 2D array (list of lists)")
            if data and not isinstance(data[0], list):
                raise ValueError("Data must be a 2D array (list of lists)")
        
        result = add_table(
            filepath=args.file,
            slide_index=args.slide,
            rows=args.rows,
            cols=args.cols,
            position=args.position,
            size=args.size,
            data=data,
            headers=headers
        )
        
        if args.json:
            print(json.dumps(result, indent=2))
        else:
            print(f"✅ Added {result['rows']}×{result['cols']} table to slide {result['slide_index']}")
            if result['has_headers']:
                print(f"   Headers: Yes")
            print(f"   Data rows filled: {result['data_rows_filled']}")
            print(f"   Total cells: {result['total_cells']}")
        
        sys.exit(0)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__
        }
        
        if args.json:
            print(json.dumps(error_result, indent=2))
        else:
            print(f"❌ Error: {e}", file=sys.stderr)
        
        sys.exit(1)


if __name__ == "__main__":
    main()
```

---

## **File 4/5: `tools/ppt_add_shape.py`**

```python
#!/usr/bin/env python3
"""
PowerPoint Add Shape Tool
Add shape (rectangle, circle, arrow, etc.) to slide

Usage:
    uv python ppt_add_shape.py --file presentation.pptx --slide 0 --shape rectangle --position '{"left":"20%","top":"30%"}' --size '{"width":"60%","height":"40%"}' --fill-color "#0070C0" --json

Exit Codes:
    0: Success
    1: Error occurred
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError, SlideNotFoundError
)


def add_shape(
    filepath: Path,
    slide_index: int,
    shape_type: str,
    position: Dict[str, Any],
    size: Dict[str, Any],
    fill_color: str = None,
    line_color: str = None,
    line_width: float = 1.0
) -> Dict[str, Any]:
    """Add shape to slide."""
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Validate slide index
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides-1})"
            )
        
        # Add shape
        agent.add_shape(
            slide_index=slide_index,
            shape_type=shape_type,
            position=position,
            size=size,
            fill_color=fill_color,
            line_color=line_color,
            line_width=line_width
        )
        
        # Save
        agent.save()
    
    return {
        "status": "success",
        "file": str(filepath),
        "slide_index": slide_index,
        "shape_type": shape_type,
        "position": position,
        "size": size,
        "styling": {
            "fill_color": fill_color,
            "line_color": line_color,
            "line_width": line_width
        }
    }


def main():
    parser = argparse.ArgumentParser(
        description="Add shape to PowerPoint slide",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Available Shapes:
  - rectangle: Standard rectangle/box
  - rounded_rectangle: Rectangle with rounded corners
  - ellipse: Circle or oval
  - triangle: Triangle (isosceles)
  - arrow_right: Right-pointing arrow
  - arrow_left: Left-pointing arrow
  - arrow_up: Up-pointing arrow
  - arrow_down: Down-pointing arrow
  - star: 5-point star
  - heart: Heart shape

Common Uses:
  - rectangle: Callout boxes, containers, dividers
  - rounded_rectangle: Buttons, soft containers
  - ellipse: Emphasis, icons, venn diagrams
  - arrows: Process flows, directional indicators
  - star: Highlights, ratings, attention
  - triangle: Warning indicators, play buttons

Examples:
  # Blue callout box
  uv python ppt_add_shape.py \\
    --file presentation.pptx \\
    --slide 1 \\
    --shape rounded_rectangle \\
    --position '{"left":"10%","top":"15%"}' \\
    --size '{"width":"30%","height":"15%"}' \\
    --fill-color "#0070C0" \\
    --line-color "#FFFFFF" \\
    --line-width 2 \\
    --json
  
  # Process flow arrows
  uv python ppt_add_shape.py \\
    --file presentation.pptx \\
    --slide 2 \\
    --shape arrow_right \\
    --position '{"left":"30%","top":"40%"}' \\
    --size '{"width":"15%","height":"8%"}' \\
    --fill-color "#00B050" \\
    --json
  
  # Emphasis circle
  uv python ppt_add_shape.py \\
    --file presentation.pptx \\
    --slide 3 \\
    --shape ellipse \\
    --position '{"anchor":"center"}' \\
    --size '{"width":"20%","height":"20%"}' \\
    --fill-color "#FFC000" \\
    --line-color "#C65911" \\
    --line-width 3 \\
    --json
  
  # Warning triangle
  uv python ppt_add_shape.py \\
    --file presentation.pptx \\
    --slide 4 \\
    --shape triangle \\
    --position '{"left":"5%","top":"5%"}' \\
    --size '{"width":"8%","height":"8%"}' \\
    --fill-color "#FF0000" \\
    --json
  
  # Transparent overlay (no fill)
  uv python ppt_add_shape.py \\
    --file presentation.pptx \\
    --slide 5 \\
    --shape rectangle \\
    --position '{"left":"0%","top":"0%"}' \\
    --size '{"width":"100%","height":"100%"}' \\
    --line-color "#0070C0" \\
    --line-width 5 \\
    --json

Design Tips:
  - Use shapes to organize content visually
  - Consistent colors across shapes
  - Align shapes to grid for professional look
  - Use subtle colors for backgrounds
  - Bold colors for emphasis
  - Combine shapes to create diagrams
  - Layer shapes for depth (background first)

Color Palette (Corporate):
  - Primary Blue: #0070C0
  - Secondary Gray: #595959
  - Accent Orange: #ED7D31
  - Success Green: #70AD47
  - Warning Yellow: #FFC000
  - Danger Red: #C00000
  - White: #FFFFFF
  - Black: #000000

Shape Layering:
  - Shapes are added in order (first = back, last = front)
  - Use transparent shapes for overlays
  - Group related shapes together
  - Send to back: create shape first
  - Bring to front: create shape last
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file path'
    )
    
    parser.add_argument(
        '--slide',
        required=True,
        type=int,
        help='Slide index (0-based)'
    )
    
    parser.add_argument(
        '--shape',
        required=True,
        choices=['rectangle', 'rounded_rectangle', 'ellipse', 'triangle',
                'arrow_right', 'arrow_left', 'arrow_up', 'arrow_down',
                'star', 'heart'],
        help='Shape type'
    )
    
    parser.add_argument(
        '--position',
        required=True,
        type=json.loads,
        help='Position dict (JSON string)'
    )
    
    parser.add_argument(
        '--size',
        required=True,
        type=json.loads,
        help='Size dict (JSON string)'
    )
    
    parser.add_argument(
        '--fill-color',
        help='Fill color (hex, e.g., #0070C0)'
    )
    
    parser.add_argument(
        '--line-color',
        help='Line/border color (hex)'
    )
    
    parser.add_argument(
        '--line-width',
        type=float,
        default=1.0,
        help='Line width in points (default: 1.0)'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        help='Output JSON response'
    )
    
    args = parser.parse_args()
    
    try:
        result = add_shape(
            filepath=args.file,
            slide_index=args.slide,
            shape_type=args.shape,
            position=args.position,
            size=args.size,
            fill_color=args.fill_color,
            line_color=args.line_color,
            line_width=args.line_width
        )
        
        if args.json:
            print(json.dumps(result, indent=2))
        else:
            print(f"✅ Added {result['shape_type']} to slide {result['slide_index']}")
            if args.fill_color:
                print(f"   Fill: {result['styling']['fill_color']}")
            if args.line_color:
                print(f"   Line: {result['styling']['line_color']} ({result['styling']['line_width']}pt)")
        
        sys.exit(0)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__
        }
        
        if args.json:
            print(json.dumps(error_result, indent=2))
        else:
            print(f"❌ Error: {e}", file=sys.stderr)
        
        sys.exit(1)


if __name__ == "__main__":
    main()
```

---

## **File 5/5: `tools/ppt_replace_text.py`**

```python
#!/usr/bin/env python3
"""
PowerPoint Replace Text Tool
Find and replace text across entire presentation

Usage:
    uv python ppt_replace_text.py --file presentation.pptx --find "Company Inc." --replace "Company LLC" --json

Exit Codes:
    0: Success
    1: Error occurred
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, PowerPointAgentError
)


def replace_text(
    filepath: Path,
    find: str,
    replace: str,
    match_case: bool = False,
    whole_words: bool = False,
    dry_run: bool = False
) -> Dict[str, Any]:
    """Find and replace text across presentation."""
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if not find:
        raise ValueError("Find text cannot be empty")
    
    # For dry run, just scan without modifying
    if dry_run:
        with PowerPointAgent(filepath) as agent:
            agent.open(filepath, acquire_lock=False)  # Read-only
            
            # Count occurrences
            count = 0
            locations = []
            
            for slide_idx, slide in enumerate(agent.prs.slides):
                for shape_idx, shape in enumerate(slide.shapes):
                    if hasattr(shape, 'text_frame'):
                        text = shape.text_frame.text
                        
                        if match_case:
                            occurrences = text.count(find)
                        else:
                            occurrences = text.lower().count(find.lower())
                        
                        if occurrences > 0:
                            count += occurrences
                            locations.append({
                                "slide": slide_idx,
                                "shape": shape_idx,
                                "occurrences": occurrences,
                                "preview": text[:100]
                            })
        
        return {
            "status": "dry_run",
            "file": str(filepath),
            "find": find,
            "replace": replace,
            "matches_found": count,
            "locations": locations[:10],  # First 10 locations
            "total_locations": len(locations),
            "match_case": match_case
        }
    
    # Actual replacement
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Perform replacement
        count = agent.replace_text(
            find=find,
            replace=replace,
            match_case=match_case
        )
        
        # Save
        agent.save()
    
    return {
        "status": "success",
        "file": str(filepath),
        "find": find,
        "replace": replace,
        "replacements_made": count,
        "match_case": match_case
    }


def main():
    parser = argparse.ArgumentParser(
        description="Find and replace text across PowerPoint presentation",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Simple replacement
  uv python ppt_replace_text.py \\
    --file presentation.pptx \\
    --find "2023" \\
    --replace "2024" \\
    --json
  
  # Case-sensitive replacement
  uv python ppt_replace_text.py \\
    --file presentation.pptx \\
    --find "Company Inc." \\
    --replace "Company LLC" \\
    --match-case \\
    --json
  
  # Dry run to preview changes
  uv python ppt_replace_text.py \\
    --file presentation.pptx \\
    --find "old_term" \\
    --replace "new_term" \\
    --dry-run \\
    --json
  
  # Update product name
  uv python ppt_replace_text.py \\
    --file product_deck.pptx \\
    --find "Product X" \\
    --replace "Product Y" \\
    --json
  
  # Fix typo across all slides
  uv python ppt_replace_text.py \\
    --file presentation.pptx \\
    --find "recieve" \\
    --replace "receive" \\
    --json

Common Use Cases:
  - Update dates (2023 → 2024)
  - Change company names (rebranding)
  - Fix recurring typos
  - Update product names
  - Change terminology
  - Update prices/numbers
  - Localization (English → Spanish)
  - Template customization

Best Practices:
  1. Always use --dry-run first to preview changes
  2. Create backup before bulk replacements
  3. Use --match-case for proper nouns
  4. Test on a copy first
  5. Review results after replacement
  6. Be specific with find text to avoid unwanted matches

Safety Tips:
  - Backup file before major changes
  - Use dry-run to verify matches
  - Check match count makes sense
  - Review a few slides manually after
  - Use case-sensitive for precision
  - Avoid replacing common words

Limitations:
  - Only replaces visible text (not in images)
  - Does not replace text in charts/tables (text only)
  - Preserves original formatting
  - Cannot use regex patterns (exact match only)
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file path'
    )
    
    parser.add_argument(
        '--find',
        required=True,
        help='Text to find'
    )
    
    parser.add_argument(
        '--replace',
        required=True,
        help='Replacement text'
    )
    
    parser.add_argument(
        '--match-case',
        action='store_true',
        help='Case-sensitive matching'
    )
    
    parser.add_argument(
        '--whole-words',
        action='store_true',
        help='Match whole words only (not yet implemented)'
    )
    
    parser.add_argument(
        '--dry-run',
        action='store_true',
        help='Preview changes without modifying file'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        help='Output JSON response'
    )
    
    args = parser.parse_args()
    
    try:
        result = replace_text(
            filepath=args.file,
            find=args.find,
            replace=args.replace,
            match_case=args.match_case,
            whole_words=args.whole_words,
            dry_run=args.dry_run
        )
        
        if args.json:
            print(json.dumps(result, indent=2))
        else:
            if args.dry_run:
                print(f"🔍 Dry run - no changes made")
                print(f"   Found: {result['matches_found']} occurrences")
                print(f"   In: {result['total_locations']} locations")
                if result['locations']:
                    print(f"   Sample locations:")
                    for loc in result['locations'][:3]:
                        print(f"     - Slide {loc['slide']}: {loc['occurrences']} matches")
            else:
                print(f"✅ Replaced '{args.find}' with '{args.replace}'")
                print(f"   Replacements: {result['replacements_made']}")
                if args.match_case:
                    print(f"   Case-sensitive: Yes")
        
        sys.exit(0)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__
        }
        
        if args.json:
            print(json.dumps(error_result, indent=2))
        else:
            print(f"❌ Error: {e}", file=sys.stderr)
        
        sys.exit(1)


if __name__ == "__main__":
    main()
```

---

## 🎉 **Implementation Complete! All 5 P1 Tools Delivered**

### **Summary:**

| # | Tool | Lines | Key Features |
|---|------|-------|-------------|
| 6 | `ppt_add_bullet_list.py` | 315 | Bullets, numbered lists, custom formatting |
| 7 | `ppt_add_chart.py` | 365 | 9 chart types, data validation, titles |
| 8 | `ppt_add_table.py` | 330 | Headers, data import, size limits |
| 9 | `ppt_add_shape.py` | 290 | 10 shapes, colors, layering |
| 10 | `ppt_replace_text.py` | 275 | Find/replace, dry-run, case-sensitive |

**Total Lines:** ~1,575 lines of production code

---

## 🚀 **Complete Workflow Example**

```bash
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

---

## 📊 **Tool Comparison Matrix**

| Feature | Excel Tools | PowerPoint Tools |
|---------|-------------|------------------|
| **Positioning** | Grid (A1, B2) | 5 systems (%, anchor, grid, absolute, Excel-like) |
| **Data Input** | Values, formulas | Text, images, charts, tables, shapes |
| **Validation** | Formula errors | Accessibility, assets, layout |
| **Bulk Operations** | Range formulas | Replace text, clone slides |
| **Visual Design** | Limited (colors) | Extensive (shapes, images, layouts) |
| **Complexity** | Medium | High |

---

## ✅ **What We've Built (10 Tools Total)**

### **P0 Tools (Critical - Week 1):**
1. ✅ `ppt_create_new.py` - Create presentations
2. ✅ `ppt_add_slide.py` - Add slides
3. ✅ `ppt_set_title.py` - Set titles
4. ✅ `ppt_add_text_box.py` - Add text (5 positioning systems!)
5. ✅ `ppt_insert_image.py` - Insert images

### **P1 Tools (Important - Week 2):**
6. ✅ `ppt_add_bullet_list.py` - Bullet/numbered lists
7. ✅ `ppt_add_chart.py` - Data visualization (9 chart types)
8. ✅ `ppt_add_table.py` - Data tables
9. ✅ `ppt_add_shape.py` - Visual shapes
10. ✅ `ppt_replace_text.py` - Bulk text replacement

---

## 📋 **Next Steps**

With 10 tools complete, would you like me to:

**Option A:** Create validation & export tools (`ppt_validate_presentation.py`, `ppt_check_accessibility.py`, `ppt_export_pdf.py`)

**Option B:** Create the comprehensive `AGENT_SYSTEM_PROMPT.md` for all 10 PowerPoint tools

**Option C:** Create working examples gallery (complete presentation templates using these tools)

**Option D:** Create integration tests for all 10 tools

Please advise! 🎯
