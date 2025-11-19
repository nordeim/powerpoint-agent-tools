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
