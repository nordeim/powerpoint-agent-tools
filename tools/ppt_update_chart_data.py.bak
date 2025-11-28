#!/usr/bin/env python3
"""
PowerPoint Update Chart Data Tool
Update the data of an existing chart

Usage:
    uv python ppt_update_chart_data.py --file deck.pptx --slide 0 --chart 0 --data new_data.json --json
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
from pptx.chart.data import CategoryChartData

def update_chart_data(
    filepath: Path,
    slide_index: int,
    chart_index: int,
    data: Dict[str, Any]
) -> Dict[str, Any]:
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
        
    if "categories" not in data or "series" not in data:
        raise ValueError("Data JSON must contain 'categories' and 'series'")

    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        total = agent.get_slide_count()
        if not 0 <= slide_index < total:
            raise SlideNotFoundError(f"Slide index {slide_index} out of range")
            
        slide = agent.prs.slides[slide_index]
        
        # Find the chart
        charts = [shape for shape in slide.shapes if shape.has_chart]
        if not 0 <= chart_index < len(charts):
             raise ValueError(f"Chart index {chart_index} out of range. Slide has {len(charts)} charts.")
             
        chart = charts[chart_index].chart
        
        # Create new chart data
        chart_data = CategoryChartData()
        chart_data.categories = data["categories"]
        
        for series in data["series"]:
            chart_data.add_series(series["name"], series["values"])
            
        # Replace data
        chart.replace_data(chart_data)
        
        agent.save()
        
    return {
        "status": "success",
        "file": str(filepath),
        "slide_index": slide_index,
        "chart_index": chart_index,
        "data_points": sum(len(s["values"]) for s in data["series"])
    }

def main():
    parser = argparse.ArgumentParser(description="Update PowerPoint chart data")
    parser.add_argument('--file', required=True, type=Path, help='PowerPoint file path')
    parser.add_argument('--slide', required=True, type=int, help='Slide index (0-based)')
    parser.add_argument('--chart', required=True, type=int, help='Chart index (0-based)')
    parser.add_argument('--data', required=True, type=Path, help='JSON data file')
    parser.add_argument('--json', action='store_true', default=True, help='Output JSON')
    
    args = parser.parse_args()
    
    try:
        with open(args.data, 'r') as f:
            data_content = json.load(f)
            
        result = update_chart_data(
            filepath=args.file, 
            slide_index=args.slide, 
            chart_index=args.chart,
            data=data_content
        )
        print(json.dumps(result, indent=2))
        sys.exit(0)
    except Exception as e:
        print(json.dumps({"status": "error", "error": str(e), "error_type": type(e).__name__}, indent=2))
        sys.exit(1)

if __name__ == "__main__":
    main()
