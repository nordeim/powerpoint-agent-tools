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
