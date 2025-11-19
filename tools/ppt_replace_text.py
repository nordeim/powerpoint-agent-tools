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
