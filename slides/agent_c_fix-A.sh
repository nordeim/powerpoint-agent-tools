#!/bin/bash
# Footer Remediation Script
# Adds manual footer text boxes to slides 1-8

set -e

FOOTER_TEXT="Bitcoin Market Report • November 2024"
PRIMARY_COLOR="#595959"  # Corporate secondary color for subtlety

echo "🔧 Applying manual footer to slides 1-8..."

# Loop through content slides (skip title slide 0)
for SLIDE in {1..8}; do
  # Add footer text box (bottom-left)
  uv run tools/ppt_add_text_box.py \
    --file "bitcoin_market_report_nov2024.pptx" \
    --slide $SLIDE \
    --text "$FOOTER_TEXT" \
    --position '{"left": "5%", "top": "92%"}' \
    --size '{"width": "60%", "height": "5%"}' \
    --font-size 10 \
    --color "$PRIMARY_COLOR" \
    --json
  
  # Add slide number (bottom-right)
  SLIDE_NUM=$((SLIDE + 1))  # Display number (1-indexed for audience)
  uv run tools/ppt_add_text_box.py \
    --file "bitcoin_market_report_nov2024.pptx" \
    --slide $SLIDE \
    --text "$SLIDE_NUM" \
    --position '{"left": "92%", "top": "92%"}' \
    --size '{"width": "5%", "height": "5%"}' \
    --font-size 10 \
    --color "$PRIMARY_COLOR" \
    --json
  
  echo "  ✅ Footer applied to slide $SLIDE"
done

echo ""
echo "✅ Footer remediation complete!"
echo "🔍 Verifying slide 1..."

uv run tools/ppt_get_slide_info.py \
  --file "bitcoin_market_report_nov2024.pptx" \
  --slide 1 \
  --json
