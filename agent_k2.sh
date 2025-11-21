#!/bin/bash
set -e

# ============================================================================
# BITCOIN MARKET REPORT Q4 2025 - PRESENTATION GENERATOR
# Professional Deck: 12 Slides | Corporate Theme | Full Validation
# ============================================================================

# CONFIGURATION
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
OUTPUT="${SCRIPT_DIR}/bitcoin_market_report_q4_2025.pptx"
DATA_DIR="${SCRIPT_DIR}/ppt_chart_data"
EXPORT_DIR="${SCRIPT_DIR}/ppt_export"  # Explicit export directory path

# Create required directories with explicit validation
mkdir -p "$DATA_DIR"
mkdir -p "$EXPORT_DIR"  # Pre-create export directory for image thumbnails

# Create chart data files with absolute paths
LIQUIDITY_DATA="${DATA_DIR}/liquidity_data.json"
INSTITUTIONAL_DATA="${DATA_DIR}/institutional_data.json"

# ============================================================================
# STEP 1: Chart Data Preparation
# ============================================================================

cat > "$LIQUIDITY_DATA" << 'EOF'
{
   "categories": ["October 2025", "November 2025"],
   "series": [
      {
         "name": "Market Depth ($ Millions)",
         "values": [700, 535]
      }
   ]
}
EOF

cat > "$INSTITUTIONAL_DATA" << 'EOF'
{
   "categories": ["Daily Average (BTC)"],
   "series": [
      {
         "name": "Institutional Purchases",
         "values": [450]
      },
      {
         "name": "Daily Mined Supply",
         "values": [900]
      }
   ]
}
EOF

echo "✓ Chart data files created"
echo "✓ Export directory validated: $EXPORT_DIR"

# ============================================================================
# STEP 2: Presentation Creation & Architecture
# ============================================================================

# Create new presentation with corporate layout
uv run tools/ppt_create_new.py --output "$OUTPUT" --layout "Title Slide" --json

# Add core content slides
uv run tools/ppt_add_slide.py --file "$OUTPUT" --layout "Title and Content" --index 1 --title "Executive Summary" --json
uv run tools/ppt_add_slide.py --file "$OUTPUT" --layout "Title and Content" --index 2 --title "Key Causes of the Downturn" --json
uv run tools/ppt_add_slide.py --file "$OUTPUT" --layout "Title and Content" --index 3 --title "Panic Selling & Forced Liquidations" --json
uv run tools/ppt_add_slide.py --file "$OUTPUT" --layout "Title and Content" --index 4 --title "Critical Liquidity Shortage" --json
uv run tools/ppt_add_slide.py --file "$OUTPUT" --layout "Title and Content" --index 5 --title "Macroeconomic Headwinds" --json
uv run tools/ppt_add_slide.py --file "$OUTPUT" --layout "Title and Content" --index 6 --title "Institutional Demand Slowdown" --json
uv run tools/ppt_add_slide.py --file "$OUTPUT" --layout "Title and Content" --index 7 --title "Technical Breakdown & Sentiment" --json
uv run tools/ppt_add_slide.py --file "$OUTPUT" --layout "Title and Content" --index 8 --title "Seasonal & Halving Cycle Dynamics" --json
uv run tools/ppt_add_slide.py --file "$OUTPUT" --layout "Title and Content" --index 9 --title "On-Chain Data Insights" --json
uv run tools/ppt_add_slide.py --file "$OUTPUT" --layout "Title and Content" --index 10 --title "Conclusion & Recovery Outlook" --json
uv run tools/ppt_add_slide.py --file "$OUTPUT" --layout "Title and Content" --index 11 --title "Sources & References" --json

echo "✓ 12-slide structure created"
echo "✓ Building content layers..."

# ============================================================================
# STEPS 3-14: Content Population (Unchanged from original)
# ============================================================================

# [Content steps 3-14 remain identical to original script]
# Using compact representation for brevity while preserving full functionality

uv run tools/ppt_set_title.py --file "$OUTPUT" --slide 0 --title "Bitcoin Market Analysis: Navigating the Downturn" --subtitle "Q4 2025 Market Report | November 20, 2025" --json

uv run tools/ppt_add_bullet_list.py --file "$OUTPUT" --slide 1 --items "Bitcoin corrected from >$120,000 to <$95,000 in November 2025,Macroeconomic pressures and liquidity crunch drove >20% decline,Rare institutional buying slowdown vs. daily mined supply,Technical breach of $100K support triggered panic cascade,Recovery depends on risk sentiment stabilization and renewed inflows" --position '{"left": "10%", "top": "30%", "width": "80%", "height": "55%"}' --font-size 20 --color "#111111" --json
uv run tools/ppt_add_shape.py --file "$OUTPUT" --slide 1 --shape rectangle --position '{"anchor": "bottom_right", "offset_x": -2, "offset_y": -2}' --size '{"width": "25%", "height": "12%"}' --fill-color "#ED7D31" --json
uv run tools/ppt_add_text_box.py --file "$OUTPUT" --slide 1 --text "Price Drop: -20%+" --position '{"anchor": "bottom_right", "offset_x": -1.5, "offset_y": -1.5}' --size '{"font-size": 18, "color": "#FFFFFF", "bold": true}' --json

uv run tools/ppt_add_bullet_list.py --file "$OUTPUT" --slide 2 --items "Panic Selling by Short-Term Holders,Critical Liquidity Crunch,Macroeconomic Uncertainty,Institutional Buying Slowdown,Technical Breakdown & Sentiment Shifts" --position '{"left": "10%", "top": "30%", "width": "80%", "height": "60%"}' --font-size 22 --color "#111111" --json

uv run tools/ppt_add_text_box.py --file "$OUTPUT" --slide 3 --text "Price Range: $120,000 → $95,000" --position '{"left": "10%", "top": "30%", "width": "80%", "height": "10%"}' --font-size 24 --color "#0070C0" --bold true --json
uv run tools/ppt_add_bullet_list.py --file "$OUTPUT" --slide 3 --items "Short-term holders accelerated selling at a loss,Forced liquidations cascaded through derivatives markets,Long-term holders took minimal profits—no bear market distribution,Deleveraging cleared speculative positions from the system" --position '{"left": "10%", "top": "45%", "width": "80%", "height": "40%"}' --font-size 20 --color "#111111" --json
uv run tools/ppt_add_shape.py --file "$OUTPUT" --slide 3 --shape arrow_down --position '{"anchor": "top_right", "offset_x": -3, "offset_y": 3}' --size '{"width": "8%", "height": "15%"}' --fill-color "#C00000" --json

uv run tools/ppt_add_chart.py --file "$OUTPUT" --slide 4 --chart-type column --data "$LIQUIDITY_DATA" --position '{"left": "10%", "top": "35%", "width": "45%", "height": "45%"}' --title "Market Depth Collapse" --json
uv run tools/ppt_add_text_box.py --file "$OUTPUT" --slide 4 --text "Market depth fell from $700M+ to ~$535M, creating extreme price vulnerability. Thinner order books amplify volatility and enable cascading sell-offs." --position '{"left": "58%", "top": "40%", "width": "35%", "height": "35%"}' --font-size 18 --color "#595959" --json
uv run tools/ppt_add_text_box.py --file "$OUTPUT" --slide 4 --text "-24% Liquidity" --position '{"left": "58%", "top": "35%", "width": "25%"}' --font-size 20 --color "#ED7D31" --bold true --json

uv run tools/ppt_add_bullet_list.py --file "$OUTPUT" --slide 5 --items "Federal Reserve caution on interest rate cuts rattled risk assets,Inflation resilience and USD strength pressured Bitcoin,US-China trade war fears suppressed global risk appetite,Capital flight from high-risk markets accelerated" --position '{"left": "10%", "top": "30%", "width": "80%", "height": "50%"}' --font-size 20 --color "#111111" --json
uv run tools/ppt_add_shape.py --file "$OUTPUT" --slide 5 --shape rectangle --position '{"left": "5%", "top": "25%", "width": "90%", "height": "60%"}' --fill-color "#F5F5F5" --json

uv run tools/ppt_add_chart.py --file "$OUTPUT" --slide 6 --chart-type bar --data "$INSTITUTIONAL_DATA" --position '{"left": "10%", "top": "35%", "width": "60%", "height": "45%"}' --title "Institutional Demand vs. Supply" --json
uv run tools/ppt_add_text_box.py --file "$OUTPUT" --slide 6 --text "For the first time in seven months, net institutional purchases fell below daily mined supply. Large players are not absorbing new issuance." --position '{"left": "15%", "top": "82%", "width": "70%"}' --font-size 18 --color "#595959" --json

uv run tools/ppt_add_text_box.py --file "$OUTPUT" --slide 7 --text "$100,000 Support Breach" --position '{"left": "10%", "top": "30%"}' --font-size 28 --color "#C00000" --bold true --json
uv run tools/ppt_add_bullet_list.py --file "$OUTPUT" --slide 7 --items "Psychological $100K level failure triggered retail panic,Options traders positioned for further downside,Market sentiment: Extreme Fear,Many holders now underwater on cost basis" --position '{"left": "10%", "top": "45%", "width": "80%", "height": "40%"}' --font-size 20 --color "#111111" --json
uv run tools/ppt_add_shape.py --file "$OUTPUT" --slide 7 --shape rectangle --position '{"anchor": "center_right", "offset_x": -5, "offset_y": 2}' --size '{"width": "15%", "height": "8%"}' --fill-color "#C00000" --json
uv run tools/ppt_add_text_box.py --file "$OUTPUT" --slide 7 --text "EXTREME FEAR" --position '{"anchor": "center_right", "offset_x": -4.5, "offset_y": 2.5}' --size '{"font-size": 12, "color": "#FFFFFF", "bold": true}' --json

uv run tools/ppt_add_bullet_list.py --file "$OUTPUT" --slide 8 --items "November historically Bitcoin's strongest month,Post-halving volatility patterns show mid-cycle dips,Current correction may be healthy position clearing,Consolidation phase likely before potential rebound" --position '{"left": "10%", "top": "30%", "width": "80%", "height": "50%"}' --font-size 20 --color "#111111" --json
uv run tools/ppt_add_shape.py --file "$OUTPUT" --slide 8 --shape rectangle --position '{"anchor": "bottom_right", "offset_x": -2, "offset_y": -2}' --size '{"width": "30%", "height": "12%"}' --fill-color "#70AD47" --json
uv run tools/ppt_add_text_box.py --file "$OUTPUT" --slide 8 --text "Historical Pattern: Reset → Rally" --position '{"anchor": "bottom_right", "offset_x": -1.5, "offset_y": -1.5}' --size '{"font-size": 16, "color": "#FFFFFF", "bold": true}' --json

uv run tools/ppt_add_bullet_list.py --file "$OUTPUT" --slide 9 --items "Coins moving from long-dormant wallets to exchanges,Indicates capitulation by weaker hands—not institutional exodus,ETF inflows from traditional finance have slowed significantly,Selling pressure concentrated in short-term holder cohort" --position '{"left": "10%", "top": "30%", "width": "80%", "height": "50%"}' --font-size 20 --color "#111111" --json

uv run tools/ppt_add_text_box.py --file "$OUTPUT" --slide 10 --text "Recovery Prerequisites" --position '{"left": "10%", "top": "30%"}' --font-size 24 --color "#0070C0" --bold true --json
uv run tools/ppt_add_bullet_list.py --file "$OUTPUT" --slide 10 --items "Stabilization of global risk sentiment,Renewed institutional capital inflows,Macroeconomic policy clarity from Fed,Absorption of current supply overhang" --position '{"left": "10%", "top": "45%", "width": "80%", "height": "40%"}' --font-size 20 --color "#111111" --json
uv run tools/ppt_add_shape.py --file "$OUTPUT" --slide 10 --shape rectangle --position '{"left": "15%", "top": "75%", "width": "70%", "height": "10%"}' --fill-color "#2A9D8F" --json
uv run tools/ppt_add_text_box.py --file "$OUTPUT" --slide 10 --text "The downturn reflects market maturation—not structural failure." --position '{"left": "20%", "top": "78%", "width": "60%"}' --size '{"font-size": 18, "color": "#FFFFFF", "bold": true}' --json

uv run tools/ppt_add_text_box.py --file "$OUTPUT" --slide 11 --text "Primary Sources" --position '{"left": "10%", "top": "25%"}' --font-size 22 --color "#111111" --bold true --json
uv run tools/ppt_add_text_box.py --file "$OUTPUT" --slide 11 --text "Yahoo Finance: Real Reason Behind Bitcoin's Drop\nCNN Business: Bitcoin Price & Crypto Stocks\nBusiness Insider: Why Bitcoin Is Falling\nKi Ecke: Bitcoin Price Outlook November 2025\nThe Motley Fool: 2 Major Factors Driving Decline\n99Bitcoins: Saylor Buys The Dip\nCoinDesk: November Strongest Month Analysis\nFinanceFeeds: Technical Analysis Report Nov 20" --position '{"left": "10%", "top": "35%", "width": "80%", "height": "50%"}' --font-size 12 --color "#595959" --json

# ============================================================================
# STEP 15: Global Design Standardization
# ============================================================================

# Set corporate background
uv run tools/ppt_set_background.py --file "$OUTPUT" --color "#FFFFFF" --json

# Set professional footer with slide numbers
uv run tools/ppt_set_footer.py --file "$OUTPUT" \
  --text "Confidential • Q4 2025 Bitcoin Analysis" \
  --show-number true --show-date false --json

echo "✓ Design standards applied"

# ============================================================================
# STEP 16: Professional Polish - Visual Layering
# ============================================================================

# Add subtle title accent shapes on key slides
for SLIDE_NUM in 3 5 7; do
  uv run tools/ppt_add_shape.py --file "$OUTPUT" --slide "$SLIDE_NUM" --shape rectangle \
    --position '{"left": "10%", "top": "28%", "width": "25%", "height": "0.3%"}' \
    --fill-color "#ED7D31" --json
done

echo "✓ Visual polish applied"

# ============================================================================
# STEP 17: Validation & Quality Assurance
# ============================================================================

echo "Running mandatory validation checks..."

# General health check
VALIDATION=$(uv run tools/ppt_validate_presentation.py --file "$OUTPUT" --json)
echo "Validation Result: $VALIDATION"

# Accessibility compliance check (WCAG 2.1)
ACCESSIBILITY=$(uv run tools/ppt_check_accessibility.py --file "$OUTPUT" --json)
echo "Accessibility Result: $ACCESSIBILITY"

# Extract slide count for confirmation
INFO=$(uv run tools/ppt_get_info.py --file "$OUTPUT" --json)
SLIDE_COUNT=$(echo "$INFO" | grep -o '"slides":[0-9]*' | grep -o '[0-9]*')
echo "✓ Final deck contains $SLIDE_COUNT slides"

# ============================================================================
# STEP 18: Export Deliverables with Pre-validated Paths
# ============================================================================

# Export to PDF (output directory pre-validated)
uv run tools/ppt_export_pdf.py --file "$OUTPUT" \
  --output "${SCRIPT_DIR}/bitcoin_report_q4_2025.pdf" --json

# Export slide thumbnails (export directory guaranteed to exist)
uv run tools/ppt_export_images.py --file "$OUTPUT" \
  --output-dir "$EXPORT_DIR" --format png --json

echo "✓ PDF and PNG exports created to validated directories"

# ============================================================================
# FINAL STATUS
# ============================================================================
echo ""
echo "=================================================="
echo "PRESENTATION CREATED SUCCESSFULLY"
echo "=================================================="
echo "File: $OUTPUT"
echo "Slides: $SLIDE_COUNT"
echo "Data Directory: $DATA_DIR"
echo "Export Directory: $EXPORT_DIR"
echo ""
echo "Next Steps:"
echo "1. Review slide 11 (Sources) for URL formatting"
echo "2. Verify chart data visualization on slides 4 & 6"
echo "3. Check accessibility alt-text for charts (autogenerated)"
echo "4. Adjust footer date if needed"
echo "=================================================="
