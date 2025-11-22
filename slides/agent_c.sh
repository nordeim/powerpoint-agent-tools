#!/bin/bash
# ═══════════════════════════════════════════════════════════════
# Bitcoin Market Report - November 2024 Downturn Analysis
# Professional PowerPoint Generation Script
# ═══════════════════════════════════════════════════════════════

set -e  # Exit on error

# Color palette (Corporate theme)
PRIMARY="#0070C0"
ACCENT="#ED7D31"

echo "🚀 Starting Bitcoin Market Report presentation generation..."

# ═══════════════════════════════════════════════════════════════
# STEP 1: Create base presentation with Title Slide
# ═══════════════════════════════════════════════════════════════
uv run tools/ppt_create_new.py \
  --output "bitcoin_market_report_nov2024.pptx" \
  --layout "Title Slide" \
  --json

# ═══════════════════════════════════════════════════════════════
# STEP 2: Set title slide content
# ═══════════════════════════════════════════════════════════════
uv run tools/ppt_set_title.py \
  --file "bitcoin_market_report_nov2024.pptx" \
  --slide 0 \
  --title "Bitcoin Market Report: November 2024 Downturn Analysis" \
  --subtitle "Macroeconomic Pressures, Market Structure, and Investor Behavior" \
  --json

# ═══════════════════════════════════════════════════════════════
# SLIDE 2: Executive Summary
# ═══════════════════════════════════════════════════════════════
uv run tools/ppt_add_slide.py \
  --file "bitcoin_market_report_nov2024.pptx" \
  --layout "Title and Content" \
  --json

uv run tools/ppt_set_title.py \
  --file "bitcoin_market_report_nov2024.pptx" \
  --slide 1 \
  --title "Executive Summary" \
  --json

uv run tools/ppt_add_bullet_list.py \
  --file "bitcoin_market_report_nov2024.pptx" \
  --slide 1 \
  --items "Recent downturn driven by macroeconomic pressures and market imbalances,Sharp decline from above \$120K to below \$95K,Key drivers: panic selling and liquidity crunch,Institutional demand slowdown with technical support breakdown,Recovery dependent on stabilized risk sentiment and renewed inflows" \
  --position '{"left": "8%", "top": "25%"}' \
  --size '{"width": "84%", "height": "60%"}' \
  --font-size 20 \
  --json

# ═══════════════════════════════════════════════════════════════
# SLIDE 3: Key Driver #1 - Panic Selling by Short-Term Holders
# ═══════════════════════════════════════════════════════════════
uv run tools/ppt_add_slide.py \
  --file "bitcoin_market_report_nov2024.pptx" \
  --layout "Title and Content" \
  --json

uv run tools/ppt_set_title.py \
  --file "bitcoin_market_report_nov2024.pptx" \
  --slide 2 \
  --title "Key Driver #1: Panic Selling by Short-Term Holders" \
  --json

uv run tools/ppt_add_bullet_list.py \
  --file "bitcoin_market_report_nov2024.pptx" \
  --slide 2 \
  --items "Price acceleration from \$120K+ to <\$95K driven by short-term holders,Selling at a loss triggered forced liquidations and deleveraging,Long-term holders took profits but NOT mass distribution,Pattern differs from typical bear market cycle tops,Short-term holder capitulation signals potential market bottom" \
  --position '{"left": "8%", "top": "25%"}' \
  --size '{"width": "84%", "height": "55%"}' \
  --font-size 18 \
  --json

# Add accent callout box for key statistic
uv run tools/ppt_add_shape.py \
  --file "bitcoin_market_report_nov2024.pptx" \
  --slide 2 \
  --shape "rectangle" \
  --position '{"left": "62%", "top": "72%"}' \
  --size '{"width": "32%", "height": "15%"}' \
  --fill-color "#ED7D31" \
  --json

uv run tools/ppt_add_text_box.py \
  --file "bitcoin_market_report_nov2024.pptx" \
  --slide 2 \
  --text "Price Drop: \$120K → \$95K" \
  --position '{"left": "63%", "top": "75%"}' \
  --size '{"width": "30%", "height": "8%"}' \
  --font-size 18 \
  --color "#FFFFFF" \
  --json

# ═══════════════════════════════════════════════════════════════
# SLIDE 4: Key Driver #2 - Liquidity Crunch
# ═══════════════════════════════════════════════════════════════
uv run tools/ppt_add_slide.py \
  --file "bitcoin_market_report_nov2024.pptx" \
  --layout "Title and Content" \
  --json

uv run tools/ppt_set_title.py \
  --file "bitcoin_market_report_nov2024.pptx" \
  --slide 3 \
  --title "Key Driver #2: Liquidity Crunch" \
  --json

uv run tools/ppt_add_bullet_list.py \
  --file "bitcoin_market_report_nov2024.pptx" \
  --slide 3 \
  --items "Market depth fell from \$700M+ in October to ~\$535M,Thinner order books increase price vulnerability to large trades,Reduced spot and institutional buying weakened market structure,Lower liquidity allows sell-offs to cascade more easily,Heightened volatility across all trading pairs" \
  --position '{"left": "8%", "top": "25%"}' \
  --size '{"width": "84%", "height": "55%"}' \
  --font-size 18 \
  --json

# Add accent callout for liquidity metric
uv run tools/ppt_add_shape.py \
  --file "bitcoin_market_report_nov2024.pptx" \
  --slide 3 \
  --shape "rectangle" \
  --position '{"left": "60%", "top": "72%"}' \
  --size '{"width": "34%", "height": "15%"}' \
  --fill-color "#ED7D31" \
  --json

uv run tools/ppt_add_text_box.py \
  --file "bitcoin_market_report_nov2024.pptx" \
  --slide 3 \
  --text "Liquidity: \$700M → \$535M" \
  --position '{"left": "61%", "top": "75%"}' \
  --size '{"width": "32%", "height": "8%"}' \
  --font-size 18 \
  --color "#FFFFFF" \
  --json

# ═══════════════════════════════════════════════════════════════
# SLIDE 5: Key Driver #3 - Macroeconomic Uncertainty
# ═══════════════════════════════════════════════════════════════
uv run tools/ppt_add_slide.py \
  --file "bitcoin_market_report_nov2024.pptx" \
  --layout "Title and Content" \
  --json

uv run tools/ppt_set_title.py \
  --file "bitcoin_market_report_nov2024.pptx" \
  --slide 4 \
  --title "Key Driver #3: Macroeconomic Uncertainty" \
  --json

uv run tools/ppt_add_bullet_list.py \
  --file "bitcoin_market_report_nov2024.pptx" \
  --slide 4 \
  --items "Federal Reserve caution on interest rate cuts rattled investors,Inflation resilience and strong dollar pressure risk assets,Trade war fears from renewed US-China tensions,Capital pullback from high-risk markets accelerating,Bitcoin weakness amplified by broader risk-off sentiment" \
  --position '{"left": "8%", "top": "25%"}' \
  --size '{"width": "84%", "height": "60%"}' \
  --font-size 18 \
  --json

# ═══════════════════════════════════════════════════════════════
# SLIDE 6: Key Driver #4 - Institutional Buying Slowdown
# ═══════════════════════════════════════════════════════════════
uv run tools/ppt_add_slide.py \
  --file "bitcoin_market_report_nov2024.pptx" \
  --layout "Title and Content" \
  --json

uv run tools/ppt_set_title.py \
  --file "bitcoin_market_report_nov2024.pptx" \
  --slide 5 \
  --title "Key Driver #4: Institutional Buying Slowdown" \
  --json

uv run tools/ppt_add_bullet_list.py \
  --file "bitcoin_market_report_nov2024.pptx" \
  --slide 5 \
  --items "Net institutional purchases fell below daily mined supply,First occurrence in seven months signals demand weakness,Large players no longer absorbing new Bitcoin supply,Institutional cash reserves showing signs of depletion,Raises risk of deeper corrections without demand recovery" \
  --position '{"left": "8%", "top": "25%"}' \
  --size '{"width": "84%", "height": "55%"}' \
  --font-size 18 \
  --json

# Add accent callout for timeframe stat
uv run tools/ppt_add_shape.py \
  --file "bitcoin_market_report_nov2024.pptx" \
  --slide 5 \
  --shape "rectangle" \
  --position '{"left": "55%", "top": "72%"}' \
  --size '{"width": "39%", "height": "15%"}' \
  --fill-color "#ED7D31" \
  --json

uv run tools/ppt_add_text_box.py \
  --file "bitcoin_market_report_nov2024.pptx" \
  --slide 5 \
  --text "First time in 7 months" \
  --position '{"left": "56%", "top": "75%"}' \
  --size '{"width": "37%", "height": "8%"}' \
  --font-size 18 \
  --color "#FFFFFF" \
  --json

# ═══════════════════════════════════════════════════════════════
# SLIDE 7: Key Driver #5 - Technical Breakdowns & Sentiment
# ═══════════════════════════════════════════════════════════════
uv run tools/ppt_add_slide.py \
  --file "bitcoin_market_report_nov2024.pptx" \
  --layout "Title and Content" \
  --json

uv run tools/ppt_set_title.py \
  --file "bitcoin_market_report_nov2024.pptx" \
  --slide 6 \
  --title "Key Driver #5: Technical Breakdowns & Sentiment Shifts" \
  --json

uv run tools/ppt_add_bullet_list.py \
  --file "bitcoin_market_report_nov2024.pptx" \
  --slide 6 \
  --items "Breach of \$100K psychological support triggered retail panic,Many investors exited positions below their cost basis,Extreme fear levels in market sentiment indicators,Options and derivatives traders positioned for further downside,Technical weakness reinforcing negative sentiment loop" \
  --position '{"left": "8%", "top": "25%"}' \
  --size '{"width": "84%", "height": "55%"}' \
  --font-size 18 \
  --json

# Add accent callout for support level
uv run tools/ppt_add_shape.py \
  --file "bitcoin_market_report_nov2024.pptx" \
  --slide 6 \
  --shape "rectangle" \
  --position '{"left": "62%", "top": "72%"}' \
  --size '{"width": "32%", "height": "15%"}' \
  --fill-color "#ED7D31" \
  --json

uv run tools/ppt_add_text_box.py \
  --file "bitcoin_market_report_nov2024.pptx" \
  --slide 6 \
  --text "Support Break: \$100K" \
  --position '{"left": "63%", "top": "75%"}' \
  --size '{"width": "30%", "height": "8%"}' \
  --font-size 18 \
  --color "#FFFFFF" \
  --json

# ═══════════════════════════════════════════════════════════════
# SLIDE 8: Market Context & Historical Patterns (Two-Column Layout)
# ═══════════════════════════════════════════════════════════════
uv run tools/ppt_add_slide.py \
  --file "bitcoin_market_report_nov2024.pptx" \
  --layout "Title and Content" \
  --json

uv run tools/ppt_set_title.py \
  --file "bitcoin_market_report_nov2024.pptx" \
  --slide 7 \
  --title "Market Context & Historical Patterns" \
  --json

# Left column header
uv run tools/ppt_add_text_box.py \
  --file "bitcoin_market_report_nov2024.pptx" \
  --slide 7 \
  --text "Seasonal & Halving Cycles" \
  --position '{"left": "8%", "top": "22%"}' \
  --size '{"width": "40%", "height": "8%"}' \
  --font-size 22 \
  --color "#0070C0" \
  --json

# Left column bullets
uv run tools/ppt_add_bullet_list.py \
  --file "bitcoin_market_report_nov2024.pptx" \
  --slide 7 \
  --items "November historically strong for Bitcoin,Post-halving volatility expected,Mid-cycle dip may be healthy reset,Clearing leveraged positions,Potential consolidation phase" \
  --position '{"left": "8%", "top": "32%"}' \
  --size '{"width": "40%", "height": "55%"}' \
  --font-size 16 \
  --json

# Right column header
uv run tools/ppt_add_text_box.py \
  --file "bitcoin_market_report_nov2024.pptx" \
  --slide 7 \
  --text "On-Chain Data Signals" \
  --position '{"left": "52%", "top": "22%"}' \
  --size '{"width": "40%", "height": "8%"}' \
  --font-size 22 \
  --color "#0070C0" \
  --json

# Right column bullets
uv run tools/ppt_add_bullet_list.py \
  --file "bitcoin_market_report_nov2024.pptx" \
  --slide 7 \
  --items "Dormant wallets moving to exchanges,Capitulation by weaker hands,NOT mass exodus by long-term holders,ETF flows and inflows slowed,Traditional finance channels weakening" \
  --position '{"left": "52%", "top": "32%"}' \
  --size '{"width": "40%", "height": "55%"}' \
  --font-size 16 \
  --json

# ═══════════════════════════════════════════════════════════════
# SLIDE 9: Conclusion - Path Forward
# ═══════════════════════════════════════════════════════════════
uv run tools/ppt_add_slide.py \
  --file "bitcoin_market_report_nov2024.pptx" \
  --layout "Title and Content" \
  --json

uv run tools/ppt_set_title.py \
  --file "bitcoin_market_report_nov2024.pptx" \
  --slide 8 \
  --title "Conclusion: Path Forward" \
  --json

uv run tools/ppt_add_bullet_list.py \
  --file "bitcoin_market_report_nov2024.pptx" \
  --slide 8 \
  --items "Downturn driven by multiple intersecting factors—not isolated events,Macro headwinds + liquidity crunch + institutional slowdown,Technical breakdown at \$100K accelerated the correction,Market absorbing losses and clearing weak positions,Recovery requires renewed institutional inflows and stabilized sentiment" \
  --position '{"left": "8%", "top": "25%"}' \
  --size '{"width": "84%", "height": "60%"}' \
  --font-size 18 \
  --json

# ═══════════════════════════════════════════════════════════════
# FINALIZATION: Footer, Validation, and Accessibility Check
# ═══════════════════════════════════════════════════════════════

# Set professional footer with slide numbers
uv run tools/ppt_set_footer.py \
  --file "bitcoin_market_report_nov2024.pptx" \
  --text "Bitcoin Market Report • November 2024" \
  --show-number \
  --json

# Validate presentation structure
echo "🔍 Validating presentation structure..."
uv run tools/ppt_validate_presentation.py \
  --file "bitcoin_market_report_nov2024.pptx" \
  --json

# Check accessibility compliance (WCAG 2.1)
echo "♿ Checking accessibility compliance..."
uv run tools/ppt_check_accessibility.py \
  --file "bitcoin_market_report_nov2024.pptx" \
  --json

# ═══════════════════════════════════════════════════════════════
# COMPLETION SUMMARY
# ═══════════════════════════════════════════════════════════════
echo ""
echo "✅ ═══════════════════════════════════════════════════════════"
echo "✅ Bitcoin Market Report Presentation Created Successfully!"
echo "✅ ═══════════════════════════════════════════════════════════"
echo ""
echo "📊 File: bitcoin_market_report_nov2024.pptx"
echo "📄 Total Slides: 9"
echo ""
echo "📋 Slide Breakdown:"
echo "   1. Title Slide - Report Cover"
echo "   2. Executive Summary"
echo "   3. Key Driver #1: Panic Selling by Short-Term Holders"
echo "   4. Key Driver #2: Liquidity Crunch"
echo "   5. Key Driver #3: Macroeconomic Uncertainty"
echo "   6. Key Driver #4: Institutional Buying Slowdown"
echo "   7. Key Driver #5: Technical Breakdowns & Sentiment Shifts"
echo "   8. Market Context & Historical Patterns (Two-Column)"
echo "   9. Conclusion: Path Forward"
echo ""
echo "🎨 Design Features:"
echo "   • Corporate color palette (#0070C0 primary, #ED7D31 accent)"
echo "   • Accent callout boxes for key statistics (4 slides)"
echo "   • Professional footer with slide numbering"
echo "   • Two-column layout on slide 8 for comparative analysis"
echo "   • 18-20pt body text for optimal readability"
echo ""
echo "✅ Validation & Accessibility checks completed"
echo "═══════════════════════════════════════════════════════════"
