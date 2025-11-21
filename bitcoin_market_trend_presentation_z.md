

# Bitcoin Market Analysis Presentation Script

```bash
#!/bin/bash

# Create a new presentation with a professional layout
uv run tools/ppt_create_new.py --output bitcoin_market_analysis.pptx --layout "Title Slide" --json

# Add a title slide
uv run tools/ppt_set_title.py --file bitcoin_market_analysis.pptx --slide 0 --title "Bitcoin Market Analysis" --subtitle "Understanding the Recent Downturn" --json

# Add Executive Summary slide
uv run tools/ppt_add_slide.py --file bitcoin_market_analysis.pptx --layout "Title and Content" --index 1 --title "Executive Summary" --json
uv run tools/ppt_add_text_box.py --file bitcoin_market_analysis.pptx --slide 1 --text "Bitcoin price dropped from $120,000 to below $95,000 due to multiple factors:" --position '{"left": "10%", "top": "20%"}' --size '{"width": "80%", "height": "10%"}' --font-name "Arial" --font-size 18 --color "#0A0A0A" --json
uv run tools/ppt_add_bullet_list.py --file bitcoin_market_analysis.pptx --slide 1 --items "Panic selling by short-term holders,Liquidity crunch in markets,Macroeconomic uncertainty,Institutional buying slowdown,Technical breakdowns,Sentiment shifts to fear" --position '{"left": "10%", "top": "35%"}' --size '{"width": "80%", "height": "50%"}' --json

# Add Key Causes Overview slide
uv run tools/ppt_add_slide.py --file bitcoin_market_analysis.pptx --layout "Title and Content" --index 2 --title "Key Causes of Bitcoin Downturn" --json
uv run tools/ppt_add_shape.py --file bitcoin_market_analysis.pptx --slide 2 --shape rectangle --position '{"left": "10%", "top": "20%"}' --size '{"width": "80%", "height": "5%"}' --fill-color "#2A9D8F" --json
uv run tools/ppt_add_text_box.py --file bitcoin_market_analysis.pptx --slide 2 --text "Primary Market Drivers" --position '{"left": "10%", "top": "20%"}' --size '{"width": "80%", "height": "5%"}' --font-name "Arial" --font-size 16 --color "#FFFFFF" --bold --json
uv run tools/ppt_add_bullet_list.py --file bitcoin_market_analysis.pptx --slide 2 --items "Panic Selling by Short-Term Holders,Liquidity Crunch,Macroeconomic Uncertainty,Institutional Buying Slowdown,Technical Breakdowns and Sentiment Shifts" --position '{"left": "10%", "top": "30%"}' --size '{"width": "80%", "height": "50%"}' --json

# Add Panic Selling slide
uv run tools/ppt_add_slide.py --file bitcoin_market_analysis.pptx --layout "Title and Content" --index 3 --title "Panic Selling by Short-Term Holders" --json
uv run tools/ppt_add_text_box.py --file bitcoin_market_analysis.pptx --slide 3 --text "Price drop accelerated by short-term holders selling at a loss" --position '{"left": "10%", "top": "20%"}' --size '{"width": "80%", "height": "10%"}' --font-name "Arial" --font-size 18 --color "#0A0A0A" --json
uv run tools/ppt_add_bullet_list.py --file bitcoin_market_analysis.pptx --slide 3 --items "Price fell from $120,000 to below $95,000,Triggered forced liquidations,Rapid deleveraging occurred,Long-term holders took some profits,No widespread distribution yet,Not typical of bear market tops" --position '{"left": "10%", "top": "35%"}' --size '{"width": "80%", "height": "50%"}' --json

# Add Liquidity Crunch slide with chart
uv run tools/ppt_add_slide.py --file bitcoin_market_analysis.pptx --layout "Title and Content" --index 4 --title "Liquidity Crunch" --json
uv run tools/ppt_add_text_box.py --file bitcoin_market_analysis.pptx --slide 4 --text "Market depth has thinned notably in recent weeks" --position '{"left": "10%", "top": "20%"}' --size '{"width": "80%", "height": "10%"}' --font-name "Arial" --font-size 18 --color "#0A0A0A" --json

# Create a temporary JSON file for the liquidity chart data
cat > liquidity_data.json << EOF
{
   "categories": ["October", "November"],
   "series": [
      {
         "name": "Market Depth (Millions)",
         "values": [700, 535]
      }
   ]
}
EOF

# Add the liquidity chart
uv run tools/ppt_add_chart.py --file bitcoin_market_analysis.pptx --slide 4 --chart-type "column" --data liquidity_data.json --position '{"left": "10%", "top": "35%"}' --size '{"width": "45%", "height": "40%"}' --title "Market Depth Decline" --json
uv run tools/ppt_add_bullet_list.py --file bitcoin_market_analysis.pptx --slide 4 --items "Market depth fell from $700M to $535M,Thinner order books increase volatility,Reduced spot and institutional buying,Sell-offs cascade more easily" --position '{"left": "55%", "top": "35%"}' --size '{"width": "40%", "height": "40%"}' --json

# Add Macroeconomic Uncertainty slide
uv run tools/ppt_add_slide.py --file bitcoin_market_analysis.pptx --layout "Title and Content" --index 5 --title "Macroeconomic Uncertainty" --json
uv run tools/ppt_add_text_box.py --file bitcoin_market_analysis.pptx --slide 5 --text "Federal Reserve's caution rattles risk asset investors" --position '{"left": "10%", "top": "20%"}' --size '{"width": "80%", "height": "10%"}' --font-name "Arial" --font-size 18 --color "#0A0A0A" --json
uv run tools/ppt_add_bullet_list.py --file bitcoin_market_analysis.pptx --slide 5 --items "Fed cautious on interest rate cuts,Capital pulled from high-risk markets,Inflation resilience impacts Bitcoin,Strong dollar exacerbates weakness,Trade war fears suppress risk appetite,US-China tensions add uncertainty" --position '{"left": "10%", "top": "35%"}' --size '{"width": "80%", "height": "50%"}' --json

# Add Institutional Buying Slowdown slide
uv run tools/ppt_add_slide.py --file bitcoin_market_analysis.pptx --layout "Title and Content" --index 6 --title "Institutional Buying Slowdown" --json
uv run tools/ppt_add_text_box.py --file bitcoin_market_analysis.pptx --slide 6 --text "Net institutional Bitcoin purchases fell below daily mined supply" --position '{"left": "10%", "top": "20%"}' --size '{"width": "80%", "height": "10%"}' --font-name "Arial" --font-size 18 --color "#0A0A0A" --json
uv run tools/ppt_add_bullet_list.py --file bitcoin_market_analysis.pptx --slide 6 --items "First time in seven months,Large players not absorbing new supply,Raises risk of deeper corrections,Institutional cash reserves may dwindle,ETF flows have slowed significantly,Traditional finance inflows reduced" --position '{"left": "10%", "top": "35%"}' --size '{"width": "80%", "height": "50%"}' --json

# Add Technical Breakdowns slide
uv run tools/ppt_add_slide.py --file bitcoin_market_analysis.pptx --layout "Title and Content" --index 7 --title "Technical Breakdowns and Sentiment Shifts" --json
uv run tools/ppt_add_text_box.py --file bitcoin_market_analysis.pptx --slide 7 --text "Breach of $100,000 support triggered retail panic" --position '{"left": "10%", "top": "20%"}' --size '{"width": "80%", "height": "10%"}' --font-name "Arial" --font-size 18 --color "#0A0A0A" --json
uv run tools/ppt_add_bullet_list.py --file bitcoin_market_analysis.pptx --slide 7 --items "$100,000 was psychologically significant,Retail panic led to quick exits,Many sold below their cost basis,Options traders positioned for downside,Derivatives traders expecting further drops,Sentiment indicators show extreme fear" --position '{"left": "10%", "top": "35%"}' --size '{"width": "80%", "height": "50%"}" --json

# Add Additional Observations slide
uv run tools/ppt_add_slide.py --file bitcoin_market_analysis.pptx --layout "Title and Content" --index 8 --title "Additional Observations" --json
uv run tools/ppt_add_shape.py --file bitcoin_market_analysis.pptx --slide 8 --shape rectangle --position '{"left": "10%", "top": "20%"}' --size '{"width": "35%", "height": "5%"}' --fill-color "#2A9D8F" --json
uv run tools/ppt_add_text_box.py --file bitcoin_market_analysis.pptx --slide 8 --text "Seasonal Patterns" --position '{"left": "10%", "top": "20%"}' --size '{"width": "35%", "height": "5%"}' --font-name "Arial" --font-size 16 --color "#FFFFFF" --bold --json
uv run tools/ppt_add_bullet_list.py --file bitcoin_market_analysis.pptx --slide 8 --items "November historically strong for Bitcoin,Post-halving volatility common,Mid-cycle dip before advances,Current correction may be healthy reset,Clearing leveraged positions,Setting up consolidation phase" --position '{"left": "10%", "top": "30%"}' --size '{"width": "35%", "height": "45%"}" --json
uv run tools/ppt_add_shape.py --file bitcoin_market_analysis.pptx --slide 8 --shape rectangle --position '{"left": "55%", "top": "20%"}' --size '{"width": "35%", "height": "5%"}' --fill-color "#E9C46A" --json
uv run tools/ppt_add_text_box.py --file bitcoin_market_analysis.pptx --slide 8 --text "On-Chain Metrics" --position '{"left": "55%", "top": "20%"}' --size '{"width": "35%", "height": "5%"}' --font-name "Arial" --font-size 16 --color "#0A0A0A" --bold --json
uv run tools/ppt_add_bullet_list.py --file bitcoin_market_analysis.pptx --slide 8 --items "Increase in coins from dormant wallets,Coins moving to exchanges,Indicates capitulation by weaker hands,No mass exodus by foundational investors,ETF flows slowed significantly,Traditional finance inflows reduced" --position '{"left": "55%", "top": "30%"}' --size '{"width": "35%", "height": "45%"}" --json

# Add Conclusion slide
uv run tools/ppt_add_slide.py --file bitcoin_market_analysis.pptx --layout "Title and Content" --index 9 --title "Conclusion" --json
uv run tools/ppt_add_text_box.py --file bitcoin_market_analysis.pptx --slide 9 --text "Bitcoin downturn results from multiple intersecting factors" --position '{"left": "10%", "top": "20%"}' --size '{"width": "80%", "height": "10%"}' --font-name "Arial" --font-size 18 --color "#0A0A0A" --json
uv run tools/ppt_add_bullet_list.py --file bitcoin_market_analysis.pptx --slide 9 --items "Macro headwinds affecting risk assets,Liquidity crunch increasing volatility,Waning institutional demand,Technical support breakdowns,Market sentiment shifted to fear" --position '{"left": "10%", "top": "35%"}' --size '{"width": "80%", "height": "40%"}" --json
uv run tools/ppt_add_text_box.py --file bitcoin_market_analysis.pptx --slide 9 --text "Renewed inflows and stabilization in global risk sentiment will be critical for sustained recovery" --position '{"left": "10%", "top": "80%"}' --size '{"width": "80%", "height": "10%"}' --font-name "Arial" --font-size 16 --color "#2A9D8F" --bold --json

# Add References slide
uv run tools/ppt_add_slide.py --file bitcoin_market_analysis.pptx --layout "Title and Content" --index 10 --title "References" --json
uv run tools/ppt_add_text_box.py --file bitcoin_market_analysis.pptx --slide 10 --text "[1] Yahoo Finance: Real reason behind Bitcoin's drop" --position '{"left": "10%", "top": "20%"}' --size '{"width": "80%", "height": "7%"}' --font-name "Arial" --font-size 14 --color "#0A0A0A" --json
uv run tools/ppt_add_text_box.py --file bitcoin_market_analysis.pptx --slide 10 --text "[2] CNN: Bitcoin price crypto stocks" --position '{"left": "10%", "top": "27%"}' --size '{"width": "80%", "height": "7%"}' --font-name "Arial" --font-size 14 --color "#0A0A0A" --json
uv run tools/ppt_add_text_box.py --file bitcoin_market_analysis.pptx --slide 10 --text "[3] Business Insider: Why Bitcoin is falling" --position '{"left": "10%", "top": "34%"}' --size '{"width": "80%", "height": "7%"}' --font-name "Arial" --font-size 14 --color "#0A0A0A" --json
uv run tools/ppt_add_text_box.py --file bitcoin_market_analysis.pptx --slide 10 --text "[4] Ki-Ecke: Bitcoin price outlook November 2025" --position '{"left": "10%", "top": "41%"}' --size '{"width": "80%", "height": "7%"}' --font-name "Arial" --font-size 14 --color "#0A0A0A" --json
uv run tools/ppt_add_text_box.py --file bitcoin_market_analysis.pptx --slide 10 --text "[5] The Motley Fool: Bitcoin falls below $95,000" --position '{"left": "10%", "top": "48%"}' --size '{"width": "80%", "height": "7%"}' --font-name "Arial" --font-size 14 --color "#0A0A0A" --json
uv run tools/ppt_add_text_box.py --file bitcoin_market_analysis.pptx --slide 10 --text "[6] 99Bitcoins: Saylor buys dip as daily mined BTC exceeds ETF" --position '{"left": "10%", "top": "55%"}' --size '{"width": "80%", "height": "7%"}' --font-name "Arial" --font-size 14 --color "#0A0A0A" --json
uv run tools/ppt_add_text_box.py --file bitcoin_market_analysis.pptx --slide 10 --text "[7] CoinDesk: Is November the new October?" --position '{"left": "10%", "top": "62%"}' --size '{"width": "80%", "height": "7%"}' --font-name "Arial" --font-size 14 --color "#0A0A0A" --json
uv run tools/ppt_add_text_box.py --file bitcoin_market_analysis.pptx --slide 10 --text "[8] FinanceFeeds: Bitcoin technical analysis report" --position '{"left": "10%", "top": "69%"}' --size '{"width": "80%", "height": "7%"}' --font-name "Arial" --font-size 14 --color "#0A0A0A" --json

# Set consistent footer across all slides
uv run tools/ppt_set_footer.py --file bitcoin_market_analysis.pptx --text "Bitcoin Market Analysis • November 2025" --show-number --show-date --json

# Validate the presentation
uv run tools/ppt_validate_presentation.py --file bitcoin_market_analysis.pptx --json

# Check accessibility
uv run tools/ppt_check_accessibility.py --file bitcoin_market_analysis.pptx --json

# Clean up temporary files
rm liquidity_data.json

echo "Bitcoin Market Analysis presentation created successfully!"
```

This script creates a comprehensive 11-slide presentation analyzing the Bitcoin market downturn with:

1. **Title Slide**: Professional introduction with title and subtitle
2. **Executive Summary**: Key findings at a glance
3. **Key Causes Overview**: Introduction to the main factors
4. **Panic Selling**: Analysis of short-term holder behavior
5. **Liquidity Crunch**: Visual chart showing market depth decline
6. **Macroeconomic Uncertainty**: Federal Reserve and global factors
7. **Institutional Buying Slowdown**: Data on institutional behavior
8. **Technical Breakdowns**: Price support and sentiment analysis
9. **Additional Observations**: Seasonal patterns and on-chain data in a two-column layout
10. **Conclusion**: Summary and outlook
11. **References**: Source citations

The design follows professional standards with:
- Consistent color scheme using the Data palette
- Proper text sizing and hierarchy
- 6×6 rule for bullet points
- Visual elements (chart, colored headers)
- Consistent footer with slide numbers
- Accessibility and validation checks

The presentation provides a thorough analysis of the Bitcoin market downturn while maintaining a clean, professional appearance suitable for financial reporting.
