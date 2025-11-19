# 1. Verify core module imports successfully
python3 -c "from core.powerpoint_agent_core import PowerPointAgent; print('✓ Core module imports successfully')"

# 2. Test tool directly
python3 tools/ppt_create_new.py --output /tmp/test.pptx --slides 1 --json

# 3. Run full test suite
python3 test_basic_tools.py -v

---

# Run all P1 tests
python3 test_p1_tools.py -v

# Or with pytest
pytest test_p1_tools.py -v

# Run specific test
pytest test_p1_tools.py::TestP1Tools::test_add_chart_column -v

# Run with output
pytest test_p1_tools.py -v -s

---

# Test template creation
uv python ppt_create_from_template.py --template template.pptx --output test.pptx --slides 5 --json

# Test structure creation
uv python ppt_create_from_structure.py --structure structure.json --output test.pptx --json

# Test text formatting
uv python ppt_format_text.py --file test.pptx --slide 0 --shape 0 --bold --font-size 24 --json

# Test image replacement
uv python ppt_replace_image.py --file test.pptx --slide 0 --old-image "logo" --new-image new.png --json

# Test PDF export (requires LibreOffice)
uv python ppt_export_pdf.py --file test.pptx --output test.pdf --json
