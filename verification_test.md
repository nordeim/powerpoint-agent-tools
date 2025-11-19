# 1. Verify core module imports successfully
python3 -c "from core.powerpoint_agent_core import PowerPointAgent; print('✓ Core module imports successfully')"

# 2. Test tool directly
python3 tools/ppt_create_new.py --output /tmp/test.pptx --slides 1 --json

# 3. Run full test suite
python3 test_basic_tools.py -v
