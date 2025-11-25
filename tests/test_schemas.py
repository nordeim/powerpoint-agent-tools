import pytest
import json
import os
import sys
from jsonschema import validate, ValidationError

# Add tools directory to path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '../tools')))
from ppt_json_adapter import map_aliases, compute_presentation_version

SCHEMAS_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '../schemas'))

def load_schema(name):
    with open(os.path.join(SCHEMAS_DIR, name), 'r') as f:
        return json.load(f)

def test_ppt_get_info_schema_valid():
    schema = load_schema('ppt_get_info.schema.json')
    valid_data = {
        "tool_name": "ppt_get_info",
        "tool_version": "1.0.0",
        "schema_version": "1.0.0",
        "file": "/tmp/test.pptx",
        "presentation_version": "sha256-dummy",
        "slide_count": 1,
        "slides": [
            {"index": 0, "id": "256", "layout": "Title Slide", "shape_count": 2}
        ]
    }
    validate(instance=valid_data, schema=schema)

def test_ppt_get_info_schema_invalid():
    schema = load_schema('ppt_get_info.schema.json')
    invalid_data = {
        "tool_name": "ppt_get_info",
        # Missing required fields
    }
    with pytest.raises(ValidationError):
        validate(instance=invalid_data, schema=schema)

def test_adapter_alias_mapping():
    raw = {
        "slidesTotal": 5,
        "slides_list": [{"index": 0}],
        "probe_time": "2025-01-01"
    }
    mapped = map_aliases(raw)
    assert mapped["slide_count"] == 5
    assert mapped["slides"] == [{"index": 0}]
    assert mapped["probe_timestamp"] == "2025-01-01"

def test_presentation_version_computation():
    info = {
        "file": "deck.pptx",
        "slide_count": 2,
        "slides": [{"id": "101"}, {"id": "102"}]
    }
    pv = compute_presentation_version(info)
    assert pv is not None
    assert len(pv) == 64  # SHA-256 hex digest length
