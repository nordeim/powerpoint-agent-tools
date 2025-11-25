#!/usr/bin/env python3
"""
ppt_json_adapter.py

Validates and normalizes JSON outputs from presentation CLI tools.
Usage:
  python ppt_json_adapter.py --schema ppt_get_info.schema.json --input raw.json

Behavior:
- Validates input JSON against provided schema.
- Maps common alias keys to canonical keys.
- Emits normalized JSON to stdout.
- On validation failure, emits structured error JSON and exits non-zero.
"""

import argparse
import json
import sys
import hashlib
from jsonschema import validate, ValidationError

# Alias mapping table for common drifted keys
ALIAS_MAP = {
    "slides_count": "slide_count",
    "slidesTotal": "slide_count",
    "slides_list": "slides",
    "probe_time": "probe_timestamp",
    "canWrite": "can_write",
    "canRead": "can_read",
    "maxImageSizeMB": "max_image_size_mb"
}

ERROR_TEMPLATE = {
    "error": {
        "error_code": None,
        "message": None,
        "details": None,
        "retryable": False
    }
}

def load_json(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

def map_aliases(obj):
    if isinstance(obj, dict):
        new = {}
        for k, v in obj.items():
            canonical = ALIAS_MAP.get(k, k)
            # recursively map nested dicts and lists
            if isinstance(v, dict):
                new[canonical] = map_aliases(v)
            elif isinstance(v, list):
                new[canonical] = [map_aliases(i) for i in v]
            else:
                new[canonical] = v
        return new
    elif isinstance(obj, list):
        return [map_aliases(i) for i in obj]
    else:
        return obj

def compute_presentation_version(info_obj):
    """
    Compute a stable presentation_version if missing.
    Uses slide ids and counts to produce a deterministic hash.
    """
    try:
        slides = info_obj.get("slides", [])
        slide_ids = ",".join([str(s.get("id", s.get("index", ""))) for s in slides])
        base = f"{info_obj.get('file','')}-{info_obj.get('slide_count',len(slides))}-{slide_ids}"
        return hashlib.sha256(base.encode("utf-8")).hexdigest()
    except Exception:
        return None

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--schema", required=True, help="Path to JSON Schema")
    parser.add_argument("--input", required=True, help="Path to raw JSON input")
    args = parser.parse_args()

    try:
        schema = load_json(args.schema)
    except Exception as e:
        print(json.dumps({**ERROR_TEMPLATE, "error": {"error_code": "SCHEMA_LOAD_ERROR", "message": str(e), "details": None, "retryable": False}}))
        sys.exit(5)

    try:
        raw = load_json(args.input)
    except Exception as e:
        print(json.dumps({**ERROR_TEMPLATE, "error": {"error_code": "INPUT_LOAD_ERROR", "message": str(e), "details": None, "retryable": True}}))
        sys.exit(3)

    # Map aliases
    normalized = map_aliases(raw)

    # If presentation_version missing for get_info, compute a best-effort version
    if "presentation_version" not in normalized and schema.get("title","").lower().find("ppt_get_info") != -1:
        pv = compute_presentation_version(normalized)
        if pv:
            normalized["presentation_version"] = pv

    # Validate
    try:
        validate(instance=normalized, schema=schema)
    except ValidationError as ve:
        err = {
            "error": {
                "error_code": "SCHEMA_VALIDATION_ERROR",
                "message": str(ve.message),
                "details": ve.schema_path,
                "retryable": False
            }
        }
        print(json.dumps(err))
        sys.exit(2)

    # Emit normalized JSON
    print(json.dumps(normalized, indent=2))
    sys.exit(0)

if __name__ == "__main__":
    main()
