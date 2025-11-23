# core/strict_validator.py
import json
from pathlib import Path
from jsonschema import Draft202012Validator, FormatChecker

def validate_against_schema(payload: dict, schema_path: str) -> None:
    """
    Strictly validate payload against JSON Schema. Raises ValueError on failure.
    """
    schema = json.loads(Path(schema_path).read_text())
    validator = Draft202012Validator(schema, format_checker=FormatChecker())
    errors = sorted(validator.iter_errors(payload), key=lambda e: e.path)
    if errors:
        msgs = []
        for e in errors:
            loc = "/".join(str(p) for p in e.path)
            msgs.append(f"{loc or '<root>'}: {e.message}")
        raise ValueError("Strict schema validation failed:\n" + "\n".join(msgs))

