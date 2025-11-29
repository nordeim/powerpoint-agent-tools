#!/usr/bin/env python3
"""
PowerPoint JSON Adapter Tool v3.1.0
Validates and normalizes JSON outputs from presentation CLI tools.

This tool provides:
- JSON Schema validation
- Key alias normalization
- Fallback presentation version computation
- Consistent error formatting

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_json_adapter.py --schema schema.json --input raw.json
    uv run tools/ppt_json_adapter.py --schema schema.json --input raw.json --output normalized.json

Exit Codes:
    0: Success (valid JSON emitted)
    1: Usage error (invalid arguments)
    2: Validation error (schema validation failed)
    3: Transient error (file I/O issue, retryable)
    5: Internal error (unexpected failure)
"""

import sys
import os

# --- HYGIENE BLOCK START ---
# CRITICAL: Redirect stderr to /dev/null immediately.
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

import argparse
import json
import hashlib
import logging
from pathlib import Path
from typing import Dict, Any, Optional
from datetime import datetime

# Configure logging to null handler
logging.basicConfig(level=logging.CRITICAL)

# ============================================================================
# CONSTANTS
# ============================================================================

__version__ = "3.1.0"

# Alias mapping table for common drifted keys
ALIAS_MAP = {
    "slides_count": "slide_count",
    "slidesTotal": "slide_count",
    "slides_list": "slides",
    "probe_time": "probe_timestamp",
    "canWrite": "can_write",
    "canRead": "can_read",
    "maxImageSizeMB": "max_image_size_mb",
    "slideCount": "slide_count",
    "presentationVersion": "presentation_version",
    "toolVersion": "tool_version"
}

# ============================================================================
# IMPORTS
# ============================================================================

try:
    from jsonschema import validate, ValidationError
    JSONSCHEMA_AVAILABLE = True
except ImportError:
    JSONSCHEMA_AVAILABLE = False
    ValidationError = Exception


# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def make_error_response(
    error_code: str, 
    message: str, 
    details: Any = None, 
    retryable: bool = False,
    suggestion: Optional[str] = None
) -> Dict[str, Any]:
    """
    Create a standardized error response.
    
    Args:
        error_code: Error type code
        message: Human-readable error message
        details: Additional error details
        retryable: Whether the operation can be retried
        suggestion: Suggested fix
        
    Returns:
        Standardized error dict
    """
    response = {
        "status": "error",
        "error": message,
        "error_type": error_code,
        "retryable": retryable,
        "tool_version": __version__,
        "processed_at": datetime.utcnow().isoformat() + "Z"
    }
    
    if details is not None:
        response["details"] = details
        
    if suggestion:
        response["suggestion"] = suggestion
        
    return response


def load_json(path: Path) -> Dict[str, Any]:
    """
    Load JSON file with error handling.
    
    Args:
        path: Path to JSON file
        
    Returns:
        Parsed JSON data
        
    Raises:
        FileNotFoundError: If file doesn't exist
        json.JSONDecodeError: If JSON is invalid
    """
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def map_aliases(obj: Any) -> Any:
    """
    Recursively map aliased keys to canonical keys.
    
    Args:
        obj: Object to process (dict, list, or primitive)
        
    Returns:
        Object with aliased keys replaced
    """
    if isinstance(obj, dict):
        new = {}
        for k, v in obj.items():
            canonical = ALIAS_MAP.get(k, k)
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


def compute_presentation_version(info_obj: Dict[str, Any]) -> Optional[str]:
    """
    Compute a stable presentation_version if missing.
    
    Uses slide ids and counts to produce a deterministic hash.
    
    Args:
        info_obj: Presentation info dict
        
    Returns:
        SHA-256 hash string or None if computation fails
    """
    try:
        slides = info_obj.get("slides", [])
        slide_ids = []
        for s in slides:
            slide_id = s.get("id") or s.get("index") or ""
            slide_ids.append(str(slide_id))
        
        slide_ids_str = ",".join(slide_ids)
        file_path = info_obj.get("file", "")
        slide_count = info_obj.get("slide_count", len(slides))
        
        base = f"{file_path}-{slide_count}-{slide_ids_str}"
        return hashlib.sha256(base.encode("utf-8")).hexdigest()[:16]
    except Exception:
        return None


def normalize_json(raw: Dict[str, Any], schema: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """
    Normalize JSON data by mapping aliases and adding missing fields.
    
    Args:
        raw: Raw JSON data
        schema: Optional schema for context-aware normalization
        
    Returns:
        Normalized JSON data
    """
    normalized = map_aliases(raw)
    
    # Add presentation_version if missing and this appears to be ppt_get_info output
    if "presentation_version" not in normalized:
        if "slides" in normalized or "slide_count" in normalized:
            pv = compute_presentation_version(normalized)
            if pv:
                normalized["presentation_version"] = pv
    
    # Ensure status field exists
    if "status" not in normalized:
        if "error" in normalized:
            normalized["status"] = "error"
        else:
            normalized["status"] = "success"
    
    return normalized


# ============================================================================
# MAIN LOGIC
# ============================================================================

def process_json(
    schema_path: Path, 
    input_path: Path, 
    output_path: Optional[Path] = None
) -> Dict[str, Any]:
    """
    Process and validate JSON input against schema.
    
    Args:
        schema_path: Path to JSON Schema file
        input_path: Path to input JSON file
        output_path: Optional path to write normalized output
        
    Returns:
        Normalized and validated JSON data
        
    Raises:
        Various exceptions on failure
    """
    # Load schema
    schema = load_json(schema_path)
    
    # Load input
    raw = load_json(input_path)
    
    # Normalize
    normalized = normalize_json(raw, schema)
    
    # Validate if jsonschema available
    if JSONSCHEMA_AVAILABLE:
        validate(instance=normalized, schema=schema)
    
    # Add processing metadata
    normalized["_adapter"] = {
        "processed_at": datetime.utcnow().isoformat() + "Z",
        "adapter_version": __version__,
        "schema_file": str(schema_path.name)
    }
    
    # Write to output file if specified
    if output_path:
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(normalized, f, indent=2)
    
    return normalized


# ============================================================================
# CLI ENTRY POINT
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description=f"Validate and normalize JSON outputs - v{__version__}",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
    # Validate and normalize JSON
    uv run tools/ppt_json_adapter.py --schema schema.json --input raw.json
    
    # Write normalized output to file
    uv run tools/ppt_json_adapter.py --schema schema.json --input raw.json --output normalized.json

Features:
    - Maps aliased keys to canonical keys
    - Computes presentation_version if missing
    - Validates against JSON Schema
    - Ensures consistent output structure

Version: """ + __version__
    )
    
    parser.add_argument(
        "--schema", 
        required=True, 
        type=Path,
        help="Path to JSON Schema file"
    )
    
    parser.add_argument(
        "--input", 
        required=True, 
        type=Path,
        help="Path to raw JSON input file"
    )
    
    parser.add_argument(
        "--output",
        type=Path,
        help="Optional path to write normalized JSON output"
    )
    
    args = parser.parse_args()
    
    # Load schema
    try:
        schema = load_json(args.schema)
    except FileNotFoundError:
        error = make_error_response(
            "FileNotFoundError",
            f"Schema file not found: {args.schema}",
            retryable=False,
            suggestion="Verify the schema file path"
        )
        sys.stdout.write(json.dumps(error, indent=2) + "\n")
        sys.exit(1)
    except json.JSONDecodeError as e:
        error = make_error_response(
            "JSONDecodeError",
            f"Invalid JSON in schema file: {e}",
            details={"file": str(args.schema), "line": e.lineno},
            retryable=False
        )
        sys.stdout.write(json.dumps(error, indent=2) + "\n")
        sys.exit(5)

    # Load input
    try:
        raw = load_json(args.input)
    except FileNotFoundError:
        error = make_error_response(
            "FileNotFoundError",
            f"Input file not found: {args.input}",
            retryable=True,
            suggestion="Verify the input file path"
        )
        sys.stdout.write(json.dumps(error, indent=2) + "\n")
        sys.exit(3)
    except json.JSONDecodeError as e:
        error = make_error_response(
            "JSONDecodeError",
            f"Invalid JSON in input file: {e}",
            details={"file": str(args.input), "line": e.lineno},
            retryable=True,
            suggestion="Check the input file for JSON syntax errors"
        )
        sys.stdout.write(json.dumps(error, indent=2) + "\n")
        sys.exit(3)

    # Normalize
    normalized = normalize_json(raw, schema)

    # Validate
    if JSONSCHEMA_AVAILABLE:
        try:
            validate(instance=normalized, schema=schema)
        except ValidationError as ve:
            error = make_error_response(
                "ValidationError",
                str(ve.message),
                details={
                    "schema_path": list(ve.schema_path) if ve.schema_path else None,
                    "instance_path": list(ve.absolute_path) if ve.absolute_path else None
                },
                retryable=False,
                suggestion="Fix the input data to match the schema requirements"
            )
            sys.stdout.write(json.dumps(error, indent=2) + "\n")
            sys.exit(2)
    else:
        normalized["_warning"] = "jsonschema not installed - validation skipped"

    # Add success metadata
    normalized["_adapter"] = {
        "status": "normalized",
        "processed_at": datetime.utcnow().isoformat() + "Z",
        "adapter_version": __version__,
        "schema_file": args.schema.name
    }

    # Write to output file if specified
    if args.output:
        try:
            with open(args.output, "w", encoding="utf-8") as f:
                json.dump(normalized, f, indent=2)
        except IOError as e:
            error = make_error_response(
                "IOError",
                f"Failed to write output file: {e}",
                retryable=True
            )
            sys.stdout.write(json.dumps(error, indent=2) + "\n")
            sys.exit(3)

    # Emit normalized JSON to stdout
    sys.stdout.write(json.dumps(normalized, indent=2) + "\n")
    sys.exit(0)


if __name__ == "__main__":
    main()
