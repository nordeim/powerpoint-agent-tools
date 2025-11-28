#!/usr/bin/env python3
"""
PowerPoint JSON Adapter Tool v3.1.0
Validates and normalizes JSON outputs from presentation CLI tools.

This utility:
- Validates input JSON against provided schema
- Maps common alias keys to canonical keys
- Computes presentation_version if missing
- Emits normalized JSON to stdout
- Returns structured error JSON on validation failure

Usage:
    uv run tools/ppt_json_adapter.py --schema schemas/ppt_get_info.schema.json --input raw.json
    
Exit Codes:
    0: Success (valid, normalized JSON emitted)
    1: Usage error (invalid arguments)
    2: Schema validation error (input doesn't match schema)
    3: Input load error (file not found, invalid JSON)
    5: Schema load error (schema file not found, invalid schema)

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0
"""

import sys
import os

# --- HYGIENE BLOCK START ---
# CRITICAL: Redirect stderr to /dev/null immediately.
# This prevents libraries from printing non-JSON text.
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

import argparse
import json
import hashlib
import logging
from pathlib import Path
from typing import Dict, Any, List, Optional
from datetime import datetime

# Configure logging to null handler
logging.basicConfig(level=logging.CRITICAL)

# ============================================================================
# CONSTANTS
# ============================================================================

__version__ = "3.1.0"

# Alias mapping table for common drifted keys
ALIAS_MAP = {
    # Slide count variations
    "slides_count": "slide_count",
    "slidesTotal": "slide_count",
    "slideCount": "slide_count",
    "num_slides": "slide_count",
    
    # Slides list variations
    "slides_list": "slides",
    "slidesList": "slides",
    
    # Probe timestamp variations
    "probe_time": "probe_timestamp",
    "probeTime": "probe_timestamp",
    "probed_at": "probe_timestamp",
    
    # Permission variations
    "canWrite": "can_write",
    "can_edit": "can_write",
    "canRead": "can_read",
    
    # Size variations
    "maxImageSizeMB": "max_image_size_mb",
    "max_image_mb": "max_image_size_mb",
    
    # Version variations
    "pptx_version": "presentation_version",
    "file_version": "presentation_version",
    "version": "presentation_version"
}

# ============================================================================
# SAFE IMPORTS
# ============================================================================

JSONSCHEMA_AVAILABLE = False
validate = None
ValidationError = None

try:
    from jsonschema import validate as _validate, ValidationError as _ValidationError
    validate = _validate
    ValidationError = _ValidationError
    JSONSCHEMA_AVAILABLE = True
except ImportError:
    pass

# ============================================================================
# ERROR HELPERS
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
        error_code: Error type code (e.g., "SCHEMA_VALIDATION_ERROR")
        message: Human-readable error message
        details: Additional error details
        retryable: Whether the operation can be retried
        suggestion: Suggested fix action
        
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

# ============================================================================
# FILE LOADING
# ============================================================================

def load_json(path: Path) -> Dict[str, Any]:
    """
    Load and parse a JSON file.
    
    Args:
        path: Path to JSON file
        
    Returns:
        Parsed JSON as dict
        
    Raises:
        FileNotFoundError: If file doesn't exist
        json.JSONDecodeError: If file isn't valid JSON
    """
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

# ============================================================================
# ALIAS MAPPING
# ============================================================================

def map_aliases(obj: Any) -> Any:
    """
    Recursively map aliased keys to canonical keys.
    
    Args:
        obj: JSON object (dict, list, or primitive)
        
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

# ============================================================================
# VERSION COMPUTATION
# ============================================================================

def compute_presentation_version(info_obj: Dict[str, Any]) -> Optional[str]:
    """
    Compute a stable presentation_version if missing.
    
    Uses file path, slide count, and slide IDs to produce a deterministic hash.
    
    Args:
        info_obj: Normalized info object
        
    Returns:
        SHA-256 hash string (first 16 chars) or None on failure
    """
    try:
        slides = info_obj.get("slides", [])
        
        # Build stable identifier from slide data
        slide_identifiers = []
        for s in slides:
            # Prefer 'id' over 'index' for stability
            slide_id = s.get("id")
            if slide_id is None:
                slide_id = s.get("index", "")
            slide_identifiers.append(str(slide_id))
        
        slide_ids_str = ",".join(slide_identifiers)
        file_path = info_obj.get("file", "")
        slide_count = info_obj.get("slide_count", len(slides))
        
        base = f"{file_path}-{slide_count}-{slide_ids_str}"
        full_hash = hashlib.sha256(base.encode("utf-8")).hexdigest()
        
        # Return first 16 characters per project convention
        return full_hash[:16]
        
    except Exception:
        return None

# ============================================================================
# MAIN LOGIC
# ============================================================================

def process_json(
    schema_path: Path, 
    input_path: Path
) -> Dict[str, Any]:
    """
    Validate and normalize JSON input against schema.
    
    Args:
        schema_path: Path to JSON Schema file
        input_path: Path to raw JSON input
        
    Returns:
        Normalized and validated JSON object
        
    Raises:
        Various exceptions for load/validation failures
    """
    # Load schema
    schema = load_json(schema_path)
    
    # Load input
    raw = load_json(input_path)
    
    # Map aliases to canonical keys
    normalized = map_aliases(raw)
    
    # Add presentation_version if missing and this looks like a get_info output
    if "presentation_version" not in normalized:
        schema_title = schema.get("title", "").lower()
        # Check if this is likely a presentation info schema
        if "get_info" in schema_title or "slides" in normalized:
            pv = compute_presentation_version(normalized)
            if pv:
                normalized["presentation_version"] = pv
    
    # Validate against schema
    if JSONSCHEMA_AVAILABLE and validate:
        validate(instance=normalized, schema=schema)
    
    # Add processing metadata
    normalized["_adapter_metadata"] = {
        "tool_version": __version__,
        "processed_at": datetime.utcnow().isoformat() + "Z",
        "schema_used": str(schema_path),
        "aliases_mapped": True,
        "validated": JSONSCHEMA_AVAILABLE
    }
    
    return normalized

# ============================================================================
# CLI INTERFACE
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description=f"Validate and normalize PowerPoint tool JSON output (v{__version__})",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
    # Validate get_info output against schema
    uv run tools/ppt_json_adapter.py --schema schemas/ppt_get_info.schema.json --input raw.json
    
    # Validate probe output
    uv run tools/ppt_json_adapter.py --schema schemas/capability_probe.schema.json --input probe_result.json

Exit Codes:
    0: Success - normalized JSON emitted to stdout
    2: Schema validation error
    3: Input load error (retryable if transient)
    5: Schema load error
        """
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
    
    args = parser.parse_args()
    
    # Validate jsonschema is available
    if not JSONSCHEMA_AVAILABLE:
        error = make_error_response(
            "DEPENDENCY_ERROR",
            "jsonschema library is not installed",
            suggestion="Install with: pip install jsonschema"
        )
        sys.stdout.write(json.dumps(error, indent=2) + "\n")
        sys.exit(1)
    
    # Load schema
    try:
        schema = load_json(args.schema)
    except FileNotFoundError:
        error = make_error_response(
            "SCHEMA_NOT_FOUND",
            f"Schema file not found: {args.schema}",
            details={"path": str(args.schema)},
            suggestion="Verify the schema 
