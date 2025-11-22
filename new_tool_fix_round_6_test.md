# Schema and CI plan for automated validation

## Overview

- **Goal:** Formalize output validation for ppt_capability_probe to ensure contract stability and detect regressions automatically.
- **Scope:** A single JSON Schema with versioning that validates both success and error outputs, plus CI jobs for GitHub Actions that run the probe, collect artifacts, and validate them.
- **Approach:** Use a top-level `oneOf` to distinguish success vs error payloads. Encode strict types, required fields, enumerations, and pattern constraints; include extensibility via additionalProperties.

---

## Schema design principles

- **Versioning:** Embed schema version in the file name and `$id`. Reference in output via `metadata.schema_version`.
- **Dual-mode validation:** Use `oneOf` for success and error structures with strong `required` sets.
- **Extensibility:** Allow non-breaking additions via `additionalProperties: true`, while enforcing core contract fields.
- **Strict typing:** Constrain primitives and arrays; enforce enums for `status`.
- **Auditability:** Validate UUID format, ISO datetime, positive durations, and counts.

---

## Schema file: capability_probe.v1.1.1.schema.json

```json
{
  "$schema": "https://json-schema.org/draft/2020-12/schema",
  "$id": "https://schemas.example.com/capability_probe/v1.1.1",
  "title": "PowerPoint capability probe output (v1.1.1)",
  "type": "object",
  "oneOf": [
    {
      "title": "Success payload",
      "type": "object",
      "required": ["status", "metadata", "slide_dimensions", "layouts", "theme", "capabilities", "warnings", "info"],
      "properties": {
        "status": { "const": "success" },
        "metadata": {
          "type": "object",
          "required": [
            "file",
            "probed_at",
            "tool_version",
            "schema_version",
            "operation_id",
            "deep_analysis",
            "atomic_verified",
            "duration_ms",
            "library_versions",
            "layout_count_total",
            "layout_count_analyzed"
          ],
          "properties": {
            "file": { "type": "string", "minLength": 1 },
            "probed_at": { "type": "string", "format": "date-time" },
            "tool_version": { "type": "string", "pattern": "^\\d+\\.\\d+\\.\\d+$" },
            "schema_version": { "type": "string", "pattern": "^capability_probe\\.v\\d+\\.\\d+\\.\\d+$" },
            "operation_id": { "type": "string", "pattern": "^[0-9a-fA-F-]{36}$" },
            "deep_analysis": { "type": "boolean" },
            "atomic_verified": { "type": "boolean" },
            "duration_ms": { "type": "integer", "minimum": 0 },
            "library_versions": {
              "type": "object",
              "required": ["python-pptx", "Pillow"],
              "properties": {
                "python-pptx": { "type": "string" },
                "Pillow": { "type": "string" }
              },
              "additionalProperties": true
            },
            "checksum": { "type": ["string", "null"], "pattern": "^[0-9a-f]{32}$" },
            "timeout_seconds": { "type": ["integer", "null"], "minimum": 0 },
            "layout_count_total": { "type": "integer", "minimum": 0 },
            "layout_count_analyzed": { "type": "integer", "minimum": 0 }
          },
          "additionalProperties": true
        },
        "slide_dimensions": {
          "type": "object",
          "required": [
            "width_inches", "height_inches",
            "width_emu", "height_emu",
            "width_pixels", "height_pixels",
            "aspect_ratio", "aspect_ratio_float", "dpi_estimate"
          ],
          "properties": {
            "width_inches": { "type": "number", "minimum": 0 },
            "height_inches": { "type": "number", "minimum": 0 },
            "width_emu": { "type": "integer", "minimum": 0 },
            "height_emu": { "type": "integer", "minimum": 0 },
            "width_pixels": { "type": "integer", "minimum": 0 },
            "height_pixels": { "type": "integer", "minimum": 0 },
            "aspect_ratio": { "type": "string", "minLength": 3 },
            "aspect_ratio_float": { "type": "number", "minimum": 0 },
            "dpi_estimate": { "type": "integer", "minimum": 1 }
          },
          "additionalProperties": false
        },
        "layouts": {
          "type": "array",
          "items": {
            "type": "object",
            "required": ["index", "original_index", "name", "placeholder_count", "master_index"],
            "properties": {
              "index": { "type": "integer", "minimum": 0 },
              "original_index": { "type": "integer", "minimum": 0 },
              "name": { "type": "string" },
              "placeholder_count": { "type": "integer", "minimum": 0 },
              "master_index": { "type": ["integer", "null"], "minimum": 0 },
              "placeholders": {
                "type": "array",
                "items": {
                  "type": "object",
                  "required": ["type", "type_code", "idx", "name", "position_source"],
                  "properties": {
                    "type": { "type": "string" },
                    "type_code": { "type": "integer" },
                    "idx": { "type": "integer", "minimum": 0 },
                    "name": { "type": "string" },
                    "position_source": { "enum": ["instantiated", "template", "error"] },
                    "position_inches": {
                      "type": "object",
                      "properties": {
                        "left": { "type": "number" },
                        "top": { "type": "number" }
                      },
                      "required": ["left", "top"],
                      "additionalProperties": false
                    },
                    "position_percent": {
                      "type": "object",
                      "properties": {
                        "left": { "type": "string", "pattern": "^\\d+(\\.\\d+)?%$" },
                        "top": { "type": "string", "pattern": "^\\d+(\\.\\d+)?%$" }
                      },
                      "required": ["left", "top"],
                      "additionalProperties": false
                    },
                    "position_emu": {
                      "type": "object",
                      "properties": {
                        "left": { "type": "integer", "minimum": 0 },
                        "top": { "type": "integer", "minimum": 0 }
                      },
                      "required": ["left", "top"],
                      "additionalProperties": false
                    },
                    "size_inches": {
                      "type": "object",
                      "properties": {
                        "width": { "type": "number", "minimum": 0 },
                        "height": { "type": "number", "minimum": 0 }
                      },
                      "required": ["width", "height"],
                      "additionalProperties": false
                    },
                    "size_percent": {
                      "type": "object",
                      "properties": {
                        "width": { "type": "string", "pattern": "^\\d+(\\.\\d+)?%$" },
                        "height": { "type": "string", "pattern": "^\\d+(\\.\\d+)?%$" }
                      },
                      "required": ["width", "height"],
                      "additionalProperties": false
                    },
                    "size_emu": {
                      "type": "object",
                      "properties": {
                        "width": { "type": "integer", "minimum": 0 },
                        "height": { "type": "integer", "minimum": 0 }
                      },
                      "required": ["width", "height"],
                      "additionalProperties": false
                    },
                    "error": { "type": "string" }
                  },
                  "additionalProperties": true
                }
              },
              "instantiation_complete": { "type": "boolean" },
              "placeholder_expected": { "type": "integer", "minimum": 0 },
              "placeholder_instantiated": { "type": "integer", "minimum": 0 },
              "placeholder_types": {
                "type": "array",
                "items": { "type": "string" }
              }
            },
            "additionalProperties": true
          }
        },
        "theme": {
          "type": "object",
          "required": ["colors", "fonts"],
          "properties": {
            "colors": { "type": "object", "additionalProperties": { "type": "string" } },
            "fonts": { "type": "object", "additionalProperties": { "type": "string" } },
            "per_master": {
              "type": "array",
              "items": {
                "type": "object",
                "required": ["master_index", "colors", "fonts"],
                "properties": {
                  "master_index": { "type": "integer", "minimum": 0 },
                  "colors": { "type": "object", "additionalProperties": { "type": "string" } },
                  "fonts": { "type": "object", "additionalProperties": { "type": "string" } }
                },
                "additionalProperties": false
              }
            }
          },
          "additionalProperties": true
        },
        "capabilities": {
          "type": "object",
          "required": [
            "has_footer_placeholders",
            "has_slide_number_placeholders",
            "has_date_placeholders",
            "layouts_with_footer",
            "layouts_with_slide_number",
            "layouts_with_date",
            "total_layouts",
            "total_master_slides",
            "per_master",
            "footer_support_mode",
            "slide_number_strategy",
            "recommendations",
            "analysis_complete"
          ],
          "properties": {
            "has_footer_placeholders": { "type": "boolean" },
            "has_slide_number_placeholders": { "type": "boolean" },
            "has_date_placeholders": { "type": "boolean" },
            "layouts_with_footer": { "type": "array", "items": { "type": "object" } },
            "layouts_with_slide_number": { "type": "array", "items": { "type": "object" } },
            "layouts_with_date": { "type": "array", "items": { "type": "object" } },
            "total_layouts": { "type": "integer", "minimum": 0 },
            "total_master_slides": { "type": "integer", "minimum": 0 },
            "per_master": {
              "type": "array",
              "items": {
                "type": "object",
                "required": [
                  "master_index",
                  "layout_count",
                  "has_footer_layouts",
                  "has_slide_number_layouts",
                  "has_date_layouts"
                ],
                "properties": {
                  "master_index": { "type": "integer", "minimum": 0 },
                  "layout_count": { "type": "integer", "minimum": 0 },
                  "has_footer_layouts": { "type": "integer", "minimum": 0 },
                  "has_slide_number_layouts": { "type": "integer", "minimum": 0 },
                  "has_date_layouts": { "type": "integer", "minimum": 0 }
                },
                "additionalProperties": false
              }
            },
            "footer_support_mode": { "enum": ["placeholder", "fallback_textbox"] },
            "slide_number_strategy": { "enum": ["placeholder", "textbox"] },
            "recommendations": { "type": "array", "items": { "type": "string" } },
            "analysis_complete": { "type": "boolean" }
          },
          "additionalProperties": true
        },
        "warnings": { "type": "array", "items": { "type": "string" } },
        "info": { "type": "array", "items": { "type": "string" } }
      },
      "additionalProperties": true
    },
    {
      "title": "Error payload",
      "type": "object",
      "required": ["status", "error", "error_type", "metadata", "warnings"],
      "properties": {
        "status": { "const": "error" },
        "error": { "type": "string" },
        "error_type": { "type": "string" },
        "metadata": {
          "type": "object",
          "required": ["file", "tool_version", "operation_id", "probed_at"],
          "properties": {
            "file": { "type": ["string", "null"] },
            "tool_version": { "type": "string", "pattern": "^\\d+\\.\\d+\\.\\d+$" },
            "operation_id": { "type": "string", "pattern": "^[0-9a-fA-F-]{36}$" },
            "probed_at": { "type": "string", "format": "date-time" }
          },
          "additionalProperties": true
        },
        "warnings": { "type": "array", "items": { "type": "string" } }
      },
      "additionalProperties": true
    }
  ]
}
```

Place this at: `schemas/capability_probe.v1.1.1.schema.json`.

---

## CI jobs for automated validation (GitHub Actions)

### 1) Python-based validation using jsonschema

```yaml
name: Validate capability probe outputs (python-jsonschema)

on:
  push:
    paths:
      - "tools/ppt_capability_probe.py"
      - "schemas/capability_probe.v1.1.1.schema.json"
  pull_request:

jobs:
  validate-probe-output:
    runs-on: ubuntu-latest
    steps:
      - name: Checkout
        uses: actions/checkout@v4

      - name: Set up Python
        uses: actions/setup-python@v5
        with:
          python-version: "3.11"

      - name: Install deps
        run: |
          python -m pip install --upgrade pip
          pip install python-pptx jsonschema jq

      - name: Run probe (JSON)
        run: |
          # Use a repository sample or generated PPTX; adjust path accordingly
          python tools/ppt_capability_probe.py --file samples/sample.pptx --json > probe.json
          cat probe.json | jq '.metadata.tool_version, .status'

      - name: Validate JSON against schema
        run: |
          python - <<'PY'
          import json, sys
          from jsonschema import Draft202012Validator
          schema = json.load(open("schemas/capability_probe.v1.1.1.schema.json"))
          payload = json.load(open("probe.json"))
          validator = Draft202012Validator(schema)
          errors = sorted(validator.iter_errors(payload), key=lambda e: e.path)
          if errors:
              for e in errors:
                  print("Validation error:", e.message, "at", list(e.path))
              sys.exit(1)
          print("Validation OK")
          PY
```

### 2) Node-based validation using ajv-cli (optional alternative)

```yaml
name: Validate capability probe outputs (ajv-cli)

on:
  workflow_dispatch:

jobs:
  validate-ajv:
    runs-on: ubuntu-latest
    steps:
      - name: Checkout
        uses: actions/checkout@v4

      - name: Install Node and ajv-cli
        uses: actions/setup-node@v4
        with:
          node-version: "18"
      - run: npm install -g ajv-cli

      - name: Generate probe output
        run: |
          python tools/ppt_capability_probe.py --file samples/sample.pptx --json > probe.json

      - name: Validate with ajv
        run: |
          ajv validate -s schemas/capability_probe.v1.1.1.schema.json -d probe.json --strict=false
```

---

## Local validation commands

- **Python jsonschema:**
  - **Validate:** python -m json.tool probe.json >/dev/null && python - <<'PY'
  - **Tip:** Use the embedded script from the CI step for quick local validation.

- **ajv-cli:**
  - **Install:** npm i -g ajv-cli
  - **Validate:** ajv validate -s schemas/capability_probe.v1.1.1.schema.json -d probe.json --strict=false

---

## Maintenance guidelines

- **Schema evolution:**
  - **Version bump:** When you change required fields or types, increment schema and tool versions together.
  - **Backwards compatibility:** Prefer adding optional fields with defaults; avoid removing required fields unless absolutely necessary.
- **Test artifacts:**
  - **Sample decks:** Maintain a small curated set of `.pptx` samples (single-master, multi-master, theme-only, large deck).
- **CI gating:**
  - **Blocking:** Make validation jobs required for merging PRs that touch the tool or schema.

If you want, I can add a makefile target (e.g., `make validate-probe`) and a small Python unit test for the validator to run in your existing CI.
