### Executive Summary
The current toolset is well-conceived for CLI-driven presentation engineering: it enforces JSON-first I/O, capability probing, cloning-before-editing, and a validation gate. However, a meticulous review reveals **implementation correctness gaps, brittle assumptions, and missing operational controls** that reduce reliability, safety, and alignment with the Presentation Architect’s needs. This review identifies concrete defects, root causes, and prioritized remediation with testable acceptance criteria so the toolset becomes robust, idempotent, auditable, and safe for automated pipelines and human-in-the-loop workflows.

---

### Key findings (high-level)
- **Correctness gaps**: inconsistent error codes and JSON schemas across tools; fragile parsing assumptions; missing versioning for structural changes.  
- **Safety gaps**: destructive operations lack enforceable approval tokens; no single-source-of-truth for approval policy; overlays and z-order operations risk obscuring content.  
- **Resilience gaps**: capability probe can block pipelines (no timeout/fallback); no retry/backoff for transient failures; no snapshot/version token to detect concurrent edits.  
- **Observability & auditability gaps**: inconsistent audit fields in command outputs; no standardized change manifest produced by tools; limited raw-output logging for debugging.  
- **Validation alignment gaps**: validation tools run but acceptance thresholds are undefined; accessibility checks produce warnings but no remediation hooks.  
- **Developer ergonomics**: inconsistent CLI flags, missing machine-readable exit codes, and no test harness for schema drift detection.

---

### Detailed issues, root causes, and precise fixes

#### 1. Inconsistent JSON outputs and fragile parsers
- **Symptom**: Tools return JSON with different key names or optional fields; parsers assume exact keys (e.g., `slide_count` vs `slides_count`).  
- **Root cause**: No enforced JSON schema or versioning for each tool output.  
- **Fix**:
  - **Define and publish a JSON Schema** for every tool output (Draft-07 or later). Include `tool_version` and `schema_version` fields in outputs.
  - **Implement a small compatibility layer** (`ppt_json_adapter.py`) that maps aliases and validates required keys; fail fast with clear error codes when required keys are missing.
  - **Acceptance criteria**: `ajv`/`jsonschema` validation passes for all tool outputs in CI; adapter logs missing/alias keys.

#### 2. Capability probe blocking pipelines
- **Symptom**: `ppt_capability_probe.py --deep` can hang or return partial data, stalling automation.  
- **Root cause**: No timeout, no fallback probes, no retry/backoff.  
- **Fix**:
  - Add **configurable timeout** and **retry policy** (3 attempts, exponential backoff) to the probe wrapper.
  - Implement **fallback probes**: if deep probe fails, run `ppt_get_info.py` and `ppt_get_slide_info.py` for slide 0 and 1 to collect minimal metadata.
  - **Acceptance criteria**: probe wrapper returns either full probe JSON or a documented minimal metadata JSON within timeout; CI simulates probe failure and verifies fallback path.

#### 3. No presentation versioning or snapshot token
- **Symptom**: Multi-step scripts race; shape indices become stale after structural edits.  
- **Root cause**: Tools mutate presentation state without returning a stable `presentation_version` or snapshot token.  
- **Fix**:
  - Add `presentation_version` (hash of slide list + timestamps + template id) to all mutating tool outputs and to `ppt_get_info.py`.
  - Require subsequent commands to include `--expected-presentation-version` and abort if mismatch.
  - **Acceptance criteria**: scripts that include the expected version abort on mismatch and re-run probe; unit tests simulate concurrent edits.

#### 4. Destructive operations lack enforceable approval
- **Symptom**: `ppt_delete_slide.py` or `ppt_remove_shape.py` can be invoked in automation without human approval.  
- **Root cause**: Approval is a policy, not an enforced technical control.  
- **Fix**:
  - Implement **approval-token** enforcement: destructive commands require `--approval-token TOKEN`. Tokens are HMAC-signed JSON (fields: `manifest_id`, `user`, `expiry`, `scope`) and validated by the CLI before execution.
  - Provide a **token generation utility** (`ppt_generate_approval_token.py`) that requires interactive confirmation or an authorized service account.
  - **Acceptance criteria**: destructive commands fail with a clear exit code when token missing/invalid; CI tests token generation and verification.

#### 5. Change manifest missing or inconsistent
- **Symptom**: No single machine-readable manifest describing intended operations, diffs, and rollback steps.  
- **Root cause**: Tools log actions but do not produce a standardized manifest.  
- **Fix**:
  - Define a **Change Manifest JSON Schema** (full, strict) and require `--manifest PATH` for multi-step automation runs. The manifest includes `manifest_id`, `source_file`, `work_copy`, `operations[]` (cmd, args, expected_effect, rollback_cmd), `diff_summary`, `created_by`, `timestamp`, `approval_token`.
  - Tools should append their actual result to the manifest (operation result, `presentation_version` after op, stdout/stderr snippet).
  - **Acceptance criteria**: manifest validates against schema; manifest contains operation results and can be used to drive rollback.

#### 6. Validation thresholds and remediation hooks missing
- **Symptom**: `ppt_check_accessibility.py` returns warnings but pipeline lacks pass/fail thresholds or automated remediation.  
- **Root cause**: Validation tools are used ad-hoc without policy.  
- **Fix**:
  - Define **validation policy** (e.g., critical failures = fail; warnings ≤ 3 allowed). Store policy in `validation_policy.json`.
  - Implement **remediation hooks**: for common failures (missing alt text, low contrast), provide `ppt_fix_alt_text.py` and `ppt_adjust_contrast.py` that can be invoked automatically or suggested in the manifest.
  - **Acceptance criteria**: pipeline enforces policy; remediation scripts reduce failures in automated tests.

#### 7. Overlay and z-order operations risk content loss
- **Symptom**: Visual transformer recipes add full-slide shapes that obscure text.  
- **Root cause**: No defaults for z-order, opacity, or anchor; no post-change readability check.  
- **Fix**:
  - Require explicit `--z-order` and `--opacity` flags for overlay commands; default to `behind_text` and `opacity=0.15` for readability overlays.
  - After overlay insertion, run a **readability check** (contrast ratio sampling) and abort if text contrast falls below threshold.
  - **Acceptance criteria**: readability check passes in automated tests; overlay commands include z-order/opacity in manifest.

#### 8. Error codes and observability
- **Symptom**: Tools return mixed exit codes and freeform stderr messages; automation cannot reliably react.  
- **Root cause**: No standardized exit code table or structured error objects.  
- **Fix**:
  - Standardize exit codes (0 success, 1 usage error, 2 validation failure, 3 transient error, 4 permission/approval error, 5 internal error).
  - Return structured error JSON on non-zero exit with `error_code`, `message`, `retryable` boolean, and `hint`.
  - **Acceptance criteria**: CI asserts exit codes and structured error JSON for simulated failures.

---

### Tests and validation plan (concrete)
- **Unit tests**
  - Validate every tool output against its JSON Schema using `ajv` or `jsonschema`.
  - Test `ppt_json_adapter.py` with alias key scenarios.
- **Integration tests**
  - Simulate full pipeline: probe → preflight → clone → modify (add overlay) → validate → export. Verify manifests, `presentation_version` checks, and approval enforcement.
  - Simulate probe failure and verify fallback probes succeed.
  - Simulate concurrent edit: mutate work copy externally and verify `--expected-presentation-version` aborts.
- **Security tests**
  - Verify approval token HMAC verification and expiry enforcement.
  - Ensure tokens cannot be reused across manifests.
- **Accessibility tests**
  - Run `ppt_check_accessibility.py` on sample decks with known issues; verify remediation scripts reduce failures.
- **Chaos tests**
  - Inject transient failures (I/O errors, timeouts) and verify retry/backoff behavior and idempotency.

---

### Implementation roadmap (prioritized, with checkpoints)
1. **Schema and adapter (2 days)**  
   - Deliver: JSON Schemas for all tool outputs; `ppt_json_adapter.py`.  
   - Checkpoint: All tool outputs validate in CI.
2. **Probe wrapper and preflight (1 day)**  
   - Deliver: `probe_wrapper.sh`/`.ps1` with timeout, retries, fallback probes; preflight script template.  
   - Checkpoint: Fallback path exercised in CI.
3. **Presentation versioning (1 day)**  
   - Deliver: `presentation_version` generation and enforcement in mutating tools.  
   - Checkpoint: Version mismatch abort test passes.
4. **Approval-token enforcement (2 days)**  
   - Deliver: token generation utility, verification library, CLI enforcement.  
   - Checkpoint: Destructive ops blocked without valid token; token expiry test passes.
5. **Change manifest and manifest wiring (1.5 days)**  
   - Deliver: `change_manifest.schema.json`, manifest append logic in tools.  
   - Checkpoint: End-to-end manifest produced and validated.
6. **Validation policy and remediation hooks (1.5 days)**  
   - Deliver: `validation_policy.json`, remediation scripts, pipeline integration.  
   - Checkpoint: Pipeline enforces policy and remediation reduces failures.
7. **Observability and error standardization (1 day)**  
   - Deliver: standardized exit codes and structured error JSON.  
   - Checkpoint: CI asserts error formats.

**Total estimated effort**: ~10–11 days (can be compressed with parallel workstreams).

---

### Short-term tactical fixes (apply immediately)
- Add probe timeout and fallback wrapper to the repo and wire it into CI.  
- Add `--approval-token` flag to destructive commands that returns exit code 4 when missing.  
- Publish a minimal `change_manifest.schema.json` and require `--manifest` for multi-step automation runs.  
- Add `presentation_version` to `ppt_get_info.py` output.

---

### Next steps and deliverables I will produce if you want me to proceed
- **Concrete artifacts**: full JSON Schemas for all tools, `ppt_json_adapter.py`, `probe_wrapper` scripts (POSIX + PowerShell), `ppt_generate_approval_token.py` and verification library, `change_manifest.schema.json`, and a CI test matrix.  
- **Runbook**: step-by-step operational runbook for approval token issuance, manifest usage, and rollback procedures.  
- **Automated tests**: unit and integration test suites and sample decks to validate behavior.

---

### Executive Summary
I will produce the first set of artifacts in the requested order and with rigorous validation: **(A)** JSON Schemas for `ppt_get_info.py` and `ppt_capability_probe.py` outputs plus a resilient Python adapter that validates and normalizes tool JSON; **(B)** a probe wrapper and preflight script templates for POSIX and PowerShell that implement timeouts, retries, fallbacks, and preflight checks; **(C)** a strict change manifest JSON Schema and a fully populated sample manifest.  

Below is a validated implementation plan followed by the executed artifacts, test guidance, and next steps. Each artifact is production-oriented, includes clear success criteria, and is designed to integrate into CI pipelines and automation flows.

---

### Implementation Plan Overview
**Goal** Produce artifacts A, B, C in sequence with validation checkpoints and testable acceptance criteria.

**Phases**
1. **Phase A Schema and Adapter**
   - Deliverables: `ppt_get_info.schema.json`, `ppt_capability_probe.schema.json`, `ppt_json_adapter.py`.
   - Key tasks: design schemas, include `tool_version` and `schema_version`, implement adapter that validates, maps aliases, and emits normalized JSON and structured errors.
   - Checkpoint: adapter validates sample outputs and maps alias keys; CI schema validation passes.

2. **Phase B Probe Wrapper and Preflight Scripts**
   - Deliverables: `probe_wrapper.sh` and `probe_wrapper.ps1`, `preflight_check.sh` and `preflight_check.ps1`.
   - Key tasks: implement timeout, retry with exponential backoff, fallback probes, file path and permission checks, disk space check, tool availability check, structured JSON output on success/failure.
   - Checkpoint: wrapper returns either full probe JSON or documented minimal metadata JSON within timeout; fallback path exercised in tests.

3. **Phase C Change Manifest Schema and Sample**
   - Deliverables: `change_manifest.schema.json`, `sample_change_manifest.json`.
   - Key tasks: design manifest schema with operations array, rollback commands, diffs, approval token field; produce a realistic sample manifest for a clone → overlay → validate → export flow.
   - Checkpoint: manifest validates against schema; sample manifest includes expected fields and a plausible approval token placeholder.

**Estimated effort** 4–6 hours for design and artifact creation, plus 1–2 hours for basic tests and validation.

**Success criteria**
- All JSON Schemas validate sample outputs.
- Adapter returns normalized JSON and clear structured errors for invalid inputs.
- Probe wrapper enforces timeout and fallback behavior and returns JSON.
- Preflight scripts perform file, permission, disk, and tool checks and return structured JSON.
- Change manifest schema validates the sample manifest.

---

---

### Phase A Artifacts JSON Schemas and Adapter

#### JSON Schema for ppt_get_info.py output
```json
{
  "$schema": "http://json-schema.org/draft-07/schema#",
  "title": "ppt_get_info Output Schema",
  "type": "object",
  "required": ["tool_name", "tool_version", "schema_version", "file", "presentation_version", "slide_count", "slides"],
  "properties": {
    "tool_name": { "type": "string" },
    "tool_version": { "type": "string" },
    "schema_version": { "type": "string" },
    "file": { "type": "string" },
    "presentation_version": { "type": "string" },
    "slide_count": { "type": "integer", "minimum": 0 },
    "slides": {
      "type": "array",
      "items": {
        "type": "object",
        "required": ["index", "id", "layout", "shape_count"],
        "properties": {
          "index": { "type": "integer", "minimum": 0 },
          "id": { "type": "string" },
          "layout": { "type": "string" },
          "shape_count": { "type": "integer", "minimum": 0 },
          "notes": { "type": "string" }
        }
      }
    },
    "template_id": { "type": ["string", "null"] },
    "theme": {
      "type": ["object", "null"],
      "properties": {
        "primary_color": { "type": "string" },
        "accent_colors": {
          "type": "array",
          "items": { "type": "string" }
        },
        "font_family": { "type": "string" }
      },
      "additionalProperties": true
    },
    "metadata": { "type": "object", "additionalProperties": true }
  },
  "additionalProperties": false
}
```

#### JSON Schema for ppt_capability_probe.py output
```json
{
  "$schema": "http://json-schema.org/draft-07/schema#",
  "title": "ppt_capability_probe Output Schema",
  "type": "object",
  "required": ["tool_name", "tool_version", "schema_version", "file", "probe_timestamp", "capabilities"],
  "properties": {
    "tool_name": { "type": "string" },
    "tool_version": { "type": "string" },
    "schema_version": { "type": "string" },
    "file": { "type": "string" },
    "probe_timestamp": { "type": "string", "format": "date-time" },
    "capabilities": {
      "type": "object",
      "required": ["can_read", "can_write", "layouts", "slide_dimensions"],
      "properties": {
        "can_read": { "type": "boolean" },
        "can_write": { "type": "boolean" },
        "layouts": {
          "type": "array",
          "items": { "type": "string" }
        },
        "slide_dimensions": {
          "type": "object",
          "required": ["width_pt", "height_pt"],
          "properties": {
            "width_pt": { "type": "number" },
            "height_pt": { "type": "number" }
          }
        },
        "max_image_size_mb": { "type": "number" },
        "supports_z_order": { "type": "boolean" }
      },
      "additionalProperties": true
    },
    "warnings": {
      "type": "array",
      "items": { "type": "string" }
    },
    "metadata": { "type": "object", "additionalProperties": true }
  },
  "additionalProperties": false
}
```

#### Python Adapter ppt_json_adapter.py
```python
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
```

**Adapter Notes**
- Uses `jsonschema` for validation. In CI, install `jsonschema` via pip.
- Maps common alias keys and computes a best-effort `presentation_version` if missing.
- Emits structured error JSON with `error_code` and `retryable` flag for automation to react.

---

---

### Phase B Artifacts Probe Wrapper and Preflight Scripts

#### POSIX Probe Wrapper probe_wrapper.sh
```bash
#!/usr/bin/env bash
# probe_wrapper.sh
# Usage: probe_wrapper.sh /absolute/path/to/presentation.pptx
# Behavior:
# - Validates absolute path and readability
# - Runs ppt_capability_probe.py with timeout and retries
# - On failure, runs fallback probes: ppt_get_info.py and ppt_get_slide_info.py
# - Emits JSON to stdout; on error emits structured JSON and non-zero exit

set -euo pipefail

FILE="${1:-}"
TIMEOUT_SECONDS=15
MAX_RETRIES=3
SLEEP_BASE=2
TMPDIR="$(mktemp -d)"
PROBE_OUT="$TMPDIR/probe.json"

function emit_error {
  jq -n --arg code "$1" --arg msg "$2" --argjson retryable "$3" \
    '{error:{error_code:$code, message:$msg, retryable:$retryable}}'
}

if [[ -z "$FILE" ]]; then
  emit_error "USAGE_ERROR" "Missing file argument" false
  exit 1
fi

if [[ "${FILE:0:1}" != "/" ]]; then
  emit_error "RELATIVE_PATH_NOT_ALLOWED" "Absolute path required" false
  exit 1
fi

if [[ ! -r "$FILE" ]]; then
  emit_error "PERMISSION_DENIED" "File not readable" false
  exit 1
fi

# Disk space check on containing filesystem
MIN_SPACE_MB=100
avail_mb=$(df --output=avail -m "$(dirname "$FILE")" | tail -1 | tr -d ' ')
if [[ -z "$avail_mb" || "$avail_mb" -lt "$MIN_SPACE_MB" ]]; then
  emit_error "LOW_DISK_SPACE" "Available space less than ${MIN_SPACE_MB}MB" false
  exit 1
fi

# Check tool availability
if ! command -v ppt_capability_probe.py >/dev/null 2>&1; then
  emit_error "TOOL_MISSING" "ppt_capability_probe.py not found in PATH" false
  exit 1
fi

# Attempt probe with retries and exponential backoff
attempt=0
while [[ $attempt -lt $MAX_RETRIES ]]; do
  attempt=$((attempt+1))
  if timeout "${TIMEOUT_SECONDS}s" ppt_capability_probe.py --file "$FILE" --deep --json > "$PROBE_OUT" 2>&1; then
    cat "$PROBE_OUT"
    rm -rf "$TMPDIR"
    exit 0
  else
    sleep_time=$((SLEEP_BASE ** attempt))
    sleep "$sleep_time"
  fi
done

# Fallback probes
if command -v ppt_get_info.py >/dev/null 2>&1 && command -v ppt_get_slide_info.py >/dev/null 2>&1; then
  info_json="$TMPDIR/info.json"
  slide0_json="$TMPDIR/slide0.json"
  if ppt_get_info.py --file "$FILE" --json > "$info_json" 2>&1 && ppt_get_slide_info.py --file "$FILE" --slide 0 --json > "$slide0_json" 2>&1; then
    # Merge minimal metadata into a single JSON object
    jq -s '.[0] + {probe_fallback:true, slide0:.[1]}' "$info_json" "$slide0_json"
    rm -rf "$TMPDIR"
    exit 0
  else
    emit_error "PROBE_FALLBACK_FAILED" "Both deep probe and fallback probes failed" true
    rm -rf "$TMPDIR"
    exit 3
  fi
else
  emit_error "PROBE_AND_FALLBACK_TOOLS_MISSING" "Fallback tools not available" false
  rm -rf "$TMPDIR"
  exit 1
fi
```

#### POSIX Preflight Script preflight_check.sh
```bash
#!/usr/bin/env bash
# preflight_check.sh
# Usage: preflight_check.sh /absolute/path/to/presentation.pptx
# Performs:
# - absolute path enforcement
# - read/write permission checks
# - disk space check
# - probe via probe_wrapper.sh
# - outputs JSON summary

set -euo pipefail

FILE="${1:-}"
MIN_SPACE_MB=100

if [[ -z "$FILE" ]]; then
  echo '{"error":"Missing file argument"}' && exit 1
fi

if [[ "${FILE:0:1}" != "/" ]]; then
  echo '{"error":"Absolute path required"}' && exit 1
fi

if [[ ! -r "$FILE" ]]; then
  echo '{"error":"File not readable"}' && exit 1
fi

if [[ ! -w "$(dirname "$FILE")" ]]; then
  echo '{"error":"No write permission to destination directory"}' && exit 1
fi

avail_mb=$(df --output=avail -m "$(dirname "$FILE")" | tail -1 | tr -d ' ')
if [[ -z "$avail_mb" || "$avail_mb" -lt "$MIN_SPACE_MB" ]]; then
  echo "{\"error\":\"Low disk space: ${avail_mb}MB available\"}" && exit 1
fi

# Run probe wrapper and capture JSON
if probe_wrapper.sh "$FILE" > /tmp/preflight_probe.json 2>&1; then
  jq -n --arg file "$FILE" --argjson probe "$(cat /tmp/preflight_probe.json | jq '.')" \
    '{file:$file, preflight: {status: "ok"}, probe: $probe}'
  exit 0
else
  echo '{"error":"Probe failed"}' && exit 2
fi
```

#### PowerShell Probe Wrapper probe_wrapper.ps1
```powershell
# probe_wrapper.ps1
# Usage: .\probe_wrapper.ps1 -File C:\path\to\presentation.pptx
param(
  [Parameter(Mandatory=$true)][string]$File,
  [int]$TimeoutSeconds = 15,
  [int]$MaxRetries = 3
)

function Emit-ErrorJson {
  param($Code, $Message, $Retryable)
  $obj = @{ error = @{ error_code = $Code; message = $Message; retryable = $Retryable } }
  $obj | ConvertTo-Json -Depth 5
}

if (-not [System.IO.Path]::IsPathRooted($File)) {
  Emit-ErrorJson "RELATIVE_PATH_NOT_ALLOWED" "Absolute path required" $false
  exit 1
}

if (-not (Test-Path -Path $File -PathType Leaf)) {
  Emit-ErrorJson "FILE_NOT_FOUND" "File does not exist" $false
  exit 1
}

if (-not (Get-Item $File).IsReadOnly -and -not (Get-Acl $File)) {
  # best-effort permission check; continue
  $null = $null
}

# Disk space check
$drive = Get-PSDrive -Name ([System.IO.Path]::GetPathRoot($File).TrimEnd('\'))
if ($drive.Free -lt 100MB) {
  Emit-ErrorJson "LOW_DISK_SPACE" "Available space less than 100MB" $false
  exit 1
}

# Tool availability
if (-not (Get-Command ppt_capability_probe.py -ErrorAction SilentlyContinue)) {
  Emit-ErrorJson "TOOL_MISSING" "ppt_capability_probe.py not found" $false
  exit 1
}

$attempt = 0
while ($attempt -lt $MaxRetries) {
  $attempt++
  try {
    $proc = Start-Process -FilePath "ppt_capability_probe.py" -ArgumentList "--file", $File, "--deep", "--json" -NoNewWindow -RedirectStandardOutput "$env:TEMP\probe_out.json" -Wait -PassThru -ErrorAction Stop -Timeout $TimeoutSeconds
    Get-Content "$env:TEMP\probe_out.json" | Out-String
    Remove-Item "$env:TEMP\probe_out.json" -ErrorAction SilentlyContinue
    exit 0
  } catch {
    Start-Sleep -Seconds ([math]::Pow(2, $attempt))
  }
}

# Fallback probes
if (Get-Command ppt_get_info.py -ErrorAction SilentlyContinue -and Get-Command ppt_get_slide_info.py -ErrorAction SilentlyContinue) {
  try {
    & ppt_get_info.py --file $File --json > "$env:TEMP\info.json"
    & ppt_get_slide_info.py --file $File --slide 0 --json > "$env:TEMP\slide0.json"
    $info = Get-Content "$env:TEMP\info.json" -Raw | ConvertFrom-Json
    $slide0 = Get-Content "$env:TEMP\slide0.json" -Raw | ConvertFrom-Json
    $merged = $info | Add-Member -PassThru -NotePropertyName probe_fallback -NotePropertyValue $true
    $merged | Add-Member -PassThru -NotePropertyName slide0 -NotePropertyValue $slide0
    $merged | ConvertTo-Json -Depth 10
    Remove-Item "$env:TEMP\info.json","$env:TEMP\slide0.json" -ErrorAction SilentlyContinue
    exit 0
  } catch {
    Emit-ErrorJson "PROBE_FALLBACK_FAILED" "Both deep probe and fallback probes failed" $true
    exit 3
  }
} else {
  Emit-ErrorJson "PROBE_AND_FALLBACK_TOOLS_MISSING" "Fallback tools not available" $false
  exit 1
}
```

**Probe Wrapper Notes**
- Both wrappers return structured JSON on success and structured error JSON on failure.
- Timeout and retry parameters are configurable.
- Fallback probes produce a minimal merged JSON to allow automation to continue with essential metadata.

---

---

### Phase C Artifacts Change Manifest Schema and Sample Manifest

#### Change Manifest JSON Schema change_manifest.schema.json
```json
{
  "$schema": "http://json-schema.org/draft-07/schema#",
  "title": "Presentation Change Manifest",
  "type": "object",
  "required": ["manifest_id", "source_file", "work_copy", "operations", "created_by", "timestamp", "approval_token"],
  "properties": {
    "manifest_id": { "type": "string" },
    "source_file": { "type": "string" },
    "work_copy": { "type": "string" },
    "created_by": { "type": "string" },
    "timestamp": { "type": "string", "format": "date-time" },
    "approval_token": { "type": "string" },
    "operations": {
      "type": "array",
      "items": {
        "type": "object",
        "required": ["op_id", "cmd", "args", "expected_effect"],
        "properties": {
          "op_id": { "type": "string" },
          "cmd": { "type": "string" },
          "args": { "type": "object" },
          "expected_effect": { "type": "string" },
          "rollback_cmd": { "type": "string" },
          "critical": { "type": "boolean" }
        },
        "additionalProperties": false
      }
    },
    "diff_summary": {
      "type": "object",
      "properties": {
        "slides_added": { "type": "integer" },
        "slides_removed": { "type": "integer" },
        "shapes_added": { "type": "integer" },
        "shapes_removed": { "type": "integer" }
      },
      "additionalProperties": false
    },
    "notes": { "type": "string" }
  },
  "additionalProperties": false
}
```

#### Sample Change Manifest sample_change_manifest.json
```json
{
  "manifest_id": "manifest-20251125-001",
  "source_file": "/data/presentations/quarterly-deck.pptx",
  "work_copy": "/tmp/manifest-20251125-001-workcopy.pptx",
  "created_by": "alice@example.com",
  "timestamp": "2025-11-25T06:57:00Z",
  "approval_token": "HMAC-SHA256:eyJtYW5pZmVzdF9pZCI6Im1hbmlmZXN0LTIwMjUxMTI1LTAwMSIsInVzZXIiOiJhbGljZUBleGFtcGxlLmNvbSIsImV4cGlyeSI6IjIwMjUtMTEtMjVUMDg6MDA6MDBaIiwic2NvcGUiOiJkZWxldGU6c2xpZGUifQ==.signature-placeholder",
  "operations": [
    {
      "op_id": "op-01-clone",
      "cmd": "ppt_clone_presentation.py",
      "args": {
        "--source": "/data/presentations/quarterly-deck.pptx",
        "--output": "/tmp/manifest-20251125-001-workcopy.pptx",
        "--json": true
      },
      "expected_effect": "Create a writable work copy for edits",
      "rollback_cmd": "rm -f /tmp/manifest-20251125-001-workcopy.pptx",
      "critical": true
    },
    {
      "op_id": "op-02-add-overlay",
      "cmd": "ppt_add_shape.py",
      "args": {
        "--file": "/tmp/manifest-20251125-001-workcopy.pptx",
        "--slide-index": 2,
        "--shape-type": "rectangle",
        "--anchor": "full-slide",
        "--opacity": 0.15,
        "--z-order": "behind_text",
        "--fill-color": "#000000"
      },
      "expected_effect": "Add a readability overlay to slide 2 without obscuring text",
      "rollback_cmd": "ppt_remove_shape.py --file /tmp/manifest-20251125-001-workcopy.pptx --slide-index 2 --shape-id <shape-id-from-result>",
      "critical": false
    },
    {
      "op_id": "op-03-run-accessibility",
      "cmd": "ppt_check_accessibility.py",
      "args": {
        "--file": "/tmp/manifest-20251125-001-workcopy.pptx",
        "--json": true
      },
      "expected_effect": "Produce accessibility report; must meet policy thresholds",
      "rollback_cmd": "",
      "critical": true
    },
    {
      "op_id": "op-04-export",
      "cmd": "ppt_export_pdf.py",
      "args": {
        "--file": "/tmp/manifest-20251125-001-workcopy.pptx",
        "--output": "/data/exports/quarterly-deck-20251125.pdf"
      },
      "expected_effect": "Export final PDF after validation",
      "rollback_cmd": "rm -f /data/exports/quarterly-deck-20251125.pdf",
      "critical": true
    }
  ],
  "diff_summary": {
    "slides_added": 0,
    "slides_removed": 0,
    "shapes_added": 1,
    "shapes_removed": 0
  },
  "notes": "Add overlay to slide 3 for readability; ensure accessibility checks pass before export."
}
```

**Manifest Notes**
- `approval_token` is a placeholder HMAC-signed token string. In production, tokens must be generated by the approved token service and verified by the CLI before executing critical or destructive operations.
- Each operation includes `rollback_cmd` where feasible. Tools should append actual `result` objects to the manifest during execution.

---

---

### Validation and Test Guidance

**Unit tests**
- Validate `ppt_get_info.schema.json` and `ppt_capability_probe.schema.json` using `jsonschema` or `ajv` against representative sample outputs.
- Run `ppt_json_adapter.py` with intentionally alias-laden JSON to confirm alias mapping and computed `presentation_version`.

**Integration tests**
- Run `probe_wrapper.sh` against a sample deck:
  - Simulate successful deep probe and verify full probe JSON returned.
  - Simulate probe failure by making `ppt_capability_probe.py` exit non-zero and verify fallback probes produce merged minimal JSON.
- Run `preflight_check.sh` to confirm file, permission, and disk checks behave as expected.

**Manifest validation**
- Validate `sample_change_manifest.json` against `change_manifest.schema.json` using `jsonschema`.
- Simulate executing operations in the manifest in a dry-run mode where each command is validated but not executed; ensure approval token presence blocks execution if invalid.

**Error handling tests**
- Feed malformed JSON to adapter and confirm structured error JSON with `error_code` and `retryable` flags.
- Simulate missing tools and confirm probe wrapper returns `TOOL_MISSING` error JSON.

---

### Documentation and Usage Notes

**Adapter usage**
- Install dependency: `pip install jsonschema`
- Run: `python ppt_json_adapter.py --schema ppt_get_info.schema.json --input raw_get_info.json`
- On success, normalized JSON is printed. On failure, structured error JSON is printed and exit code indicates type.

**Probe wrapper usage**
- Ensure `probe_wrapper.sh` is executable and in PATH or call with full path.
- Call: `probe_wrapper.sh /absolute/path/to/deck.pptx`
- On success, wrapper prints JSON; on failure, prints structured error JSON.

**Manifest usage**
- CI or orchestration should require `--manifest` for multi-step runs.
- Before executing critical operations, verify `approval_token` via the token verification library.
- Tools should append operation results to the manifest after each step for auditability.

---

### Next Steps and Recommendations
- **Immediate** Add these artifacts to the repository and wire basic CI checks:
  - Schema validation job that validates all tool outputs against schemas.
  - Linting and unit tests for `ppt_json_adapter.py`.
  - Integration test that runs `probe_wrapper` and `preflight_check` against sample decks.
- **Short term** Implement token generation and verification utilities and integrate them into destructive CLI commands.
- **Medium term** Extend adapter to support schema version negotiation and automatic migration helpers for older tool outputs.
- **Optional** Add a manifest runner utility that executes operations in the manifest, appends results, enforces `presentation_version` checks, and supports rollback.

---

### Final Checklist Before Handoff
- [ ] Add schemas to repo and document schema versions.
- [ ] Add adapter and include dependency installation instructions.
- [ ] Add probe wrapper and preflight scripts for POSIX and PowerShell.
- [ ] Add change manifest schema and sample manifest.
- [ ] Add CI jobs to validate schemas and run basic integration tests.

---

https://copilot.microsoft.com/shares/s1Yt54f7T8JyC7KqVw2fC  
