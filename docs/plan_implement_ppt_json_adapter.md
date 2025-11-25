# Implementation Plan - Meticulous Improvements (2025-11-25)

## Goal Description
Implement the "Meticulous Improvements" identified in `improvement_suggestions_20251125.md` to address correctness, safety, and resilience gaps. This includes introducing JSON schemas, a resilience wrapper for the capability probe, and a standardized change manifest system.

## User Review Required
> [!IMPORTANT]
> **New Safety Controls**:
> *   **Approval Tokens**: Destructive operations will eventually require signed tokens.
> *   **Change Manifests**: Complex workflows will require a manifest file.
> *   **Strict Schemas**: Tool outputs will be strictly validated.

## Proposed Changes

### Phase A: Schema & Adapter
Standardize and validate tool outputs.

#### [NEW] [ppt_get_info.schema.json](file:///home/project/powerpoint-agent-tools/schemas/ppt_get_info.schema.json)
- **Checklist**:
    - [ ] Define `tool_version`, `schema_version`, `presentation_version`.
    - [ ] Define `slides` array with `index`, `id`, `layout`.
    - [ ] Validate against Draft-07.

#### [NEW] [ppt_capability_probe.schema.json](file:///home/project/powerpoint-agent-tools/schemas/ppt_capability_probe.schema.json)
- **Checklist**:
    - [ ] Define `capabilities` object (`can_read`, `can_write`, `layouts`).
    - [ ] Define `probe_timestamp`.
    - [ ] Validate against Draft-07.

#### [NEW] [ppt_json_adapter.py](file:///home/project/powerpoint-agent-tools/tools/ppt_json_adapter.py)
- **Checklist**:
    - [ ] Implement alias mapping (`slidesTotal` -> `slide_count`).
    - [ ] Implement `presentation_version` computation if missing.
    - [ ] Implement `jsonschema` validation.
    - [ ] Return structured error JSON on failure.

### Phase B: Probe Wrapper & Preflight
Enhance resilience and operational safety.

#### [NEW] [probe_wrapper.sh](file:///home/project/powerpoint-agent-tools/scripts/probe_wrapper.sh)
- **Checklist**:
    - [ ] Implement timeout (default 15s).
    - [ ] Implement retry with exponential backoff.
    - [ ] Implement fallback to `ppt_get_info` + `ppt_get_slide_info` (slide 0, 1).
    - [ ] Output structured JSON.

#### [NEW] [probe_wrapper.ps1](file:///home/project/powerpoint-agent-tools/scripts/probe_wrapper.ps1)
- **Checklist**:
    - [ ] PowerShell equivalent of `probe_wrapper.sh`.

#### [NEW] [preflight_check.sh](file:///home/project/powerpoint-agent-tools/scripts/preflight_check.sh)
- **Checklist**:
    - [ ] Check absolute path.
    - [ ] Check read/write permissions.
    - [ ] Check disk space (>100MB).
    - [ ] Run `probe_wrapper.sh`.

#### [NEW] [preflight_check.ps1](file:///home/project/powerpoint-agent-tools/scripts/preflight_check.ps1)
- **Checklist**:
    - [ ] PowerShell equivalent of `preflight_check.sh`.

### Phase C: Change Manifest
Standardize workflow definition and auditing.

#### [NEW] [change_manifest.schema.json](file:///home/project/powerpoint-agent-tools/schemas/change_manifest.schema.json)
- **Checklist**:
    - [ ] Define `manifest_id`, `source_file`, `work_copy`.
    - [ ] Define `operations` array with `cmd`, `args`, `expected_effect`.
    - [ ] Define `approval_token` field.

#### [NEW] [sample_change_manifest.json](file:///home/project/powerpoint-agent-tools/examples/sample_change_manifest.json)
- **Checklist**:
    - [ ] Create realistic example (Clone -> Overlay -> Validate -> Export).
    - [ ] Include placeholder approval token.
    - [ ] Validate against schema.

## Verification Plan

### Automated Tests
1.  **Schema Validation**:
    *   Create `tests/test_schemas.py`.
    *   Test `ppt_json_adapter.py` with valid/invalid inputs.
    *   Test alias mapping logic.
2.  **Wrapper Resilience**:
    *   Create `tests/test_probe_wrapper.sh`.
    *   Mock `ppt_capability_probe.py` failure to verify fallback.
    *   Verify timeout behavior.
3.  **Manifest Compliance**:
    *   Validate `sample_change_manifest.json` against `change_manifest.schema.json`.

### Manual Verification
*   Run `scripts/preflight_check.sh` on a sample deck.
*   Run `tools/ppt_json_adapter.py` against existing tool outputs.

---

# Task: Implement Improvements (2025-11-25)

- [/] Plan & Analyze
    - [/] Read `improvement_suggestions_20251125.md`
    - [ ] Create detailed `implementation_plan.md`
    - [ ] Review and validate plan with user
- [ ] Implementation
    - [/] Phase A: Schema & Adapter
        - [x] `schemas/ppt_get_info.schema.json`
        - [x] `schemas/ppt_capability_probe.schema.json`
        - [x] `tools/ppt_json_adapter.py`
    - [/] Phase B: Probe Wrapper & Preflight
        - [x] `scripts/probe_wrapper.sh`
        - [x] `scripts/probe_wrapper.ps1`
        - [x] `scripts/preflight_check.sh`
        - [x] `scripts/preflight_check.ps1`
    - [/] Phase C: Change Manifest
        - [x] `schemas/change_manifest.schema.json`
        - [x] `examples/sample_change_manifest.json`
- [/] Verification & Regression Testing
    - [x] Create `tests/test_schemas.py`
    - [x] Create `tests/test_probe_wrapper.sh`
    - [x] Verify all changes

