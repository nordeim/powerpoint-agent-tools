#!/usr/bin/env bash
set -euo pipefail

TOOL="tools/ppt_capability_probe.py"
SCHEMA="schemas/capability_probe.v1.1.1.schema.json"

echo "═══════════════════════════════════════════════════════════════"
echo "Capability Probe v1.1.1 - Comprehensive Test Matrix"
echo "═══════════════════════════════════════════════════════════════"

# 1. Essential success
python "$TOOL" --file samples/sample.pptx --json > out_essential.json
python - <<'PY'
import json; d=json.load(open("out_essential.json"))
assert d["status"]=="success"
assert d["metadata"]["analysis_mode"]=="essential"
assert isinstance(d["metadata"]["warnings_count"], int)
assert d["metadata"]["layout_count_total"] >= d["metadata"]["layout_count_analyzed"]
PY

# 2. Deep mode success (positions present)
python "$TOOL" --file samples/sample.pptx --deep --json > out_deep.json
python - <<'PY'
import json; d=json.load(open("out_deep.json"))
assert d["metadata"]["analysis_mode"]=="deep"
# Assert placeholders present for at least one layout
assert any("placeholders" in l and isinstance(l["placeholders"], list) for l in d["layouts"])
PY

# 3. Missing file error path
if python "$TOOL" --file samples/missing.pptx --json > out_missing.json; then
  echo "❌ FAIL: expected error for missing file"
  exit 1
else
  echo "✅ PASS: Missing file error path"
fi

# 4. Multi-master template
python "$TOOL" --file samples/multi_master.pptx --json > out_multi.json
python - <<'PY'
import json; d=json.load(open("out_multi.json"))
assert d["capabilities"]["total_master_slides"] >= 2
assert len(d["capabilities"]["per_master"]) >= 2
PY

# 5. Theme color scheme fallback
python "$TOOL" --file samples/scheme_colors.pptx --json > out_scheme.json
python - <<'PY'
import json; d=json.load(open("out_scheme.json"))
colors=d["theme"]["colors"]
if colors:
  assert any(str(v).startswith("schemeColor:") for v in colors.values())
  assert any("scheme references" in w for w in d["warnings"])
PY

# 6. Timeout partial analysis
python "$TOOL" --file samples/large.pptx --deep --timeout 1 --json > out_timeout.json
python - <<'PY'
import json; d=json.load(open("out_timeout.json"))
assert d["capabilities"]["analysis_complete"] in (True, False)
# If not complete, expect a warning message
if not d["capabilities"]["analysis_complete"]:
  assert any("timeout" in w.lower() for w in d["warnings"])
PY

# 7. Strict schema validation
python - <<'PY'
import json
from core.strict_validator import validate_against_schema
d=json.load(open("out_essential.json"))
validate_against_schema(d, "schemas/capability_probe.v1.1.1.schema.json")
d=json.load(open("out_deep.json"))
validate_against_schema(d, "schemas/capability_probe.v1.1.1.schema.json")
PY

echo "═══════════════════════════════════════════════════════════════"
echo "All tests passed"

