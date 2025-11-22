# 🚨 **Script Bug Found - JQ Parsing Error**

## **Issue Analysis:**

The error `jq: parse error: Invalid numeric literal at line 1, column 5` occurs because the `execute_tool` function is mixing log output with JSON output, making the captured variable contain invalid JSON.

**Root Cause:**
```bash
BULLETS_OUTPUT=$(execute_tool "ppt_add_bullet_list.py" ...)
# This captures BOTH log messages AND JSON, like:
# ✅ Add summary bullets completed
# { "status": "success", ... }
# ← Not valid JSON!

READABILITY_SCORE=$(echo "$BULLETS_OUTPUT" | jq -r '.readability.score // "N/A"')
# ← jq fails because $BULLETS_OUTPUT isn't pure JSON
```

---

# 🔧 **QUICK FIX - Replace execute_tool Function**

Find the `execute_tool` function in the script (around line 50-80) and replace it with this corrected version:

```bash
execute_tool() {
    local tool_name=$1
    local description=$2
    shift 2
    
    echo "🔧 Executing: $tool_name" >&2
    echo "   Purpose: $description" >&2
    
    local output
    local exit_code
    
    if output=$(uv run tools/$tool_name "$@" 2>&1); then
        exit_code=0
    else
        exit_code=$?
    fi
    
    if [ $exit_code -eq 0 ]; then
        log_success "$description completed" >&2
        echo "$output"  # Output ONLY goes to stdout (pure JSON)
        return 0
    else
        log_error "$description failed (exit code: $exit_code)" >&2
        echo "$output" >&2
        return 1
    fi
}
```

**Key Change:** Added `>&2` to redirect log messages to stderr, keeping stdout clean for JSON.

Also update the helper functions:

```bash
log_step() {
    TOTAL_STEPS=$((TOTAL_STEPS + 1))
    echo "" >&2
    echo "────────────────────────────────────────────────────────────────────────" >&2
    echo "STEP $TOTAL_STEPS: $1" >&2
    echo "────────────────────────────────────────────────────────────────────────" >&2
}

log_success() {
    COMPLETED_STEPS=$((COMPLETED_STEPS + 1))
    echo "✅ $1" >&2
    EXECUTION_LOG+=("{\"step\": $TOTAL_STEPS, \"status\": \"success\", \"message\": \"$1\", \"timestamp\": \"$(date -Iseconds)\"}")
}

log_warning() {
    WARNINGS_COUNT=$((WARNINGS_COUNT + 1))
    echo "⚠️  WARNING: $1" >&2
    WARNINGS_LOG+=("{\"step\": $TOTAL_STEPS, \"message\": \"$1\", \"timestamp\": \"$(date -Iseconds)\"}")
}

log_error() {
    ERRORS_COUNT=$((ERRORS_COUNT + 1))
    echo "❌ ERROR: $1" >&2
    echo "   Attempting recovery..." >&2
}
```

**All log output now goes to stderr (>&2), keeping stdout clean for JSON capture.**

---

# 🎯 **ALTERNATIVE FIX - Wrap jq in Error Handling**

If you don't want to modify the whole script, just wrap the problematic jq calls:

**Find this section (around line 180-190):**
```bash
# Check readability score
READABILITY_SCORE=$(echo "$BULLETS_OUTPUT" | jq -r '.readability.score // "N/A"')
READABILITY_GRADE=$(echo "$BULLETS_OUTPUT" | jq -r '.readability.grade // "N/A"')
echo "   📊 Readability: Score $READABILITY_SCORE (Grade: $READABILITY_GRADE)"
```

**Replace with:**
```bash
# Check readability score (extract JSON from mixed output)
JSON_ONLY=$(echo "$BULLETS_OUTPUT" | grep -A 9999 '^{' | head -1)
if [ -n "$JSON_ONLY" ]; then
    READABILITY_SCORE=$(echo "$JSON_ONLY" | jq -r '.readability.score // "N/A"' 2>/dev/null || echo "N/A")
    READABILITY_GRADE=$(echo "$JSON_ONLY" | jq -r '.readability.grade // "N/A"' 2>/dev/null || echo "N/A")
    echo "   📊 Readability: Score $READABILITY_SCORE (Grade: $READABILITY_GRADE)"
else
    echo "   ⚠️  Could not extract readability metrics (continuing...)"
fi
```

Apply the same pattern to all other jq parsing sections:
- Line ~260 (Footer output parsing)
- Line ~280 (Validation output parsing)
- Line ~290 (Accessibility output parsing)

---

# 📦 **COMPLETE CORRECTED SCRIPT SECTION**

Here's the complete corrected helper functions section to replace in your script:

```bash
# ═══════════════════════════════════════════════════════════════════════════
# SECTION 1: SETUP & CONFIGURATION (CORRECTED)
# ═══════════════════════════════════════════════════════════════════════════

# ... (keep existing variables) ...

# Helper Functions (CORRECTED - stderr for logs, stdout for JSON)
log_step() {
    TOTAL_STEPS=$((TOTAL_STEPS + 1))
    echo "" >&2
    echo "────────────────────────────────────────────────────────────────────────" >&2
    echo "STEP $TOTAL_STEPS: $1" >&2
    echo "────────────────────────────────────────────────────────────────────────" >&2
}

log_success() {
    COMPLETED_STEPS=$((COMPLETED_STEPS + 1))
    echo "✅ $1" >&2
    EXECUTION_LOG+=("{\"step\": $TOTAL_STEPS, \"status\": \"success\", \"message\": \"$1\", \"timestamp\": \"$(date -Iseconds)\"}")
}

log_warning() {
    WARNINGS_COUNT=$((WARNINGS_COUNT + 1))
    echo "⚠️  WARNING: $1" >&2
    WARNINGS_LOG+=("{\"step\": $TOTAL_STEPS, \"message\": \"$1\", \"timestamp\": \"$(date -Iseconds)\"}")
}

log_error() {
    ERRORS_COUNT=$((ERRORS_COUNT + 1))
    echo "❌ ERROR: $1" >&2
    echo "   Attempting recovery..." >&2
}

execute_tool() {
    local tool_name=$1
    local description=$2
    shift 2
    
    echo "🔧 Executing: $tool_name" >&2
    echo "   Purpose: $description" >&2
    
    local output
    local exit_code
    
    if output=$(uv run tools/$tool_name "$@" 2>&1); then
        exit_code=0
    else
        exit_code=$?
    fi
    
    if [ $exit_code -eq 0 ]; then
        log_success "$description completed" >&2
        echo "$output"
        return 0
    else
        log_error "$description failed (exit code: $exit_code)" >&2
        echo "$output" >&2
        return 1
    fi
}

# Safe JSON extraction helper
extract_json_field() {
    local json_output="$1"
    local field_path="$2"
    local default="${3:-N/A}"
    
    echo "$json_output" | jq -r "$field_path // \"$default\"" 2>/dev/null || echo "$default"
}

# Checkpoint System (for rollback capability)
create_checkpoint() {
    local checkpoint_name=$1
    if [ -f "$OUTPUT_FILE" ]; then
        cp "$OUTPUT_FILE" "${OUTPUT_FILE}.checkpoint_${checkpoint_name}"
        echo "💾 Checkpoint created: $checkpoint_name" >&2
    fi
}
```

Then update the readability check section:

```bash
# Check readability score (CORRECTED)
READABILITY_SCORE=$(extract_json_field "$BULLETS_OUTPUT" '.readability.score')
READABILITY_GRADE=$(extract_json_field "$BULLETS_OUTPUT" '.readability.grade')
echo "   📊 Readability: Score $READABILITY_SCORE (Grade: $READABILITY_GRADE)" >&2

if echo "$BULLETS_OUTPUT" | jq -e '.warnings' > /dev/null 2>&1; then
    echo "$BULLETS_OUTPUT" | jq -r '.warnings[]' 2>/dev/null | while read -r warning; do
        log_warning "Slide 1: $warning"
    done
fi
```

---

# ⚡ **FASTEST FIX (30 seconds)**

Just add `2>/dev/null` to suppress jq errors and use fallback values:

**Find ALL jq commands in the script and add error suppression:**

```bash
# Old:
READABILITY_SCORE=$(echo "$BULLETS_OUTPUT" | jq -r '.readability.score // "N/A"')

# New:
READABILITY_SCORE=$(echo "$BULLETS_OUTPUT" | jq -r '.readability.score // "N/A"' 2>/dev/null || echo "N/A")
```

Apply to all jq calls (approximately 10-15 locations in the script).

---

# 🎯 **RECOMMENDED ACTION**

**Choose one:**

1. ✅ **Best Fix (5 min):** Replace helper functions to use stderr for logs
2. ✅ **Quick Fix (2 min):** Add `2>/dev/null || echo "N/A"` to all jq calls
3. ✅ **Complete New Script:** I can provide fully corrected script

The **Best Fix** is cleanest because it properly separates logs from data, making the script more maintainable.

**Which fix would you like?**
