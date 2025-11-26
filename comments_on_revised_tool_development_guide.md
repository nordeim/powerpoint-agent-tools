## 📋 Remediation Recommendations

### Immediate Actions (Critical)
1. **Standardize Exit Codes**: Align template with actual codebase implementation or update all tools
2. **Add Probe Resilience**: Include probe wrapper pattern and timeout handling
3. **Implement Schema Validation**: Add complete `strict_validator.py` usage examples

### Short-Term Actions (High Priority)
5. **Platform-Aware Path Safety**: Replace string-based path checks with `PathValidator` usage
6. **Approval Token Consistency**: Update template to match actual token enforcement patterns
7. **Shape Index Nuances**: Document which operations automatically refresh indices vs. require explicit refresh

### Long-Term Actions (Medium Priority)
8. **Add Preflight Checklist**: Integrate system prompt's preflight requirements
9. **Document Fallback Patterns**: Add graceful degradation strategies for failed operations
10. **Include Real Tool Examples**: Reference actual tool implementations as canonical examples

## 🔄 Validation Protocol

Before finalizing this guide, execute this validation checklist:

- [ ] **Cross-Reference Check**: Verify every code example against actual tool implementations
- [ ] **Platform Testing**: Test all path operations on Windows, macOS, and Linux
- [ ] **Error Code Audit**: Confirm exit codes match across all 39 tools
- [ ] **Schema Validation**: Test all JSON schema examples against actual schemas
- [ ] **Version Consistency**: Ensure all version references match (v3.1.0 throughout)
- [ ] **Tool Name Verification**: Confirm all referenced tool names exist in codebase
- [ ] **Parameter Validation**: Verify all CLI parameters match actual tool implementations
- [ ] **Exception Hierarchy**: Confirm all custom exceptions are properly imported and used

## 📝 Detailed Section-by-Section Critique

### **Section 2: Governance Principles**
✅ **Strengths**: Comprehensive coverage of core governance concepts
⚠️ **Gap**: Missing `PathValidator` class usage for platform-independent path safety
✅ **Accuracy**: Approval token structure matches system prompt exactly

### **Section 3: Master Template**
⚠️ **Critical Issue**: Exit code handling doesn't match codebase reality
✅ **Improvement**: Added comprehensive exception handling for core exceptions
⚠️ **Platform Issue**: Path safety check fails on Windows systems
✅ **Governance**: Clone-before-edit enforcement is correctly implemented

### **Section 6: Error Handling Standards**
✅ **Framework**: Exit code matrix is theoretically sound
⚠️ **Implementation Gap**: No examples of how to categorize exceptions in practice
✅ **Structure**: Error response format matches system prompt requirements
⚠️ **Missing**: No guidance on logging to STDERR vs. including in JSON

### **Section 8: Workflow Context**
✅ **Concept**: 5-phase workflow accurately reflects system architecture
⚠️ **Detail Gap**: No examples of how tools interact across phases
✅ **Classification**: Phase requirements match system prompt documentation
⚠️ **Integration**: No guidance on manifest updates during workflow execution

### **Section 10: Testing Requirements**
✅ **Structure**: Test directory layout matches actual repository
✅ **Coverage**: Test categories reflect real testing priorities
⚠️ **Detail Gap**: No examples of how to test approval token enforcement
✅ **Patterns**: Pytest fixture examples match actual test implementations

## 🚀 Recommended Next Steps

1. **High Priority**: Create a validation script that checks guide examples against actual tool code
2. **Medium Priority**: Develop platform-specific examples for Windows path handling

This revised guide represents a **substantial improvement** but requires critical alignment with the actual codebase implementation before it can serve as the definitive reference. The governance foundation is excellent, but practical implementation details must match the codebase's actual behavior to prevent developer confusion and ensure tool compatibility.

The most urgent priority is resolving the **exit code inconsistency**, as these create immediate barriers to effective tool development. Once these are addressed, the guide will serve as an excellent foundation for the PowerPoint Agent tool ecosystem.

---

# Implementation Plan - Refine Tool Development Guide

## Goal Description
Refine `PowerPoint_Tool_Development_Guide.md` based on review comments to ensure perfect alignment with the codebase and v3.1.0 standards.

## User Review Required
> [!IMPORTANT]
> **PathValidator Rejection**: The review comment suggested adding `PathValidator`, but this class does not exist in the codebase. I will instead recommend standard `pathlib.Path` usage for platform safety.
> **Exit Code Alignment**: I will update the Master Template to use Exit Codes 0/1 as the default to match the codebase, while keeping the 0-5 matrix as an advanced standard.

## Proposed Changes

### Documentation
#### [MODIFY] [PowerPoint_Tool_Development_Guide.md](file:///home/project/powerpoint-agent-tools/PowerPoint_Tool_Development_Guide.md)

1.  **Master Template Updates**:
    - **Action**: Update the template code block.
    - **Change**: Revert default exit codes to 0/1 (Success/Error) to match codebase reality. Add comment about advanced codes.
    - **Change**: Add `validate_against_schema` import and usage example.

2.  **Governance Principles**:
    - **Action**: Update "2.3 Approval Token System".
    - **Change**: Clarify that this is a *requirement for new destructive tools* and mandated by System Prompt v3.0.

3.  **Shape Index Nuances**:
    - **Action**: Enhance "2.4 Shape Index Management".
    - **Change**: Explicitly list operations that invalidate indices (`add_shape`, `remove_shape`, `delete_slide`, `set_z_order`) and require refresh.

4.  **Probe Resilience**:
    - **Action**: Add "8.1 Probe Resilience Pattern" under Workflow Context.
    - **Change**: Document the timeout and transient slide pattern used in `ppt_capability_probe.py`.

5.  **Platform Safety**:
    - **Action**: Add "6.1 Platform-Independent Paths" under Error Handling.
    - **Change**: Recommend `pathlib.Path` over string manipulation.

## Checklist of Sections to Merge
- [ ] 2.3 Approval Token System (Clarification)
- [ ] 2.4 Shape Index Management (Enhanced)
- [ ] 3. The Master Template (Exit Codes 0/1, Schema Validation)
- [ ] 6.1 Platform-Independent Paths (New)
- [ ] 8.1 Probe Resilience Pattern (New)

## Verification Plan

### Manual Verification
- Review the updated guide to ensure:
    - Exit codes in template match 0/1.
    - No references to non-existent `PathValidator`.
    - `validate_against_schema` example is correct.
    - Probe resilience pattern matches `ppt_capability_probe.py`.

---

# Task: Refine Tool Development Guide

- [x] Analyze `comments_on_revised_tool_development_guide.md` <!-- id: 0 -->
- [x] Validate comments against codebase <!-- id: 1 -->
- [x] Create Implementation Plan <!-- id: 2 -->
- [x] Merge changes into `PowerPoint_Tool_Development_Guide.md` <!-- id: 3 -->
- [x] Verify updated guide <!-- id: 4 -->

