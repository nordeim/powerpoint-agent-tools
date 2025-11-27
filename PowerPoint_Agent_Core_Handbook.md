# PowerPoint Agent Core: Programming Handbook (v3.1.0)

**Version:** 3.1.0  
**Library:** `core/powerpoint_agent_core.py`  
**License:** MIT  

---

## 1. Introduction

The `PowerPointAgent` core library is the foundational engine for the PowerPoint Agent Tools ecosystem. It provides a **stateless, atomic, and security-hardened** interface for manipulating `.pptx` files. Unlike the raw `python-pptx` library, this core handles file locking, complex positioning logic, accessibility compliance, and operation auditing (versioning).

### 1.1 Key Capabilities (v3.1.0)
*   **Context-Safe**: Handles file opening/closing/locking automatically.
*   **Observability**: Tracks presentation state via deterministic SHA-256 hashing (Geometry + Content).
*   **Governance**: Enforces "Approval Tokens" for destructive actions (`delete_slide`, `remove_shape`).
*   **Visual Fidelity**: Implements XML hacks for features missing in `python-pptx` (Opacity, Z-Order).
*   **Accessibility**: Built-in WCAG 2.1 AA checking and Color Contrast calculation.

---

## 2. Usage Pattern (The "Hub" Model)

Tools interacting with this core **must** use the Context Manager pattern to ensure file safety.

```python
from core.powerpoint_agent_core import PowerPointAgent

# Atomic Operation Pattern
with PowerPointAgent(filepath) as agent:
    # 1. Acquire Lock & Load
    agent.open(filepath)
    
    # 2. Mutate (Capture return dict)
    result = agent.add_shape(...)
    
    # 3. Save (Atomic Write)
    agent.save()
    
    # 4. Release Lock (Automatic on exit)