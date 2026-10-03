---
name: code-spec-validator
description: Strict contract-auditing and test-execution protocol for validating generated VBA and algorithmic code against SPEC requirements. Execute before marking any task, module, or checklist item as completed.
version: 1.0.0
author: "Ícaro Caminoto"
---

# Code & Spec Contract Validator

## Core Philosophy: The Silent Contractor Rule
Code is never evaluated by "intent" or "cosmetic correctness." Code is evaluated strictly as a deterministic contract implementation. If code deviates from specified table schemas, makes assumptions about undefined states, leaks memory/application locks, or implements features not declared in the SPEC, it is defective and must be rejected.

---

## 4-Stage Verification Protocol

Before finalizing any implementation, the agent must run this audit sequentially:

### Stage 1: Contract & Schema Fidelity Check
- [ ] **Exact Entity Matching:** Compare all hardcoded names (Worksheets, ListObjects, Dictionaries, Output Columns) against the SPEC. Zero tolerance for renaming or drift (e.g., `Relatorio_Final` must never become `Tabela_Preco`).
- [ ] **Data Types & Signatures:** Confirm every function signature matches the architectural blueprint. No undeclared `Variant` parameters where strict types (`Long`, `Double`, `String`) were contracted.
- [ ] **I/O Schema Alignment:** Verify input readers ingest only specified columns and output dumpers write strictly the specified column order.

### Stage 2: Scope Creep & Dangling Reference Audit
- [ ] **Zero Hallucinated Features:** Confirm no helper functions, UI elements, or logic branches exist unless explicitly demanded by the SPEC.
- [ ] **No Incomplete Stubs:** Ensure no procedures contain placeholder comments (`' TODO: Implement here`), incomplete loops, or unhandled `Else` branches.

### Stage 3: Resilience & Edge-Case Guardrails
- [ ] **Empty Set Defense:** Verify code handles zero-row datasets gracefully (e.g., empty source tables, no transactions within the 90-day window) without throwing run-time errors.
- [ ] **Missing Resource Recovery:** Confirm parameter/lookup references invoke dynamic auto-provisioning if the source sheet/table does not exist.
- [ ] **State Restoration Trap:** Verify that EVERY exit path (normal or runtime error) passes through a cleanup block restoring `ScreenUpdating`, `Calculation`, and `EnableEvents`.

### Stage 4: Execution of Automated Test Harness
- [ ] Compile check: Ensure zero syntax or reference errors under `Option Explicit`.
- [ ] Execute the dedicated test harness (from `templates/`) against synthetic/mock data.
- [ ] Confirm all assertion tests pass with zero manual intervention.

---

## Test Execution Directives

When building and executing tests for verified code, adhere strictly to these testing principles:

1. **Non-Destructive In-Memory Isolation:**
   - Tests must run in isolated test workbooks, temporary memory structures, or dedicated sandboxed worksheets prefixed with `tmp_test_`.
   - Production sheets (`Relatorio_Final`, `Param_*`) must never be modified by a test run without a complete automated rollback.

2. **Deterministic Synthetic Fixtures:**
   - Tests must inject known, predictable datasets into memory arrays to validate calculation routines.
   - For rolling-window logic (e.g., 90-day average), fixtures must include:
     - Boundary cases: Transactions on `Date - 89`, `Date - 90`, and `Date - 91`.
     - Zero-case: SKUs with no transactions in the window (fallback validation).
     - Multi-case: SKUs with identical dates but different values (average resolution).

3. **State Assertion:**
   - Tests must assert that `Application.ScreenUpdating` and `Application.Calculation` return to their baseline states even when the target routine encounters an artificial error.