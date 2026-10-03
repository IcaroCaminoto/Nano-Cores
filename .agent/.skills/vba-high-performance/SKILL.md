---
name: vba-high-performance
description: Mandatory engineering standards and constraints for developing high-performance VBA 7.1 (64-bit) routines in Microsoft Excel. Use whenever generating, refactoring, or auditing VBA code.
---
# VBA Performance Guidelines

## Critical Execution Constraints

1. **Zero UI Blocking:** All public routines must wrap execution within state handlers, ensuring mandatory state restoration in an exit/cleanup block (`ErrorHandler` / `Finally`):
   - `Application.ScreenUpdating = False`
   - `Application.Calculation = xlCalculationManual`
   - `Application.EnableEvents = False`
2. **In-Memory Operations:**
   - Direct cell iteration (`For Each Cell In Range...`) is strictly prohibited.
   - Raw data reading must be performed via a direct dump into a `Variant` array (`DataArray = Range.Value2`).
   - Mappings, parent-child relationships, and lookups must strictly utilize hash tables via `Scripting.Dictionary`.
   - Writing transformed data must be consolidated in memory and flushed via a single-shot block assignment (`Range.Resize.Value = OutputArray`).
3. **Strict Typing:** `Option Explicit` is mandatory at the top of every module, class, or UserForm. Every variable must have an explicitly declared data type. Avoid `Variant` unless strictly necessary for array operations.
4. **Closing Resources:** Always close and destroy ADO connections and recordsets after use to prevent memory leaks and ensure optimal performance.

## High-Volume Array Processing

When handling ranges with more than 1,000 rows, strictly follow this procedure:

1. **State Locking:** Toggle application state (`ScreenUpdating = False`, `Calculation = xlCalculationManual`).
2. **Bulk Read:** Transfer entire range to a 2D `Variant` array using `.Value2`.
3. **Allocation:** Pre-allocate the output array using `ReDim OutputArray(1 To UBound(InputArray, 1), 1 To TargetColumns)`.
4. **Traversal:** Iterate sequentially through the array indices (`For i = 1 To UBound(...)`), avoiding any COM object access inside the loop.
5. **Single-Shot Flush:** Paste data back in one operation using `TargetRange.Resize(UBound(OutputArray, 1), UBound(OutputArray, 2)).Value2 = OutputArray`.
6. **State Cleanup:** Always restore application state in the cleanup/exit block.