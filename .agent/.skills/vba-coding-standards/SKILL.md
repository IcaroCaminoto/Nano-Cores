---
name: vba-coding-standards
description: Mandatory engineering standards and practices for developing VBA 7.1 (64-bit) routines in Microsoft Excel. Use whenever generating, refactoring, or auditing VBA code.
version: 1.0.0
author: "Ícaro Caminoto"
---

# VBA Process Guidelines

## Engineering Standards & Mandatory Methods
1. **Resilience & Auto-Provisioning:** If parameter/lookup sheets do not exist, create them dynamically with defaults—based on the sources provided in the spec.md file-rather than throwing unhandled exceptions.

2. **Integration with Existing Codebase:** Integrate with and respect existing modules located inside `Nano Cores > src > vba`.

3. **Value Conversion:** Implement consistent value conversion logic to ensure data integrity across different data types and sources.
  - **`,` to `.` Conversion:** Convert all decimal commas to decimal points for numeric calculations when ingesting data and exporting results. Remember to convert back to commas when writing to Excel cells and printed reports, if the locale requires it.

4. **Network Path Validation:** If the network path does not exist, abort the routine and raise `Err.Raise` with a descriptive message requesting network connection.