### Airborne Particle Counts – Measurement Calculator (Excel/VBA)

This workbook is an Excel-based measurement and reporting tool designed to standardize airborne particle count calculations and compliance evaluation across users and sites. It combines structured Excel tables, formula-driven calculations, and VBA automation to reduce manual data processing and minimize reporting errors.

### What the calculator does

- Converts raw airborne particle count data into normalized results (e.g., counts per cubic meter) using workbook-defined sampling parameters.

- Evaluates results against acceptance criteria and produces clear PASS/FAIL outputs (with “N/A” handling when inputs are incomplete or invalid).

- Standardizes the data entry workflow so users input data only in designated fields, while calculation logic remains protected from accidental edits.

### Workbook structure

**Instructions**

User-facing guidance for operating the tool correctly (intended to reduce training burden and ensure consistent usage).

**CleanroomData**

Centralized input/configuration area. This typically stores shared parameters (e.g., sampling volumes, number of samples, test metadata).
Notably, the workbook uses a single source of truth for the number of rows/samples required, which then drives downstream tables.

**AirborneCounts**

Primary working sheet where users review/enter particle results and receive computed outputs and compliance status.
This sheet is also where the workbook applies strong guardrails to prevent editing calculation cells.

**CalcData**

Backend calculations and supporting logic. This layer keeps calculation formulas organized and separated from user-facing workflow, improving maintainability and reducing the chance of accidental changes.

### Automation & guardrails (VBA)
**1) Controlled editing (worksheet event code)**

The AirborneCounts sheet uses event handlers to enforce intentional data entry:

- Allows input only in highlighted (yellow) entry cells
        If a cell appears yellow due to conditional formatting, the change is accepted.

- Reverts edits to protected (black) cells
        If a user tries to change a black-filled cell (#262626), the macro restores the prior value (originalValue). This prevents accidental modification of calculated or controlled fields.

- Prevents selection of restricted regions
        If the user selects A1 or A4:G9, the macro redirects selection to B1. This blocks interaction with key control areas and keeps the workflow on track.

Purpose: These controls reduce user error and help ensure that calculated outputs remain consistent across different users and sites.

**2) Dynamic table resizing + standard formatting (NumberTablesAndResize)**

The workbook includes a macro that scales multiple Excel Tables to match a required number of samples/rows.

Core behaviors:

- Reads the desired row count from AirborneCounts!A1 (linked to an upstream input in CleanroomData).

- Iterates through all ListObject tables on AirborneCounts.

- For each table:

-       Resizes the table to the specified row count (expand or shrink).

-       Clears contents and removes formatting outside the resized table, but within defined maximum bounds (per table), keeping the worksheet clean.

-       Auto-numbers the first column from 1…N to preserve indexing consistency.

-       Propagates formulas down the resized table so newly created rows calculate correctly.

-       Copies formatting from the first row to all rows so the table stays visually standardized.

It includes explicit max ranges for each table (Table7 through Table13) so expansion/shrinking is predictable and doesn’t impact unrelated regions of the worksheet.

Purpose: Users can change the number of samples without manually inserting rows, copying formulas, or rebuilding formatting—reducing manual processing time and improving repeatability.

### Calculation logic

From the formula patterns, the tool performs:

- Normalization / conversion from raw counts to standardized volumetric results (commonly using a *1000 / sample_volume_L pattern to produce counts per cubic meter).

- Compliance evaluation by comparing computed values to limits/criteria and returning:

-      PASS if results meet the threshold

-      FAIL if results exceed the threshold

-      N/A when inputs are missing/invalid (using error handling in formulas)

This combination of normalization + automated compliance flags is what makes the workbook effective as a standardized reporting tool.

### How to use

1. Open the workbook in desktop Excel (macros must be enabled).
2. Follow the Instructions tab to enter required sampling setup values (typically in CleanroomData).
3. Set the required number of samples/rows (drives AirborneCounts!A1).
4. Run the table sizing macro (if not triggered by a button/automation in your version):
           NumberTablesAndResize
5. Enter data only in the designated input cells (yellow-highlighted fields).
6. Review calculated outputs and PASS/FAIL results.


### Notes / constraints

This workbook relies on VBA events and Excel Table objects (ListObject). It is intended for desktop Excel, not Excel Online.

Guardrails depend on cell formatting (yellow via conditional formatting, black fills for protected areas). Changes to formatting rules may change behavior.

If you distribute the file, consider adding:

a “version” cell and changelog

a signed macro / trusted location guidance (if your org supports it)
