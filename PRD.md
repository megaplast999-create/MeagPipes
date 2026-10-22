# Mega Pipes — Ralph Loop Task Specification

## Objective
Perform a read-only technical audit of the existing Mega Pipes quotation project.

## Constraints
- Do not modify, create, rename, or delete application source files.
- Do not modify the Excel workbook or embedded VBA.
- Do not modify existing backups, exports, quotes, or scratch files.
- Do not run Git reset, clean, checkout, restore, or commit commands.
- Do not change quotation formulas, pricing, GST, freight, margins, rounding, plant selection, or offer terms.
- Follow PROJECT_RULES.md and the existing project-specific instructions.
- Treat existing uncommitted changes as intentional until verified otherwise.

## Task 1: Inspect the architecture
- Identify the Python calculation engine, VBA modules, workbook, tests, and project instructions.
- Explain the relationship between the Python tests and the Excel/VBA quotation workflow.
- Report important entry points and dependencies where they can be established from source.

## Task 2: Review test coverage and risks
- Review the existing tests without editing them.
- Identify critical quotation rules that appear covered and those that may lack coverage.
- Separate verified findings from assumptions.
- Do not claim that workbook behavior is verified solely because Python tests pass.

## Task 3: Produce an audit report
Report:
1. Project architecture.
2. Test files and their purposes.
3. Baseline test command and result.
4. Potential correctness and regression risks.
5. Recommended next steps, ranked by priority.

## Acceptance criteria
- All three tasks are addressed.
- Findings reference actual files or test names where possible.
- Unknowns and limitations are explicitly identified.
- No application files, workbook files, VBA modules, backups, or Git state are changed.
- Existing tests remain runnable.

## Completion
Stop after producing the audit report. Do not implement recommendations.
