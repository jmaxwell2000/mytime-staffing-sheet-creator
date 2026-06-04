# Parsing Research Guide

The parser is the riskiest part of this project. The future implementation agent must analyze provided source export examples before writing final parser logic.

The output generator is also example-driven. If desired output examples are provided as `.xlsx` workbooks, inspect them directly. If desired output examples are provided as PDFs, use them as visual print references. If both are available, use the workbook for exact structure and the PDF for visual confirmation.

Do not assume the raw workbook structure is stable. The source export layout may change, and the app should be resilient where practical.

## Research Goals

Determine how to identify:

- Schedule dates.
- Shift groupings.
- Employee display names.
- Job codes.
- Start times.
- End times.
- Lead-pay indicators.
- Rows or cells that should be ignored.
- Multiple sheets, hidden sheets, merged cells, formulas, and formatting patterns.

Document the discovered structure in implementation notes using generic descriptions only. Do not quote real employee names, real organization labels, or exact real workbook filenames.

## Recommended Process

1. Inspect workbook metadata and sheet names.
2. List visible sheets, hidden sheets, merged ranges, dimensions, and non-empty cell counts.
3. Identify candidate date, name, job-code, and time fields.
4. Compare more than one source example if available.
5. Identify which structure is stable and which structure varies.
6. Inspect any desired output workbook to identify exact sheet structure, styles, dimensions, print settings, and tab order.
7. Compare any desired output PDF against the workbook structure as a printed visual reference.
8. Build synthetic fixtures that reproduce the discovered patterns without real data.
9. Write parser and workbook tests before broadening the implementation.
10. Add warnings for ambiguous or unsupported shapes.

## Parser Design Principles

- Prefer evidence-based detection over fixed cell coordinates.
- Use headers, date-like values, time-like values, job-code sets, and nearby context together.
- Preserve source evidence in internal records for debugging.
- Fail clearly when required fields cannot be identified.
- Do not silently skip unknown job codes.
- Do not silently coerce unclear times.
- Keep parsing separate from workbook generation.

## Suggested Inspection Helpers

Build small local scripts or tests that can summarize a workbook without printing sensitive values. For example, summarize:

- Sheet count.
- Sheet visibility.
- Used range dimensions.
- Count of date-like cells.
- Count of time-like cells.
- Count of known role-code cells.
- Merged range count.

If values must be shown during local debugging, do not commit that output. Prefer redacted summaries in docs.

For desired output `.xlsx` examples, summarize:

- Sheet names and order.
- Print area.
- Page setup and margins.
- Column widths and row heights.
- Merged ranges.
- Freeze panes, if any.
- Cell fills, borders, fonts, and alignments.
- Representative formula counts, if any.

Do not commit a real desired output workbook unless it has been fully anonymized.

## Required Parser Outputs

The parser should produce normalized records shaped like:

```text
employee_display_name
work_date
shift_kind
start_time
end_time
job_code
role_group
is_clinical_leader
has_lead_pay
source_evidence
warnings
```

`employee_display_name` can contain real uploaded data at runtime, but no committed fixture should use real names.

## Ambiguity Handling

If parsing confidence is low, return a user-facing error or warning. The message should be useful without leaking sensitive values.

Examples of safe messages:

```text
Could not identify a schedule date in one worksheet.
Found a row with a known job code but no readable start time.
Some records used unknown job codes and were placed in Other.
```

Do not include raw names, organization labels, or full row contents in these messages.
