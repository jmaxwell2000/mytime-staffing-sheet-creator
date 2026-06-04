# Implementation Guide

Use this guide after completing the parsing and desired-output research described in [parsing-research-guide.md](parsing-research-guide.md).

## Default Stack

- Python 3.12+
- FastAPI
- Uvicorn
- openpyxl
- pytest
- Plain HTML/CSS/JS

Use no database for the initial version.

## Suggested App Structure

```text
app/
  main.py
  parser/
    models.py
    parse_workbook.py
    normalize.py
  workbook/
    build_workbook.py
    styles.py
    sheets.py
  templates/
    index.html
  static/
    app.css
    app.js
tests/
  fixtures/
```

All fixture data must be synthetic.

## Backend Flow

`GET /`

- Render the upload page.
- Include collapsed options menu.
- Load previously selected offset from browser local storage only.

`POST /convert`

- Accept one `.xlsx` file.
- Accept `include_back_pages`.
- Accept `back_page_offset`.
- Validate offset range when back pages are enabled.
- Read workbook into request-scoped memory or a temporary file.
- Parse schedule records.
- Generate staffing workbook.
- Return generated workbook as a download.
- Delete temporary files before the request finishes.

`GET /healthz`

- Return a small JSON response for deployment checks.

## Internal Record Model

Create a normalized internal model before workbook generation:

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

`role_group` should be one of:

```text
RN
PSA
OTHER
```

Runtime records may contain employee names from the upload. Source code, docs, fixtures, tests, sample files, and logs must not.

## Parsing Layer

Keep parser logic independent from output formatting.

Parser responsibilities:

- Identify dates, staff, job codes, start times, end times, and lead indicators.
- Normalize times to internal time objects.
- Classify role groups.
- Attach warnings when source information is incomplete or ambiguous.
- Preserve enough source evidence for debugging without logging sensitive values.

The parser should not decide workbook styling.

## Workbook Layer

Workbook builder responsibilities:

- Sort sheet order newest to oldest.
- For each date, emit `NIGHT` before `DAY`.
- Build front sheets.
- Build back sheets only when enabled.
- Apply sheet names, titles, sections, row capacity, fall form, styles, margins, and print settings.
- Return bytes suitable for an `.xlsx` response.

If a desired output workbook is available, create inspection tests or scripts that compare generated synthetic output against the example's non-sensitive structural properties: sheet ordering, print areas, margins, page setup, row heights, column widths, merged ranges, and style patterns. Do not copy real names or schedule values from the example into committed fixtures.

## Filename Generation

Generated download filenames should use schedule date range only.

Safe examples:

```text
staffing-sheets-2026-03-29-to-2026-04-04.xlsx
staffing-sheets-2026-03-29.xlsx
```

Do not include:

- Uploaded source filename.
- Organization names.
- Unit names.
- Employee names.

## UI Notes

The UI should be simple and direct:

- Upload control.
- Convert button.
- Processing state.
- Collapsed options area.
- Offset control shown only when back pages are enabled.

Do not include in-app text that exposes internal implementation details. Keep user-facing copy short and operational.

## Error Handling

Return clear messages for:

- Missing file.
- Non-`.xlsx` file.
- Unsupported workbook structure.
- Missing required schedule fields.
- Ambiguous parser output.
- Offset outside allowed range.

Error messages must not include names, full row values, organization labels, or raw source workbook content.

## Privacy-Safe Logging

Safe logs:

```text
conversion_started
conversion_completed
parsed_record_count=24
warning_count=2
duration_ms=850
```

Unsafe logs:

```text
employee names
organization labels
raw row contents
uploaded filenames containing identifying text
generated workbook cell values
```

## Render Notes

Use environment variables only for non-sensitive runtime settings such as max upload size. No secrets are expected for the initial app.

The app should work as a small web service with a single process and request-scoped conversion.
