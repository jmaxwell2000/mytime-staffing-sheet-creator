# Requirements

Current release: v2.0. Updated June 3, 2026.

## Input

The app accepts a source scheduling `.xlsx` workbook. The future implementation agent must research the workbook structure from provided examples before writing parser logic.

Do not hardcode assumptions about the exact raw workbook layout until the examples have been analyzed and the discovered structure has been documented.

Provided desired output examples may be `.xlsx` workbooks, PDFs, or both. If a desired output workbook is provided, use it to inspect exact workbook structure, sheet names, styles, merged cells, dimensions, margins, row heights, column widths, page setup, and tab ordering. If a PDF is provided, use it as a print-layout reference.

## Shift Times

- Day shift: 7 a.m. to 7 p.m.
- Night shift: 7 p.m. to 7 a.m.

Display assignment periods as:

- Day: `0700-1100`, `1100-1500`, `1500-1900`
- Night: `1900-2300`, `2300-0300`, `0300-0700`

## Role Codes

RN codes:

```text
317N
317W
3178
8536
```

PSA codes:

```text
5001
5056
5057
```

The code `8536` is treated as RN and also flagged as `is_clinical_leader`.

The code `4120` (`Staff LPN`) is explicitly excluded from generated staffing sheets.

Unknown job codes must not be discarded. They should appear in an `Other` section when present.

## Lead Rules

Clinical leader and lead-pay display rules:

- If exactly one `8536` clinical leader is working on a shift, that person displays `LEAD`.
- If two or more `8536` clinical leaders are working on a shift, each displays `(CL)` after the name instead of `LEAD`.
- If any clinical leader is working, lead-pay indicators do not create `LEAD` for other staff.
- If no clinical leader is working, staff with `LP`, `LP - LEAD PAY`, or `LEAD PAY` display `LEAD`.

The parser should preserve the source evidence used to determine clinical leader and lead-pay status so tests and debugging can explain why a label was applied.

## Time Formatting

- Times display as four-digit 24-hour values: `0700`, `1930`, `1400`.
- Exact short-shift start and end times must be preserved.
- Non-standard start or end times should keep standard size 11 text and stand out with yellow highlighting only.
- If an end time is exactly `1900`, display it as `1930`.
- If an end time is exactly `0700`, display it as `0730`.
- Standard 12-hour day shifts should display as `0700` to `1930`.
- Standard 12-hour night shifts should display as `1900` to `0730`.

## Web UI

- Single upload control for a source scheduling `.xlsx`.
- Primary conversion button.
- Collapsed options menu by default.
- `Back page offset` control inside options.
- Offset feature disabled by default.
- When offset is disabled, generate front sheets only.
- When offset is enabled, generate paired `BACK` sheets.

Offset control:

```text
Label: Back page offset
Default: 0.00
Minimum: -0.10
Maximum: 0.30
Step: 0.01
Unit: inches
```

Help text:

```text
Moves the back-page lines up or down so they align with the front sheet when double-sided pages are viewed through paper.
```

Calibration guidance:

```text
If back lines print too high, increase the offset.
If back lines print too low, decrease the offset.
```

Persist the offset locally in the browser. Do not persist it on the server.

## Backend Endpoints

- `GET /` serves the upload page.
- `POST /convert` accepts the workbook upload plus `include_back_pages` and `back_page_offset`.
- `GET /healthz` returns a basic health response for deployment checks.

The `/convert` response should be an `.xlsx` download. The filename should be based on the schedule date range only and must not include organization names, employee names, or source workbook filenames.

## Data Handling

- Do not store uploaded files after conversion.
- Do not store generated files after the response is sent.
- Do not log employee names, organization identifiers, schedule cell contents, uploaded filenames, or generated workbook contents.
- Committed fixtures and sample outputs must use synthetic neutral placeholders only.
