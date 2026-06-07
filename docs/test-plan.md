# Test Plan

Testing should prioritize parser accuracy, workbook print-readiness, desired-output fidelity, and privacy.

All test data must be synthetic. Do not commit real source exports or real generated staffing workbooks.

## Parser Tests

Use synthetic fixtures that reproduce discovered source workbook structures.

Cover:

- Date extraction.
- Shift classification.
- Employee display name extraction using synthetic placeholders.
- Job-code extraction.
- Start and end time extraction.
- Hidden sheets or irrelevant sheets when present.
- Merged cells or formulas when present.
- Ambiguous structure failures.
- Unknown job codes routed to `OTHER`.

## Role And Lead Tests

Verify RN codes:

```text
317N
317W
3178
8536
```

Verify PSA codes:

```text
5001
5056
5057
```

Verify excluded codes:

```text
4120
```

Verify lead rules:

- One `8536` clinical leader displays `LEAD`.
- Multiple `8536` clinical leaders display `(CL)` after each name.
- Lead-pay indicators do not create `LEAD` when a clinical leader is working.
- Lead-pay indicators create `LEAD` when no clinical leader is working.

## Time Tests

Verify:

- Four-digit 24-hour display formatting.
- Standard day shift display as `0700` to `1930`.
- Standard night shift display as `1900` to `0730`.
- End time `1900` displays as `1930`.
- End time `0700` displays as `0730`.
- Non-standard start times are preserved.
- Non-standard end times are preserved.
- Non-standard time cells are yellow highlighted with standard size 14, non-bold text in the generated workbook.

## Workbook Tests

Check generated `.xlsx` files with `openpyxl`.

Verify:

- Sheet order newest to oldest.
- For each date, `NIGHT` comes before `DAY`.
- `BACK` sheets are omitted when back pages are disabled.
- `BACK` sheets immediately follow matching front sheets when enabled.
- Print area uses the expected visible columns.
- Column `B` is wide enough for square initials cells while total `B:H` width stays fixed.
- Column `C` is widened for names and Start/End columns are widened for 14pt four-digit times.
- Front and back date titles use rich text with full weekday/shift wording, day/night circle, and mirrored back-page ordering.
- RN section has 16 staff/blank rows total.
- PSA section has 11 staff/blank rows total.
- RN and PSA sections add one extra blank writable row when filled to capacity.
- Front and matching back-page row heights are identical so writing lines align.
- Back writing lines are solid thin and light gray.
- Back-page text is light gray, including top title, reminder, notes label, footer, and bottom title.
- Back-page blank-space message appears centered immediately below the writing lines.
- Back-page footer contains the MyTime creator/version label in small text.
- `Other` section appears only when unknown codes exist.
- Fall information form is present on every front sheet.
- Fall prevention options are evenly spaced, contain no underscores, and leave open write-in space after `Other`.
- Portrait letter print settings are applied.
- Fit-to-page settings are applied.
- Binder-safe and printer-safe margin expectations are met.
- Synthetic verification samples include day and night front/back sheets, long names that trigger automatic shrink-to-fit, and non-standard start/end times.

## Desired Output Example Tests

If a desired output `.xlsx` workbook is provided, compare generated synthetic output against its non-sensitive structural properties:

- Sheet naming pattern.
- Sheet ordering.
- Print areas.
- Page setup.
- Margins.
- Column widths.
- Row heights.
- Merged ranges.
- Core style patterns.
- Front/back sheet pairing.

If a desired output PDF is provided, use it for visual print checks only. Do not rely on the PDF as the sole source for workbook internals when a workbook example is also available.

Never commit the real desired output workbook or PDF unless it has been fully anonymized.

## Offset Tests

Verify:

- Default offset is `0.00`.
- Minimum accepted offset is `-0.10`.
- Maximum accepted offset is `0.30`.
- Offset step in UI is `0.01`.
- Effective back sheet top margin is `0.15 + offset`.
- Negative offset moves content up by reducing the margin.
- Positive offset moves content down by increasing the margin.
- Offset is stored in browser local storage, not server state.

## Endpoint Tests

Cover:

- `GET /` returns the upload page.
- `GET /healthz` returns a basic success response.
- `POST /convert` rejects missing files.
- `POST /convert` rejects non-`.xlsx` files.
- `POST /convert` returns an `.xlsx` download for valid synthetic input.
- `POST /convert` does not include sensitive values in error messages.

## Privacy Regression Tests

Add tests or checks to prevent accidental sensitive fixture content.

Recommended checks:

- Search committed fixtures for realistic proper-name patterns.
- Search docs and code for disallowed real-data markers before commit.
- Verify logs from tests do not contain synthetic employee display values.
- Verify generated sample output fixtures use only neutral placeholders.

If real examples are used locally, verify they are ignored and not staged before committing.
