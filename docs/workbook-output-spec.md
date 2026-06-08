# Workbook Output Spec

The generated workbook contains one printable front sheet for each date/shift pair. If back pages are enabled, add a matching lined `BACK` sheet immediately after each front sheet.

Tab order is newest to oldest. For each date, `NIGHT` comes before `DAY`.

If a desired output `.xlsx` workbook is provided during implementation, use it as the primary reference for workbook structure and print settings. This spec defines the intended behavior, but the example workbook can clarify dimensions, styling, and exact print-readiness details. If a desired output PDF is provided, use it as the visual print reference.

## Sheet Names

Use this pattern:

```text
Sun Mar29 NIGHT
Sun Mar29 NIGHT BACK
Sun Mar29 DAY
Sun Mar29 DAY BACK
```

When back pages are disabled, omit the `BACK` sheets.

## Front Sheet Layout

Each front sheet is portrait, letter-size, and printable. The visible print area is columns `B:H`.

Column layout:

```text
B = initial column
C = employee name
D = Start
E = End
F = first assignment/notes period
G = second assignment/notes period
H = third assignment/notes period
```

Column `B` should be wide enough that ordinary staffing-row cells read as square. Its section-header label is `Initial`, using small shrink-to-fit text so it fits the narrow box without changing the column width. Managers and Charge RNs use this column to initial staffing changes. Column `C` should be wide enough for long names, Start/End columns should fit larger four-digit times, and assignment-period columns should adjust evenly so the overall `B:H` print width stays `141.69` width units.

At the top of each front sheet, show a left-side reminder:

```text
Use pencil. Write and initial notes on the back of this sheet.
```

The date/shift title remains on the top row:

```text
Front: Saturday Day Shift ◯ 5/23
Front: Saturday Night Shift ⬤ 5/23
Back: 5/23 ◯ Saturday Day Shift
Back: 5/23 ⬤ Saturday Night Shift
```

Use the actual full weekday, shift, light/dark shift circle, and date for each sheet. The shift wording is regular 16pt, the circle is bold 22pt, and the date is bold 22pt, with exactly one space before and after the circle. The same front title appears at the top and bottom of each front sheet; the matching back sheet mirrors the order as date, circle, readable shift wording.

## Section Headers

All front-sheet table borders use the same thin black line weight, including section headers.

Day shift header:

```text
RN
Start | End | 0700-1100 | 1100-1500 | 1500-1900
```

Night shift header:

```text
RN
Start | End | 1900-2300 | 2300-0300 | 0300-0700
```

The PSA section uses the same time headers for the sheet's shift.

## Staff Sections

Use fixed blank capacity so the printed page has writing space:

```text
RN section: 16 staff/blank rows total
PSA section: 11 staff/blank rows total
Other section: only included if unknown job codes exist
```

If the RN or PSA section is filled to capacity, add one extra blank writable row below that section.
Writable staff rows should be taller than the section headers so pencil edits have enough vertical room, and matching back-page rows must copy the same heights for front/back alignment.

Employee rows show:

- Employee display name.
- Start time.
- End time.
- Optional centered `LEAD` in the first assignment/notes period column.
- Optional `(CL)` after the name when multiple clinical leaders are present.

Employee display-name cells calculate font size from the final displayed text, including `(CL)` when present, so long names stay inside column `C` without changing column width or row height. Single clinical leaders display `LEAD` separately, so their names do not include `(CL)`.

Committed generated examples must use synthetic employee placeholders only.

## Time Display

- Format times as four-digit 24-hour strings, such as `0700`, `1930`, `1400`.
- Preserve exact non-standard start and end times.
- Make non-standard start/end values yellow highlighted while keeping standard size 14, non-bold text.
- If an end time is exactly `1900`, display it as `1930`.
- If an end time is exactly `0700`, display it as `0730`.

## Patient Fall Review Form

At the bottom of every front sheet, include a compact, readable, light-yellow `Patient Fall Review` form with:

- Date and brief Charge RN instructions combined in the title row.
- A wide question column with bullet-point questions.
- A wide answer/notes space for written responses.
- The following questions:

```text
- Bathroom or bedside commode?
- Staff or family with patient at time of fall?
- If no, should someone have been present?
- Any injury?
```

Add the final prevention-opportunities row:

```text
Post-fall measures (circle):         Alarm          Nonskid socks          Fall sign          Stay with pt toileting          Fall risk score          Other
```

Keep the prevention-opportunities row on one line because the user only needs to circle items. Separate options with generous spacing rather than pipe characters, and leave blank space after `Other` instead of using underscores.
Question and answer text should be indented enough that it does not sit against the table lines. Do not shrink the question text to solve spacing issues; give rows enough height and the question column enough width to read cleanly.

## Back Sheets

Back sheets are generated only when the user enables back pages.

Each `BACK` sheet is a lined notes page paired with the front sheet. It should include:

- Same date/shift title in gray near the top, aligned to the left edge.
- `Use pencil` in gray near the top-right.
- Centered `Notes (Please initial)` line near the top.
- Light solid horizontal writing lines.
- Same date/shift title repeated at the bottom, aligned to the left edge.
- A small footer below the writing lines and above the bottom title:

```text
Created by MyTime Staffing Sheet Creator v2.1
```

Back-page text should be light gray so it remains readable without competing with handwritten notes.

Back sheet top margin:

```text
0.15 inches + offset
```

Offset range:

```text
0.05 inches to 0.45 inches effective top margin
```

Positive offset values move notes content down. Negative values move notes content up.

## Print Settings

Apply to front and back sheets:

- Portrait orientation.
- Letter paper.
- Fit to one page wide.
- Fit to one page tall.
- Binder-safe left margin around `0.75"` on front sheets.
- Mirrored binder-safe right margin around `0.75"` on back sheets.
- Printer-safe non-binder side margin around `0.25"` on front and back sheets.
- Narrow top and bottom margins.
- Print area should fit columns `B:H` for front sheets.
- Assignment columns should be wide enough that Excel 2016 and later prints the full-height sheet from the binder-safe edge to near the opposite page edge, while leaving enough non-binder margin to avoid printer clipping.
- Front-sheet fall review and bottom date should sit low on the printed page, using available vertical space for taller staff rows before the fall review.

The generated workbook should be immediately printable without manual Excel adjustments.
