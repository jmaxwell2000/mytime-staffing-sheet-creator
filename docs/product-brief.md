# Product Brief

Current release: v2.1. Updated June 8, 2026.

## Goal

Build a lightweight web app that turns a difficult-to-read source scheduling workbook into a clear, print-ready staffing workbook for a hospital unit workflow.

The user should be able to upload a scheduling export, run the conversion, and receive a finished `.xlsx` workbook that can be opened and printed immediately.

Development examples may include desired output as a workbook, a PDF, or both. A desired output workbook should be used to verify workbook structure and print settings directly; a PDF should be used to verify the printed visual result.

## Primary User

The primary user is a nurse manager or charge nurse who needs to understand staffing by date, shift, role, start time, end time, and lead responsibility.

The app should feel simple, calm, and operational. It should not behave like a dashboard or reporting suite. The main screen should focus on the upload, conversion, and a small collapsed options area.

## Success Criteria

- The converted workbook accurately reflects who is working, which role group they belong to, and their exact start/end times.
- The generated workbook matches the provided desired output examples in structure, sheet order, print behavior, and visual layout.
- Non-standard start/end times are preserved and made visually noticeable.
- Standard 12-hour shifts are normalized to the expected shift-change display times.
- The generated workbook is formatted for portrait letter printing and binder use.
- The user can print the output without manually resizing columns, adjusting margins, or editing sheet contents.
- Uploaded and generated files are not stored after the request completes.
- The open-source repository contains only synthetic neutral data.

## Safety And Accuracy

Incorrect schedules are dangerous. The implementation must treat ambiguity as a failure condition, not as something to guess through silently.

When the parser cannot confidently identify key fields, dates, names, job codes, or time ranges, it should return a clear error or warning that helps the user provide feedback and helps the developer improve the parser.

## User Workflow

1. User opens the web app.
2. User uploads a source scheduling `.xlsx`.
3. Optional: user opens the collapsed options menu.
4. Optional: user enables back pages and adjusts the back page offset.
5. User starts conversion after the button changes to `Create staffing sheets`.
6. App shows a clear `Creating...` processing state.
7. Browser downloads the generated staffing workbook and the button changes to `Downloaded`.
8. User can pick a new file to reset the interface for another conversion.

## Non-Goals For The Initial Version

- No database.
- No user accounts.
- No server-side schedule history.
- No stored uploaded files.
- No real sample names or real schedule fixtures in the repository.
- No broad analytics or reporting features beyond the printable workbook.
