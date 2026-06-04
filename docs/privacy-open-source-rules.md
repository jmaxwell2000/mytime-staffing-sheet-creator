# Privacy And Open-Source Rules

This project is intended to be open source. Treat privacy and data minimization as product requirements, not cleanup tasks.

## Never Commit

Do not commit:

- Real organization names.
- Real department, unit, floor, or location names.
- Real company names from customer data.
- Real employee names.
- Real patient names.
- Real schedule data.
- Real uploaded workbooks.
- Real generated staffing workbooks.
- Real desired output workbooks.
- Real desired output PDFs.
- Screenshots containing real data.
- Logs containing source workbook values.
- Notes that quote real source workbook content.

## Synthetic Data Only

All committed examples must use synthetic neutral placeholders, such as:

```text
Employee 001
Employee 002
Unit Placeholder
Source Workbook A
```

Avoid realistic proper names in tests and fixtures. Neutral placeholders make accidental leakage easier to detect.

## Local Research Rules

Real examples may be inspected locally only when needed to understand parsing behavior. When doing that work:

- Keep real files outside the committed repository, or mark them ignored before use.
- Do not copy real names or schedule rows into docs, code, tests, issue notes, or comments.
- When documenting discovered parser behavior, describe structure generically.
- If a fixture is needed, create a synthetic workbook that reproduces the structure without real values.

## Logging Rules

Production and test logs must not include:

- Employee display names.
- Organization identifiers.
- Full row contents.
- Full cell contents.
- Uploaded filenames if they may contain identifying details.
- Generated workbook filenames if they include identifying details.

Safe logs may include:

- Request started/completed.
- File size bucket or exact byte size.
- Number of parsed records.
- Number of warnings.
- Conversion duration.
- Generic parser error codes.

## Generated Output Rules

At runtime, the generated workbook may include employee display names from the uploaded workbook because the user needs that output. That data must exist only in request-scoped memory or temporary files and must be deleted after the response completes.

Committed generated examples must contain synthetic placeholders only.

## Review Checklist

Before committing, inspect changed files for accidental sensitive content:

```text
rg -n "Employee [A-Z][a-z]+|Patient|Unit|Hospital|Clinic|Company" .
```

The exact scan should be adjusted for the repository, but every commit should include some privacy review. If a real example was used during development, verify it is not staged.
