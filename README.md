# MyTime Staffing Sheet Creator

Current release: v2.1. Updated June 8, 2026.

This project is a lightweight open-source web app that converts a source scheduling export workbook into a print-ready staffing workbook. The app is intended for a nurse manager or charge nurse workflow: upload an `.xlsx`, convert it, download the finished staffing workbook, and print it without manual editing.

The implementation uses Python/FastAPI, `openpyxl`, simple HTML/CSS/JS, and no database.

## What The App Should Do

- Accept one uploaded source scheduling `.xlsx` workbook.
- Parse staff, dates, job codes, start times, and end times accurately.
- Generate one printable front staffing sheet per date/shift pair.
- Optionally generate paired lined `BACK` sheets for duplex printing.
- Automatically download the generated workbook to the user's machine.
- Avoid storing uploads, generated workbooks, employee names, schedule data, or organization data.

Accuracy is the central product requirement. Incorrect staffing information can create unsafe operational decisions, so the parser and output generator must be researched against provided examples before implementation and must fail clearly when the input cannot be interpreted with confidence.

Provided examples may include source scheduling exports, desired output workbooks, desired output PDFs, or any combination of those. Treat desired output `.xlsx` workbooks as the strongest reference for sheet structure, styles, margins, print settings, formulas, merged ranges, row heights, and tab order. Treat PDFs as visual print references.

## Open-Source Privacy Rule

This repository must never include real organization names, unit names, company names, employee names, patient names, vendor-identifying customer data, or real schedule content.

All committed code, tests, fixtures, docs, screenshots, generated examples, and sample workbooks must use synthetic neutral placeholders only. Real examples may be used locally for research, but they must remain uncommitted unless fully anonymized.

See [privacy-open-source-rules.md](docs/privacy-open-source-rules.md) before implementing anything.

## Suggested Project Shape

```text
app/
  main.py
  parser/
  workbook/
  templates/
  static/
tests/
  fixtures/
docs/
```

This repository contains the FastAPI app, parser, workbook generator, browser UI, deployment config, synthetic tests, and implementation docs.

## Handoff Docs

- [Product Brief](docs/product-brief.md)
- [Requirements](docs/requirements.md)
- [Privacy And Open-Source Rules](docs/privacy-open-source-rules.md)
- [Parsing Research Guide](docs/parsing-research-guide.md)
- [Workbook Output Spec](docs/workbook-output-spec.md)
- [Implementation Guide](docs/implementation-guide.md)
- [Test Plan](docs/test-plan.md)
- [Render Deployment](docs/render-deployment.md)

## Default Stack

- Python 3.12+
- FastAPI
- Uvicorn
- openpyxl
- pytest
- Plain HTML/CSS/JS

Use a database only if a later requirement explicitly adds persistent accounts, audit history, or saved settings. The initial version should keep all server-side processing request-scoped.

## License

MIT. See [LICENSE](LICENSE).
