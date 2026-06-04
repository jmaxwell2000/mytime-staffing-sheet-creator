# Render Deployment

The initial app should deploy as a small Render web service with no database.

## Runtime

Suggested runtime:

```text
Python 3.12+
FastAPI
Uvicorn
```

Suggested start command:

```text
uvicorn app.main:app --host 0.0.0.0 --port $PORT
```

Suggested build command:

```text
pip install -r requirements.txt
```

## Required Endpoints

- `GET /`
- `POST /convert`
- `GET /healthz`

The health endpoint should not inspect files or depend on external services.

## Environment Variables

Initial version should not require secrets.

Optional non-sensitive settings:

```text
MAX_UPLOAD_MB
REQUEST_TIMEOUT_SECONDS
```

If settings are added, document defaults and keep the app runnable locally without secret configuration.

## Storage

Do not configure persistent disks for the initial version.

Uploaded files and generated workbooks should be held in memory or temporary request-scoped files only. Temporary files must be deleted after conversion.

## Resource Limits

The implementation should define a maximum upload size to avoid memory pressure.

When the upload is too large, return a clear error that does not echo the filename or workbook content.

## Deployment Verification

After deployment, verify:

- `/healthz` returns success.
- Upload page loads.
- Synthetic fixture workbook converts successfully.
- Generated workbook downloads automatically.
- Workbook opens locally.
- Generated workbook structure matches anonymized desired output expectations when such a reference exists.
- Front sheets print to one portrait letter page.
- Back sheets are omitted by default.
- Back sheets appear only when enabled.
- Offset changes back sheet top margin.
- Logs do not contain employee names, organization labels, uploaded filenames, or workbook cell values.

## Open-Source Release Check

Before publishing:

- Confirm no real input or output workbooks are committed.
- Confirm no screenshots contain real data.
- Confirm sample files use synthetic placeholders.
- Confirm generated filenames do not include organization names or employee names.
- Confirm docs avoid specific real example filenames.
