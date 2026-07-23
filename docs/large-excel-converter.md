# Large Excel Converter deployment

Cargo Statement Analyser keeps normal `.xlsx` and `.csv` analysis in the browser. The optional Large Excel Converter is only used when the browser preflight safety limit is exceeded.

## Service

Run the lightweight Node service:

```bash
npm ci
npm run converter
```

The service accepts `POST /convert-xlsx`, extracts only the case-insensitive `Digital Sales` worksheet, returns UTF-8 CSV, and deletes its temporary upload directory in a `finally` cleanup. It does not analyse statements, calculate totals, use a database, authenticate users, or keep historical uploads.

## Environment variables

Browser app:

- `VITE_LARGE_XLSX_CONVERTER_URL` - converter endpoint, default `/api/convert-xlsx`.
- `VITE_LARGE_XLSX_BROWSER_SIZE_LIMIT` - browser XLSX byte safety limit, default `10485760`.
- `VITE_LARGE_XLSX_BROWSER_ROW_LIMIT` - browser row safety limit, default `50000`.
- `VITE_LARGE_XLSX_CONVERTER_TIMEOUT_MS` - client timeout, default `120000`.

Converter service:

- `CONVERTER_PORT` - default `8787`.
- `CONVERTER_MAX_UPLOAD_BYTES` - default `78643200`.
- `CONVERTER_TIMEOUT_MS` - default `120000`.

## Cargo-friendly deployment

Deploy the static Vite app as usual, and run `server/converter.mjs` as a small adjacent Node process behind HTTPS. Route `/api/convert-xlsx` to `http://127.0.0.1:8787/convert-xlsx` and set matching proxy body-size and timeout limits. Keep proxy buffering disabled where possible so CSV responses can be returned immediately.

## Cleanup and privacy

Uploads are written to an OS temporary directory named `cargo-xlsx-*` only for the current request. The directory is removed after success or failure. Add normal platform-level temp directory sweeping as defence-in-depth, but the application does not rely on scheduled cleanup for request data.

## Future utilities

The converter is isolated under `server/` so future server-side utilities such as validation, XLSX repair, or metadata extraction can be added as new endpoints without changing the browser analyser import architecture.
