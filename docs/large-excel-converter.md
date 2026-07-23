# Large Excel Converter deployment

Cargo Statement Analyser keeps normal `.xlsx` and `.csv` analysis in the browser. The optional Large Excel Converter is only used when the browser preflight safety limit is exceeded.

## Service

Run the lightweight Node service:

```bash
npm ci
npm run converter
```

The service accepts `POST /convert-xlsx`, extracts only the case-insensitive `Digital Sales` worksheet, returns UTF-8 CSV, and deletes its temporary upload directory in a `finally` cleanup. It does not analyse statements, calculate totals, use a database, authenticate users, or keep historical uploads.

## Client workflow and timeouts

The browser uploads the original `File` directly with `XMLHttpRequest` so upload progress is byte-based and advances while the request body is being sent. The converter workflow reports these distinct stages:

1. Reading file locally
2. Uploading workbook
3. Server received workbook
4. Converting Digital Sales worksheet
5. Downloading CSV
6. Importing CSV
7. Complete

Configured client-side safeguards:

- local file access/read check timeout: `VITE_LARGE_XLSX_LOCAL_READ_TIMEOUT_MS`, default `30000`.
- upload inactivity timeout: `VITE_LARGE_XLSX_UPLOAD_INACTIVITY_TIMEOUT_MS`, default `30000`.
- converter response timeout: `VITE_LARGE_XLSX_RESPONSE_TIMEOUT_MS`, default `120000`.
- total conversion timeout: `VITE_LARGE_XLSX_TOTAL_TIMEOUT_MS`, default `180000`.

If no upload progress is reported for 30 seconds, the browser aborts the XHR, releases the request handle and local read handle, and shows a retryable diagnostic error. Cancel import aborts the FileReader/XHR and invalidates stale callbacks so an older upload cannot update a later import.

## Environment variables

Browser app:

- `VITE_LARGE_XLSX_CONVERTER_URL` - converter endpoint, default `/api/convert-xlsx`.
- `VITE_LARGE_XLSX_BROWSER_SIZE_LIMIT` - browser XLSX byte safety limit, default `10485760`.
- `VITE_LARGE_XLSX_BROWSER_ROW_LIMIT` - browser row safety limit, default `50000`.
- `VITE_LARGE_XLSX_LOCAL_READ_TIMEOUT_MS` - local file access timeout, default `30000`.
- `VITE_LARGE_XLSX_UPLOAD_INACTIVITY_TIMEOUT_MS` - abort if uploaded bytes do not advance, default `30000`.
- `VITE_LARGE_XLSX_RESPONSE_TIMEOUT_MS` - wait for converter response/CSV, default `120000`.
- `VITE_LARGE_XLSX_TOTAL_TIMEOUT_MS` - whole conversion budget, default `180000`.
- `VITE_LARGE_XLSX_CONVERTER_TIMEOUT_MS` - legacy total-timeout fallback if `VITE_LARGE_XLSX_TOTAL_TIMEOUT_MS` is not set.

Converter service:

- `CONVERTER_PORT` - default `8787`.
- `CONVERTER_MAX_UPLOAD_BYTES` - default `78643200` (75 MiB, enough for a 50 MB XLSX with headroom when the hosting layer permits it).
- `CONVERTER_TIMEOUT_MS` - default `120000`.

## Deployment requirements for 50 MB XLSX files

Do **not** deploy the converter path on Netlify Functions for 50 MB workbook uploads. Netlify Functions are not a reliable fit for this workflow because function request bodies are handled by the platform before user code runs, and the platform has request payload, execution-time, and memory ceilings that can reject or buffer large XLSX bodies before this converter can stream or report useful progress.

For production, use one of these instead:

- a dedicated Node service/container/VM behind HTTPS with request body size at least `75 MiB`, timeout at least the configured total conversion timeout, and memory sized for SheetJS conversion of the largest expected workbook;
- or a direct-to-object-storage upload (for example S3-compatible presigned PUT) followed by a conversion job/service that reads the uploaded object and returns the CSV.

The static Vite app can still be hosted on Netlify, but `/api/convert-xlsx` must be routed to infrastructure that can reliably accept a 50 MB request. Do not repeatedly retry 50 MB workbooks through a serverless environment that buffers or rejects them before the converter starts.

## Cargo-friendly deployment

Deploy the static Vite app as usual, and run `server/converter.mjs` as a small adjacent Node process behind HTTPS. Route `/api/convert-xlsx` to `http://127.0.0.1:8787/convert-xlsx` and set matching proxy body-size and timeout limits. Keep proxy buffering disabled where possible so CSV responses can be returned immediately.

Recommended proxy/service settings:

- request body size: at least `75 MiB`;
- upload/read timeout: greater than `VITE_LARGE_XLSX_TOTAL_TIMEOUT_MS`;
- response timeout: greater than `VITE_LARGE_XLSX_RESPONSE_TIMEOUT_MS`;
- memory: enough for the XLSX parse plus generated CSV; test with the largest real workbook before enabling production traffic.

## Diagnostics

Browser-side technical details include endpoint URL, bytes read locally, bytes uploaded, HTTP status, elapsed upload time, whether conversion response headers were received, response size, and timeout/abort reason. Server responses also include `x-upload-bytes-received` and `x-conversion-started` headers on successful conversions.

## Cleanup and privacy

Uploads are written to an OS temporary directory named `cargo-xlsx-*` only for the current request. The directory is removed after success or failure. Add normal platform-level temp directory sweeping as defence-in-depth, but the application does not rely on scheduled cleanup for request data.

## Future utilities

The converter is isolated under `server/` so future server-side utilities such as validation, XLSX repair, or metadata extraction can be added as new endpoints without changing the browser analyser import architecture.
