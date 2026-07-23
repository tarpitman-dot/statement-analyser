# Large Excel Converter deployment

Cargo Statement Analyser keeps normal `.xlsx` and `.csv` analysis in the browser. The optional Large Excel Converter is only available after `GET /api/converter-health` returns healthy service metadata.

## Production audit

The repository currently contains a standalone Node converter at `server/converter.mjs`. It is **not** a Netlify Function, **not** a Netlify Edge Function, and `netlify.toml` does not deploy any function bundle. The static Netlify app has only the SPA catch-all redirect, so `/api/convert-xlsx` is not backed by Netlify unless production infrastructure adds an external proxy/rewrite in front of the site.

Do not assume frontend `fetch`/`XMLHttpRequest` code creates `/api` routes. If no external converter is configured and healthy, the browser disables **Convert and continue** and shows manual CSV instructions.

## Selected architecture

For the real 49,943,633-byte Winter Garden workbook, use a dedicated converter service/container or VM behind HTTPS, not standard Netlify Functions. The static app may remain on Netlify, but production must route:

- `GET /api/converter-health` to the converter service health endpoint;
- `POST /api/convert-xlsx` to the converter service conversion endpoint.

The converter service accepts both `/api/*` paths and direct `/health`/`/convert-xlsx` paths so a reverse proxy can preserve or strip the prefix.

Recommended proxy order if the same domain fronts both the API and SPA:

```text
/api/converter-health  https://<converter-host>/api/converter-health  200
/api/convert-xlsx      https://<converter-host>/api/convert-xlsx      200
/*                     /index.html                                    200
```

The two specific API routes must be evaluated before the SPA fallback (`/* /index.html 200`).

## Service

Run the lightweight Node service:

```bash
npm ci
npm run converter
```

The service accepts `POST /api/convert-xlsx` or `POST /convert-xlsx`, extracts only the case-insensitive `Digital Sales` worksheet, returns UTF-8 CSV, and deletes its temporary upload directory in a `finally` cleanup. It does not analyse statements, calculate totals, use a database, authenticate users, or keep historical uploads.

Health endpoint response (`GET /api/converter-health` or `GET /health`):

```json
{
  "status": "ok",
  "maxUploadBytes": 104857600,
  "converterVersion": "node-xlsx-2026-07-23",
  "supportedWorksheet": "Digital Sales"
}
```

## Platform requirements

Do **not** deploy the converter path on standard Netlify Functions for 50 MB workbook uploads. This workflow needs:

- request body size: at least 60 MB; configured default is 100 MiB (`104857600` bytes);
- execution timeout: enough for upload, XLSX decompression, CSV creation and download; default converter timeout is 300 seconds;
- memory: sized for the 50 MB ZIP plus the decompressed worksheet XML and generated CSV; validate with the real workbook because the worksheet XML can be hundreds of MB;
- temporary disk: OS temp directory with room for the uploaded workbook and transient conversion files;
- response handling: `text/csv` responses may be large and should not be buffered by a low-limit proxy.

A Python/FastAPI/openpyxl read-only container or a direct object-storage upload plus conversion worker is also suitable. Do not retry this file through a platform that buffers/rejects large bodies before user code runs.

## Environment variables

Browser app:

- `VITE_CONVERTER_API_URL` - preferred converter endpoint, for example `/api/convert-xlsx` on same origin or `https://converter.example.com/api/convert-xlsx` cross-origin.
- `VITE_CONVERTER_HEALTH_URL` - optional explicit health endpoint. If omitted, the app derives it by replacing `/convert-xlsx` with `/converter-health`.
- `VITE_LARGE_XLSX_CONVERTER_URL` - legacy converter endpoint fallback.
- `VITE_LARGE_XLSX_BROWSER_SIZE_LIMIT` - browser XLSX byte safety limit, default `10485760`.
- `VITE_LARGE_XLSX_BROWSER_ROW_LIMIT` - browser row safety limit, default `50000`.
- `VITE_LARGE_XLSX_LOCAL_READ_TIMEOUT_MS` - local file access timeout, default `30000`.
- `VITE_LARGE_XLSX_UPLOAD_INACTIVITY_TIMEOUT_MS` - abort if uploaded bytes do not advance, default `30000`.
- `VITE_LARGE_XLSX_RESPONSE_TIMEOUT_MS` - wait for converter response/CSV, default `120000`.
- `VITE_LARGE_XLSX_TOTAL_TIMEOUT_MS` - whole conversion budget, default `180000`.

Converter service:

- `CONVERTER_PORT` - default `8787`.
- `CONVERTER_MAX_UPLOAD_BYTES` - default `104857600` (100 MiB).
- `CONVERTER_TIMEOUT_MS` - default `300000`.
- `CONVERTER_VERSION` - version string returned by health checks.
- `CONVERTER_ALLOWED_ORIGINS` - comma-separated analyser production and preview origins. Empty means reflect any origin; set this in production.

## CORS

For cross-origin deployment, set `CONVERTER_ALLOWED_ORIGINS` to the production analyser origin and required preview origins. The converter handles `OPTIONS`, allows `GET`, `POST`, `OPTIONS`, allows `Content-Type` and `Accept`, and exposes conversion headers including `x-conversion-started`, `x-source-worksheet`, `x-converted-row-count`, `x-upload-bytes-received`, and `x-converter-cleanup`.

## Responses and errors

Success returns HTTP 200, `content-type: text/csv; charset=utf-8`, `x-source-worksheet`, `x-converted-row-count`, `x-upload-bytes-received`, and `x-conversion-started: true`.

Failures return JSON with HTTP status and this shape:

```json
{ "error": "...", "stage": "...", "details": "..." }
```

Handled stages include routing/404, receiving-upload/413, locating-worksheet/422, conversion/500, and timeout/504.

## Upload method

The browser uploads the original `File` directly using `XMLHttpRequest` with XLSX content type. It reads only a small local slice first to prove file access for progress diagnostics; it does not read the complete workbook into an `ArrayBuffer` before upload. Upload progress and cancellation are preserved.

## Cleanup and privacy

Uploads are written to an OS temporary directory named `cargo-xlsx-*` for the current request. The directory is removed after success or failure. Add platform-level temp sweeping as defence-in-depth.

## Production acceptance test

Before enabling **Convert and continue**, test with `WINTER_GARDEN_PRODUCTIONS___2026_04_01___2026_06_30.xlsx` and confirm health succeeds, all 49,943,633 bytes upload, HTTP status is received, `Digital Sales` is found, CSV imports in the browser, dashboard opens, Safari does not reload, temp files are deleted, totals reconcile, and ZIP exports reconcile.
