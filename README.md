# Cargo Statement Analyser

A standalone React, TypeScript and Vite static web application for record labels that receive sales and royalty statements exported from the Details reporting platform.

## Privacy model

All statement processing happens locally in the user's browser. Files are parsed with SheetJS in the page, are not uploaded, and are not stored in local storage, browser persistence, a database, analytics, or any external API. Refreshing or closing the page clears the imported statement data.

Do not commit real commercial statements to this repository.

## Supported file types

- `.xlsx`
- `.xls`
- `.csv`

## Known supported Details format

The first version supports the Details digital sales statement structure with a worksheet such as `Digital Sales` and columns including Artist, Album Title, Title, Amount, Royalty Amount, sales/usage fields, release identifiers, shops, countries, sales periods, deductions, line charges, and royalty rates.

The parser detects the header row in the first 30 rows, ignores blank worksheets and blank rows, preserves Source Sheet and Source Row, and keeps identifiers such as catalogue numbers and barcodes as text.

## Local setup

```bash
npm install
npm run dev
```

## Development commands

- `npm run dev` starts the Vite development server.
- `npm run build` runs TypeScript and creates a static production build in `dist/`.
- `npm run preview` previews the static build locally.
- `npm run test` runs Vitest.
- `npm run lint` runs ESLint.

## Production build

```bash
npm run build
```

The app is static and can be deployed from `dist/`.

## GitHub Pages deployment

Build command: `npm run build`  
Publish directory: `dist`  
The Vite `base` is relative (`./`) so the app can be served from a repository subpath.

## Cloudflare Pages deployment

Build command: `npm run build`  
Build output directory: `dist`

## Netlify deployment

Build command: `npm run build`  
Publish directory: `dist`

## Column aliases

Aliases live in `src/lib/columnAliases.ts`. The required fields are Artist, Album Title, Title, Amount, and Royalty Amount. Minor variations such as Catalogue No, UPC/EAN, Release Title, Track Title, Label Earnings, Net Royalty, and Revenue Amount are supported.

## Release grouping

Releases are grouped by this priority:

1. Barcode
2. Catalog No
3. Release Code
4. Artist plus Album Title

This prevents unrelated releases with the same title being merged unless no stronger identifier is present.

## Track grouping

Rows are treated as tracks only when ISRC is populated, Asset Type indicates Track, or Usage Type is track-level. Bundle rows with blank ISRC are not counted as tracks. Tracks are grouped by ISRC first, then by release identifier plus Artist plus Title.

## Sales-period interpretation

Sales Period is when the underlying usage or sale occurred. It may be earlier than the statement reporting month. The app preserves the source value and derives a sortable date for common formats such as `YYYY-MM`, `YYYY_MM`, `YYYYMM`, date strings, and Excel date serials.

## Current limitations

- Currency detection is conservative and displays currency-neutral totals unless a future enhancement confidently reads workbook currency metadata.
- The first version does not include a mapping wizard; aliases must be configured in code.
- Charts are intentionally minimal; tables are the primary inspection interface.
- Uploaded real statements must be manually verified before any public deployment.
