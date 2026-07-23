import * as XLSX from 'xlsx';
import { groupArtists, groupBy, groupReleases, groupTracks, totals, trackRows } from '../src/lib/analytics';
import { buildBreakdowns, buildExportEntries, fullDetailRows } from '../src/lib/exportZip';
import { parseArrayBuffer, parseWorkbook } from '../src/lib/parser';
import type { StatementData } from '../src/lib/types';

const HEADERS = [
  'Artist',
  'Album Title',
  'Track Title',
  'Asset Type',
  'Usage Type',
  'Country',
  'Shop',
  'Sales Period',
  'Sales',
  'Returns',
  'Amount',
  'Royalty Amount',
  'Royalty Rate',
  'Deduction 1',
  'Barcode',
  'Catalog Number',
  'ISRC',
];

type FixtureKind = 'xlsx' | 'csv';

type FixtureResult = {
  kind: FixtureKind;
  requestedRows: number;
  fileSize: number;
  importMs: number;
  dashboardReadyMs: number;
  rowsImported: number;
  rowsSkipped: number;
  largeFileMode: boolean;
  reconciles: boolean | 'not-run';
  fullDetailRows: number;
};

function assert(condition: unknown, message: string): asserts condition {
  if (!condition) throw new Error(message);
}

function row(index: number) {
  return [
    `Artist ${index % 17}`,
    `Album ${index % 101}`,
    `Track ${index}`,
    'Track',
    index % 3 === 0 ? 'Download' : 'Stream',
    index % 2 === 0 ? 'GB' : 'US',
    `Shop ${index % 5}`,
    `2026-${String((index % 12) + 1).padStart(2, '0')}`,
    '1',
    '0',
    '1.23',
    '0.50',
    '50%',
    '0',
    String(1_000_000_000_000 + index).padStart(13, '0'),
    `CAT${String(index).padStart(8, '0')}`,
    `ISRC${String(index).padStart(8, '0')}`,
  ];
}

function fixtureRows(count: number) {
  return [HEADERS, ...Array.from({ length: count }, (_, index) => row(index))];
}

function fixtureBuffer(kind: FixtureKind, count: number) {
  const rows = fixtureRows(count);
  if (kind === 'csv') {
    const csv = rows.map((cells) => cells.map(String).join(',')).join('\n');
    return new TextEncoder().encode(csv).buffer;
  }
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), 'Digital Sales');
  return XLSX.write(workbook, { bookType: 'xlsx', type: 'array' }) as ArrayBuffer;
}

function completeBreakdownsReconcile(data: StatementData) {
  return buildBreakdowns(data)
    .filter((breakdown) => breakdown.complete !== false)
    .every((breakdown) => breakdown.recon?.Reconciles === 'Yes');
}

function dashboardReady(data: StatementData) {
  totals(data.rows);
  groupArtists(data.rows);
  groupReleases(data.rows);
  groupTracks(data.rows);
  groupBy(data.rows, 'shop');
  groupBy(data.rows, 'country');
  groupBy(data.rows, 'salesPeriod');
  groupBy(data.rows, 'usageType');
}

async function exerciseFixture(kind: FixtureKind, requestedRows: number): Promise<FixtureResult> {
  const buffer = fixtureBuffer(kind, requestedRows);
  const started = performance.now();
  const data = await parseArrayBuffer(buffer, `release-health-${requestedRows}.${kind}`, buffer.byteLength);
  const imported = performance.now();
  if (requestedRows <= 50_000) dashboardReady(data);
  const dashboardReadyMs = performance.now() - imported;
  return {
    kind,
    requestedRows,
    fileSize: buffer.byteLength,
    importMs: Math.round(imported - started),
    dashboardReadyMs: Math.round(dashboardReadyMs),
    rowsImported: data.rows.length,
    rowsSkipped: data.diagnostics.blankRowsIgnored,
    largeFileMode: !!data.diagnostics.largeFileMode,
    reconciles: requestedRows <= 50_000 ? completeBreakdownsReconcile(data) : 'not-run',
    fullDetailRows: fullDetailRows(data.rows).length,
  };
}

function assertHealthyFixture(result: FixtureResult) {
  assert(result.rowsImported === result.requestedRows, `${result.kind} imported row count mismatch`);
  assert(result.rowsSkipped === 0, `${result.kind} skipped rows unexpectedly`);
  assert(
    result.largeFileMode === (result.fileSize > 10 * 1024 * 1024 || result.requestedRows > 50_000),
    `${result.kind} large-file mode mismatch`,
  );
  if (result.requestedRows <= 50_000) assert(result.reconciles === true, `${result.kind} did not reconcile`);
  assert(result.fullDetailRows === result.requestedRows, `${result.kind} full detail row count mismatch`);
}

function assertZipManifest() {
  const buffer = fixtureBuffer('csv', 5_000);
  return parseArrayBuffer(buffer, 'release-health-zip.csv', buffer.byteLength).then((data) => {
    const names = buildExportEntries(data).map((entry) => entry.name);
    for (const expected of [
      '00-reconciliation-report.csv',
      '01-statement-summary.csv',
      '02-artists.csv',
      '03-releases.csv',
      '04-tracks.csv',
      '05-shops.csv',
      '06-countries.csv',
      '07-sales-periods.csv',
      '08-usage-types-or-formats.csv',
      '09-full-detail.csv',
    ]) {
      assert(names.includes(expected), `ZIP entry missing: ${expected}`);
    }
  });
}

function assertIdentifierRegressions() {
  const worksheet = XLSX.utils.aoa_to_sheet([HEADERS, row(0), row(1), row(2)]);
  (worksheet['O2'] as XLSX.CellObject).v = 888831328476;
  (worksheet['O2'] as XLSX.CellObject).w = '8.88831E+11';
  (worksheet['O3'] as XLSX.CellObject).v = 888831328477;
  (worksheet['O3'] as XLSX.CellObject).w = '8.88831E+11';
  (worksheet['O4'] as XLSX.CellObject).v = '0888831328478';
  (worksheet['O4'] as XLSX.CellObject).t = 's';
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Digital Sales');
  const data = parseWorkbook(workbook, 'exact-barcodes.xlsx', 1000);
  assert(
    JSON.stringify(data.rows.map((transaction) => transaction.barcode)) ===
      JSON.stringify(['888831328476', '888831328477', '0888831328478']),
    'Exact/leading-zero barcode parsing regression',
  );
  assert(groupReleases(data.rows).length === 3, 'Exact barcode releases should remain separate');
}

function assertBundleTrackException() {
  const rows = fixtureRows(2);
  rows[2][2] = '';
  rows[2][3] = 'Bundle';
  rows[2][4] = 'Bundle Download';
  rows[2][16] = '';
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(rows), 'Digital Sales');
  const data = parseWorkbook(workbook, 'bundle.xlsx', 1000);
  assert(trackRows(data.rows).length === 1, 'Bundle row entered track grouping unexpectedly');
  const tracks = buildBreakdowns(data).find((breakdown) => breakdown.label === 'Tracks');
  assert(tracks?.recon?.Reconciles === 'Exception', 'Track bundle exclusion was not documented as exception');
  assert(
    String(tracks?.recon?.Notes ?? '').includes('Tracks exclude bundle/non-track rows'),
    'Track bundle exclusion note missing',
  );
}

for (const count of [5_000, 50_000, 100_000]) {
  for (const kind of ['xlsx', 'csv'] as const) {
    const result = await exerciseFixture(kind, count);
    console.info(JSON.stringify(result));
    assertHealthyFixture(result);
  }
}
await assertZipManifest();
assertIdentifierRegressions();
assertBundleTrackException();
console.info('release-health checks passed');
