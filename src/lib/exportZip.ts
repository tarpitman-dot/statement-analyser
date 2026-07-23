import Decimal from 'decimal.js';
import {
  groupArtists,
  groupBy,
  groupReleases,
  groupTracks,
  periodSortValue,
  totals,
  trackRows,
} from './analytics';
import { D } from './format';
import type { StatementData, Transaction } from './types';

export type ExportProgressStage =
  | 'Validating totals'
  | 'Preparing reconciliation report'
  | 'Preparing summary'
  | 'Preparing artists'
  | 'Preparing releases'
  | 'Preparing tracks'
  | 'Preparing shops'
  | 'Preparing countries'
  | 'Preparing sales periods'
  | 'Preparing usage types'
  | 'Preparing full detail'
  | 'Preparing import checks'
  | 'Creating ZIP'
  | 'Complete';
export type ExportProgress = { stage: ExportProgressStage; index: number; total: number };
export type ZipEntry = { name: string; content: string | Uint8Array };
export type ReconciliationRow = {
  Breakdown: string;
  'Row Count': number;
  'Statement Royalty Amount': string;
  'Breakdown Royalty Amount': string;
  'Royalty Difference': string;
  'Statement Amount': string;
  'Breakdown Amount': string;
  'Amount Difference': string;
  'Statement Sales': string;
  'Breakdown Sales': string;
  'Sales Difference': string;
  Reconciles: 'Yes' | 'No' | 'Exception';
  Notes: string;
};
type TotallingGroup = Record<string, unknown> & {
  royaltyAmount: unknown;
  amount: unknown;
  sales: unknown;
  returns?: unknown;
  transactionRows?: number;
};
type TotalsAccumulator = { royaltyAmount: Decimal; amount: Decimal; sales: Decimal };
type Breakdown = {
  file: string;
  label: string;
  headers: string[];
  rows: Record<string, unknown>[];
  recon?: ReconciliationRow;
  complete?: boolean;
};
const tolerance = new Decimal('0.0001');
const textIdFields = new Set([
  'Barcode',
  'Catalog Number',
  'Catalogue Number',
  'Release Code',
  'ISRC',
  'Source Row',
  'sourceRow',
  'barcode',
  'catalogNumber',
  'releaseCode',
  'isrc',
]);
export function csvEscape(v: unknown, header = '') {
  let s = v instanceof Decimal ? v.toString() : String(v ?? '');
  if (textIdFields.has(header) && s) s = `="${s.replace(/"/g, '""')}"`;
  return `"${s.replace(/"/g, '""')}"`;
}
export function toExportCsv(rows: Record<string, unknown>[], headers: string[]) {
  return `\uFEFF${[
    headers.join(','),
    ...rows.map((r) => headers.map((h) => csvEscape(r[h], h)).join(',')),
  ].join('\n')}`;
}
function pct(part: unknown, total: Decimal) {
  return total.isZero() ? '0' : D(part).div(total).mul(100).toString();
}
function hasField(rows: Transaction[], field: keyof Transaction) {
  return rows.some((r) => String(r[field] ?? '').trim());
}
function ds(v: unknown) {
  return D(v).toString();
}
function reconcile(
  label: string,
  statement: Transaction[],
  groups: TotallingGroup[],
  notes = '',
  exception = false,
): ReconciliationRow {
  const st = totals(statement);
  const gt = groups.reduce(
    (a: TotalsAccumulator, g) => ({
      royaltyAmount: a.royaltyAmount.plus(D(g.royaltyAmount)),
      amount: a.amount.plus(D(g.amount)),
      sales: a.sales.plus(D(g.sales)),
    }),
    { royaltyAmount: D(0), amount: D(0), sales: D(0) } satisfies TotalsAccumulator,
  );
  const rd = gt.royaltyAmount.minus(st.royaltyAmount);
  const ad = gt.amount.minus(st.amount);
  const sd = gt.sales.minus(st.sales);
  const ok = rd.abs().lte(tolerance) && ad.abs().lte(tolerance) && sd.isZero();
  return {
    Breakdown: label,
    'Row Count': groups.length,
    'Statement Royalty Amount': st.royaltyAmount.toString(),
    'Breakdown Royalty Amount': gt.royaltyAmount.toString(),
    'Royalty Difference': rd.toString(),
    'Statement Amount': st.amount.toString(),
    'Breakdown Amount': gt.amount.toString(),
    'Amount Difference': ad.toString(),
    'Statement Sales': st.sales.toString(),
    'Breakdown Sales': gt.sales.toString(),
    'Sales Difference': sd.toString(),
    Reconciles: exception ? 'Exception' : ok ? 'Yes' : 'No',
    Notes: notes || 'Percentages are exported as 0-100 decimal percentage values.',
  };
}
function exportGroupForReconciliation(r: Record<string, unknown>, i: number): TotallingGroup {
  return {
    ...r,
    royaltyAmount: r['Royalty Amount'],
    amount: r.Amount,
    sales: r.Sales,
    transactionRows: i,
  };
}
export function buildBreakdowns(data: StatementData) {
  const rows = data.rows;
  const st = totals(rows);
  const breakdowns: Breakdown[] = [];
  const add = (
    file: string,
    label: string,
    headers: string[],
    rowsOut: Record<string, unknown>[],
    recon?: ReconciliationRow,
    complete = true,
  ) => {
    if (rowsOut.length) breakdowns.push({ file, label, headers, rows: rowsOut, recon, complete });
  };
  const common = (g: TotallingGroup) => ({
    'Royalty Amount': ds(g.royaltyAmount),
    Amount: ds(g.amount),
    Sales: ds(g.sales),
    'Percentage of total Royalty Amount': pct(g.royaltyAmount, st.royaltyAmount),
    'Percentage of total Amount': pct(g.amount, st.amount),
    'Percentage of total Sales': pct(g.sales, st.sales),
  });
  add(
    '02-artists.csv',
    'Artists',
    [
      'Artist',
      'Royalty Amount',
      'Amount',
      'Sales',
      'Percentage of total Royalty Amount',
      'Percentage of total Amount',
      'Percentage of total Sales',
      'Release Count',
      'Track Count',
      'Transaction Count',
    ],
    groupArtists(rows).map((g) => ({
      Artist: g.artist,
      ...common(g),
      'Release Count': g.releaseCount,
      'Track Count': g.trackCount,
      'Transaction Count': g.transactionRows,
    })),
  );
  add(
    '03-releases.csv',
    'Releases',
    [
      'Artist',
      'Release Title',
      'Barcode',
      'Catalog Number',
      'Release Code',
      'Royalty Amount',
      'Amount',
      'Sales',
      'Percentage of total Royalty Amount',
      'Percentage of total Amount',
      'Percentage of total Sales',
      'Track Count',
      'Transaction Count',
    ],
    groupReleases(rows).map((g) => ({
      Artist: g.artist,
      'Release Title': g.albumTitle,
      Barcode: g.barcode,
      'Catalog Number': g.catalogNumber,
      'Release Code': g.releaseCode,
      ...common(g),
      'Track Count': g.trackCount,
      'Transaction Count': g.transactionRows,
    })),
  );
  const tr = groupTracks(rows);
  const trackSource = trackRows(rows);
  add(
    '04-tracks.csv',
    'Tracks',
    [
      'Artist',
      'Track Title',
      'ISRC',
      'Release Title',
      'Barcode',
      'Royalty Amount',
      'Amount',
      'Sales',
      'Percentage of total Royalty Amount',
      'Transaction Count',
    ],
    tr.map((g) => ({
      Artist: g.artist,
      'Track Title': g.trackTitle,
      ISRC: g.isrc,
      'Release Title': g.albumTitle,
      Barcode: g.barcode,
      'Royalty Amount': ds(g.royaltyAmount),
      Amount: ds(g.amount),
      Sales: ds(g.sales),
      'Percentage of total Royalty Amount': pct(g.royaltyAmount, st.royaltyAmount),
      'Transaction Count': g.transactionRows,
    })),
    reconcile(
      'Tracks',
      trackSource,
      tr,
      trackSource.length === rows.length
        ? 'Track rows reconcile to all statement rows.'
        : 'Tracks exclude bundle/non-track rows under existing track rules.',
      trackSource.length !== rows.length,
    ),
    false,
  );
  if (hasField(rows, 'shop')) {
    const g = groupBy(rows, 'shop');
    add(
      '05-shops.csv',
      'Shops',
      [
        'Shop',
        'Royalty Amount',
        'Amount',
        'Sales',
        'Percentage of total Royalty Amount',
        'Percentage of total Amount',
        'Percentage of total Sales',
        'Artist Count',
        'Release Count',
        'Track Count',
        'Transaction Count',
      ],
      g.map((x) => ({
        Shop: x.shop,
        ...common(x),
        'Artist Count': x.artistCount,
        'Release Count': x.releaseCount,
        'Track Count': x.trackCount,
        'Transaction Count': x.transactionRows,
      })),
    );
  }
  if (hasField(rows, 'country')) {
    const g = groupBy(rows, 'country');
    add(
      '06-countries.csv',
      'Countries',
      [
        'Country',
        'Royalty Amount',
        'Amount',
        'Sales',
        'Percentage of total Royalty Amount',
        'Percentage of total Amount',
        'Percentage of total Sales',
        'Artist Count',
        'Release Count',
        'Track Count',
        'Shop Count',
        'Transaction Count',
      ],
      g.map((x) => ({
        Country: x.country,
        ...common(x),
        'Artist Count': x.artistCount,
        'Release Count': x.releaseCount,
        'Track Count': x.trackCount,
        'Shop Count': x.shopCount,
        'Transaction Count': x.transactionRows,
      })),
    );
  }
  if (hasField(rows, 'salesPeriod')) {
    const g = groupBy(rows, 'salesPeriod').sort((a, b) =>
      periodSortValue(a.salesPeriod).localeCompare(periodSortValue(b.salesPeriod)),
    );
    add(
      '07-sales-periods.csv',
      'Sales Periods',
      ['Sales Period', 'Royalty Amount', 'Amount', 'Sales', 'Shop Count', 'Country Count', 'Transaction Count'],
      g.map((x) => ({
        'Sales Period': x.salesPeriod,
        'Royalty Amount': ds(x.royaltyAmount),
        Amount: ds(x.amount),
        Sales: ds(x.sales),
        'Shop Count': x.shopCount,
        'Country Count': x.countryCount,
        'Transaction Count': x.transactionRows,
      })),
    );
  }
  if (hasField(rows, 'usageType')) {
    const g = groupBy(rows, 'usageType');
    add(
      '08-usage-types-or-formats.csv',
      'Usage Types',
      [
        'Usage Type',
        'Royalty Amount',
        'Amount',
        'Sales',
        'Percentage of total Royalty Amount',
        'Percentage of total Amount',
        'Percentage of total Sales',
        'Artist Count',
        'Release Count',
        'Track Count',
        'Shop Count',
        'Country Count',
        'Transaction Count',
      ],
      g.map((x) => ({
        'Usage Type': x.usageType,
        ...common(x),
        'Artist Count': x.artistCount,
        'Release Count': x.releaseCount,
        'Track Count': x.trackCount,
        'Shop Count': x.shopCount,
        'Country Count': x.countryCount,
        'Transaction Count': x.transactionRows,
      })),
    );
  }
  for (const b of breakdowns.filter((b) => b.complete !== false)) {
    b.recon = reconcile(b.label, rows, b.rows.map(exportGroupForReconciliation), b.recon?.Notes);
  }
  return breakdowns;
}
export function statementSummaryRows(data: StatementData) {
  const rows = data.rows;
  const st = totals(rows);
  const d = data.diagnostics;
  const vals: Record<string, unknown> = {
    'Imported filename': data.filename,
    'Statement period': d.reportingPeriod,
    'Contract name': rows.find((r) => r.contractName)?.contractName,
    'Contract ID': rows.find((r) => r.contractId)?.contractId,
    Currency: d.currency?.code || d.currency?.symbol,
    'Transaction rows': rows.length,
    'Artist count': groupArtists(rows).length,
    'Release count': groupReleases(rows).length,
    'Track count': groupTracks(rows).length,
    'Shop count': hasField(rows, 'shop') ? groupBy(rows, 'shop').length : undefined,
    'Country count': hasField(rows, 'country') ? groupBy(rows, 'country').length : undefined,
    'Sales-period range': hasField(rows, 'salesPeriod')
      ? [...new Set(rows.map((r) => r.salesPeriod).filter(Boolean))]
          .sort((a, b) => periodSortValue(a).localeCompare(periodSortValue(b)))
          .filter(Boolean)
          .join(' to ')
      : undefined,
    'Royalty Amount': st.royaltyAmount.toString(),
    Amount: st.amount.toString(),
    Sales: st.sales.toString(),
    'Average royalty per sale': st.sales.isZero() ? undefined : st.royaltyAmount.div(st.sales).toString(),
    'Effective royalty share': st.amount.isZero()
      ? undefined
      : st.royaltyAmount.div(st.amount).mul(100).toString(),
    'Statement Health': d.statementHealth?.dataQuality,
    'Barcode warning count': d.statementHealth?.barcodeWarnings,
    'Rows requiring review': d.statementHealth?.rowsRequiringReview,
  };
  return [Object.fromEntries(Object.entries(vals).filter(([, v]) => v !== undefined && v !== ''))];
}
export function fullDetailRows(rows: Transaction[]) {
  return rows.map(({ originalRow, ...r }) => {
    void originalRow;
    return Object.fromEntries(
      Object.entries(r).map(([k, v]) => [k, v instanceof Object ? JSON.stringify(v) : String(v ?? '')]),
    );
  });
}
export function buildExportEntries(data: StatementData, b = buildBreakdowns(data)) {
  const report = b.map((x) => x.recon).filter((r): r is ReconciliationRow => !!r);
  const bad = report.find((r) => r.Reconciles === 'No');
  if (bad)
    throw new Error(
      `Cannot create ZIP because the ${bad.Breakdown} breakdown does not reconcile. Royalty Amount difference: ${bad['Royalty Difference']}; Amount difference: ${bad['Amount Difference']}; Sales difference: ${bad['Sales Difference']}`,
    );
  const entries: ZipEntry[] = [];
  entries.push({
    name: '00-reconciliation-report.csv',
    content: toExportCsv(report as unknown as Record<string, unknown>[], Object.keys(report[0] ?? {})),
  });
  const summary = statementSummaryRows(data);
  entries.push({ name: '01-statement-summary.csv', content: toExportCsv(summary, Object.keys(summary[0] ?? {})) });
  for (const x of b) entries.push({ name: x.file, content: toExportCsv(x.rows, x.headers) });
  const detail = fullDetailRows(data.rows);
  entries.push({ name: '09-full-detail.csv', content: toExportCsv(detail, Object.keys(detail[0] ?? {})) });
  const d = data.diagnostics;
  if (d.barcodeIntegrity?.warnings?.length)
    entries.push({
      name: '10-import-checks.csv',
      content: toExportCsv(
        d.barcodeIntegrity.warnings.map((w) => ({
          sourceSheet: w.sourceSheet,
          sourceRow: w.sourceRow,
          barcodeValue: w.barcodeValue,
          warning: w.warning,
          suggestedReason: w.suggestedReason,
        })),
        ['sourceSheet', 'sourceRow', 'barcodeValue', 'warning', 'suggestedReason'],
      ),
    });
  return entries;
}
let crcTable: Uint32Array | undefined;
function crc32(bytes: Uint8Array) {
  crcTable ??= new Uint32Array(
    Array.from({ length: 256 }, (_, n) => {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      return c >>> 0;
    }),
  );
  let c = 0xffffffff;
  for (const b of bytes) c = crcTable[(c ^ b) & 255] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}
function u16(n: number) {
  return new Uint8Array([n & 255, (n >>> 8) & 255]);
}
function u32(n: number) {
  return new Uint8Array([n & 255, (n >>> 8) & 255, (n >>> 16) & 255, (n >>> 24) & 255]);
}
function concat(parts: Uint8Array[]) {
  const len = parts.reduce((a, p) => a + p.length, 0);
  const out = new Uint8Array(len);
  let o = 0;
  for (const p of parts) {
    out.set(p, o);
    o += p.length;
  }
  return out;
}
export function createZip(entries: ZipEntry[]) {
  const enc = new TextEncoder();
  const locals: Uint8Array[] = [];
  const central: Uint8Array[] = [];
  let offset = 0;
  for (const e of entries) {
    const name = enc.encode(e.name);
    const data = typeof e.content === 'string' ? enc.encode(e.content) : e.content;
    const crc = crc32(data);
    const local = concat([
      u32(0x04034b50),
      u16(20),
      u16(0x0800),
      u16(0),
      u16(0),
      u16(0),
      u32(crc),
      u32(data.length),
      u32(data.length),
      u16(name.length),
      u16(0),
      name,
      data,
    ]);
    locals.push(local);
    central.push(
      concat([
        u32(0x02014b50),
        u16(20),
        u16(20),
        u16(0x0800),
        u16(0),
        u16(0),
        u16(0),
        u32(crc),
        u32(data.length),
        u32(data.length),
        u16(name.length),
        u16(0),
        u16(0),
        u16(0),
        u16(0),
        u32(0),
        u32(offset),
        name,
      ]),
    );
    offset += local.length;
  }
  const cd = concat(central);
  return concat([
    ...locals,
    cd,
    u32(0x06054b50),
    u16(0),
    u16(0),
    u16(entries.length),
    u16(entries.length),
    u32(cd.length),
    u32(offset),
    u16(0),
  ]);
}
export function exportFilename(data: StatementData) {
  const base = (data.diagnostics.reportingPeriod || data.label || data.filename || 'statement')
    .replace(/\.[^.]+$/, '')
    .replace(/[^a-z0-9._-]+/gi, '-')
    .replace(/^-+|-+$/g, '') || 'statement';
  return `cargo-statement-breakdowns-${base}.zip`;
}
export async function downloadBreakdownsZip(data: StatementData, onProgress?: (p: ExportProgress) => void) {
  const stages: ExportProgressStage[] = [
    'Validating totals',
    'Preparing reconciliation report',
    'Preparing summary',
    'Preparing artists',
    'Preparing releases',
    'Preparing tracks',
    'Preparing shops',
    'Preparing countries',
    'Preparing sales periods',
    'Preparing usage types',
    'Preparing full detail',
    'Preparing import checks',
    'Creating ZIP',
    'Complete',
  ];
  let i = 0;
  const prog = (stage: ExportProgressStage) => onProgress?.({ stage, index: ++i, total: stages.length });
  prog('Validating totals');
  await Promise.resolve();
  let entries = buildExportEntries(data);
  prog('Creating ZIP');
  const zip = createZip(entries);
  entries = [];
  const blob = new Blob([zip], { type: 'application/zip' });
  const url = URL.createObjectURL(blob);
  try {
    const a = document.createElement('a');
    a.href = url;
    a.download = exportFilename(data);
    a.click();
  } finally {
    setTimeout(() => URL.revokeObjectURL(url), 0);
  }
  prog('Complete');
  return zip.length;
}
