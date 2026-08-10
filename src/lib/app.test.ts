import { describe, it, expect, vi } from 'vitest';
import * as XLSX from 'xlsx';
import { detectHeaderRow, parseWorkbook, normaliseIdentifier, parseRoyaltyRate } from './parser';
import { clampProgress, isLargeFile, IMPORT_STAGES } from './importProgress';
import { fieldLabels } from './columnAliases';
import { toCsv } from './exportCsv';
import { D, fmtRate, fmtMoney } from './format';
import { sampleStatement } from './sampleData';
import { ensureImportTimings, recordTiming, recordTimingValue, safeNow } from './timing';
import {
  groupArtists,
  groupReleases,
  groupTracks,
  groupBy,
  groupRoyaltyRates,
  rateSummary,
  importCheckGroups,
  periodSortValue,
  searchRows,
  totals,
  groupedReconciliation,
  isTrackRow,
  prepareOverviewChartData,
  chartDatasetCacheKey,
  prepareOverviewChartDataFromRows,
  overviewChartDataForRows,
} from './analytics';
const headers = Object.values(fieldLabels);
const row = [
  'C1',
  'Contract',
  '',
  'Track',
  'R1',
  'Release A',
  'CAT-001',
  '1234567890123',
  'ISRC1',
  'Artist A',
  'Track A',
  'Track Stream',
  'GB',
  'Shop A',
  '2026-06',
  '10',
  '1',
  '0.1',
  '1.20',
  '',
  '',
  '',
  '',
  '',
  '',
  '',
  '',
  '',
  '',
  '0.85',
  '1.02',
];
function wb(aoa: any[][], name = 'Digital Sales') {
  const w = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(w, XLSX.utils.aoa_to_sheet(aoa), name);
  return w;
}
describe('parsing', () => {
  it('detects header row and known Details columns', () => {
    const h = detectHeaderRow([['title'], headers]);
    expect(h?.idx).toBe(1);
    expect(h?.map.get('royaltyAmount')).toBe(30);
  });
  it('handles aliases, blank sheets, reordered columns, numeric/currency strings and csv-like workbooks', () => {
    const w = wb([[]], 'Blank');
    XLSX.utils.book_append_sheet(
      w,
      XLSX.utils.aoa_to_sheet([
        ['x'],
        [
          'Artist',
          'Release Title',
          'Track Title',
          'Revenue Amount',
          'Label Earnings',
          'Cat Number',
          'EAN',
        ],
        ['A', 'R', 'T', '£1,000.10', '850.085', '00001', '1.23457E+12'],
      ]),
      'Other',
    );
    const s = parseWorkbook(w, 'KIKA___2026_06_01___2026_06_30-4.xlsx', 123);
    expect(s.diagnostics.worksheetsIgnored[0]).toContain('blank');
    expect(s.diagnostics.account).toBe('KIKA');
    expect(s.rows[0].catalogNumber).toBe('00001');
    expect(s.rows[0].barcode).toBe('1234570000000');
    expect(D(s.rows[0].amount).toString()).toBe('1000.1');
  });
  it('throws clear missing required error', () =>
    expect(() => parseWorkbook(wb([['Artist'], ['A']]))).toThrow(/missing/));
  it('normalises scientific identifiers', () =>
    expect(normaliseIdentifier('1.23E+5')).toBe('123000'));
  it('parses royalty rates from decimals, numbers and percentages', () => {
    expect(parseRoyaltyRate('100%')?.toString()).toBe('1');
    expect(parseRoyaltyRate('100.00%')?.toString()).toBe('1');
    expect(parseRoyaltyRate('85%')?.toString()).toBe('0.85');
    expect(parseRoyaltyRate('85.0%')?.toString()).toBe('0.85');
    expect(parseRoyaltyRate(' 1,000.00% ')?.toString()).toBe('10');
    expect(parseRoyaltyRate(0.85)?.toString()).toBe('0.85');
    expect(parseRoyaltyRate(1)?.toString()).toBe('1');
    expect(parseRoyaltyRate(85)?.toString()).toBe('0.85');
    expect(parseRoyaltyRate('not a rate')).toBeNull();
  });
  it('normalises percentage royalty rates and records invalid rates without throwing', () => {
    const percent = [...row];
    percent[29] = '100.00%';
    percent[30] = '1.20';
    const invalid = [...row];
    invalid[29] = 'bad%';
    const s = parseWorkbook(wb([headers, percent, invalid]));
    expect(s.rows[0].royaltyRate).toBe('1');
    expect(s.rows[1].royaltyRate).toBe('');
    expect(s.diagnostics.invalidNumericValues[0]).toContain('Royalty Rate could not be parsed');
  });
});
describe('calculations and grouping', () => {
  const s = parseWorkbook(
    wb([
      headers,
      row,
      [
        ...row.slice(0, 7),
        '1234567890123',
        'ISRC2',
        'Artist A',
        'Track B',
        'Bundle DL',
        'US',
        'Shop B',
        '2026_05',
        '5',
        '0',
        '0.2',
        '2.00',
        '',
        '',
        '',
        '-0.1',
        '',
        '',
        '',
        '',
        '',
        '-0.2',
        '0.80',
        '1.60',
      ],
      [
        ...row.slice(0, 6),
        'CAT-002',
        '',
        '',
        'Artist B',
        'Bundle',
        'Bundle DL',
        'GB',
        'Shop A',
        '202605',
        '1',
        '0',
        '',
        '3.00',
        '',
        '',
        '',
        '',
        '',
        '',
        '',
        '',
        '',
        '',
        '0.85',
        '2.55',
      ],
    ]),
  );
  it('uses decimal-safe totals', () => {
    const t = totals(s.rows);
    expect(t.amount.toString()).toBe('6.2');
    expect(t.royaltyAmount.toString()).toBe('5.17');
    expect(t.sales.toString()).toBe('16');
    expect(t.returns.toString()).toBe('1');
    expect(t.deductions.toString()).toBe('-0.3');
    expect(t.lineCharges.toString()).toBe('-0.2');
  });
  it('formats rates', () => {
    expect(fmtRate('0.85')).toBe('85%');
    expect(new Set(s.rows.map((r) => r.royaltyRate)).size).toBe(2);
  });
  it('groups artists/releases/tracks/shops/countries and excludes blank bundle ISRC track', () => {
    expect(groupArtists(s.rows)[0].artist).toBe('Artist A');
    expect(groupReleases(s.rows).length).toBe(2);
    expect(groupTracks(s.rows).length).toBe(2);
    expect(groupBy(s.rows, 'shop').length).toBe(2);
    expect(groupBy(s.rows, 'country').length).toBe(2);
    expect(['2026-05', '2026-06']).toEqual(
      ['2026_05', '2026-06']
        .sort((a, b) => periodSortValue(a).localeCompare(periodSortValue(b)))
        .map(periodSortValue),
    );
  });
  it('searches only the selected source field and supports exact identifiers', () => {
    expect(searchRows(s.rows, 'artist', 'Release A').length).toBe(0);
    expect(searchRows(s.rows, 'albumTitle', 'Release A').length).toBe(3);
    expect(searchRows(s.rows, 'trackTitle', 'Track A').length).toBe(1);
    expect(searchRows(s.rows, 'catalogNumber', 'CAT-001', true).length).toBe(2);
    expect(searchRows(s.rows, 'barcode', '1234567890123', true).length).toBe(2);
    expect(searchRows(s.rows, 'isrc', 'ISRC1', true).length).toBe(1);
  });
  it('classifies validation and duplicates', () => {
    expect(s.diagnostics.royaltyValidation.matches).toBeGreaterThan(0);
    expect(s.diagnostics.royaltyValidation.notChecked).toBeGreaterThan(0);
    const dup = parseWorkbook(wb([headers, row, row]));
    expect(dup.diagnostics.duplicateLookingRows).toBe(2);
  });
});

describe('large import progress helpers', () => {
  it('keeps stages ordered and progress monotonic', () => {
    let p = 0;
    for (const stage of IMPORT_STAGES) {
      const n = clampProgress(p, stage, 1);
      expect(n).toBeGreaterThanOrEqual(p);
      p = n;
    }
    expect(p).toBe(100);
  });
  it('enables Large File Mode by size or estimated rows', () => {
    expect(isLargeFile(11 * 1024 * 1024, 1)).toBe(true);
    expect(isLargeFile(1, 50001)).toBe(true);
    expect(isLargeFile(1, 50000)).toBe(false);
  });
  it('chunks large rows, defers duplicates and drops raw row copies', () => {
    const rows = Array.from({ length: 5001 }, (_, i) => [
      ...row.slice(0, 15),
      String(i + 1),
      '0',
      '0.1',
      '1.00',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '',
      '0.85',
      '0.85',
    ]);
    const events: any[] = [];
    const s = parseWorkbook(wb([headers, ...rows]), 'large.xlsx', 11 * 1024 * 1024, {
      onProgress: (e) => events.push(e),
    });
    expect(s.diagnostics.largeFileMode).toBe(true);
    expect(events.filter((e) => e.stage === 'Importing transactions').length).toBeGreaterThan(1);
    expect(s.rows[0].originalRow).toBeUndefined();
    expect(s.diagnostics.duplicateLookingRows).toBe(0);
    expect(s.diagnostics.initialSummary?.transactionCount).toBe(5001);
    expect(events.at(-1).percent).toBe(100);
  });
  it('row-level problems do not stop full import', () => {
    const bad = [...row];
    bad[29] = 'bad%';
    const s = parseWorkbook(wb([headers, bad, row]));
    expect(s.rows.length).toBe(2);
    expect(s.diagnostics.invalidNumericValues.length).toBeGreaterThan(0);
  });
});

describe('grouped export reconciliation regressions', () => {
  const tx = (i: number, over: Partial<any> = {}) => ({
    sourceSheet: 's',
    sourceRow: i,
    contractId: '',
    contractName: '',
    shareContract: '',
    assetType: 'Track',
    releaseCode: `R${i % 5}`,
    albumTitle: `Release ${i % 7}`,
    catalogNumber: `CAT${i % 7}`,
    barcode: `BC${i % 7}`,
    isrc: `ISRC${i}`,
    artist: `Artist ${i % 23}`,
    trackTitle: `Track ${i % 9}`,
    usageType: i % 2 ? 'Stream' : 'Download',
    country: i % 3 === 0 ? 'GB' : 'US',
    shop: `Shop ${i % 25}`,
    salesPeriod: `2026-${String((i % 4) + 1).padStart(2, '0')}`,
    sales: String(i + 1),
    returns: String(i % 2),
    ppu: '0.0001',
    amount: (0.10001 + i / 1000).toString(),
    share: '',
    rata1: '',
    rata2: '',
    deduction1: '',
    deduction2: '',
    deduction3: '',
    contractDeductions: '',
    deduction4: '',
    deduction5: '',
    lineCharges: '',
    royaltyRate: '0.85',
    royaltyAmount: (0.0850085 + i / 10000).toString(),
    ...over,
  });
  const rows = Array.from({ length: 55 }, (_, i) =>
    tx(
      i,
      i === 3
        ? { shop: '' }
        : i === 4
          ? { artist: '', country: '', salesPeriod: '', usageType: '' }
          : {},
    ),
  );
  function expectRecon(groups: any[], subset = rows) {
    const r = groupedReconciliation(subset, groups);
    expect(r.reconciled).toBe(true);
    expect(r.grouped.royaltyAmount.toString()).toBe(r.dashboard.royaltyAmount.toString());
    expect(r.grouped.amount.toString()).toBe(r.dashboard.amount.toString());
    expect(r.grouped.sales.toString()).toBe(r.dashboard.sales.toString());
    expect(r.grouped.returns.toString()).toBe(r.dashboard.returns.toString());
  }
  it('shop export source contains every shop beyond top 10 and reconciles independently of chart data', () => {
    const all = groupBy(rows, 'shop');
    const top10 = all.slice(0, 10);
    expect(all.length).toBeGreaterThan(20);
    expect(top10.length).toBe(10);
    expect(all.map((g) => g.shop)).toContain('Unspecified shop');
    expectRecon(all);
    expect(groupedReconciliation(rows, top10).reconciled).toBe(false);
  });
  it('all grouped summaries reconcile to dashboard totals including blanks', () => {
    for (const groups of [
      groupArtists(rows),
      groupReleases(rows),
      groupTracks(rows),
      groupBy(rows, 'country'),
      groupBy(rows, 'salesPeriod'),
      groupBy(rows, 'usageType'),
      groupBy(rows, 'royaltyRate'),
    ])
      expectRecon(groups);
    expect(groupArtists(rows).map((g) => g.artist)).toContain('Unspecified artist');
    expect(groupBy(rows, 'country').map((g) => g.country)).toContain('Unspecified country');
    expect(groupBy(rows, 'salesPeriod').map((g) => g.salesPeriod)).toContain(
      'Unspecified sales period',
    );
    expect(groupBy(rows, 'usageType').map((g) => g.usageType)).toContain('Unspecified usage type');
  });
  it('pagination and visual table limits do not limit complete-summary exports', () => {
    const all = groupBy(rows, 'shop');
    const page = all.slice(0, 10);
    expect(page.length).toBeLessThan(all.length);
    expectRecon(all);
    expect(groupedReconciliation(rows, page).grouped.royaltyAmount.toString()).not.toBe(
      totals(rows).royaltyAmount.toString(),
    );
  });
  it('global filters affect dashboard and exports consistently while preserving decimal precision', () => {
    const filtered = rows.filter((r) => r.country === 'GB');
    const groups = groupBy(filtered, 'shop');
    expect(filtered.length).toBeGreaterThan(0);
    expectRecon(groups, filtered);
    expect(totals(filtered).royaltyAmount.decimalPlaces()).toBeGreaterThan(2);
  });
  it('CSV grouped totals are numeric strings without currency symbols', () => {
    const groups = groupBy(rows, 'shop');
    const csv = toCsv(
      [
        ...groups.map((g) => ({
          shop: g.shop,
          royaltyAmount: g.royaltyAmount.toString(),
          amount: g.amount.toString(),
          sales: g.sales.toString(),
          returns: g.returns.toString(),
        })),
        {
          shop: 'TOTAL',
          royaltyAmount: totals(rows).royaltyAmount.toString(),
          amount: totals(rows).amount.toString(),
          sales: totals(rows).sales.toString(),
          returns: totals(rows).returns.toString(),
        },
      ],
      ['shop', 'royaltyAmount', 'amount', 'sales', 'returns'],
    );
    expect(csv).toContain('TOTAL');
    expect(csv).not.toContain('£');
    expect(csv).toMatch(/"\d+\.\d+"/);
  });
});

describe('review correctness regressions', () => {
  const make = (over: any = {}) => ({
    sourceSheet: 's',
    sourceRow: 1,
    contractId: '',
    contractName: '',
    shareContract: '',
    assetType: 'Track',
    releaseCode: 'R',
    albumTitle: 'Release',
    catalogNumber: 'CAT',
    barcode: 'BC',
    isrc: 'ISRC',
    artist: 'Artist',
    trackTitle: 'Track',
    usageType: 'Track Stream',
    country: 'GB',
    shop: 'Shop',
    salesPeriod: '2026-06',
    sales: '1',
    returns: '0',
    ppu: '',
    amount: '10',
    share: '',
    rata1: '',
    rata2: '',
    deduction1: '',
    deduction2: '',
    deduction3: '',
    contractDeductions: '',
    deduction4: '',
    deduction5: '',
    lineCharges: '',
    royaltyRate: '0.85',
    royaltyAmount: '8.5',
    ...over,
  });
  it('uses isTrackRow consistently so bundle/release/service rows are excluded from tracks', () => {
    const rows = [
      make({ isrc: 'ISRC1' }),
      make({
        isrc: '',
        assetType: 'Track',
        usageType: 'Track Download',
        trackTitle: 'Blank ISRC Track',
      }),
      make({ isrc: '', assetType: 'Bundle', usageType: 'Bundle DL', trackTitle: 'Bundle Asset' }),
      make({ isrc: '', assetType: 'Album', usageType: 'Bundle DL', trackTitle: 'Bundle DL' }),
      make({ isrc: '', assetType: 'Release', usageType: 'Service Fee', trackTitle: 'Fee' }),
    ];
    const tracks = groupTracks(rows);
    expect(rows.filter(isTrackRow).length).toBe(2);
    expect(tracks.length).toBe(2);
    expect(tracks.map((t) => t.trackTitle)).not.toContain('Bundle Asset');
    expect(tracks.map((t) => t.trackTitle)).not.toContain('Fee');
    expect(
      parseWorkbook(
        wb([
          headers,
          ...rows.map((r) => [
            r.contractId,
            r.contractName,
            r.shareContract,
            r.assetType,
            r.releaseCode,
            r.albumTitle,
            r.catalogNumber,
            r.barcode,
            r.isrc,
            r.artist,
            r.trackTitle,
            r.usageType,
            r.country,
            r.shop,
            r.salesPeriod,
            r.sales,
            r.returns,
            r.ppu,
            r.amount,
            r.share,
            r.rata1,
            r.rata2,
            r.deduction1,
            r.deduction2,
            r.deduction3,
            r.contractDeductions,
            r.deduction4,
            r.deduction5,
            r.lineCharges,
            r.royaltyRate,
            r.royaltyAmount,
          ]),
        ]),
      ).diagnostics.initialSummary?.uniqueTrackCount,
    ).toBe(tracks.length);
    const csv = toCsv(
      tracks.map((t) => ({
        trackTitle: t.trackTitle,
        isrc: t.isrc,
        royaltyAmount: t.royaltyAmount.toString(),
      })),
      ['trackTitle', 'isrc', 'royaltyAmount'],
    );
    expect(csv).not.toContain('Bundle');
  });
  it('ignores blank and invalid royalty rates for labels and grouped breakdowns', () => {
    expect(
      rateSummary([
        make({ royaltyRate: '0.85' }),
        make({ royaltyRate: '' }),
        make({ royaltyRate: '' }),
      ]).label,
    ).toBe('85%');
    expect(
      rateSummary([
        make({ royaltyRate: '0.85' }),
        make({ royaltyRate: '0.85' }),
        make({ royaltyRate: 'bad' }),
      ]).label,
    ).toBe('85%');
    expect(
      rateSummary([
        make({ royaltyRate: '0.85' }),
        make({ royaltyRate: '1' }),
        make({ royaltyRate: '' }),
      ]).label,
    ).toBe('Multiple rates');
    expect(rateSummary([make({ royaltyRate: '' }), make({ royaltyRate: '' })]).label).toBe(
      'Rate not supplied',
    );
    expect(groupRoyaltyRates([make({ royaltyRate: '' })]).length).toBe(0);
  });
  it('formats visible money as GBP while CSV values remain numeric', () => {
    expect(fmtMoney('11416.23')).toBe('£11,416.23');
    expect(fmtMoney('-125.4')).toBe('-£125.40');
    expect(fmtMoney(0)).toBe('£0.00');
    const csv = toCsv([{ amount: D('11416.23').toString() }], ['amount']);
    expect(csv).toContain('11416.23');
    expect(csv).not.toContain('£');
  });
});

describe('chart cache key correctness regressions', () => {
  const tx = (i: number, over: any = {}) => ({
    sourceSheet: 'Sheet 1',
    sourceRow: i,
    contractId: '',
    contractName: '',
    shareContract: '',
    assetType: 'Track',
    releaseCode: `R${i % 17}`,
    albumTitle: `Release ${i % 17}`,
    catalogNumber: `CAT${i % 17}`,
    barcode: `BC${i % 17}`,
    isrc: `ISRC${i}`,
    artist: `Artist ${i % 13}`,
    trackTitle: `Track ${i}`,
    usageType: 'Track Stream',
    country: 'GB',
    shop: i > 2000 ? 'Tail Shop' : 'Shared Shop',
    salesPeriod: '2026-06',
    sales: '1',
    returns: '0',
    ppu: '',
    amount: '1',
    share: '',
    rata1: '',
    rata2: '',
    deduction1: '',
    deduction2: '',
    deduction3: '',
    contractDeductions: '',
    deduction4: '',
    deduction5: '',
    lineCharges: '',
    royaltyRate: '0.85',
    royaltyAmount: '1',
    ...over,
  });
  function datasets() {
    const shared = Array.from({ length: 2000 }, (_, i) => tx(i + 1));
    const a = [
      ...shared,
      ...Array.from({ length: 25 }, (_, i) =>
        tx(2001 + i, {
          sourceSheet: 'Tail A',
          sourceRow: 2001 + i,
          shop: 'Tail A Shop',
          royaltyAmount: '10',
          amount: '10',
        }),
      ),
    ];
    const b = [
      ...shared,
      ...Array.from({ length: 25 }, (_, i) =>
        tx(2001 + i, {
          sourceSheet: 'Tail B',
          sourceRow: 2001 + i,
          shop: 'Tail B Shop',
          royaltyAmount: '20',
          amount: '20',
        }),
      ),
    ];
    return { a, b };
  }
  it('uses every stable row identifier so datasets differing only after row 2000 get different keys', async () => {
    const { a, b } = datasets();
    expect(a.slice(0, 2000)).toEqual(b.slice(0, 2000));
    expect(a.length).toBe(b.length);
    await expect(chartDatasetCacheKey(a)).resolves.not.toBe(await chartDatasetCacheKey(b));
  });
  it('does not reuse stale chart aggregates when switching large filtered datasets', async () => {
    const { a, b } = datasets();
    const cache: Record<string, any> = {};
    const firstA = await overviewChartDataForRows(a, cache);
    const firstB = await overviewChartDataForRows(b, cache);
    const secondA = await overviewChartDataForRows(a, cache);
    expect(firstA.fromCache).toBe(false);
    expect(firstB.fromCache).toBe(false);
    expect(secondA.fromCache).toBe(true);
    expect(firstA.key).not.toBe(firstB.key);
    expect(firstA.data.filteredRoyaltyTotal.toString()).toBe(totals(a).royaltyAmount.toString());
    expect(firstB.data.filteredRoyaltyTotal.toString()).toBe(totals(b).royaltyAmount.toString());
    expect(firstA.data.filteredRoyaltyTotal.toString()).not.toBe(
      firstB.data.filteredRoyaltyTotal.toString(),
    );
    expect(firstA.data.topShops.map((x: any) => x.shop)).toContain('Tail A Shop');
    expect(firstB.data.topShops.map((x: any) => x.shop)).toContain('Tail B Shop');
  });
  it('calculates a large complete-dataset key and aligned chart totals within a reasonable time', async () => {
    const rows = Array.from({ length: 10000 }, (_, i) =>
      tx(i + 1, {
        sourceSheet: `S${i % 4}`,
        sourceRow: i + 1,
        royaltyAmount: String((i % 5) + 1),
        amount: String((i % 5) + 1),
      }),
    );
    const st = safeNow();
    const key = await chartDatasetCacheKey(rows);
    const charts = prepareOverviewChartDataFromRows(rows);
    expect(safeNow() - st).toBeLessThan(5000);
    expect(key).toMatch(/^chart-v2:10000:/);
    expect(charts.filteredRoyaltyTotal.toString()).toBe(totals(rows).royaltyAmount.toString());
  });
});

describe('lazy import check and timing regressions', () => {
  const tx = (i: number, over: any = {}) => ({
    sourceSheet: 's',
    sourceRow: i,
    contractId: '',
    contractName: '',
    shareContract: '',
    assetType: 'Track',
    releaseCode: `R${i % 2}`,
    albumTitle: `Release ${i % 3}`,
    catalogNumber: `CAT${i % 3}`,
    barcode: `BC${i % 3}`,
    isrc: `ISRC${i}`,
    artist: `Artist ${i % 4}`,
    trackTitle: `Track ${i}`,
    usageType: i % 2 ? 'Track Stream' : 'Track DL',
    country: i % 2 ? 'GB' : 'US',
    shop: `Shop ${i % 5}`,
    salesPeriod: `2026-0${(i % 3) + 1}`,
    sales: String(i + 1),
    returns: String(i % 2),
    ppu: '',
    amount: (i + 1).toString(),
    share: '',
    rata1: '',
    rata2: '',
    deduction1: '',
    deduction2: '',
    deduction3: '',
    contractDeductions: '',
    deduction4: '',
    deduction5: '',
    lineCharges: '',
    royaltyRate: i % 2 ? '0.85' : '',
    royaltyAmount: ((i + 1) * 0.85).toString(),
    ...over,
  });
  const rows = Array.from({ length: 12 }, (_, i) => tx(i));
  it('calculates every Import Checks group without visiting lazy tabs first', () => {
    const groups = importCheckGroups(rows);
    expect(groups.shops.length).toBeGreaterThan(0);
    expect(groups.countries.length).toBeGreaterThan(0);
    expect(groups.salesPeriods.length).toBeGreaterThan(0);
    expect(groups.usageTypes.length).toBeGreaterThan(0);
    expect(groups.royaltyRates.length).toBe(1);
    expect(groupedReconciliation(rows, groups.shops).reconciled).toBe(true);
    expect(groupedReconciliation(rows, groups.countries).reconciled).toBe(true);
    expect(groupedReconciliation(rows, groups.salesPeriods).reconciled).toBe(true);
    expect(groupedReconciliation(rows, groups.usageTypes).reconciled).toBe(true);
    expect(groupedReconciliation(rows, []).reconciled).toBe(false);
    expect(importCheckGroups(rows)).toEqual(groups);
  });
  it('creates optional timing diagnostics safely for sample, missing, partial, and unavailable performance cases', () => {
    const sample = sampleStatement();
    expect(sample.diagnostics.statementHealth.fileStatus).toBe('Original statement appears intact');
    expect(sample.diagnostics.statementHealth.barcodeWarnings).toBe(0);
    expect(sample.diagnostics.importTimings).toBeUndefined();
    recordTiming(sample.diagnostics, 'tabFirstCalculationMs', 'Shops', 1);
    recordTiming(sample.diagnostics, 'tabFirstCalculationMs', 'Countries', 2);
    recordTiming(sample.diagnostics, 'tabFirstCalculationMs', 'Sales Periods', 3);
    recordTiming(sample.diagnostics, 'tabFirstCalculationMs', 'Import Checks', 4);
    recordTimingValue(sample.diagnostics, 'dashboardFirstRenderMs', 5);
    expect(sample.diagnostics.importTimings?.tabFirstCalculationMs?.Shops).toBe(1);
    expect(sample.diagnostics.importTimings?.tabFirstCalculationMs?.Countries).toBe(2);
    expect(sample.diagnostics.importTimings?.tabFirstCalculationMs?.['Sales Periods']).toBe(3);
    expect(sample.diagnostics.importTimings?.tabFirstCalculationMs?.['Import Checks']).toBe(4);
    expect(sample.diagnostics.importTimings?.dashboardFirstRenderMs).toBe(5);
    const partial: any = { importTimings: { searchIndexMs: { artist: 10 } } };
    expect(() => recordTiming(partial, 'searchIndexMs', 'shop', 6)).not.toThrow();
    expect(partial.importTimings.searchIndexMs.shop).toBe(6);
    expect(() => ensureImportTimings(null)).not.toThrow();
    const original = (globalThis as any).performance;
    (globalThis as any).performance = {
      now() {
        throw new Error('timer unavailable');
      },
    };
    expect(() => safeNow()).not.toThrow();
    (globalThis as any).performance = original;
  });
});

describe('workbook opening regression coverage', () => {
  it('opens a generated workbook from a non-empty ArrayBuffer', async () => {
    const bytes = XLSX.write(wb([headers, row]), { type: 'array', bookType: 'xlsx' });
    const { parseArrayBuffer } = await import('./parser');
    const data = await parseArrayBuffer(bytes, 'open.xlsx', bytes.byteLength);
    expect(data.rows.length).toBe(1);
    expect(data.diagnostics.importTimings?.workbookOpeningMs).toBeGreaterThanOrEqual(0);
  });
  it('preserves worker errors and serialises non-Error values with non-blank technical details', async () => {
    const { serialiseThrown, technicalDetails } = await import('./importDiagnostics');
    const thrown = { code: 'BOOM', nested: { why: 'test' } };
    const s = serialiseThrown(thrown);
    expect(s.message).toContain('BOOM');
    const details = technicalDetails(
      {
        processingStage: 'Opening workbook',
        workerEvent: 'error',
        workbookFilename: 'bad.xlsx',
        fileSize: 123,
        arrayBufferSizeBeforeTransfer: 123,
        arrayBufferSizeInsideWorker: 123,
        sheetJsReadAttempt: 'conservative',
        retryAttemptNumber: 2,
      },
      thrown,
    );
    for (const label of [
      'error name',
      'error message',
      'stack trace',
      'processing stage',
      'worker event',
      'workbook filename',
      'file size',
      'ArrayBuffer size before transfer',
      'ArrayBuffer size inside the worker',
      'SheetJS read attempt',
      'retry attempt number',
    ])
      expect(details).toContain(label);
    for (const label of [
      'line number',
      'worker started',
      'ArrayBuffer transfer succeeded',
      'worker received expected payload',
      'SheetJS imported inside worker',
      'XLSX.read invoked',
    ])
      expect(details).toContain(label);
    expect(details.trim().length).toBeGreaterThan(0);
  });
  it('rejects empty ArrayBuffers before SheetJS read', async () => {
    const { parseArrayBuffer } = await import('./parser');
    await expect(parseArrayBuffer(new ArrayBuffer(0), 'empty.xlsx', 0)).rejects.toThrow(/empty/);
  });
  it('supports conservative retry options after a failed first opening attempt', async () => {
    const parser = await import('./parser');
    const bytes = XLSX.write(wb([headers, row]), { type: 'array', bookType: 'xlsx' });
    await expect(
      parser.parseArrayBuffer(bytes.slice(0), 'retry.xlsx', bytes.byteLength, {
        readOptions: parser.CONSERVATIVE_SHEETJS_READ_OPTIONS,
        sheetJsReadAttempt: 'conservative',
        retryAttemptNumber: 2,
      }),
    ).resolves.toMatchObject({ rows: [expect.objectContaining({ artist: 'Artist A' })] });
  });
});

describe('barcode-first release grouping regressions', () => {
  const make = (over: any = {}) => ({
    sourceSheet: 's',
    sourceRow: 1,
    contractId: '',
    contractName: '',
    shareContract: '',
    assetType: 'Track',
    releaseCode: 'RC1',
    albumTitle: 'Album',
    catalogNumber: 'CAT1',
    barcode: 'BC1',
    isrc: 'ISRC1',
    artist: 'Artist',
    trackTitle: 'Track',
    usageType: 'Track Stream',
    country: 'GB',
    shop: 'Shop',
    salesPeriod: '2026-06',
    sales: '1',
    returns: '0',
    ppu: '',
    amount: '10.123456',
    share: '',
    rata1: '',
    rata2: '',
    deduction1: '',
    deduction2: '',
    deduction3: '',
    contractDeductions: '',
    deduction4: '',
    deduction5: '',
    lineCharges: '',
    royaltyRate: '0.85',
    royaltyAmount: '8.605',
    ...over,
  });
  const releaseCsv = (groups: any[]) =>
    toCsv(
      groups.map((g) => ({
        barcode: g.barcode,
        albumTitle: g.albumTitle,
        artist: g.artist,
        catalogNumber: g.catalogNumber,
        releaseCode: g.releaseCode,
        royaltyAmount: g.royaltyAmount.toString(),
        amount: g.amount.toString(),
        sales: g.sales.toString(),
        returns: g.returns.toString(),
        trackCount: g.trackCount.toString(),
        transactionRows: g.transactionRows.toString(),
      })),
      [
        'barcode',
        'albumTitle',
        'artist',
        'catalogNumber',
        'releaseCode',
        'royaltyAmount',
        'amount',
        'sales',
        'returns',
        'trackCount',
        'transactionRows',
      ],
    );

  it('preserves customer-facing fallbacks for blank release descriptions without filling identifiers', () => {
    const rows = [
      make({
        barcode: 'BC-BLANK',
        artist: '',
        albumTitle: '',
        catalogNumber: '',
        releaseCode: '',
        amount: '4',
        royaltyAmount: '3',
        sales: '7',
      }),
      make({
        barcode: 'BC-BLANK',
        artist: '',
        albumTitle: '',
        catalogNumber: '',
        releaseCode: '',
        amount: '6',
        royaltyAmount: '5',
        sales: '9',
      }),
    ];
    const groups = groupReleases(rows);
    expect(groups).toHaveLength(1);
    expect(groups[0].artist).toBe('Unspecified artist');
    expect(groups[0].albumTitle).toBe('Unspecified release');
    expect(groups[0].barcode).toBe('BC-BLANK');
    expect(groups[0].catalogNumber).toBe('');
    expect(groups[0].releaseCode).toBe('');
    expect(groupedReconciliation(rows, groups).reconciled).toBe(true);
    const csv = releaseCsv(groups);
    expect(csv).toContain('Unspecified artist');
    expect(csv).toContain('Unspecified release');
  });
  it('keeps multiple populated release descriptions visible in a single barcode-led group', () => {
    const rows = [
      make({ barcode: 'BC-MULTI', artist: 'Artist A', albumTitle: 'Album A' }),
      make({
        barcode: 'BC-MULTI',
        artist: 'Artist B',
        albumTitle: 'Album B',
        amount: '2',
        royaltyAmount: '1',
      }),
    ];
    const groups = groupReleases(rows);
    expect(groups).toHaveLength(1);
    expect(groups[0].artist).toBe('Artist A | Artist B');
    expect(groups[0].albumTitle).toBe('Album A | Album B');
    expect(groups[0].artists).toEqual(['Artist A', 'Artist B']);
    expect(groups[0].albumTitles).toEqual(['Album A', 'Album B']);
    expect(groupedReconciliation(rows, groups).dashboard.royaltyAmount.toString()).toBe(
      groupedReconciliation(rows, groups).grouped.royaltyAmount.toString(),
    );
  });
  it('groups multiple rows with the same Barcode despite different Catalog Numbers', () => {
    const groups = groupReleases([
      make({ catalogNumber: 'CAT-A' }),
      make({ catalogNumber: 'CAT-B', amount: '2', royaltyAmount: '1' }),
    ]);
    expect(groups).toHaveLength(1);
    expect(groups[0].barcode).toBe('BC1');
    expect(groups[0].catalogNumber).toBe('CAT-A | CAT-B');
    expect(groups[0].catalogNumbers).toEqual(['CAT-A', 'CAT-B']);
    expect(groups[0].transactionRows).toBe(2);
  });
  it('groups multiple rows with the same Barcode despite different Release Codes', () => {
    const groups = groupReleases([make({ releaseCode: 'R-A' }), make({ releaseCode: 'R-B' })]);
    expect(groups).toHaveLength(1);
    expect(groups[0].releaseCode).toBe('R-A | R-B');
    expect(groups[0].releaseCodes).toEqual(['R-A', 'R-B']);
  });
  it('groups track and bundle rows with the same Barcode together', () => {
    const groups = groupReleases([
      make({ assetType: 'Track', isrc: 'ISRC-T' }),
      make({
        assetType: 'Bundle',
        usageType: 'Bundle DL',
        isrc: '',
        catalogNumber: 'BUNDLE-CAT',
        releaseCode: 'BUNDLE-RC',
      }),
    ]);
    expect(groups).toHaveLength(1);
    expect(groups[0].trackCount).toBe(1);
    expect(groups[0].transactionRows).toBe(2);
  });
  it('prioritises populated Barcode over Catalog Number', () => {
    const groups = groupReleases([
      make({ barcode: 'BC-A', catalogNumber: 'CAT-SAME' }),
      make({ barcode: 'BC-B', catalogNumber: 'CAT-SAME' }),
    ]);
    expect(groups).toHaveLength(2);
    expect(groups.map((g) => g.barcode).sort()).toEqual(['BC-A', 'BC-B']);
  });
  it('falls back to Release Code before Catalog Number when Barcode is blank', () => {
    const groups = groupReleases([
      make({ barcode: '', catalogNumber: 'CAT-FALLBACK', releaseCode: 'R1' }),
      make({ barcode: '', catalogNumber: 'CAT-FALLBACK', releaseCode: 'R2' }),
    ]);
    expect(groups).toHaveLength(2);
    expect(groups.map((g) => g.releaseCode).sort()).toEqual(['R1', 'R2']);
  });
  it('falls back to Release Code when Barcode is blank', () => {
    const groups = groupReleases([
      make({ barcode: '', catalogNumber: '', releaseCode: 'RC-FALLBACK', albumTitle: 'A' }),
      make({ barcode: '', catalogNumber: '', releaseCode: 'RC-FALLBACK', albumTitle: 'B' }),
    ]);
    expect(groups).toHaveLength(1);
    expect(groups[0].releaseCode).toBe('RC-FALLBACK');
  });

  it('falls back to Catalog Number when Barcode and Release Code are blank', () => {
    const groups = groupReleases([
      make({ barcode: '', releaseCode: '', catalogNumber: 'CAT-FALLBACK', albumTitle: 'A' }),
      make({ barcode: '', releaseCode: '', catalogNumber: 'CAT-FALLBACK', albumTitle: 'B' }),
    ]);
    expect(groups).toHaveLength(1);
    expect(groups[0].catalogNumber).toBe('CAT-FALLBACK');
  });
  it('uses Artist and Album Title as the final fallback', () => {
    const groups = groupReleases([
      make({
        barcode: '',
        catalogNumber: '',
        releaseCode: '',
        artist: 'Artist X',
        albumTitle: 'Album X',
      }),
      make({
        barcode: '',
        catalogNumber: '',
        releaseCode: '',
        artist: 'Artist X',
        albumTitle: 'Album X',
        trackTitle: 'Other',
      }),
    ]);
    expect(groups).toHaveLength(1);
    expect(groups[0].artist).toBe('Artist X');
    expect(groups[0].albumTitle).toBe('Album X');
  });
  it('release export contains one row per Barcode group and reconciles to dashboard totals', () => {
    const rows = [
      make({
        barcode: 'BC-A',
        catalogNumber: 'CAT1',
        amount: '1.111',
        royaltyAmount: '0.1234',
        sales: '2',
      }),
      make({
        barcode: 'BC-A',
        catalogNumber: 'CAT2',
        amount: '2.222',
        royaltyAmount: '0.2345',
        sales: '3',
      }),
      make({ barcode: 'BC-B', amount: '3.333', royaltyAmount: '0.3456', returns: '1' }),
    ];
    const groups = groupReleases(rows);
    expect(groups).toHaveLength(2);
    const recon = groupedReconciliation(rows, groups);
    expect(recon.reconciled).toBe(true);
    expect(recon.grouped.royaltyAmount.toString()).toBe(totals(rows).royaltyAmount.toString());
    const csv = releaseCsv(groups);
    expect(csv.split('\n')).toHaveLength(groups.length + 1);
    expect(csv).not.toContain('TOTAL');
  });
  it('pagination and top-N charts do not limit release export source data', () => {
    const rows = Array.from({ length: 30 }, (_, i) =>
      make({
        barcode: `BC-${i}`,
        amount: String(i + 1),
        royaltyAmount: String((i + 1) / 10),
        sales: String(i + 1),
      }),
    );
    const groups = groupReleases(rows);
    const page = groups.slice(0, 10);
    expect(groups.length).toBe(30);
    expect(groupedReconciliation(rows, groups).reconciled).toBe(true);
    expect(groupedReconciliation(rows, page).reconciled).toBe(false);
    expect(releaseCsv(groups).split('\n')).toHaveLength(31);
  });
});

describe('overview chart data regressions', () => {
  const make = (i: number, over: any = {}) => ({
    sourceSheet: 's',
    sourceRow: i,
    contractId: '',
    contractName: '',
    shareContract: '',
    assetType: 'Track',
    releaseCode: `R${i}`,
    albumTitle: `Album ${i}`,
    catalogNumber: `CAT${i}`,
    barcode: `BC${i}`,
    isrc: `ISRC${i}`,
    artist: `Artist ${i}`,
    trackTitle: `Track ${i}`,
    usageType: `Usage ${i}`,
    country: 'GB',
    shop: `Shop ${i}`,
    salesPeriod: `2026-${String((i % 12) + 1).padStart(2, '0')}`,
    sales: String(i + 1),
    returns: String(99 + i),
    ppu: '',
    amount: String(i + 1),
    share: '',
    rata1: '',
    rata2: '',
    deduction1: '',
    deduction2: '',
    deduction3: '',
    contractDeductions: '',
    deduction4: '',
    deduction5: '',
    lineCharges: '',
    royaltyRate: '0.85',
    royaltyAmount: String(i + 1),
    ...over,
  });
  it('limits chart datasets without limiting complete grouped tables and never exposes returns', () => {
    const rows = Array.from({ length: 15 }, (_, i) =>
      make(i, i < 2 ? { barcode: 'SAME', albumTitle: '', artist: '' } : {}),
    );
    const releases = groupReleases(rows),
      artists = groupArtists(rows),
      shops = groupBy(rows, 'shop'),
      salesPeriods = groupBy(rows, 'salesPeriod'),
      usageTypes = groupBy(rows, 'usageType');
    const charts = prepareOverviewChartData({ releases, artists, shops, salesPeriods, usageTypes });
    expect(releases.length).toBe(14);
    expect(charts.topReleases.length).toBe(14);
    expect(charts.topArtists.length).toBe(14);
    expect(charts.topShops.length).toBe(15);
    expect(charts.usageTypes.length).toBe(15);
    expect(charts.topReleases[0]).not.toHaveProperty('returns');
    expect(charts.topShops[0]).not.toHaveProperty('returns');
    expect(JSON.stringify(charts)).not.toContain('Returns');
    expect(groupedReconciliation(rows, shops).reconciled).toBe(true);
  });
  it('sorts sales periods chronologically and preserves labels', () => {
    const rows = [
      make(1, { salesPeriod: '2026-12' }),
      make(2, { salesPeriod: '2026_01' }),
      make(3, { salesPeriod: '2026 03' }),
    ];
    const charts = prepareOverviewChartData({
      releases: groupReleases(rows),
      artists: groupArtists(rows),
      shops: groupBy(rows, 'shop'),
      salesPeriods: groupBy(rows, 'salesPeriod'),
      usageTypes: groupBy(rows, 'usageType'),
    });
    expect(charts.salesPeriods.map((x) => x.salesPeriod)).toEqual([
      '2026_01',
      '2026 03',
      '2026-12',
    ]);
  });
  it('keeps returns internal while customer-facing CSV columns can omit it without changing other totals', () => {
    const rows = [make(1, { returns: '12', amount: '10', royaltyAmount: '8', sales: '4' })];
    const groups = groupBy(rows, 'shop');
    const csv = toCsv(
      groups.map((g) => ({
        shop: g.shop,
        royaltyAmount: g.royaltyAmount.toString(),
        amount: g.amount.toString(),
        sales: g.sales.toString(),
      })),
      ['shop', 'royaltyAmount', 'amount', 'sales'],
    );
    expect(csv).not.toContain('returns');
    expect(totals(rows).returns.toString()).toBe('12');
    expect(totals(rows).amount.toString()).toBe('10');
    expect(totals(rows).royaltyAmount.toString()).toBe('8');
    expect(totals(rows).sales.toString()).toBe('4');
  });
});

describe('barcode integrity and statement health', () => {
  it('fails an import when 13-digit barcodes show likely precision-collapse corruption', () => {
    const corrupted = Array.from({ length: 24 }, (_, index) => {
      const item = [...row];
      item[5] = `Distinct Album ${index}`;
      item[6] = `CAT-${index}`;
      item[7] = ['5056320000000', '5056690000000', '5055870000000'][index % 3];
      item[9] = `Distinct Artist ${index}`;
      return item;
    });
    expect(() => parseWorkbook(wb([headers, ...corrupted]), 'corrupted.xlsx')).toThrow(
      /likely barcode precision loss/i,
    );
  });

  it('keeps similar exact 13-digit barcodes as separate releases', () => {
    const barcodes = ['5055869508216', '5055869508353', '5055869508407', '5055869517324'];
    const exactRows = barcodes.map((barcode, index) => {
      const item = [...row];
      item[5] = `Album ${index}`;
      item[6] = `CAT-${index}`;
      item[7] = barcode;
      item[9] = `Artist ${index}`;
      return item;
    });
    const statement = parseWorkbook(wb([headers, ...exactRows]), 'exact-barcodes.xlsx');
    expect(statement.rows.map((item) => item.barcode)).toEqual(barcodes);
    expect(groupReleases(statement.rows)).toHaveLength(barcodes.length);
    expect(statement.diagnostics.initialSummary?.uniqueReleaseCount).toBe(barcodes.length);
  });

  it('clean statement shows intact Excellent health with no barcode warnings', () => {
    const s = parseWorkbook(wb([headers, row]));
    expect(s.diagnostics.statementHealth.fileStatus).toBe('Original statement appears intact');
    expect(s.diagnostics.statementHealth.dataQuality).toBe('Excellent');
    expect(s.diagnostics.statementHealth.barcodeWarnings).toBe(0);
    expect(s.diagnostics.statementHealth.rowsRequiringReview).toBe(0);
    expect(s.diagnostics.barcodeIntegrity.populatedBarcodeRows).toBe(1);
  });

  it('numeric barcode cells alone do not create actionable warnings or review rows', () => {
    const numeric: any[] = [...row];
    numeric[7] = 1234567890123;
    const s = parseWorkbook(wb([headers, numeric]));
    expect(s.diagnostics.barcodeIntegrity.numericBarcodeCells).toBe(1);
    expect(s.diagnostics.barcodeIntegrity.warnings).toHaveLength(0);
    expect(s.diagnostics.barcodeIntegrity.rowsRequiringReview).toBe(0);
    expect(s.diagnostics.statementHealth.fileStatus).toBe('Original statement appears intact');
    expect(s.diagnostics.statementHealth.dataQuality).toBe('Excellent');
  });
  it('real barcode corruption still produces an actionable warning', () => {
    const malformed = [...row];
    malformed[7] = '12345ABC';
    const s = parseWorkbook(wb([headers, malformed]));
    expect(
      s.diagnostics.barcodeIntegrity.warnings.some((w) => w.warning === 'Malformed barcode'),
    ).toBe(true);
    expect(s.diagnostics.statementHealth.dataQuality).toBe('Review recommended');
    expect(s.diagnostics.statementHealth.rowsRequiringReview).toBe(1);
  });

  it('CSV and Excel produce identical totals while preserving CSV leading zeroes', () => {
    const csvWb = XLSX.read(
      [headers.join(','), [...row.slice(0, 7), '0123456789012', ...row.slice(8)].join(',')].join(
        '\n',
      ),
      { type: 'string' },
    );
    const xlsxWb = wb([headers, [...row.slice(0, 7), '0123456789012', ...row.slice(8)]]);
    const csv = parseWorkbook(csvWb, 'statement.csv', 100);
    const excel = parseWorkbook(xlsxWb, 'statement.xlsx', 100);
    expect(csv.rows[0].barcode).toBe('0123456789012');
    expect(excel.rows[0].barcode).toBe('0123456789012');
    expect(typeof csv.rows[0].barcode).toBe('string');
    expect(csv.diagnostics.initialSummary?.totalRoyaltyAmount).toBe(
      excel.diagnostics.initialSummary?.totalRoyaltyAmount,
    );
    expect(csv.diagnostics.initialSummary?.totalAmount).toBe(
      excel.diagnostics.initialSummary?.totalAmount,
    );
  });
  it('minor recoverable scientific notation and decimal suffix warnings show Good', () => {
    const s = parseWorkbook(
      wb([
        headers,
        [...row.slice(0, 7), '1.234E+12', ...row.slice(8)],
        [...row.slice(0, 7), '1234567890123.0', ...row.slice(8)],
      ]),
    );
    expect(s.rows[0].barcode).toBe('1234000000000');
    expect(s.rows[1].barcode).toBe('1234567890123');
    expect(s.diagnostics.barcodeIntegrity.scientificNotationValuesConverted).toBe(1);
    expect(s.diagnostics.barcodeIntegrity.decimalSuffixesRemoved).toBe(1);
    expect(s.diagnostics.statementHealth.dataQuality).toBe('Good');
  });
  it('unsafe precision, short values, blanks, and duplicate conflicts are reported without guessing digits', () => {
    const numeric: any[] = [...row];
    numeric[7] = 1234567890123;
    const unsafe = [...row];
    unsafe[7] = '12345678901234567';
    const short = [...row];
    short[7] = '12345';
    const blank = [...row];
    blank[7] = '';
    const conflictA = [...row];
    conflictA[5] = 'Conflict';
    conflictA[6] = 'CATX';
    conflictA[7] = '0012345678901';
    const conflictB = [...conflictA];
    conflictB[7] = '12345678901';
    const s = parseWorkbook(wb([headers, numeric, unsafe, short, blank, conflictA, conflictB]));
    expect(s.rows.find((r) => r.sourceRow === 4)?.barcode).toBe('12345');
    expect(s.diagnostics.barcodeIntegrity.blankBarcodeRows).toBe(1);
    expect(s.diagnostics.barcodeIntegrity.numericBarcodeCells).toBeGreaterThan(0);
    expect(s.diagnostics.barcodeIntegrity.unsafePrecisionWarnings).toBeGreaterThan(0);
    expect(s.diagnostics.barcodeIntegrity.possibleLostLeadingZeroWarnings).toBeGreaterThan(0);
    expect(s.diagnostics.barcodeIntegrity.duplicateBarcodeConflicts).toBeGreaterThan(0);
    expect(s.diagnostics.statementHealth.dataQuality).toBe('Review recommended');
    expect(s.diagnostics.statementHealth.rowsRequiringReview).toBe(
      new Set(
        s.diagnostics.barcodeIntegrity.warnings
          .filter((w) => w.severity === 'review')
          .map((w) => `${w.sourceSheet}:${w.sourceRow}`),
      ).size,
    );
  });

  it('uses underlying Excel values instead of scientific display text for barcode identifiers and grouping', () => {
    const rows = [
      [...row.slice(0, 5), 'Album Sci A', 'CAT-A', '888831328476', ...row.slice(8)],
      [...row.slice(0, 5), 'Album Sci B', 'CAT-B', '5055301234567', ...row.slice(8)],
      [...row.slice(0, 5), 'Album Sci C', 'CAT-C', '5055309876543', ...row.slice(8)],
    ];
    const book = wb([headers, ...rows]);
    const sheet = book.Sheets['Digital Sales'];
    sheet['H2'] = { t: 'n', v: 888831328476, w: '8.88831E+11' } as XLSX.CellObject;
    sheet['H3'] = { t: 'n', v: 5055301234567, w: '5.0553E+12' } as XLSX.CellObject;
    sheet['H4'] = { t: 'n', v: 5055309876543, w: '5.0553E+12' } as XLSX.CellObject;
    const s = parseWorkbook(book, 'scientific-display.xlsx', 100);
    expect(s.rows.map((r) => r.barcode)).toEqual([
      '888831328476',
      '5055301234567',
      '5055309876543',
    ]);
    expect(s.rows.map((r) => typeof r.barcode)).toEqual(['string', 'string', 'string']);
    const releases = groupReleases(s.rows);
    expect(releases).toHaveLength(3);
    expect(releases.map((g) => g.barcode).sort()).toEqual([
      '5055301234567',
      '5055309876543',
      '888831328476',
    ]);
    expect(releases.find((g) => g.barcode === '5055301234567')?.albumTitle).toBe('Album Sci B');
    expect(releases.find((g) => g.barcode === '5055309876543')?.albumTitle).toBe('Album Sci C');
    expect(s.diagnostics.barcodeIntegrity.scientificNotationValuesConverted).toBe(0);
  });
  it('uses underlying Excel values for release identifier fields when formatted display text is rounded', () => {
    const idRow = [...row];
    idRow[4] = '123456789012';
    idRow[6] = '5055301234567';
    idRow[7] = '888831328476';
    idRow[8] = '123456789012';
    const book = wb([headers, idRow]);
    const sheet = book.Sheets['Digital Sales'];
    sheet['E2'] = { t: 'n', v: 123456789012, w: '1.23457E+11' } as XLSX.CellObject;
    sheet['G2'] = { t: 'n', v: 5055301234567, w: '5.0553E+12' } as XLSX.CellObject;
    sheet['H2'] = { t: 'n', v: 888831328476, w: '8.88831E+11' } as XLSX.CellObject;
    sheet['I2'] = { t: 'n', v: 123456789012, w: '1.23457E+11' } as XLSX.CellObject;
    const s = parseWorkbook(book, 'identifier-display.xlsx', 100);
    expect(s.rows[0].releaseCode).toBe('123456789012');
    expect(s.rows[0].catalogNumber).toBe('5055301234567');
    expect(s.rows[0].barcode).toBe('888831328476');
    expect(s.rows[0].isrc).toBe('123456789012');
  });

  it('keeps barcode as text for leading-zero, numeric, scientific, grouping, and export paths', () => {
    const numeric = [...row];
    numeric[7] = '1234567890123';
    const leading = [...row];
    leading[7] = '0012345678901';
    const sci = [...row];
    sci[7] = '1.234E+12';
    const s = parseWorkbook(wb([headers, leading, numeric, sci]));
    expect(s.rows.map((r) => typeof r.barcode)).toEqual(['string', 'string', 'string']);
    expect(s.rows[0].barcode).toBe('0012345678901');
    expect(s.rows[1].barcode).toBe('1234567890123');
    expect(s.rows[2].barcode).toBe('1234000000000');
    expect(groupBy(s.rows, 'barcode').map((g) => g.barcode)).toEqual([
      '0012345678901',
      '1234567890123',
      '1234000000000',
    ]);
    const csv = toCsv(
      s.rows.map((r) => ({ barcode: r.barcode })),
      ['barcode'],
    );
    expect(csv).toContain('0012345678901');
    expect(csv).not.toMatch(/e\+?\d+/i);
  });
  it('finalises 100,000 barcode rows without repeatedly filtering one full rows array per release signature', () => {
    const original = Array.prototype.filter;
    let rowArrayFilterCalls = 0;
    (Array.prototype as any).filter = function (this: any[], ...args: any[]) {
      if (this.length === 100000 && this[0]?.sourceSheet === 'Digital Sales') rowArrayFilterCalls++;
      return original.apply(this, args as any);
    };
    try {
      const rows = Array.from({ length: 100000 }, (_, i) => [
        ...row.slice(0, 5),
        `Album ${i % 1000}`,
        `CAT${i % 1000}`,
        i % 2 ? `00${String(10000000000 + (i % 1000))}` : String(10000000000 + (i % 1000)),
        ...row.slice(8),
      ]);
      const st = safeNow();
      const s = parseWorkbook(wb([headers, ...rows]), 'huge.xlsx', 11 * 1024 * 1024);
      expect(s.rows).toHaveLength(100000);
      expect(s.diagnostics.barcodeIntegrity.populatedBarcodeRows).toBe(100000);
      expect(safeNow() - st).toBeLessThan(30000);
      expect(rowArrayFilterCalls).toBe(0);
      expect(JSON.stringify(s.rows[0])).not.toContain('normalisedBarcodes');
    } finally {
      Array.prototype.filter = original;
    }
  }, 30000);
  it('exports barcode warning report as CSV', () => {
    const s = parseWorkbook(wb([headers, [...row.slice(0, 7), '12345', ...row.slice(8)]]));
    const csv = toCsv(
      s.diagnostics.barcodeIntegrity.warnings.map((w) => ({
        sourceSheet: w.sourceSheet,
        sourceRow: String(w.sourceRow),
        barcodeValue: w.barcodeValue,
        warning: w.warning,
        suggestedReason: w.suggestedReason,
      })),
      ['sourceSheet', 'sourceRow', 'barcodeValue', 'warning', 'suggestedReason'],
    );
    expect(csv).toContain('Possible leading zero removed');
    expect(csv).toContain('12345');
  });
});

describe('default ZIP breakdown export', () => {
  const makeZipRow = (i: number, over: any = {}) => ({
    sourceSheet: 'Digital Sales',
    sourceRow: i,
    contractId: 'C1',
    contractName: 'Contract',
    shareContract: '',
    assetType: 'Track',
    releaseCode: `R${i}`,
    albumTitle: `Album ${i}`,
    catalogNumber: `CAT${i}`,
    barcode: `00123456789${String(i).padStart(2, '0')}`,
    isrc: `ISRC${i}`,
    artist: `Artist ${i}`,
    trackTitle: `Track ${i}`,
    usageType: i % 2 ? 'Stream' : 'Download',
    country: i % 2 ? 'GB' : 'US',
    shop: `Shop ${i % 3}`,
    salesPeriod: `2026-${String((i % 2) + 1).padStart(2, '0')}`,
    sales: String(i + 1),
    returns: '0',
    ppu: '',
    amount: String(i + 1),
    share: '',
    rata1: '',
    rata2: '',
    deduction1: '',
    deduction2: '',
    deduction3: '',
    contractDeductions: '',
    deduction4: '',
    deduction5: '',
    lineCharges: '',
    royaltyRate: '0.85',
    royaltyAmount: String(i + 1),
    ...over,
  });
  function zipNames(bytes: Uint8Array) {
    const s = new TextDecoder().decode(bytes);
    return [...s.matchAll(/\d\d-[a-z-]+\.csv/g)]
      .map((m) => m[0])
      .filter((v, i, a) => a.indexOf(v) === i)
      .sort();
  }
  function fileText(bytes: Uint8Array, name: string) {
    const all = new TextDecoder().decode(bytes);
    const i = all.indexOf(name);
    expect(i).toBeGreaterThanOrEqual(0);
    const after = all.indexOf('\ufeff', i);
    const next = all.indexOf('PK\x01\x02', after);
    return all.slice(after, next > after ? next : undefined);
  }
  it('creates an on-demand ZIP with exactly the five default breakdown CSVs', async () => {
    const { buildExportEntries, createZip } = await import('./exportZip');
    const sample = sampleStatement();
    const expectedNames = [
      '01-artists.csv',
      '02-releases.csv',
      '03-tracks.csv',
      '04-shops.csv',
      '05-countries.csv',
    ];
    const entries = buildExportEntries(sample);
    expect(entries.map((e) => e.name)).toEqual(expectedNames);
    expect(new Set(entries.map((e) => e.name)).size).toBe(entries.length);
    expect(entries.every((e) => String(e.content).split('\n').length > 1)).toBe(true);
    expect(zipNames(createZip(entries))).toEqual(expectedNames);
  });
  it('excludes unsupported empty field breakdowns and exports full grouped data, not top ten rows', async () => {
    const { buildExportEntries } = await import('./exportZip');
    const rows = Array.from({ length: 15 }, (_, i) =>
      makeZipRow(i, { country: '', shop: '', salesPeriod: '' }),
    );
    const data: any = {
      label: 'July 2026',
      filename: 'sample.csv',
      fileSize: 1,
      rows,
      diagnostics: {
        filename: 'sample.csv',
        barcodeIntegrity: { warnings: [] },
        statementHealth: { dataQuality: 'Excellent', barcodeWarnings: 0, rowsRequiringReview: 0 },
        reportingPeriod: 'July 2026',
      },
    };
    const entries = buildExportEntries(data);
    expect(entries.map((e) => e.name)).not.toEqual(
      expect.arrayContaining(['04-shops.csv', '05-countries.csv']),
    );
    const artists = String(entries.find((e) => e.name === '01-artists.csv')?.content);
    expect(artists.split('\n')).toHaveLength(16);
    expect(artists).toContain('Artist 14');
  });
  it('reconciles all default grouped exports and documents track bundle exceptions', async () => {
    const { buildBreakdowns } = await import('./exportZip');
    const rows = [
      makeZipRow(1),
      makeZipRow(2, {
        assetType: 'Bundle',
        usageType: 'Bundle Download',
        isrc: '',
        trackTitle: '',
      }),
    ];
    const data: any = {
      label: 'period',
      filename: 'statement.csv',
      fileSize: 1,
      rows,
      diagnostics: {
        filename: 'statement.csv',
        barcodeIntegrity: { warnings: [] },
        statementHealth: { dataQuality: 'Good', barcodeWarnings: 0, rowsRequiringReview: 0 },
        reportingPeriod: 'period',
      },
    };
    const breakdowns = buildBreakdowns(data);
    for (const label of ['Artists', 'Releases', 'Shops', 'Countries'])
      expect(breakdowns.find((b) => b.label === label)?.recon?.Reconciles).toBe('Yes');
    expect(breakdowns.find((b) => b.label === 'Tracks')?.recon?.Notes).toBe(
      'Tracks exclude bundle/non-track rows under existing track rules.',
    );
  });
  it('blocks ZIP creation on reconciliation failure and reports exact differences', async () => {
    const mod = await import('./exportZip');
    const rows = [makeZipRow(1)];
    const data: any = {
      label: 'period',
      filename: 'statement.csv',
      fileSize: 1,
      rows,
      diagnostics: {
        filename: 'statement.csv',
        barcodeIntegrity: { warnings: [] },
        statementHealth: { dataQuality: 'Good', barcodeWarnings: 0, rowsRequiringReview: 0 },
        reportingPeriod: 'period',
      },
    };
    const original = mod.buildBreakdowns(data);
    original[0].recon = {
      Breakdown: 'Artists',
      'Row Count': 1,
      'Statement Royalty Amount': '2',
      'Breakdown Royalty Amount': '0',
      'Royalty Difference': '-2',
      'Statement Amount': '2',
      'Breakdown Amount': '2',
      'Amount Difference': '0',
      'Statement Sales': '2',
      'Breakdown Sales': '2',
      'Sales Difference': '0',
      Reconciles: 'No',
      Notes: 'test',
    };
    expect(() => mod.buildExportEntries(data, original as any)).toThrow(/does not reconcile.*-2/);
  });
  it('preserves retained CSV barcode text, identifiers and plain money values without scientific notation', async () => {
    const { buildExportEntries } = await import('./exportZip');
    const rows = [
      makeZipRow(1, { barcode: '0012345678901', amount: '10.123456', royaltyAmount: '8.7654321' }),
    ];
    const data: any = {
      label: 'period',
      filename: 'statement.csv',
      fileSize: 1,
      rows,
      diagnostics: {
        filename: 'statement.csv',
        barcodeIntegrity: { warnings: [] },
        statementHealth: { dataQuality: 'Excellent', barcodeWarnings: 0, rowsRequiringReview: 0 },
        reportingPeriod: 'period',
      },
    };
    const entries = buildExportEntries(data);
    expect(entries.map((e) => e.name)).not.toContain('09-full-detail.csv');
    const releases = String(entries.find((e) => e.name === '02-releases.csv')?.content);
    expect(releases.split('\n')).toHaveLength(2);
    expect(releases).toContain('"=""0012345678901"""');
    expect(releases).not.toMatch(/\dE\+\d/i);
    expect(releases).not.toContain('£');
    expect(releases).toContain('10.123456');
    expect(releases).toContain('8.7654321');
  });
  it('downloads only when clicked, revokes Blob URLs, and clears entry references after creating the ZIP', async () => {
    const { downloadBreakdownsZip } = await import('./exportZip');
    const sample = sampleStatement();
    let created = 0,
      revoked = 0,
      clicked = 0;
    (globalThis as any).URL.createObjectURL = () => {
      created++;
      return 'blob:test';
    };
    (globalThis as any).URL.revokeObjectURL = () => {
      revoked++;
    };
    const old = document.createElement.bind(document);
    vi.spyOn(document, 'createElement').mockImplementation(((tag: string) => {
      const el = old(tag);
      if (tag === 'a')
        Object.defineProperty(el, 'click', {
          value: () => {
            clicked++;
          },
        });
      return el;
    }) as any);
    expect(created).toBe(0);
    const progress: any[] = [];
    await downloadBreakdownsZip(sample, (p) => progress.push(p));
    await new Promise((r) => setTimeout(r, 1));
    expect(created).toBe(1);
    expect(clicked).toBe(1);
    expect(revoked).toBe(1);
    expect(progress.map((p) => p.stage)).toEqual([
      'Validating totals',
      'Preparing artists',
      'Preparing releases',
      'Preparing tracks',
      'Preparing shops',
      'Preparing countries',
      'Creating ZIP',
      'Complete',
    ]);
    expect(progress.every((p) => p.total === 8)).toBe(true);
    vi.restoreAllMocks();
  });
  it('creates a large synthetic statement ZIP without retaining every CSV as a separate generated ZIP', async () => {
    const { buildExportEntries, createZip } = await import('./exportZip');
    const rows = Array.from({ length: 1000 }, (_, i) => makeZipRow(i));
    const data: any = {
      label: 'large',
      filename: 'large.csv',
      fileSize: 1,
      rows,
      diagnostics: {
        filename: 'large.csv',
        barcodeIntegrity: { warnings: [] },
        statementHealth: { dataQuality: 'Good', barcodeWarnings: 0, rowsRequiringReview: 0 },
        reportingPeriod: 'large',
      },
    };
    const entries = buildExportEntries(data);
    expect(entries.map((e) => e.name)).toEqual([
      '01-artists.csv',
      '02-releases.csv',
      '03-tracks.csv',
      '04-shops.csv',
      '05-countries.csv',
    ]);
    expect(createZip(entries).length).toBeGreaterThan(1000);
  });
});

describe('large XLSX converter workflow', () => {
  it('keeps small XLSX in browser and offers conversion for large XLSX', async () => {
    const { shouldOfferLargeXlsxConversion } = await import('./largeExcel');
    expect(shouldOfferLargeXlsxConversion({ name: 'small.xlsx', size: 1000 }).shouldConvert).toBe(
      false,
    );
    expect(
      shouldOfferLargeXlsxConversion({ name: 'large.xlsx', size: 11 * 1024 * 1024 }).shouldConvert,
    ).toBe(true);
    expect(
      shouldOfferLargeXlsxConversion({ name: 'large.csv', size: 50 * 1024 * 1024 }).shouldConvert,
    ).toBe(false);
  });

  it('browser automatically imports returned CSV with exact identifiers and reconciled totals', async () => {
    const { parseCsvFileStreaming } = await import('./streamingCsvParser');
    const csv = toCsv([Object.fromEntries(headers.map((h, i) => [h, row[i]]))], headers);
    const data = await parseCsvFileStreaming({
      name: 'converted.csv',
      size: csv.length,
      stream: () => new Response(csv).body,
    } as File);
    expect(data.rows[0].barcode).toBe('1234567890123');
    expect(data.rows[0].catalogNumber).toBe('CAT-001');
    expect(data.rows[0].releaseCode).toBe('R1');
    expect(data.rows[0].isrc).toBe('ISRC1');
    expect(totals(data.rows).royaltyAmount.toString()).toBe(
      data.diagnostics.initialSummary?.totalRoyaltyAmount,
    );
    expect(groupedReconciliation(data.rows, groupArtists(data.rows)).reconciled).toBe(true);
  });
});

class FakeConverterXhr {
  static instances: FakeConverterXhr[] = [];
  upload: { onprogress: ((event: ProgressEvent) => void) | null; onload: (() => void) | null } = {
    onprogress: null,
    onload: null,
  };
  onreadystatechange: (() => void) | null = null;
  onprogress: ((event: ProgressEvent) => void) | null = null;
  onload: (() => void) | null = null;
  onerror: (() => void) | null = null;
  ontimeout: (() => void) | null = null;
  onabort: (() => void) | null = null;
  readyState = 0;
  status = 0;
  responseType = '';
  response: Blob | null = null;
  aborted = false;
  requestHeaders: Record<string, string> = {};
  constructor() {
    FakeConverterXhr.instances.push(this);
  }
  open() {}
  send() {}
  abort() {
    this.aborted = true;
    this.onabort?.();
  }
  setRequestHeader(name: string, value: string) {
    this.requestHeaders[name] = value;
  }
  getResponseHeader(name: string) {
    return name.toLowerCase() === 'x-conversion-started' ? 'true' : null;
  }
}

describe('large XLSX converter response timeout', () => {
  it('uses response inactivity rather than a fixed wall-clock timeout for streaming downloads', async () => {
    vi.useFakeTimers();
    const originalXhr = globalThis.XMLHttpRequest;
    FakeConverterXhr.instances = [];
    vi.stubGlobal('XMLHttpRequest', FakeConverterXhr);
    try {
      const { uploadWorkbookWithProgress } = await import('./converterUpload');
      const diagnostics: Record<string, unknown> = {};
      const stages: string[] = [];
      const promise = uploadWorkbookWithProgress({
        file: new File(['xlsx'], 'large.xlsx', {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        }),
        converterUrl: '/api/convert-xlsx',
        activeConversion: { xhr: null },
        diagnostics,
        onProgress: (stage) => stages.push(stage),
        uploadInactivityMs: 100,
        responseInactivityMs: 100,
      });
      const xhr = FakeConverterXhr.instances[0];
      xhr.upload.onprogress?.({ loaded: 4, total: 4, lengthComputable: true } as ProgressEvent);
      xhr.upload.onload?.();
      xhr.status = 200;
      xhr.readyState = XMLHttpRequest.HEADERS_RECEIVED;
      xhr.onreadystatechange?.();
      for (let loaded = 1; loaded <= 5; loaded++) {
        await vi.advanceTimersByTimeAsync(90);
        xhr.onprogress?.({ loaded, total: 10, lengthComputable: true } as ProgressEvent);
        expect(xhr.aborted).toBe(false);
      }
      xhr.response = new Blob(['converted,csv']);
      xhr.onload?.();
      await expect(promise).resolves.toBeUndefined();
      expect(diagnostics.abortReason).toBeUndefined();
      expect(diagnostics.responseSize).toBe(xhr.response.size);
      expect(stages).toContain('Receiving converted data');
    } finally {
      vi.stubGlobal('XMLHttpRequest', originalXhr);
      vi.useRealTimers();
    }
  });

  it('aborts when a streaming response becomes inactive', async () => {
    vi.useFakeTimers();
    const originalXhr = globalThis.XMLHttpRequest;
    FakeConverterXhr.instances = [];
    vi.stubGlobal('XMLHttpRequest', FakeConverterXhr);
    try {
      const { uploadWorkbookWithProgress } = await import('./converterUpload');
      const diagnostics: Record<string, unknown> = {};
      const promise = uploadWorkbookWithProgress({
        file: new File(['xlsx'], 'large.xlsx'),
        converterUrl: '/api/convert-xlsx',
        activeConversion: { xhr: null },
        diagnostics,
        onProgress: () => {},
        uploadInactivityMs: 100,
        responseInactivityMs: 100,
      });
      const rejection = expect(promise).rejects.toThrow('The converter response timed out.');
      const xhr = FakeConverterXhr.instances[0];
      xhr.upload.onprogress?.({ loaded: 4, total: 4, lengthComputable: true } as ProgressEvent);
      xhr.upload.onload?.();
      xhr.status = 200;
      xhr.readyState = XMLHttpRequest.HEADERS_RECEIVED;
      xhr.onreadystatechange?.();
      xhr.onprogress?.({ loaded: 1, total: 10, lengthComputable: true } as ProgressEvent);
      await vi.advanceTimersByTimeAsync(101);
      await rejection;
      expect(diagnostics.abortReason).toBe('converter-response-timeout');
      expect(xhr.aborted).toBe(true);
    } finally {
      vi.stubGlobal('XMLHttpRequest', originalXhr);
      vi.useRealTimers();
    }
  });
});
