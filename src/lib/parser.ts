import Decimal from 'decimal.js';
import * as XLSX from 'xlsx';
import {
  columnAliases,
  fieldLabels,
  requiredFields,
  resolveHeader,
  type FieldKey,
} from './columnAliases';
import { D, text } from './format';
import { recordTimingValue, safeNow } from './timing';
import { clampProgress, isLargeFile, type ImportProgress } from './importProgress';
import { releaseKey, isTrackRow } from './analytics';
import type {
  BarcodeWarning,
  ImportDiagnostics,
  ImportSummary,
  StatementData,
  Transaction,
} from './types';
const textFields = new Set<FieldKey>([
  'contractId',
  'releaseCode',
  'catalogNumber',
  'barcode',
  'isrc',
  'country',
]);
const DEFAULT_CHUNK_SIZE = 1000;
const LARGE_FILE_CHUNK_SIZE = 500;
export interface ParseOptions {
  onProgress?: (p: ImportProgress) => void;
  onBeforeSheetJsRead?: () => void;
  largeFileMode?: boolean;
  fileReadMs?: number;
  estimatedRows?: number;
  readOptions?: XLSX.ParsingOptions;
  sheetJsReadAttempt?: string;
  retryAttemptNumber?: number;
}
export function parseFilename(filename: string) {
  const m = filename.match(/^([^_]+)___(\d{4})_(\d{2})_(\d{2})___(\d{4})_(\d{2})_(\d{2})/);
  if (!m) return {};
  const f = (y: string, mo: string, d: string) =>
    new Date(Date.UTC(+y, +mo - 1, +d)).toLocaleDateString('en-GB', {
      day: 'numeric',
      month: 'long',
      year: 'numeric',
      timeZone: 'UTC',
    });
  return { account: m[1], reportingPeriod: `${f(m[2], m[3], m[4])} to ${f(m[5], m[6], m[7])}` };
}
export function normaliseIdentifier(v: unknown) {
  const s = text(v);
  if (/^\d+(\.\d+)?e\+?\d+$/i.test(s)) {
    try {
      return new Decimal(s).toFixed(0);
    } catch {
      return s;
    }
  }
  return s;
}
function emptyBarcodeReport() {
  return {
    populatedBarcodeRows: 0,
    blankBarcodeRows: 0,
    uniqueBarcodeCount: 0,
    numericBarcodeCells: 0,
    textBarcodeCells: 0,
    scientificNotationValuesConverted: 0,
    decimalSuffixesRemoved: 0,
    possibleLostLeadingZeroWarnings: 0,
    unsafePrecisionWarnings: 0,
    duplicateBarcodeConflicts: 0,
    rowsRequiringReview: 0,
    warnings: [] as BarcodeWarning[],
  };
}
function emptyHealth() {
  return {
    fileStatus: 'Original statement appears intact' as const,
    dataQuality: 'Excellent' as const,
    barcodeWarnings: 0,
    rowsRequiringReview: 0,
  };
}
function addBarcodeWarning(
  d: ImportDiagnostics,
  sourceSheet: string,
  sourceRow: number,
  barcodeValue: string,
  warning: string,
  suggestedReason: string,
  severity: 'minor' | 'review',
) {
  d.barcodeIntegrity.warnings.push({
    sourceSheet,
    sourceRow,
    barcodeValue,
    warning,
    suggestedReason,
    severity,
  });
}
function identifierCellText(ws: XLSX.WorkSheet, rowIndex: number, colIndex: number) {
  if (colIndex < 0) return '';
  const cell = ws[XLSX.utils.encode_cell({ r: rowIndex, c: colIndex })] as
    XLSX.CellObject | undefined;
  if (!cell) return '';
  const rawText = text(cell.v);
  const displayText = text(cell.w);
  let value = rawText || displayText;
  if (/^0+\d+$/.test(displayText) && /^\d+$/.test(rawText) && displayText.endsWith(rawText))
    value = displayText;
  if (!value) return '';
  if (/^\d+(\.\d+)?e\+?\d+$/i.test(value)) value = normaliseIdentifier(value);
  if (/^\d+\.0+$/.test(value)) value = value.replace(/\.0+$/, '');
  return value;
}
function barcodeCellText(
  ws: XLSX.WorkSheet,
  rowIndex: number,
  colIndex: number,
  sheet: string,
  sourceRow: number,
  d: ImportDiagnostics,
) {
  if (colIndex < 0) return '';
  const cell = ws[XLSX.utils.encode_cell({ r: rowIndex, c: colIndex })] as
    XLSX.CellObject | undefined;
  const rawText = text(cell?.v);
  const displayText = text(cell?.w);
  let value = rawText || displayText;
  if (/^0+\d+$/.test(displayText) && /^\d+$/.test(rawText) && displayText.endsWith(rawText))
    value = displayText;
  if (!value) return '';
  if (cell?.t === 'n') d.barcodeIntegrity.numericBarcodeCells++;
  else d.barcodeIntegrity.textBarcodeCells++;
  const usedDisplay = !rawText && !!displayText;
  if (/^\d+(\.\d+)?e\+?\d+$/i.test(value)) {
    const converted = normaliseIdentifier(value);
    d.barcodeIntegrity.scientificNotationValuesConverted++;
    addBarcodeWarning(
      d,
      sheet,
      sourceRow,
      value,
      usedDisplay ? 'Displayed as scientific notation' : 'Stored as scientific notation',
      usedDisplay
        ? 'Only the formatted display text was available for this barcode.'
        : 'Spreadsheet software may have converted this barcode into scientific notation.',
      'minor',
    );
    value = converted;
  }
  if (/^\d+\.0+$/.test(value)) {
    d.barcodeIntegrity.decimalSuffixesRemoved++;
    addBarcodeWarning(
      d,
      sheet,
      sourceRow,
      value,
      'Decimal suffix removed',
      'Spreadsheet software may have treated this barcode as a number.',
      'minor',
    );
    value = value.replace(/\.0+$/, '');
  }
  if (/^\d+$/.test(value) && value.length < 12) {
    d.barcodeIntegrity.possibleLostLeadingZeroWarnings++;
    addBarcodeWarning(
      d,
      sheet,
      sourceRow,
      value,
      'Possible leading zero removed',
      'This barcode is unusually short and may have lost leading zeroes before upload.',
      'review',
    );
  }
  if (/^\d+$/.test(value) && value.length > 14) {
    addBarcodeWarning(
      d,
      sheet,
      sourceRow,
      value,
      'Unusually long barcode',
      'This identifier is longer than common barcode formats and should be reviewed.',
      'review',
    );
  }
  if (value && !/^\d+$/.test(value)) {
    addBarcodeWarning(
      d,
      sheet,
      sourceRow,
      value,
      'Malformed barcode',
      'Barcode contains characters outside the expected numeric format.',
      'review',
    );
  }
  if (/^\d+$/.test(value) && value.length > 15) {
    d.barcodeIntegrity.unsafePrecisionWarnings++;
    addBarcodeWarning(
      d,
      sheet,
      sourceRow,
      value,
      'Precision may have been lost',
      'Very long numeric identifiers can exceed safe spreadsheet or JavaScript integer precision.',
      'review',
    );
  }
  return value;
}
type BarcodeRowRef = { sourceSheet: string; sourceRow: number; barcode: string };
type ReleaseBarcodeDiagnostics = {
  normalisedBarcodes: Set<string>;
  rawBarcodes: Set<string>;
  rows: BarcodeRowRef[];
};
type BarcodeImportState = {
  unique: Set<string>;
  lengthRows: Map<number, BarcodeRowRef[]>;
  releaseFormats: Map<string, ReleaseBarcodeDiagnostics>;
  releaseSignaturesByBarcode: Map<string, Set<string>>;
  thirteenDigitRows: number;
  suspiciousZeroSuffixRows: number;
};
function releaseSignature(r: Transaction) {
  return [r.artist, r.albumTitle, r.catalogNumber, r.releaseCode].join('|').toLowerCase();
}
export function collectBarcodeDiagnostics(
  r: Transaction,
  state: BarcodeImportState,
  d: ImportDiagnostics,
) {
  if (r.barcode) {
    d.barcodeIntegrity.populatedBarcodeRows++;
    state.unique.add(r.barcode);
    if (/^\d{13}$/.test(r.barcode)) {
      state.thirteenDigitRows++;
      if (/0{6,}$/.test(r.barcode)) state.suspiciousZeroSuffixRows++;
    }
    if (/^\d+$/.test(r.barcode)) {
      const refs = state.lengthRows.get(r.barcode.length) ?? [];
      refs.push({ sourceSheet: r.sourceSheet, sourceRow: r.sourceRow, barcode: r.barcode });
      state.lengthRows.set(r.barcode.length, refs);
    }
    const sig = releaseSignature(r);
    const barcodeSignatures = state.releaseSignaturesByBarcode.get(r.barcode) ?? new Set<string>();
    barcodeSignatures.add(sig);
    state.releaseSignaturesByBarcode.set(r.barcode, barcodeSignatures);
    if (sig.trim()) {
      let group = state.releaseFormats.get(sig);
      if (!group) {
        group = { normalisedBarcodes: new Set(), rawBarcodes: new Set(), rows: [] };
        state.releaseFormats.set(sig, group);
      }
      group.normalisedBarcodes.add(r.barcode.replace(/^0+/, '').replace(/\.0+$/, ''));
      group.rawBarcodes.add(r.barcode);
      group.rows.push({ sourceSheet: r.sourceSheet, sourceRow: r.sourceRow, barcode: r.barcode });
    }
  } else d.barcodeIntegrity.blankBarcodeRows++;
}
function finaliseBarcodeDiagnostics(state: BarcodeImportState, d: ImportDiagnostics) {
  const b = d.barcodeIntegrity;
  b.uniqueBarcodeCount = state.unique.size;
  const zeroSuffixRatio = state.thirteenDigitRows
    ? state.suspiciousZeroSuffixRows / state.thirteenDigitRows
    : 0;
  const collapsedValues = [...state.releaseSignaturesByBarcode.values()].filter(
    (signatures) => signatures.size >= 5,
  ).length;
  if (
    state.thirteenDigitRows >= 20 &&
    zeroSuffixRatio >= 0.5 &&
    state.unique.size <= Math.max(5, Math.ceil(state.thirteenDigitRows / 20)) &&
    collapsedValues > 0
  ) {
    throw new Error(
      'Import stopped: likely barcode precision loss was detected. Many 13-digit barcodes were reduced to a few values ending in long strings of zeroes. Re-export the source with Barcode stored as text and try again.',
    );
  }
  if (state.lengthRows.size > 1) {
    for (const refs of state.lengthRows.values())
      for (const r of refs)
        addBarcodeWarning(
          d,
          r.sourceSheet,
          r.sourceRow,
          r.barcode,
          'Inconsistent barcode length',
          'Barcode lengths vary across the uploaded file; review for formatting changes.',
          'review',
        );
  }
  for (const group of state.releaseFormats.values()) {
    if (group.normalisedBarcodes.size === 1 && group.rawBarcodes.size > 1) {
      b.duplicateBarcodeConflicts += group.rawBarcodes.size;
      for (const r of group.rows)
        addBarcodeWarning(
          d,
          r.sourceSheet,
          r.sourceRow,
          r.barcode,
          'Duplicate barcode conflict',
          'Duplicate releases differ only by barcode formatting.',
          'review',
        );
    }
  }
  b.rowsRequiringReview = new Set(
    b.warnings.filter((w) => w.severity === 'review').map((w) => `${w.sourceSheet}:${w.sourceRow}`),
  ).size;
  const hasReview = b.rowsRequiringReview > 0;
  const hasEvidence =
    b.scientificNotationValuesConverted > 0 ||
    b.decimalSuffixesRemoved > 0 ||
    b.unsafePrecisionWarnings > 0 ||
    b.possibleLostLeadingZeroWarnings > 0 ||
    b.duplicateBarcodeConflicts > 0;
  d.statementHealth = {
    fileStatus: hasEvidence
      ? 'This statement may have been edited in spreadsheet software'
      : 'Original statement appears intact',
    dataQuality: hasReview ? 'Review recommended' : b.warnings.length ? 'Good' : 'Excellent',
    barcodeWarnings: b.warnings.length,
    rowsRequiringReview: b.rowsRequiringReview,
  };
  state.unique.clear();
  state.lengthRows.clear();
  state.releaseFormats.clear();
  state.releaseSignaturesByBarcode.clear();
}
export function parseRoyaltyRate(value: unknown) {
  if (value === null || value === undefined) return null;
  const raw = typeof value === 'string' ? value.trim() : value;
  if (raw === '') return null;
  const stringValue = String(raw).trim();
  const isPercentage = stringValue.includes('%');
  const cleaned = stringValue.replace(/,/g, '').replace(/%/g, '').trim();
  if (!cleaned) return null;
  try {
    const rate = new Decimal(cleaned);
    if (!rate.isFinite()) return null;
    return isPercentage || rate.abs().gt(1) ? rate.div(100) : rate;
  } catch {
    return null;
  }
}
function isBlankRow(row: unknown[]) {
  return row.every((c) => text(c) === '');
}
export function detectHeaderRow(rows: unknown[][]) {
  let best = { idx: -1, count: 0, map: new Map<FieldKey, number>() };
  for (let i = 0; i < Math.min(30, rows.length); i++) {
    const map = new Map<FieldKey, number>();
    (rows[i] ?? []).forEach((h, j) => {
      const k = resolveHeader(text(h));
      if (k && !map.has(k)) map.set(k, j);
    });
    const req = requiredFields.filter((f) => map.has(f)).length;
    const score = map.size + req * 3;
    if (score > best.count) best = { idx: i, count: score, map };
  }
  return requiredFields.every((f) => best.map.has(f)) ? best : null;
}
function tx(
  row: unknown[],
  map: Map<FieldKey, number>,
  sheet: string,
  sourceRow: number,
  diag: ImportDiagnostics,
  keepOriginal: boolean,
  ws?: XLSX.WorkSheet,
  rowIndex?: number,
): Transaction {
  const rowWarnings: string[] = [];
  const get = (k: FieldKey) => {
    const col = map.get(k) ?? -1;
    const v =
      k === 'barcode' && ws && rowIndex !== undefined
        ? barcodeCellText(ws, rowIndex, col, sheet, sourceRow, diag)
        : textFields.has(k) && ws && rowIndex !== undefined
          ? identifierCellText(ws, rowIndex, col)
          : row[col];
    if (k === 'royaltyRate') {
      const parsed = parseRoyaltyRate(v);
      if (parsed === null && text(v) !== '') {
        const msg = `${sheet} row ${sourceRow}: Royalty Rate could not be parsed (${text(v)})`;
        diag.invalidNumericValues.push(msg);
        rowWarnings.push(msg);
      }
      return parsed?.toString() ?? '';
    }
    return textFields.has(k) ? normaliseIdentifier(v) : text(v);
  };
  const out: Transaction = {
    sourceSheet: sheet,
    sourceRow,
    contractId: get('contractId'),
    contractName: get('contractName'),
    shareContract: get('shareContract'),
    assetType: get('assetType'),
    releaseCode: get('releaseCode'),
    albumTitle: get('albumTitle'),
    catalogNumber: get('catalogNumber'),
    barcode: get('barcode'),
    isrc: get('isrc'),
    artist: get('artist'),
    trackTitle: get('trackTitle'),
    usageType: get('usageType'),
    country: get('country'),
    shop: get('shop'),
    salesPeriod: get('salesPeriod'),
    sales: get('sales'),
    returns: get('returns'),
    ppu: get('ppu'),
    amount: get('amount'),
    share: get('share'),
    rata1: get('rata1'),
    rata2: get('rata2'),
    deduction1: get('deduction1'),
    deduction2: get('deduction2'),
    deduction3: get('deduction3'),
    contractDeductions: get('contractDeductions'),
    deduction4: get('deduction4'),
    deduction5: get('deduction5'),
    lineCharges: get('lineCharges'),
    royaltyRate: get('royaltyRate'),
    royaltyAmount: get('royaltyAmount'),
  };
  if (keepOriginal) out.originalRow = Object.fromEntries(row.map((v, i) => [String(i), v]));
  if (rowWarnings.length) out.rowWarnings = rowWarnings;
  return out;
}
function rangeRows(ws: XLSX.WorkSheet, start: number, end: number) {
  const ref = ws['!ref'];
  if (!ref) return [];
  const range = XLSX.utils.decode_range(ref);
  range.s.r = start;
  range.e.r = Math.max(start, end - 1);
  return XLSX.utils.sheet_to_json(ws, {
    header: 1,
    defval: '',
    raw: false,
    range,
    blankrows: false,
  }) as unknown[][];
}
function sheetLastRow(ws: XLSX.WorkSheet) {
  const ref = ws['!ref'];
  if (!ref) return 0;
  return XLSX.utils.decode_range(ref).e.r + 1;
}
function summary(rows: Transaction[]): ImportSummary {
  let totalAmount = D(0),
    totalRoyaltyAmount = D(0),
    totalSales = D(0),
    totalReturns = D(0),
    totalDeductions = D(0),
    totalLineCharges = D(0);
  const artists = new Set<string>(),
    releases = new Set<string>(),
    tracks = new Set<string>(),
    rates = new Set<string>();
  for (const r of rows) {
    totalAmount = totalAmount.plus(D(r.amount));
    totalRoyaltyAmount = totalRoyaltyAmount.plus(D(r.royaltyAmount));
    totalSales = totalSales.plus(D(r.sales));
    totalReturns = totalReturns.plus(D(r.returns));
    totalDeductions = totalDeductions.plus(
      [
        'deduction1',
        'deduction2',
        'deduction3',
        'contractDeductions',
        'deduction4',
        'deduction5',
        'lineCharges',
      ].reduce((s, k) => s.plus(D((r as any)[k])), D(0)),
    );
    totalLineCharges = totalLineCharges.plus(D(r.lineCharges));
    if (r.artist) artists.add(r.artist);
    releases.add(releaseKey(r));
    if (isTrackRow(r)) tracks.add(r.isrc || `${r.artist}|${r.albumTitle}|${r.trackTitle}`);
    if (r.royaltyRate) rates.add(r.royaltyRate);
  }
  return {
    totalAmount: totalAmount.toString(),
    totalRoyaltyAmount: totalRoyaltyAmount.toString(),
    totalSales: totalSales.toString(),
    totalReturns: totalReturns.toString(),
    totalDeductions: totalDeductions.toString(),
    totalLineCharges: totalLineCharges.toString(),
    transactionCount: rows.length,
    uniqueArtistCount: artists.size,
    uniqueReleaseCount: releases.size,
    uniqueTrackCount: tracks.size,
    royaltyRateSummary:
      rates.size === 0
        ? 'Rate not supplied'
        : rates.size === 1
          ? `${D([...rates][0])
              .mul(100)
              .toString()}%`
          : 'Multiple rates',
  };
}
export function parseWorkbook(
  wb: XLSX.WorkBook,
  filename = 'statement',
  fileSize = 0,
  opts: ParseOptions = {},
): StatementData {
  let pct = 25;
  const started = safeNow();
  const emit = (stage: any, f = 0, extra = {}) => {
    pct = clampProgress(pct, stage, f);
    opts.onProgress?.({
      stage,
      percent: pct,
      rowsExamined: diag.transactionRows + diag.blankRowsIgnored,
      rowsImported: rows.length,
      rowsSkipped: diag.blankRowsIgnored,
      ...extra,
    });
  };
  const rows: Transaction[] = [];
  const barcodeState = {
    unique: new Set<string>(),
    lengthRows: new Map<number, BarcodeRowRef[]>(),
    releaseFormats: new Map<string, ReleaseBarcodeDiagnostics>(),
    releaseSignaturesByBarcode: new Map<string, Set<string>>(),
    thirteenDigitRows: 0,
    suspiciousZeroSuffixRows: 0,
  };
  const diag: ImportDiagnostics = {
    filename,
    fileSize,
    ...parseFilename(filename),
    worksheetsDetected: wb.SheetNames,
    worksheetsImported: [],
    worksheetsIgnored: [],
    headerRows: {},
    transactionRows: 0,
    blankRowsIgnored: 0,
    detectedColumns: [],
    missingOptionalColumns: [],
    missingRequiredColumns: [],
    invalidNumericValues: [],
    unclassifiedDateValues: [],
    blankArtistValues: 0,
    blankReleaseValues: 0,
    blankIsrcValues: 0,
    bundleRowsWithBlankIsrc: 0,
    duplicateLookingRows: 0,
    royaltyValidation: { matches: 0, requiresReview: 0, notChecked: 0 },
    barcodeIntegrity: emptyBarcodeReport(),
    statementHealth: emptyHealth(),
    largeFileMode: opts.largeFileMode,
    estimatedRows: opts.estimatedRows,
    memory: {
      workbookRetained: 'Unknown',
      worksheetRetained: 'Unknown',
      rawRowsRetained: 'Unknown',
      workerTerminated: 'Unknown',
      multipleFullDatasetsRetained: 'Unknown',
      automaticSummaries: [],
    },
    importTimings: {
      fileReadingMs: opts.fileReadMs,
      totalImportMs: 0,
      tabFirstCalculationMs: {},
      searchIndexMs: {},
    },
  };
  const names = [...wb.SheetNames].sort((a, b) =>
    a === 'Digital Sales' ? -1 : b === 'Digital Sales' ? 1 : 0,
  );
  const tHead = safeNow();
  const imports: { name: string; h: ReturnType<typeof detectHeaderRow>; last: number }[] = [];
  names.forEach((name, i) => {
    const ws = wb.Sheets[name];
    const last = sheetLastRow(ws);
    const preview = rangeRows(ws, 0, Math.min(30, last));
    const h = preview.length && !preview.every(isBlankRow) ? detectHeaderRow(preview) : null;
    if (!h)
      diag.worksheetsIgnored.push(`${name} (${last ? 'no recognised statement table' : 'blank'})`);
    else imports.push({ name, h, last });
    emit('Detecting worksheets and headers', (i + 1) / names.length, { currentWorksheet: name });
  });
  recordTimingValue(diag, 'headerDetectionMs', safeNow() - tHead);
  if (!imports.length) {
    const anyHeader = names.some((name) =>
      rangeRows(wb.Sheets[name], 0, Math.min(30, sheetLastRow(wb.Sheets[name]))).some((r) =>
        r.some((c) => resolveHeader(text(c))),
      ),
    );
    if (anyHeader)
      throw new Error(
        `The statement is missing the following required columns: ${requiredFields.map((f) => fieldLabels[f]).join(', ')}`,
      );
    throw new Error('The statement structure could not be recognised.');
  }
  const estimated = imports.reduce((a, s) => a + Math.max(0, s.last - (s.h?.idx ?? 0) - 1), 0);
  diag.estimatedRows = estimated;
  diag.largeFileMode = isLargeFile(fileSize, estimated) || !!opts.largeFileMode;
  const keepOriginal = false;
  diag.memory!.rawRowsRetained = 'No';
  const chunkSize = diag.largeFileMode ? LARGE_FILE_CHUNK_SIZE : DEFAULT_CHUNK_SIZE;
  const tRows = safeNow();
  let processed = 0;
  for (const it of imports) {
    diag.worksheetsImported.push(it.name);
    diag.headerRows[it.name] = (it.h?.idx ?? 0) + 1;
    diag.detectedColumns = [
      ...new Set([...diag.detectedColumns, ...[...it.h!.map.keys()].map((k) => fieldLabels[k])]),
    ];
    for (let start = it.h!.idx + 1; start < it.last; start += chunkSize) {
      const chunk = rangeRows(wb.Sheets[it.name], start, Math.min(it.last, start + chunkSize));
      for (let j = 0; j < chunk.length; j++) {
        const sourceRow = start + j + 1;
        try {
          const row = chunk[j];
          processed++;
          if (isBlankRow(row)) {
            diag.blankRowsIgnored++;
            continue;
          }
          const transaction = tx(
            row,
            it.h!.map,
            it.name,
            sourceRow,
            diag,
            keepOriginal,
            wb.Sheets[it.name],
            start + j,
          );
          rows.push(transaction);
          collectBarcodeDiagnostics(transaction, barcodeState, diag);
        } catch (e) {
          diag.invalidNumericValues.push(
            `${it.name} row ${sourceRow}: ${e instanceof Error ? e.message : String(e)}`,
          );
        }
      }
      diag.transactionRows = rows.length;
      emit('Importing transactions', estimated ? processed / estimated : 1, {
        currentWorksheet: it.name,
      });
    }
  }
  recordTimingValue(diag, 'rowNormalisationMs', safeNow() - tRows);
  diag.missingRequiredColumns = requiredFields
    .filter((f) => !diag.detectedColumns.includes(fieldLabels[f]))
    .map((f) => fieldLabels[f]);
  diag.missingOptionalColumns = (Object.keys(columnAliases) as FieldKey[])
    .filter((f) => !requiredFields.includes(f) && !diag.detectedColumns.includes(fieldLabels[f]))
    .map((f) => fieldLabels[f]);
  if (diag.missingRequiredColumns.length)
    throw new Error(
      `The statement is missing the following required columns: ${diag.missingRequiredColumns.join(', ')}`,
    );
  const tSum = safeNow();
  enrichDiagnostics(rows, diag, { skipDuplicates: diag.largeFileMode });
  finaliseBarcodeDiagnostics(barcodeState, diag);
  diag.initialSummary = summary(rows);
  diag.memory!.automaticSummaries = ['initialSummary', 'statementHealth'];
  recordTimingValue(diag, 'initialSummaryMs', safeNow() - tSum);
  emit('Calculating initial summaries', 1);
  recordTimingValue(diag, 'dashboardPreparationMs', 1);
  diag.memory!.workbookRetained = 'No';
  diag.memory!.worksheetRetained = 'No';
  diag.memory!.multipleFullDatasetsRetained = 'No';
  emit('Preparing dashboard', 1, { message: 'workbook retained: No' });
  recordTimingValue(diag, 'totalImportMs', safeNow() - started + (opts.fileReadMs ?? 0));
  emit('Complete', 1);
  return { label: 'Uploaded statement', filename, fileSize, rows, diagnostics: diag };
}
export const DEFAULT_SHEETJS_READ_OPTIONS: XLSX.ParsingOptions = {
  type: 'array',
  cellDates: false,
  cellStyles: false,
  cellHTML: false,
  bookVBA: false,
  bookDeps: false,
  bookFiles: false,
  sheetStubs: false,
};
export const CONSERVATIVE_SHEETJS_READ_OPTIONS: XLSX.ParsingOptions = {
  type: 'array',
  cellDates: false,
  cellStyles: false,
  cellHTML: false,
  bookVBA: false,
  bookDeps: false,
  bookFiles: false,
  sheetStubs: true,
  dense: false,
};
export const sheetJsModuleLoaded = !!XLSX?.read;
export async function parseArrayBuffer(
  buf: ArrayBuffer,
  filename: string,
  fileSize: number,
  opts: ParseOptions = {},
) {
  const t = safeNow();
  if (!buf || buf.byteLength === 0)
    throw new Error('Workbook ArrayBuffer is empty before SheetJS read.');
  opts.onBeforeSheetJsRead?.();
  opts.onProgress?.({
    stage: 'Opening workbook',
    percent: 15,
    rowsExamined: 0,
    rowsImported: 0,
    rowsSkipped: 0,
    message: opts.sheetJsReadAttempt ?? 'Opening workbook',
  });
  let wb: XLSX.WorkBook | null = XLSX.read(buf, {
    ...DEFAULT_SHEETJS_READ_OPTIONS,
    ...opts.readOptions,
    type: 'array',
  });
  const opening = safeNow() - t;
  try {
    const data = parseWorkbook(wb, filename, fileSize, opts);
    recordTimingValue(data.diagnostics, 'workbookOpeningMs', opening);
    data.diagnostics.memory = {
      ...(data.diagnostics.memory ?? {}),
      workbookRetained: 'No',
      worksheetRetained: 'No',
    };
    return data;
  } finally {
    if (wb) {
      for (const name of wb.SheetNames ?? []) delete wb.Sheets[name];
      wb.SheetNames.length = 0;
    }
    wb = null;
  }
}
export async function parseFile(file: File, opts: ParseOptions = {}) {
  if (/\.csv$/i.test(file.name)) {
    const { parseCsvFileStreaming } = await import('./streamingCsvParser');
    return parseCsvFileStreaming(file, opts);
  }
  const t = safeNow();
  const buf = await file.arrayBuffer();
  opts.onProgress?.({
    stage: 'Reading file',
    percent: 15,
    rowsExamined: 0,
    rowsImported: 0,
    rowsSkipped: 0,
  });
  return parseArrayBuffer(buf, file.name, file.size, { ...opts, fileReadMs: safeNow() - t });
}
export function enrichDiagnostics(
  rows: Transaction[],
  d: ImportDiagnostics,
  opts: { skipDuplicates?: boolean } = {},
) {
  const seen = new Map<string, number>();
  for (const r of rows) {
    try {
      if (!r.artist) d.blankArtistValues++;
      if (!r.albumTitle && !r.catalogNumber && !r.barcode && !r.releaseCode) d.blankReleaseValues++;
      if (!r.isrc) d.blankIsrcValues++;
      if (!r.isrc && /bundle/i.test(`${r.assetType} ${r.usageType}`)) d.bundleRowsWithBlankIsrc++;
      if (!opts.skipDuplicates) {
        const key = [
          r.artist,
          r.albumTitle,
          r.trackTitle,
          r.isrc,
          r.shop,
          r.country,
          r.salesPeriod,
          r.usageType,
          r.sales,
          r.amount,
          r.royaltyAmount,
        ].join('|');
        seen.set(key, (seen.get(key) ?? 0) + 1);
      }
      const adjusted = [
        'share',
        'rata1',
        'rata2',
        'deduction1',
        'deduction2',
        'deduction3',
        'contractDeductions',
        'deduction4',
        'deduction5',
        'lineCharges',
      ].some((k) => !D((r as any)[k]).isZero());
      if (adjusted || !r.royaltyRate) {
        d.royaltyValidation.notChecked++;
      } else if (D(r.amount).mul(D(r.royaltyRate)).minus(D(r.royaltyAmount)).abs().lte(0.0001))
        d.royaltyValidation.matches++;
      else d.royaltyValidation.requiresReview++;
    } catch (e) {
      d.invalidNumericValues.push(
        `${r.sourceSheet} row ${r.sourceRow}: ${e instanceof Error ? e.message : String(e)}`,
      );
      d.royaltyValidation.notChecked++;
    }
  }
  d.duplicateLookingRows = opts.skipDuplicates
    ? 0
    : [...seen.values()].filter((v) => v > 1).reduce((a, v) => a + v, 0);
}
