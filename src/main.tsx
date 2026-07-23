import React, { useEffect, useMemo, useRef, useState } from 'react';
import Decimal from 'decimal.js';
import { createRoot } from 'react-dom/client';
import {
  useReactTable,
  getCoreRowModel,
  getSortedRowModel,
  getFilteredRowModel,
  getPaginationRowModel,
  flexRender,
  type ColumnDef,
} from '@tanstack/react-table';
import { sampleStatement } from './lib/sampleData';
import { fmtMoney, fmtInt, D } from './lib/format';
import {
  groupArtists,
  groupReleases,
  groupTracks,
  trackRows,
  groupBy,
  importCheckGroups,
  totals,
  rateSummary,
  periodSortValue,
  searchRows,
  deductionFields,
  groupedReconciliation,
  chartDatasetCacheKey,
  prepareOverviewChartDataFromRows,
} from './lib/analytics';
import { downloadCsv, toCsv, normalisedRows } from './lib/exportCsv';
import { downloadBreakdownsZip, type ExportProgress } from './lib/exportZip';
import { recordTiming, recordTimingValue, safeNow } from './lib/timing';
import type { StatementData, Transaction } from './lib/types';
import type { ImportProgress } from './lib/importProgress';
import { isLargeFile } from './lib/importProgress';
import { parseFile } from './lib/parser';
import { LARGE_XLSX_CONVERTER_URL, shouldOfferLargeXlsxConversion } from './lib/largeExcel';
import { technicalDetails, type ImportDebugContext } from './lib/importDiagnostics';
import './styles.css';
const tabs = [
  'Overview',
  'Artists',
  'Releases',
  'Tracks',
  'Shops',
  'Countries',
  'Sales Periods',
  'Deductions',
  'Full Detail',
  'Import Checks',
];
function App() {
  const [data, setData] = useState<StatementData>();
  const [err, setErr] = useState<ImportErrorState | null>(null);
  const [progress, setProgress] = useState<ImportProgress | null>(null);
  const [fileMeta, setFileMeta] = useState<{ name: string; size: number } | null>(null);
  const [elapsed, setElapsed] = useState(0);
  const [conversionPrompt, setConversionPrompt] = useState<{
    file: File;
    estimatedRows?: number;
  } | null>(null);
  const [stalled, setStalled] = useState(false);
  const worker = useRef<Worker | null>(null);
  const lastProgress = useRef(0);
  useEffect(() => {
    if (!progress) return;
    const id = setInterval(() => {
      setElapsed(Date.now() - lastProgress.current);
      setStalled(Date.now() - lastProgress.current > 20000);
    }, 500);
    return () => clearInterval(id);
  }, [progress]);
  async function load(f: File) {
    cancel(false);
    const largeXlsx = shouldOfferLargeXlsxConversion(f);
    if (largeXlsx.shouldConvert) {
      setErr(null);
      setProgress(null);
      setFileMeta({ name: f.name, size: f.size });
      setConversionPrompt({ file: f, estimatedRows: largeXlsx.estimatedRows });
      return;
    }
    setErr(null);
    setFileMeta({ name: f.name, size: f.size });
    lastProgress.current = Date.now();
    setElapsed(0);
    setStalled(false);
    setProgress({
      stage: 'Reading file',
      percent: 0,
      rowsExamined: 0,
      rowsImported: 0,
      rowsSkipped: 0,
      largeFileMode: isLargeFile(f.size),
    });
    try {
      const t = safeNow();
      const reader = new FileReader();
      reader.onprogress = (e) => {
        if (e.lengthComputable)
          setProgress((p) =>
            monotonic(p, {
              stage: 'Reading file',
              percent: Math.round((e.loaded / e.total) * 15),
              rowsExamined: 0,
              rowsImported: 0,
              rowsSkipped: 0,
              largeFileMode: isLargeFile(f.size),
            }),
          );
      };
      reader.onerror = () => {
        const ctx = debugContext('Reading file', 'file-read', f, 0, undefined, undefined, 0);
        setErr({
          stage: 'Reading file',
          message: 'This statement could not be fully processed.',
          details: technicalDetails(ctx, reader.error),
          rowsProcessed: 0,
          canCompatibility: false,
        });
      };
      reader.onload = () => {
        let w: Worker | null = null;
        try {
          const buffer = reader.result as ArrayBuffer;
          const before = buffer?.byteLength ?? 0;
          if (!buffer || before === 0) {
            const ctx = debugContext(
              'Opening workbook',
              'file-read',
              f,
              before,
              0,
              'not attempted',
              0,
            );
            setErr({
              stage: 'Opening workbook',
              message: 'This statement could not be fully processed.',
              details: technicalDetails(
                ctx,
                new Error('FileReader produced an empty ArrayBuffer.'),
              ),
              rowsProcessed: 0,
              canCompatibility: false,
            });
            setProgress(null);
            return;
          }
          const workerUrl = 'Vite-built worker asset';
          w = new Worker(new URL('./workers/importWorker.ts', import.meta.url), {
            type: 'module',
            name: 'statement-import-worker',
          });
          worker.current = w;
          const activeWorker = w;
          let latest = debugContext(
            'Opening workbook',
            'worker-created',
            f,
            before,
            undefined,
            'not started',
            0,
            workerUrl,
          );
          let workerReady = false;
          let firstWorkerMessageReceived = false;
          const startupTimeout = window.setTimeout(() => {
            if (workerReady) return;
            const ctx = {
              ...latest,
              processingStage: 'Opening workbook',
              workerEvent: 'worker-startup-timeout',
              workerStarted: 'No' as const,
              workerStartupCompleted: 'No' as const,
              firstWorkerMessageReceived: firstWorkerMessageReceived
                ? ('Yes' as const)
                : ('No' as const),
              sheetJsReadAttempt: 'not started',
            };
            setErr({
              stage: ctx.processingStage,
              message: 'The import worker did not start, so the workbook was not transferred.',
              details: technicalDetails(ctx, new Error('Timed out waiting for worker-ready.')),
              rowsProcessed: 0,
              canCompatibility: true,
              file: f,
            });
            setProgress(null);
            activeWorker.terminate();
            worker.current = null;
          }, 5000);
          w.onmessage = (e) => {
            firstWorkerMessageReceived = true;
            latest = {
              ...latest,
              ...e.data.context,
              workerEvent: e.data.context?.workerEvent ?? 'worker-message',
              workerStartupCompleted:
                e.data.type === 'worker-ready' ? 'Yes' : latest.workerStartupCompleted,
              firstWorkerMessageReceived: 'Yes',
            };
            if (e.data.type === 'worker-ready') {
              workerReady = true;
              window.clearTimeout(startupTimeout);
              activeWorker.postMessage(
                {
                  buffer,
                  filename: f.name,
                  fileSize: f.size,
                  fileReadMs: safeNow() - t,
                  arrayBufferSizeBeforeTransfer: before,
                },
                [buffer],
              );
              latest = {
                ...latest,
                workerEvent: 'post-message',
                arrayBufferTransferSucceeded: buffer.byteLength === 0 ? 'Yes' : 'No',
              };
              return;
            }
            if (e.data.type === 'worker-received' || e.data.type === 'sheetjs-read-start') {
              return;
            }
            if (e.data.type === 'progress') {
              lastProgress.current = Date.now();
              setStalled(false);
              setProgress((p) =>
                monotonic(p, {
                  ...e.data.progress,
                  largeFileMode: e.data.progress.largeFileMode ?? isLargeFile(f.size),
                }),
              );
            } else if (e.data.type === 'complete') {
              window.clearTimeout(startupTimeout);
              lastProgress.current = Date.now();
              setProgress((p) =>
                monotonic(p, {
                  stage: 'Complete',
                  percent: 100,
                  rowsExamined: e.data.data.rows.length,
                  rowsImported: e.data.data.rows.length,
                  rowsSkipped: e.data.data.diagnostics.blankRowsIgnored,
                  largeFileMode: e.data.data.diagnostics.largeFileMode,
                  message: 'Statement ready',
                }),
              );
              setTimeout(() => {
                const imported = e.data.data;
                imported.diagnostics.memory = {
                  ...(imported.diagnostics.memory ?? {}),
                  workerTerminated: 'Yes',
                };
                activeWorker.terminate();
                worker.current = null;
                setData(imported);
                setProgress(null);
                setFileMeta(null);
              }, 500);
            } else if (e.data.type === 'error') {
              window.clearTimeout(startupTimeout);
              const details =
                e.data.error?.details ||
                technicalDetails(
                  {
                    ...latest,
                    processingStage: progress?.stage ?? 'Opening workbook',
                    workerEvent: 'worker-message',
                  },
                  e.data.error,
                );
              setErr({
                stage: progress?.stage ?? 'Opening workbook',
                message: friendly(e.data.error?.message),
                details,
                rowsProcessed: progress?.rowsExamined ?? 0,
                canCompatibility: true,
                file: f,
              });
              setProgress(null);
              activeWorker.terminate();
              worker.current = null;
            }
          };
          w.onerror = (e) => {
            window.clearTimeout(startupTimeout);
            const err = new Error(e.message);
            err.name = e.type || 'ErrorEvent';
            err.stack = [e.message, e.filename ? `at ${e.filename}:${e.lineno}:${e.colno}` : '']
              .filter(Boolean)
              .join('\n');
            const ctx = {
              ...latest,
              processingStage: progress?.stage ?? 'Opening workbook',
              workerEvent: 'worker-error',
              errorName: err.name,
              errorMessage: err.message,
              stack: err.stack,
              filename: e.filename,
              lineNumber: e.lineno,
              columnNumber: e.colno,
              workerUrl: e.filename || workerUrl,
              workerStartupCompleted: workerReady ? ('Yes' as const) : ('No' as const),
              firstWorkerMessageReceived: firstWorkerMessageReceived
                ? ('Yes' as const)
                : ('No' as const),
            };
            setErr({
              stage: ctx.processingStage,
              message: 'This statement could not be fully processed.',
              details: technicalDetails(ctx, err),
              rowsProcessed: progress?.rowsExamined ?? 0,
              canCompatibility: true,
              file: f,
            });
            setProgress(null);
            activeWorker.terminate();
            worker.current = null;
          };
          w.onmessageerror = () => {
            window.clearTimeout(startupTimeout);
            const ctx = {
              ...latest,
              processingStage: progress?.stage ?? 'Opening workbook',
              workerEvent: 'worker-messageerror',
              workerStartupCompleted: workerReady ? ('Yes' as const) : ('No' as const),
              firstWorkerMessageReceived: firstWorkerMessageReceived
                ? ('Yes' as const)
                : ('No' as const),
            };
            setErr({
              stage: ctx.processingStage,
              message: 'This statement could not be fully processed.',
              details: technicalDetails(
                ctx,
                new Error('Worker message could not be deserialised.'),
              ),
              rowsProcessed: progress?.rowsExamined ?? 0,
              canCompatibility: true,
              file: f,
            });
            setProgress(null);
            activeWorker.terminate();
            worker.current = null;
          };
        } catch (e) {
          const ctx = debugContext(
            'Opening workbook',
            'post-message',
            f,
            undefined,
            undefined,
            'not started',
            0,
          );
          setErr({
            stage: ctx.processingStage,
            message: 'This statement could not be fully processed.',
            details: technicalDetails(ctx, e),
            rowsProcessed: 0,
            canCompatibility: true,
            file: f,
          });
          setProgress(null);
          w?.terminate();
          worker.current = null;
        }
      };
      reader.readAsArrayBuffer(f);
    } catch (e) {
      const ctx = debugContext(
        progress?.stage ?? 'Reading file',
        'error',
        f,
        undefined,
        undefined,
        undefined,
        0,
      );
      setErr({
        stage: ctx.processingStage,
        message: 'This statement could not be fully processed.',
        details: technicalDetails(ctx, e),
        rowsProcessed: progress?.rowsExamined ?? 0,
        canCompatibility: false,
      });
      setProgress(null);
    }
  }
  async function compatibilityLoad(f: File) {
    setErr(null);
    setProgress({
      stage: 'Opening workbook',
      percent: 15,
      rowsExamined: 0,
      rowsImported: 0,
      rowsSkipped: 0,
      largeFileMode: isLargeFile(f.size),
      message: 'Compatibility mode',
    });
    try {
      const data = await parseFile(f, {
        onProgress: (p) =>
          setProgress((prev) =>
            monotonic(prev, { ...p, largeFileMode: p.largeFileMode ?? isLargeFile(f.size) }),
          ),
      });
      setData(data);
      setProgress(null);
      setFileMeta(null);
    } catch (e) {
      const ctx = debugContext(
        progress?.stage ?? 'Opening workbook',
        'compatibility-mode',
        f,
        undefined,
        undefined,
        'main-thread parser',
        0,
      );
      setErr({
        stage: ctx.processingStage,
        message: 'This statement could not be fully processed.',
        details: technicalDetails(ctx, e),
        rowsProcessed: progress?.rowsExamined ?? 0,
        canCompatibility: false,
      });
      setProgress(null);
    }
  }

  async function convertLargeWorkbook(f: File) {
    setConversionPrompt(null);
    setErr(null);
    setFileMeta({ name: f.name, size: f.size });
    lastProgress.current = Date.now();
    setProgress({
      stage: 'Reading file',
      percent: 5,
      rowsExamined: 0,
      rowsImported: 0,
      rowsSkipped: 0,
      largeFileMode: true,
      message: 'Uploading workbook',
    });
    try {
      const controller = new AbortController();
      const timeout = window.setTimeout(
        () => controller.abort(),
        Number(import.meta.env.VITE_LARGE_XLSX_CONVERTER_TIMEOUT_MS || 120000),
      );
      const response = await fetch(LARGE_XLSX_CONVERTER_URL, {
        method: 'POST',
        body: f,
        signal: controller.signal,
        headers: {
          'content-type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        },
      });
      window.clearTimeout(timeout);
      if (!response.ok) {
        let message = 'The workbook could not be converted.';
        try {
          message = (await response.json()).error || message;
        } catch {
          message = message || 'The workbook could not be converted.';
        }
        throw new Error(message);
      }
      setProgress((p) =>
        monotonic(p, {
          stage: 'Opening workbook',
          percent: 30,
          rowsExamined: 0,
          rowsImported: 0,
          rowsSkipped: 0,
          largeFileMode: true,
          message: 'Downloading CSV',
        }),
      );
      const blob = await response.blob();
      setProgress((p) =>
        monotonic(p, {
          stage: 'Detecting worksheets and headers',
          percent: 35,
          rowsExamined: 0,
          rowsImported: 0,
          rowsSkipped: 0,
          largeFileMode: true,
          message: 'Preparing analysis',
        }),
      );
      const csvFile = new File([blob], f.name.replace(/\.xlsx$/i, '.csv'), { type: 'text/csv' });
      const data = await parseFile(csvFile, {
        largeFileMode: true,
        onProgress: (p) => setProgress((prev) => monotonic(prev, { ...p, largeFileMode: true })),
      });
      data.filename = f.name;
      data.diagnostics.filename = f.name;
      setData(data);
      setProgress(null);
      setFileMeta(null);
    } catch (e) {
      const aborted = e instanceof DOMException && e.name === 'AbortError';
      setErr({
        stage: aborted ? 'Converting worksheet' : 'Opening workbook',
        message: aborted
          ? 'The conversion timed out. Please try again or choose another file.'
          : friendly(e instanceof Error ? e.message : String(e)),
        details: e instanceof Error ? e.stack || e.message : String(e),
        rowsProcessed: 0,
        canCompatibility: false,
        file: f,
      });
      setProgress(null);
    }
  }

  function cancel(clear = true) {
    worker.current?.terminate();
    worker.current = null;
    if (clear) {
      setProgress(null);
      setFileMeta(null);
      setErr(null);
      setStalled(false);
      setConversionPrompt(null);
    }
  }
  return (
    <>
      {!data ? (
        <Upload
          onFile={load}
          onCompatibility={compatibilityLoad}
          onSample={() => setData(sampleStatement())}
          err={err}
          progress={progress}
          file={fileMeta}
          stalled={stalled}
          elapsed={elapsed}
          cancel={() => cancel(true)}
          conversionPrompt={conversionPrompt}
          onConvert={convertLargeWorkbook}
        />
      ) : (
        <Dashboard data={data} remove={() => setData(undefined)} />
      )}
    </>
  );
}
type ImportErrorState = {
  stage: string;
  message: string;
  details: string;
  rowsProcessed: number;
  canCompatibility?: boolean;
  file?: File;
};
function debugContext(
  processingStage: string,
  workerEvent: string,
  f: File,
  arrayBufferSizeBeforeTransfer?: number,
  arrayBufferSizeInsideWorker?: number,
  sheetJsReadAttempt?: string,
  retryAttemptNumber?: number,
  workerUrl?: string,
): ImportDebugContext {
  return {
    processingStage,
    workerEvent,
    workbookFilename: f.name,
    fileSize: f.size,
    arrayBufferSizeBeforeTransfer,
    arrayBufferSizeInsideWorker,
    sheetJsReadAttempt,
    retryAttemptNumber,
    workerUrl,
    workerStartupCompleted: 'Unknown',
    firstWorkerMessageReceived: 'Unknown',
    workerStarted: 'Unknown',
    arrayBufferTransferSucceeded: 'Unknown',
    workerPayloadReceived: 'Unknown',
    sheetJsImported: 'Unknown',
    xlsxReadInvoked: 'No',
  };
}
function monotonic(prev: ImportProgress | null, next: ImportProgress) {
  return { ...next, percent: Math.max(prev?.percent ?? 0, next.percent) };
}
function friendly(message: string) {
  if (/memory|allocation/i.test(message))
    return 'The browser ran out of memory while opening this statement. No data was uploaded or stored.';
  if (/recognised|missing/i.test(message))
    return 'The statement structure could not be recognised.';
  return message || 'This statement could not be fully processed.';
}
function Upload(p: {
  onFile: (f: File) => void;
  onCompatibility: (f: File) => void;
  onSample: () => void;
  err: ImportErrorState | null;
  progress: ImportProgress | null;
  file: { name: string; size: number } | null;
  stalled: boolean;
  elapsed: number;
  cancel: () => void;
  conversionPrompt: { file: File; estimatedRows?: number } | null;
  onConvert: (f: File) => void;
}) {
  if (p.conversionPrompt && p.file)
    return (
      <main className="upload">
        <section className="processing">
          <h1>This Excel statement is too large to analyse safely in your browser.</h1>
          <p>Would you like us to convert it to CSV automatically?</p>
          <dl>
            <dt>Filename</dt>
            <dd>{p.file.name}</dd>
            <dt>File size</dt>
            <dd>{formatBytes(p.file.size)}</dd>
            <dt>Estimated rows</dt>
            <dd>{p.conversionPrompt.estimatedRows ?? 'Not available from preflight'}</dd>
          </dl>
          <aside className="privacy compact-privacy">
            <p>Large Excel files can optionally be converted to CSV on our secure server.</p>
            <p>
              The original Excel file is used only to create a CSV of the Digital Sales worksheet
              and is deleted immediately afterwards.
            </p>
            <p>No statement analysis is performed on the server.</p>
            <p>All summaries and calculations continue to happen locally in your browser.</p>
          </aside>
          <button onClick={() => p.onConvert(p.conversionPrompt!.file)}>
            Convert and continue
          </button>
          <button onClick={p.cancel}>Cancel</button>
        </section>
      </main>
    );
  if (p.progress && p.file)
    return (
      <main className="upload">
        <section className="processing">
          <h1>Processing your statement…</h1>
          <p>Large statements may take a little longer. Please keep this page open.</p>
          {p.progress.largeFileMode && (
            <aside className="mode">
              <b>Large File Mode</b>
              <br />
              This statement will be processed in smaller sections to reduce browser memory use.
              <br />
              For very large statements, CSV may import faster than Excel.
            </aside>
          )}
          <p aria-live="polite">
            <b>{p.progress.message ?? p.progress.stage}</b>
          </p>
          <div
            className="bar"
            role="progressbar"
            aria-valuemin={0}
            aria-valuemax={100}
            aria-valuenow={p.progress.percent}
            aria-label="Import progress"
          >
            <span style={{ width: `${p.progress.percent}%` }} />
          </div>
          <dl>
            <dt>Filename</dt>
            <dd>{p.file.name}</dd>
            <dt>File size</dt>
            <dd>{formatBytes(p.file.size)}</dd>
            <dt>Current stage</dt>
            <dd>{p.progress.stage}</dd>
            <dt>Progress percentage</dt>
            <dd>{p.progress.percent}%</dd>
            <dt>Rows examined</dt>
            <dd>{p.progress.rowsExamined}</dd>
            <dt>Rows imported</dt>
            <dd>{p.progress.rowsImported}</dd>
            <dt>Rows skipped</dt>
            <dd>{p.progress.rowsSkipped}</dd>
            <dt>Current worksheet</dt>
            <dd>{p.progress.currentWorksheet ?? '—'}</dd>
            <dt>Elapsed time</dt>
            <dd>{Math.round(p.elapsed / 1000)}s</dd>
          </dl>
          {p.stalled && (
            <p className="note">
              This is taking longer than expected, but processing is still running.
            </p>
          )}
          <button onClick={p.cancel}>Cancel import</button>
        </section>
      </main>
    );
  return (
    <main className="upload">
      <section>
        <h1>Cargo Statement Analyser</h1>
        <p>Please upload your digital statement for summaries and analysis.</p>
        <div
          className="drop"
          onDragOver={(e) => e.preventDefault()}
          onDrop={(e) => {
            e.preventDefault();
            const f = e.dataTransfer.files[0];
            if (f) p.onFile(f);
          }}
        >
          <strong>Drop your Cargo statement here</strong>
          <span>
            Accepted formats:
            <br />• Excel (.xlsx, .xls)
            <br />• CSV (.csv)
          </span>
          <p className="note">
            For best results, upload the original statement downloaded from Cargo. CSV files may
            open faster for very large statements.
          </p>
          <label className="button">
            Choose file
            <input
              type="file"
              accept=".xlsx,.xls,.csv"
              onChange={(e) => {
                const f = e.target.files?.[0];
                if (f) p.onFile(f);
              }}
            />
          </label>
        </div>
        <aside className="privacy">
          <h2>Recommended</h2>
          <p>Upload the original statement downloaded from Cargo.</p>
          <p>Avoid opening and re-saving the statement in Excel before uploading if possible.</p>
          <details>
            <summary>Why?</summary>
            <p>
              Spreadsheet software can sometimes convert long identifiers such as barcodes into
              numbers.
            </p>
            <p>
              This may remove leading zeroes or alter very long identifiers before the analyser ever
              sees the file.
            </p>
            <p>Uploading the original downloaded statement gives the most reliable results.</p>
          </details>
          <details>
            <summary>How can Excel change my barcode?</summary>
            <p>Excel sometimes assumes long numeric identifiers are ordinary numbers.</p>
            <p>
              If a statement is opened and saved again, Excel may remove leading zeroes or round
              very long numbers.
            </p>
            <p>
              The analyser preserves the identifiers it receives, but it cannot restore digits that
              are no longer present in the uploaded file.
            </p>
          </details>
        </aside>
        <aside className="privacy compact-privacy">
          <p>
            <strong>
              Your statement is processed locally in your browser and is not uploaded.
            </strong>
          </p>
          <details>
            <summary>How?</summary>
            <p>
              Modern web browsers can open Excel files directly on your computer, just like
              Microsoft Excel can.
            </p>
            <p>
              This analyser runs entirely inside your browser, so your statement is never uploaded
              to Cargo or stored online.
            </p>
          </details>
          <details>
            <summary>Tell me more</summary>
            <p>When you choose a file:</p>
            <ol>
              <li>Your browser asks your computer for permission to open that one file.</li>
              <li>
                Your browser reads the file directly from your computer into its temporary working
                memory (RAM).
              </li>
              <li>
                The analyser reads the Excel workbook and creates the summaries you see on screen.
              </li>
              <li>Everything happens inside your browser.</li>
              <li>Nothing is uploaded to Cargo or sent anywhere else.</li>
              <li>When you close or refresh this page, the statement is removed from memory.</li>
            </ol>
            <p>
              <strong>We never receive, store or keep a copy of your statement.</strong>
            </p>
          </details>
        </aside>
        {p.err && (
          <section className="error">
            <h2>This statement could not be fully processed.</h2>
            <p>{p.err.message}</p>
            <p>
              Stage where processing failed: {p.err.stage} · Rows processed before failure:{' '}
              {p.err.rowsProcessed}
            </p>
            <details>
              <summary>Technical details</summary>
              <pre>{p.err.details || 'Technical details were not provided.'}</pre>
            </details>
            {p.err.canCompatibility && p.err.file && (
              <button onClick={() => p.onCompatibility(p.err!.file!)}>
                Try compatibility mode
              </button>
            )}
            <button onClick={() => location.reload()}>Try another file</button>
          </section>
        )}
        <button onClick={p.onSample}>Load sample statement</button>
      </section>
    </main>
  );
}
function formatBytes(size: number) {
  return size > 1024 * 1024
    ? `${(size / 1024 / 1024).toFixed(1)} MB`
    : `${(size / 1024).toFixed(1)} KB`;
}
function Dashboard({ data, remove }: { data: StatementData; remove: () => void }) {
  useEffect(() => {
    if (!data.diagnostics.importTimings?.dashboardFirstRenderMs)
      recordTimingValue(data.diagnostics, 'dashboardFirstRenderMs', safeNow());
  }, [data]);
  const [tab, setTab] = useState('Overview');
  const [exportProgress, setExportProgress] = useState<ExportProgress | null>(null);
  const [exportError, setExportError] = useState('');
  async function zipExport() {
    setExportError('');
    setExportProgress({ stage: 'Validating totals', index: 0, total: 14 });
    try {
      await downloadBreakdownsZip(data, setExportProgress);
    } catch (e) {
      setExportError(e instanceof Error ? e.message : String(e));
    } finally {
      setTimeout(() => setExportProgress(null), 1200);
    }
  }
  const [filters, setFilters] = useState<Record<string, string>>({});
  const [quick, setQuick] = useState<{
    field: keyof Transaction;
    q: string;
    exact?: boolean;
  } | null>(null);
  const filtered = useMemo(
    () =>
      data.rows.filter((r) =>
        Object.entries(filters).every(
          ([k, v]) =>
            !v ||
            String((r as any)[k])
              .toLowerCase()
              .includes(v.toLowerCase()),
        ),
      ),
    [data, filters],
  );
  const isUnfiltered = !Object.values(filters).some(Boolean);
  const initial = data.diagnostics.initialSummary;
  const t = useMemo(() => totals(filtered), [filtered]);
  const core = useMemo(
    () => ({
      artists: groupArtists(filtered),
      releases: groupReleases(filtered),
      tracks: groupTracks(filtered),
    }),
    [filtered],
  );
  const [cache, setCache] = useState<Record<string, any>>({});
  useEffect(() => setCache({}), [filtered]);
  function getTabData(name: string, calc: () => any) {
    if (cache[name]) return cache[name];
    if (tab !== name && name !== 'Overview') return null;
    const st = safeNow();
    const v = calc();
    recordTiming(data.diagnostics, 'tabFirstCalculationMs', name, safeNow() - st);
    setTimeout(() => setCache((c) => (c[name] ? c : { ...c, [name]: v })), 0);
    return v;
  }
  const artists = core.artists,
    releases = core.releases,
    tracks = core.tracks;
  const shops = getTabData('Shops', () => groupBy(filtered, 'shop'));
  const countries = getTabData('Countries', () => groupBy(filtered, 'country'));
  const salesPeriods = getTabData('Sales Periods', () =>
    groupBy(filtered, 'salesPeriod').sort((a, b) =>
      periodSortValue(a.salesPeriod).localeCompare(periodSortValue(b.salesPeriod)),
    ),
  );
  const importChecks = getTabData('Import Checks', () => importCheckGroups(filtered));
  const usageTypes = importChecks?.usageTypes;
  const royaltyRates = importChecks?.royaltyRates;
  const rates = useMemo(
    () =>
      isUnfiltered && initial
        ? { label: initial.royaltyRateSummary, rates: [] }
        : rateSummary(filtered),
    [filtered, isUnfiltered, initial],
  );
  const searchIndexes = useProgressiveSearchIndexes(filtered, data.diagnostics);
  const focused = quick?.q
    ? searchIndexes[String(quick.field)]
      ? searchRows(filtered, quick.field, quick.q, quick.exact, searchIndexes)
      : []
    : [];
  return (
    <main className="app">
      <header>
        <div>
          <h1>Cargo Statement Analyser</h1>
          <p>Please upload your digital statement for summaries and analysis.</p>
          <p>
            <b>{data.label}</b> · {data.filename}{' '}
            {data.diagnostics.account && `· Account: ${data.diagnostics.account}`}{' '}
            {data.diagnostics.reportingPeriod &&
              `· Reporting period: ${data.diagnostics.reportingPeriod}`}
          </p>
        </div>
        <button onClick={remove}>Remove statement</button>
      </header>
      <StatementHealth diagnostics={data.diagnostics} openChecks={() => setTab('Import Checks')} />
      <section className="zip-export">
        <button className="primary" onClick={zipExport} disabled={!!exportProgress}>
          Download all breakdowns as ZIP
        </button>
        {exportProgress && (
          <p aria-live="polite">
            {exportProgress.stage} ({exportProgress.index}/{exportProgress.total})
          </p>
        )}
        {exportError && <p className="error">{exportError}</p>}
      </section>
      <Filters filters={filters} setFilters={setFilters} rows={data.rows} />
      <Quick setQuick={setQuick} />
      {quick &&
        (searchIndexes[String(quick.field)] ? (
          <Focused rows={focused} clear={() => setQuick(null)} />
        ) : (
          <section className="focused">
            <button onClick={() => setQuick(null)}>Clear selected result</button>
            <p>Preparing search…</p>
          </section>
        ))}
      <section className="cards">
        <Card
          title="Your earnings"
          value={
            isUnfiltered && initial
              ? fmtMoney(initial.totalRoyaltyAmount)
              : fmtMoney(t.royaltyAmount)
          }
        />
        <Card
          title="Revenue before your royalty rate"
          value={isUnfiltered && initial ? fmtMoney(initial.totalAmount) : fmtMoney(t.amount)}
        />
        <Card
          title="Sales / usages"
          value={isUnfiltered && initial ? fmtInt(initial.totalSales) : fmtInt(t.sales)}
        />
        <Card
          title="Deductions and line charges"
          value={
            isUnfiltered && initial ? fmtMoney(initial.totalDeductions) : fmtMoney(t.deductions)
          }
        />
        <Card title="Transaction rows" value={String(filtered.length)} />
        <Card
          title="Releases"
          value={String(isUnfiltered && initial ? initial.uniqueReleaseCount : releases.length)}
        />
        <Card
          title="Tracks"
          value={String(isUnfiltered && initial ? initial.uniqueTrackCount : tracks.length)}
        />
        <Card
          title="Artists"
          value={String(isUnfiltered && initial ? initial.uniqueArtistCount : artists.length)}
        />
        <Card title="Royalty rate" value={rates.label} />
      </section>
      <nav>
        {tabs.map((x) => (
          <button className={tab === x ? 'active' : ''} onClick={() => setTab(x)} key={x}>
            {x}
          </button>
        ))}
      </nav>
      {tab === 'Overview' && <Overview rows={filtered} diagnostics={data.diagnostics} />}{' '}
      {tab === 'Artists' && (
        <Summary
          sourceRows={filtered}
          rows={artists}
          cols={[
            'artist',
            'royaltyAmount',
            'amount',
            'sales',
            'releaseCount',
            'trackCount',
            'shopCount',
            'countryCount',
            'transactionRows',
          ]}
          name="artist summary"
        />
      )}
      {tab === 'Releases' && (
        <Summary
          sourceRows={filtered}
          rows={releases}
          cols={[
            'barcode',
            'albumTitle',
            'artist',
            'catalogNumber',
            'releaseCode',
            'royaltyAmount',
            'amount',
            'sales',
            'trackCount',
            'transactionRows',
          ]}
          name="release summary"
        />
      )}
      {tab === 'Tracks' && (
        <Summary
          sourceRows={trackRows(filtered)}
          rows={tracks}
          cols={[
            'artist',
            'trackTitle',
            'isrc',
            'albumTitle',
            'catalogNumber',
            'barcode',
            'royaltyAmount',
            'amount',
            'sales',
            'shopCount',
            'countryCount',
            'transactionRows',
          ]}
          name="track summary"
        />
      )}
      {tab === 'Shops' &&
        (shops ? (
          <Summary
            sourceRows={filtered}
            rows={shops}
            cols={[
              'shop',
              'royaltyAmount',
              'amount',
              'sales',
              'artistCount',
              'releaseCount',
              'trackCount',
              'transactionRows',
            ]}
            name="shop summary"
          />
        ) : (
          <p className="note">Loading Shops…</p>
        ))}
      {tab === 'Countries' &&
        (countries ? (
          <Summary
            sourceRows={filtered}
            rows={countries}
            cols={[
              'country',
              'royaltyAmount',
              'amount',
              'sales',
              'artistCount',
              'releaseCount',
              'trackCount',
              'transactionRows',
            ]}
            name="country summary"
          />
        ) : (
          <p className="note">Loading Countries…</p>
        ))}
      {tab === 'Sales Periods' && (
        <>
          <p className="note">
            Sales Period is when the underlying usage or sale occurred. It may be earlier than the
            statement reporting month.
          </p>
          {salesPeriods ? (
            <Summary
              sourceRows={filtered}
              rows={salesPeriods}
              cols={[
                'salesPeriod',
                'royaltyAmount',
                'amount',
                'sales',
                'shopCount',
                'countryCount',
                'transactionRows',
              ]}
              name="sales-period summary"
            />
          ) : (
            <p className="note">Loading Sales Periods…</p>
          )}
        </>
      )}
      {tab === 'Deductions' && <Deductions rows={filtered} />}{' '}
      {tab === 'Full Detail' && <Detail rows={filtered} />}{' '}
      {tab === 'Import Checks' &&
        (importChecks ? (
          <Checks data={data} filteredRows={filtered} groups={importChecks} />
        ) : (
          <p className="note">Preparing reconciliation checks…</p>
        ))}
    </main>
  );
}

const defaultStatementHealth = {
  fileStatus: 'Original statement appears intact',
  dataQuality: 'Excellent',
  barcodeWarnings: 0,
  rowsRequiringReview: 0,
};
function StatementHealth({
  diagnostics,
  openChecks,
}: {
  diagnostics: any;
  openChecks: () => void;
}) {
  const h = { ...defaultStatementHealth, ...(diagnostics?.statementHealth ?? {}) };
  return (
    <section className="statement-health card">
      <h2>Statement Health</h2>
      <p>{h.fileStatus}</p>
      <p>
        Data Quality: <b>{h.dataQuality}</b>
      </p>
      <p>Barcode warnings: {h.barcodeWarnings}</p>
      <p>Rows requiring review: {h.rowsRequiringReview}</p>
      <button onClick={openChecks}>View data quality checks</button>
    </section>
  );
}
function Card(p: { title: string; value: string }) {
  return (
    <div className="card">
      <span>{p.title}</span>
      <b>{p.value}</b>
    </div>
  );
}
function Filters({
  filters,
  setFilters,
  rows,
}: {
  filters: Record<string, string>;
  setFilters: (f: Record<string, string>) => void;
  rows: Transaction[];
}) {
  const fields = [
    ['artist', 'Artist'],
    ['albumTitle', 'Release'],
    ['trackTitle', 'Track'],
    ['catalogNumber', 'Catalogue number'],
    ['barcode', 'Barcode'],
    ['isrc', 'ISRC'],
    ['shop', 'Shop'],
    ['country', 'Country'],
    ['usageType', 'Usage Type'],
    ['assetType', 'Asset Type'],
    ['salesPeriod', 'Sales Period'],
    ['royaltyRate', 'Royalty Rate'],
  ];
  return (
    <details className="filters">
      <summary>Global filters · {Object.values(filters).filter(Boolean).length} active</summary>
      <div>
        {fields.map(([k, l]) => (
          <label key={k}>
            {l}
            <input
              value={filters[k] ?? ''}
              onChange={(e) => setFilters({ ...filters, [k]: e.target.value })}
            />
          </label>
        ))}
        <button onClick={() => setFilters({})}>Clear filters</button>
      </div>
    </details>
  );
}
function Quick({ setQuick }: { setQuick: (q: any) => void }) {
  const qs = [
    ['artist', 'Search artist'],
    ['albumTitle', 'Search release'],
    ['catalogNumber', 'Search catalogue number'],
    ['barcode', 'Search barcode'],
    ['isrc', 'Search ISRC'],
    ['trackTitle', 'Search track'],
    ['shop', 'Search shop'],
    ['country', 'Search country'],
    ['salesPeriod', 'Search sales period'],
  ];
  return (
    <section className="quick">
      {qs.map(([f, l]) => (
        <label key={f}>
          {l}
          <input
            onKeyDown={(e) => {
              if (e.key === 'Enter')
                setQuick({
                  field: f,
                  q: (e.target as HTMLInputElement).value,
                  exact: ['catalogNumber', 'barcode', 'isrc'].includes(f),
                });
            }}
            placeholder="Press Enter"
          />
        </label>
      ))}
    </section>
  );
}

function useProgressiveSearchIndexes(rows: Transaction[], diagnostics: any, enabled = true) {
  const [indexes, setIndexes] = useState<Record<string, Map<string, number[]>>>({});
  useEffect(() => {
    let cancelled = false;
    setIndexes({});
    if (!enabled || !rows.length)
      return () => {
        cancelled = true;
      };
    const fields: (keyof Transaction)[] = [
      'artist',
      'albumTitle',
      'catalogNumber',
      'barcode',
      'isrc',
      'trackTitle',
      'shop',
      'country',
      'salesPeriod',
    ];
    let i = 0;
    function buildNext() {
      if (cancelled || i >= fields.length) return;
      const field = fields[i++];
      const st = safeNow();
      const m = new Map<string, number[]>();
      rows.forEach((r, idx) => {
        const v = String(r[field] ?? '').toLowerCase();
        if (!v) return;
        if (!m.has(v)) m.set(v, []);
        m.get(v)!.push(idx);
      });
      recordTiming(diagnostics, 'searchIndexMs', String(field), safeNow() - st);
      setIndexes((prev) => ({ ...prev, [String(field)]: m }));
      setTimeout(buildNext, 0);
    }
    setTimeout(buildNext, 0);
    return () => {
      cancelled = true;
    };
  }, [rows, diagnostics, enabled]);
  return indexes;
}
function Focused({ rows, clear }: { rows: Transaction[]; clear: () => void }) {
  const t = totals(rows);
  return (
    <section className="focused">
      <button onClick={clear}>Clear selected result</button>
      <h2>Focused result</h2>
      <div className="cards">
        <Card title="Your earnings" value={fmtMoney(t!.royaltyAmount)} />
        <Card title="Revenue before royalty rate" value={fmtMoney(t!.amount)} />
        <Card title="Sales / usages" value={fmtInt(t!.sales)} />
        <Card title="Transaction count" value={String(rows.length)} />
        <Card title="Releases included" value={String(groupReleases(rows).length)} />
        <Card title="Tracks included" value={String(groupTracks(rows).length)} />
      </div>
      <Overview rows={rows} />
    </section>
  );
}
function Overview({ rows, diagnostics }: { rows: Transaction[]; diagnostics?: any }) {
  const [prepare, setPrepare] = useState(!diagnostics?.largeFileMode);
  const [chartData, setChartData] = useState<any | null>(null);
  const [chartError, setChartError] = useState<string | null>(null);
  const cache = useRef<Record<string, any>>({});
  const [key, setKey] = useState<string | null>(null);
  useEffect(() => {
    let cancelled = false;
    setKey(null);
    if (!prepare) {
      setChartData(null);
      return () => {
        cancelled = true;
      };
    }
    chartDatasetCacheKey(rows)
      .then((k) => {
        if (!cancelled) setKey(k);
      })
      .catch((e) => {
        if (!cancelled) setChartError(e instanceof Error ? e.message : String(e));
      });
    return () => {
      cancelled = true;
    };
  }, [rows, prepare]);
  useEffect(() => {
    setChartError(null);
    if (!key) {
      setChartData(null);
      return;
    }
    if (cache.current[key]) {
      setChartData(cache.current[key]);
      return;
    }
    setChartData(null);
    let cancelled = false;
    const schedule = (cb: () => void) => {
      const ric = (window as any).requestIdleCallback as
        undefined | ((f: () => void, o?: any) => number);
      return ric ? ric(cb, { timeout: 800 }) : window.setTimeout(cb, 0);
    };
    const cancel = (id: number) => {
      const cic = (window as any).cancelIdleCallback as undefined | ((n: number) => void);
      if (cic) cic(id);
      else clearTimeout(id);
    };
    const id = schedule(() => {
      try {
        const st = safeNow();
        const data = prepareOverviewChartDataFromRows(rows);
        cache.current[key] = data;
        if (!cancelled) setChartData(data);
        recordTimingValue(diagnostics, 'chartsMs', safeNow() - st);
      } catch (e) {
        const msg = e instanceof Error ? e.message : String(e);
        if (diagnostics?.invalidNumericValues)
          diagnostics.invalidNumericValues.push(`Chart aggregation failed: ${msg}`);
        if (!cancelled) setChartError(msg);
      }
    });
    return () => {
      cancelled = true;
      cancel(id);
    };
  }, [key, rows, diagnostics]);
  return (
    <section className="overview" aria-label="Overview charts">
      {!prepare && (
        <div className="chart-card">
          <p className="note">Large File Mode is active to reduce browser memory use.</p>
          <button onClick={() => setPrepare(true)}>Prepare detailed analysis</button>
        </div>
      )}
      {prepare && chartError && <p className="error">Charts could not be prepared: {chartError}</p>}
      {prepare &&
        (!chartData ? (
          <div className="chart-card">
            <p className="note">Preparing chart…</p>
          </div>
        ) : (
          <>
            <BarChart
              title="Top releases by earnings"
              items={chartData.topReleases}
              label={(x) => x.albumTitle || 'Unspecified release'}
              tooltip={(x) =>
                [
                  `Full Album Title: ${x.albumTitle || 'Unspecified release'}`,
                  `Artist: ${x.artist || 'Unspecified artist'}`,
                  `Barcode: ${x.barcode || ''}`,
                  `Catalog Number: ${x.catalogNumber || ''}`,
                  `Release Code: ${x.releaseCode || ''}`,
                  `Your earnings: ${fmtMoney(x.royaltyAmount)}`,
                  `Revenue before your royalty rate: ${fmtMoney(x.amount)}`,
                  `Sales / usages: ${fmtInt(x.sales)}`,
                ].join('\n')
              }
            />
            <BarChart
              title="Top artists by earnings"
              items={chartData.topArtists}
              label={(x) => x.artist || 'Unspecified artist'}
              tooltip={(x) =>
                [
                  `Artist: ${x.artist || 'Unspecified artist'}`,
                  `Your earnings: ${fmtMoney(x.royaltyAmount)}`,
                  `Revenue before your royalty rate: ${fmtMoney(x.amount)}`,
                  `Sales / usages: ${fmtInt(x.sales)}`,
                  `Release count: ${fmtInt(x.releaseCount)}`,
                ].join('\n')
              }
            />
            <LineChart title="Earnings by sales period" items={chartData.salesPeriods} />
            <BarChart
              title="Top shops by earnings"
              items={chartData.topShops}
              label={(x) => x.shop || 'Unspecified shop'}
              tooltip={(x) => {
                const total = D(chartData.filteredRoyaltyTotal);
                const pct = total.isZero()
                  ? '0.0'
                  : D(x.royaltyAmount).div(total).mul(100).toFixed(1);
                return [
                  `Shop: ${x.shop || 'Unspecified shop'}`,
                  `Your earnings: ${fmtMoney(x.royaltyAmount)}`,
                  `Revenue before your royalty rate: ${fmtMoney(x.amount)}`,
                  `Sales / usages: ${fmtInt(x.sales)}`,
                  `Percentage of total filtered earnings: ${pct}%`,
                ].join('\n');
              }}
            />
            <BarChart
              title="Earnings by usage type"
              items={chartData.usageTypes}
              label={(x) => x.usageType || 'Unspecified usage type'}
              tooltip={(x) =>
                [
                  `Usage Type: ${x.usageType || 'Unspecified usage type'}`,
                  `Your earnings: ${fmtMoney(x.royaltyAmount)}`,
                  `Revenue before your royalty rate: ${fmtMoney(x.amount)}`,
                  `Sales / usages: ${fmtInt(x.sales)}`,
                ].join('\n')
              }
            />
          </>
        ))}
    </section>
  );
}
function BarChart({
  title,
  items,
  label,
  tooltip,
}: {
  title: string;
  items: any[];
  label: (x: any) => string;
  tooltip: (x: any) => string;
}) {
  const max = items.reduce((m, x) => Decimal.max(m, D(x.royaltyAmount).abs()), D(0));
  return (
    <section className="chart-card" aria-label={title}>
      <h2>{title}</h2>
      {items.map((x, i) => {
        const width = max.isZero() ? 0 : D(x.royaltyAmount).abs().div(max).mul(100).toNumber();
        return (
          <div className="bar-row" key={i} title={tooltip(x)}>
            <span>{label(x)}</span>
            <div>
              <i style={{ width: `${width}%` }} />
              <b>{fmtMoney(x.royaltyAmount)}</b>
            </div>
          </div>
        );
      })}
      {!items.length && <p className="empty">No earnings to chart.</p>}
    </section>
  );
}
function LineChart({ title, items }: { title: string; items: any[] }) {
  const max = items.reduce((m, x) => Decimal.max(m, D(x.royaltyAmount).abs()), D(0));
  const points = items
    .map((x, i) => {
      const xp = items.length <= 1 ? 50 : (i / (items.length - 1)) * 100,
        yp = max.isZero() ? 90 : 100 - D(x.royaltyAmount).abs().div(max).mul(90).toNumber();
      return `${xp},${yp}`;
    })
    .join(' ');
  return (
    <section className="chart-card" aria-label={title}>
      <h2>{title}</h2>
      <p className="note">
        Sales Period is when the underlying usage or sale occurred. It may be earlier than the
        statement reporting month.
      </p>
      <svg viewBox="0 0 100 100" role="img" aria-label="Earnings by sales period line chart">
        <polyline fill="none" stroke="#1f4b99" strokeWidth="3" points={points} />
        {items.map((x, i) => {
          const [cx, cy] = (points.split(' ')[i] || '0,0').split(',');
          return (
            <circle key={i} cx={cx} cy={cy} r="2.5">
              <title>
                {[
                  `Sales Period: ${x.salesPeriod}`,
                  `Your earnings: ${fmtMoney(x.royaltyAmount)}`,
                  `Revenue before your royalty rate: ${fmtMoney(x.amount)}`,
                  `Sales / usages: ${fmtInt(x.sales)}`,
                ].join('\n')}
              </title>
            </circle>
          );
        })}
      </svg>
      <div className="period-labels">
        {items.map((x: any) => (
          <span key={x.salesPeriod}>
            {x.salesPeriod}
            <br />
            {fmtMoney(x.royaltyAmount)}
          </span>
        ))}
      </div>
      {!items.length && <p className="empty">No valid sales periods to chart.</p>}
    </section>
  );
}
function csvRows(rows: any[], cols: string[], includeTotal = true) {
  const body = rows.map((r) =>
    Object.fromEntries(
      cols.map((c) => [
        c,
        r[c] instanceof Object && r[c].toString ? r[c].toString() : (r[c] ?? ''),
      ]),
    ),
  );
  if (!includeTotal) return body;
  const total = rows.reduce(
    (a, r) => ({
      royaltyAmount: a.royaltyAmount.plus(D(r.royaltyAmount)),
      amount: a.amount.plus(D(r.amount)),
      sales: a.sales.plus(D(r.sales)),
      returns: a.returns.plus(D(r.returns)),
    }),
    { royaltyAmount: D(0), amount: D(0), sales: D(0), returns: D(0) },
  );
  return [
    ...body,
    Object.fromEntries(
      cols.map((c, i) => [
        c,
        i === 0
          ? 'TOTAL'
          : ['royaltyAmount', 'amount', 'sales', 'returns'].includes(c)
            ? total[c as keyof typeof total].toString()
            : '',
      ]),
    ),
  ];
}
function Summary({
  rows,
  exportRows,
  sourceRows,
  cols,
  name,
}: {
  rows: any[];
  exportRows?: any[];
  sourceRows?: Transaction[];
  cols: string[];
  name: string;
}) {
  const completeRows = exportRows ?? rows;
  const recon = sourceRows ? groupedReconciliation(sourceRows, completeRows) : null;
  return (
    <section className="panel">
      <button
        onClick={() =>
          downloadCsv(
            `${name}.csv`,
            toCsv(csvRows(completeRows, cols, name !== 'release summary'), cols),
          )
        }
      >
        Export complete summary
      </button>
      {recon && (
        <p className={recon.reconciled ? 'note' : 'error'}>
          {name} totals reconcile to statement:{' '}
          {recon.reconciled
            ? 'Yes'
            : `No, royalty difference ${fmtMoney(recon.diff.royaltyAmount)}`}
        </p>
      )}
      <table>
        <thead>
          <tr>
            {cols.map((c) => (
              <th key={c}>{c}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((r, i) => (
            <React.Fragment key={i}>
              <tr>
                {cols.map((c) => (
                  <td key={c}>
                    {['amount', 'royaltyAmount', 'deductions', 'lineCharges', 'total'].includes(c)
                      ? fmtMoney(r[c])
                      : ['sales', 'returns'].includes(c)
                        ? fmtInt(r[c])
                        : String(r[c] ?? '')}
                  </td>
                ))}
              </tr>
              {r.hasMultipleIdentifiers && (
                <tr>
                  <td colSpan={cols.length}>
                    <details>
                      <summary>Contributing release identifiers</summary>
                      <dl>
                        <dt>Barcodes</dt>
                        <dd>{r.barcodes?.join(' | ') || 'Blank'}</dd>
                        <dt>Catalog Numbers</dt>
                        <dd>{r.catalogNumbers?.join(' | ') || 'Blank'}</dd>
                        <dt>Release Codes</dt>
                        <dd>{r.releaseCodes?.join(' | ') || 'Blank'}</dd>
                        <dt>Album Titles</dt>
                        <dd>{r.albumTitles?.join(' | ') || 'Blank'}</dd>
                        <dt>Artists</dt>
                        <dd>{r.artists?.join(' | ') || 'Blank'}</dd>
                      </dl>
                    </details>
                  </td>
                </tr>
              )}
            </React.Fragment>
          ))}
        </tbody>
      </table>
      {!rows.length && <p className="empty">No rows to show.</p>}
    </section>
  );
}
function Deductions({ rows }: { rows: Transaction[] }) {
  const totalsBy = deductionFields.map((f) => ({
    field: f,
    total: rows.reduce((a, r) => a.plus(D(r[f])), D(0)),
  }));
  const any = totalsBy.some((x) => !x.total.isZero());
  return (
    <section>
      {!any && <p>No deductions or line charges were found in this statement.</p>}
      <Summary rows={totalsBy} cols={['field', 'total']} name="deductions summary" />
      <Summary
        rows={groupArtists(rows.filter((r) => deductionFields.some((f) => !D(r[f]).isZero())))}
        cols={['artist', 'deductions', 'transactionRows']}
        name="deductions by artist"
      />
    </section>
  );
}
function Detail({ rows }: { rows: Transaction[] }) {
  const [globalFilter, setGlobalFilter] = useState('');
  const cols = useMemo<ColumnDef<Transaction>[]>(
    () =>
      Object.keys(rows[0] ?? { sourceSheet: '', sourceRow: 0 })
        .filter((k) => k !== 'originalRow')
        .map((k) => ({ accessorKey: k, header: k, cell: (i) => String(i.getValue() ?? '') })),
    [rows],
  );
  const table = useReactTable({
    data: rows,
    columns: cols,
    state: { globalFilter, pagination: { pageIndex: 0, pageSize: 50 } },
    onGlobalFilterChange: setGlobalFilter,
    getCoreRowModel: getCoreRowModel(),
    getSortedRowModel: getSortedRowModel(),
    getFilteredRowModel: getFilteredRowModel(),
    getPaginationRowModel: getPaginationRowModel(),
  });
  const visible = table.getFilteredRowModel().rows.map((r) => r.original);
  const t = totals(visible);
  return (
    <section className="panel">
      <input
        placeholder="Global search"
        value={globalFilter}
        onChange={(e) => setGlobalFilter(e.target.value)}
      />
      <button onClick={() => setGlobalFilter('')}>Clear filters</button>
      <button onClick={() => downloadCsv('filtered-results.csv', toCsv(normalisedRows(visible)))}>
        Export filtered results
      </button>
      <button onClick={() => downloadCsv('full-detail.csv', toCsv(normalisedRows(rows)))}>
        Export full detail
      </button>
      <p>
        Transaction count: {visible.length} · Filtered Your earnings: {fmtMoney(t!.royaltyAmount)} ·
        Amount: {fmtMoney(t!.amount)} · Sales: {fmtInt(t!.sales)} · Returns: {fmtInt(t.returns)}
      </p>
      <table className="detail">
        <thead>
          {table.getHeaderGroups().map((hg) => (
            <tr key={hg.id}>
              {hg.headers.map((h) => (
                <th key={h.id} onClick={h.column.getToggleSortingHandler()}>
                  {flexRender(h.column.columnDef.header, h.getContext())}
                </th>
              ))}
            </tr>
          ))}
        </thead>
        <tbody>
          {table.getRowModel().rows.map((r) => (
            <tr key={r.id}>
              {r.getVisibleCells().map((c) => (
                <td key={c.id}>{flexRender(c.column.columnDef.cell, c.getContext())}</td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
      <button onClick={() => table.previousPage()}>Previous</button>
      <button onClick={() => table.nextPage()}>Next</button>
    </section>
  );
}
function Checks({
  data,
  filteredRows,
  groups,
}: {
  data: StatementData;
  filteredRows: Transaction[];
  groups: Record<string, any[]>;
}) {
  const d = data.diagnostics,
    t = totals(filteredRows),
    tracksSource = trackRows(filteredRows);
  const [dups, setDups] = useState<number | null>(d.largeFileMode ? null : d.duplicateLookingRows);
  const checks = [
    ['Artists', groups.artists],
    ['Releases', groups.releases],
    ['Tracks', groups.tracks, tracksSource],
    ['Shops', groups.shops],
    ['Countries', groups.countries],
    ['Sales Periods', groups.salesPeriods],
    ['Usage Types', groups.usageTypes],
    ['Royalty Rates', groups.royaltyRates],
  ].map(([label, g, source]) => ({
    label: String(label),
    ...groupedReconciliation((source as Transaction[]) ?? filteredRows, g as any[]),
  }));
  const barcodeRows = d.barcodeIntegrity.warnings.map((w) => ({
    sourceSheet: w.sourceSheet,
    sourceRow: String(w.sourceRow),
    barcodeValue: w.barcodeValue,
    warning: w.warning,
    suggestedReason: w.suggestedReason,
  }));
  return (
    <section className="panel">
      <button onClick={() => setDups(findDuplicates(data.rows))}>Run duplicate check</button>
      <p className="note">This check may take longer on large statements.</p>
      {d.barcodeIntegrity.warnings.length > 0 && (
        <aside className="error">
          <p>⚠ Some barcode values may have been altered before upload.</p>
          <p>Possible causes include:</p>
          <ul>
            <li>Spreadsheet software converting barcodes into numbers.</li>
            <li>Leading zeroes removed.</li>
            <li>Scientific notation.</li>
            <li>Very long numeric identifiers.</li>
          </ul>
          <p>If possible, upload the original statement downloaded from Cargo.</p>
          <p>
            These warnings relate to the uploaded file itself; the analyser has not changed the
            data.
          </p>
        </aside>
      )}
      <h2>Barcode Integrity</h2>
      <button
        onClick={() =>
          downloadCsv(
            'barcode-integrity-warnings.csv',
            toCsv(barcodeRows, [
              'sourceSheet',
              'sourceRow',
              'barcodeValue',
              'warning',
              'suggestedReason',
            ]),
          )
        }
      >
        Export barcode warnings
      </button>
      <dl>
        <dt>Populated barcode rows</dt>
        <dd>{d.barcodeIntegrity.populatedBarcodeRows}</dd>
        <dt>Blank barcode rows</dt>
        <dd>{d.barcodeIntegrity.blankBarcodeRows}</dd>
        <dt>Unique barcode count</dt>
        <dd>{d.barcodeIntegrity.uniqueBarcodeCount}</dd>
        <dt>Numeric barcode cells</dt>
        <dd>{d.barcodeIntegrity.numericBarcodeCells}</dd>
        <dt>Text barcode cells</dt>
        <dd>{d.barcodeIntegrity.textBarcodeCells}</dd>
        <dt>Scientific notation values converted</dt>
        <dd>{d.barcodeIntegrity.scientificNotationValuesConverted}</dd>
        <dt>Decimal suffixes removed</dt>
        <dd>{d.barcodeIntegrity.decimalSuffixesRemoved}</dd>
        <dt>Possible lost-leading-zero warnings</dt>
        <dd>{d.barcodeIntegrity.possibleLostLeadingZeroWarnings}</dd>
        <dt>Unsafe precision warnings</dt>
        <dd>{d.barcodeIntegrity.unsafePrecisionWarnings}</dd>
        <dt>Duplicate barcode conflicts</dt>
        <dd>{d.barcodeIntegrity.duplicateBarcodeConflicts}</dd>
        <dt>Rows requiring review</dt>
        <dd>{d.barcodeIntegrity.rowsRequiringReview}</dd>
      </dl>
      <table>
        <thead>
          <tr>
            <th>Source sheet</th>
            <th>Source row</th>
            <th>Barcode value</th>
            <th>Warning</th>
            <th>Suggested reason</th>
          </tr>
        </thead>
        <tbody>
          {d.barcodeIntegrity.warnings.map((w: any, i: number) => (
            <tr key={i}>
              <td>{w.sourceSheet}</td>
              <td>{w.sourceRow}</td>
              <td>{w.barcodeValue}</td>
              <td>{w.warning}</td>
              <td>{w.suggestedReason}</td>
            </tr>
          ))}
        </tbody>
      </table>
      <h2>Grouped total reconciliation</h2>
      <table>
        <thead>
          <tr>
            {[
              'Group',
              'Dashboard Royalty Amount',
              'Grouped Royalty Amount',
              'Difference',
              'Dashboard Amount',
              'Grouped Amount',
              'Difference',
              'Dashboard Sales',
              'Grouped Sales',
              'Difference',
              'Dashboard Returns',
              'Grouped Returns',
              'Difference',
              'Reconciled',
            ].map((h) => (
              <th key={h}>{h}</th>
            ))}
          </tr>
        </thead>
        <tbody>
          {checks.map((c) => (
            <tr key={c.label}>
              <td>{c.label}</td>
              <td>{fmtMoney(c.dashboard.royaltyAmount)}</td>
              <td>{fmtMoney(c.grouped.royaltyAmount)}</td>
              <td>{fmtMoney(c.diff.royaltyAmount)}</td>
              <td>{fmtMoney(c.dashboard.amount)}</td>
              <td>{fmtMoney(c.grouped.amount)}</td>
              <td>{fmtMoney(c.diff.amount)}</td>
              <td>{fmtInt(c.dashboard.sales)}</td>
              <td>{fmtInt(c.grouped.sales)}</td>
              <td>{c.diff.sales.toString()}</td>
              <td>{fmtInt(c.dashboard.returns)}</td>
              <td>{fmtInt(c.grouped.returns)}</td>
              <td>{c.diff.returns.toString()}</td>
              <td>{c.reconciled ? 'Yes' : 'No'}</td>
            </tr>
          ))}
        </tbody>
      </table>
      <dl>
        <dt>duplicateLookingRows</dt>
        <dd>{dups ?? 'Deferred'}</dd>
        {Object.entries(d).map(([k, v]) => (
          <React.Fragment key={k}>
            <dt>{k}</dt>
            <dd>{typeof v === 'object' ? JSON.stringify(v) : String(v)}</dd>
          </React.Fragment>
        ))}
        <dt>Total Amount</dt>
        <dd>{t.amount.toString()}</dd>
        <dt>Total Royalty Amount</dt>
        <dd>{t.royaltyAmount.toString()}</dd>
        <dt>Total Sales</dt>
        <dd>{t.sales.toString()}</dd>
        <dt>Total Returns</dt>
        <dd>{t.returns.toString()}</dd>
        <dt>Total deductions</dt>
        <dd>{t.deductions.toString()}</dd>
        <dt>Total line charges</dt>
        <dd>{t.lineCharges.toString()}</dd>
      </dl>
    </section>
  );
}
function findDuplicates(rows: Transaction[]) {
  const seen = new Map<string, number>();
  for (const r of rows) {
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
  return [...seen.values()].filter((v) => v > 1).reduce((a, v) => a + v, 0);
}
createRoot(document.getElementById('root')!).render(<App />);
