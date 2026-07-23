export type WorkerEventName =
  | 'file-read'
  | 'worker-created'
  | 'worker-ready'
  | 'worker-startup-timeout'
  | 'worker-message'
  | 'worker-error'
  | 'worker-messageerror'
  | 'post-message'
  | 'parse-start'
  | 'sheetjs-read'
  | 'complete'
  | 'error'
  | 'compatibility-mode';
export interface ImportDebugContext {
  errorName?: string;
  errorMessage?: string;
  stack?: string;
  filename?: string;
  lineNumber?: number;
  columnNumber?: number;
  workerUrl?: string;
  workerStartupCompleted?: 'Yes' | 'No' | 'Unknown';
  firstWorkerMessageReceived?: 'Yes' | 'No' | 'Unknown';
  processingStage: string;
  workerEvent: WorkerEventName | string;
  workbookFilename: string;
  fileSize: number;
  arrayBufferSizeBeforeTransfer?: number;
  arrayBufferSizeInsideWorker?: number;
  sheetJsReadAttempt?: string;
  retryAttemptNumber?: number;
  workerStarted?: 'Yes' | 'No' | 'Unknown';
  arrayBufferTransferSucceeded?: 'Yes' | 'No' | 'Unknown';
  workerPayloadReceived?: 'Yes' | 'No' | 'Unknown';
  sheetJsImported?: 'Yes' | 'No' | 'Unknown';
  xlsxReadInvoked?: 'Yes' | 'No' | 'Unknown';
}
export function serialiseThrown(value: unknown) {
  if (value instanceof Error)
    return {
      name: value.name || 'Error',
      message: value.message || String(value),
      stack: value.stack,
    };
  let message: string;
  try {
    message = typeof value === 'string' ? value : JSON.stringify(value);
  } catch {
    message = String(value);
  }
  return {
    name: Object.prototype.toString.call(value),
    message: message || String(value),
    stack: undefined,
  };
}
export function technicalDetails(ctx: ImportDebugContext, thrown?: unknown) {
  const s: ReturnType<typeof serialiseThrown> | Record<string, never> =
    thrown === undefined ? {} : serialiseThrown(thrown);
  const lines = [
    `error name: ${ctx.errorName ?? s.name ?? 'UnknownError'}`,
    `error message: ${ctx.errorMessage ?? s.message ?? 'No error message supplied.'}`,
    `stack trace: ${ctx.stack ?? s.stack ?? 'Not available'}`,
    `filename: ${ctx.filename ?? 'Not available'}`,
    `line number: ${ctx.lineNumber ?? 'Not available'}`,
    `column number: ${ctx.columnNumber ?? 'Not available'}`,
    `worker URL: ${ctx.workerUrl ?? 'Unknown'}`,
    `worker startup completed: ${ctx.workerStartupCompleted ?? 'Unknown'}`,
    `first worker message received: ${ctx.firstWorkerMessageReceived ?? 'Unknown'}`,
    `processing stage: ${ctx.processingStage}`,
    `worker started: ${ctx.workerStarted ?? 'Unknown'}`,
    `ArrayBuffer transfer succeeded: ${ctx.arrayBufferTransferSucceeded ?? 'Unknown'}`,
    `worker received expected payload: ${ctx.workerPayloadReceived ?? 'Unknown'}`,
    `SheetJS imported inside worker: ${ctx.sheetJsImported ?? 'Unknown'}`,
    `XLSX.read invoked: ${ctx.xlsxReadInvoked ?? 'Unknown'}`,
    `worker event: ${ctx.workerEvent}`,
    `workbook filename: ${ctx.workbookFilename}`,
    `file size: ${ctx.fileSize}`,
    `ArrayBuffer size before transfer: ${ctx.arrayBufferSizeBeforeTransfer ?? 'Unknown'}`,
    `ArrayBuffer size inside the worker: ${ctx.arrayBufferSizeInsideWorker ?? 'Unknown'}`,
    `SheetJS read attempt: ${ctx.sheetJsReadAttempt ?? 'Not started'}`,
    `retry attempt number: ${ctx.retryAttemptNumber ?? 0}`,
  ];
  return lines.join('\n');
}
