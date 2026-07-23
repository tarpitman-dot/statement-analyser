import {
  CONSERVATIVE_SHEETJS_READ_OPTIONS,
  parseArrayBuffer,
  sheetJsModuleLoaded,
} from './lib/parser';
import {
  serialiseThrown,
  technicalDetails,
  type ImportDebugContext,
} from './lib/importDiagnostics';

const workerScope = self as unknown as Worker;
const startedContext = {
  processingStage: 'Opening workbook',
  workerEvent: 'worker-started',
  workbookFilename: 'Unknown',
  fileSize: 0,
  workerStarted: 'Yes' as const,
  sheetJsImported: sheetJsModuleLoaded ? ('Yes' as const) : ('No' as const),
  xlsxReadInvoked: 'No' as const,
};

workerScope.postMessage({ type: 'worker-ready', context: startedContext });

function postError(error: unknown, context: ImportDebugContext) {
  const serialised = serialiseThrown(error);
  workerScope.postMessage({
    type: 'error',
    error: {
      name: serialised.name,
      message: serialised.message || 'This statement could not be fully processed.',
      stack: serialised.stack,
      details: technicalDetails(
        {
          errorName: serialised.name,
          errorMessage: serialised.message,
          stack: serialised.stack,
          ...context,
        },
        error,
      ),
      context,
    },
  });
}

self.onmessage = async (e: MessageEvent) => {
  const payload = e.data;
  let { buffer } = payload ?? {};
  const { filename, fileSize, arrayBufferSizeBeforeTransfer, fileReadMs } = payload ?? {};
  const transferOk =
    buffer instanceof ArrayBuffer &&
    buffer.byteLength === Number(arrayBufferSizeBeforeTransfer ?? buffer.byteLength);
  const base = {
    processingStage: 'Opening workbook',
    workerEvent: 'parse-start',
    workbookFilename: filename ?? 'Unknown',
    fileSize: Number(fileSize ?? 0),
    arrayBufferSizeBeforeTransfer: Number(arrayBufferSizeBeforeTransfer ?? 0),
    arrayBufferSizeInsideWorker: buffer?.byteLength ?? 0,
    workerStarted: 'Yes' as const,
    arrayBufferTransferSucceeded: transferOk ? ('Yes' as const) : ('No' as const),
    workerPayloadReceived:
      payload && buffer instanceof ArrayBuffer && typeof filename === 'string'
        ? ('Yes' as const)
        : ('No' as const),
    sheetJsImported: sheetJsModuleLoaded ? ('Yes' as const) : ('No' as const),
    xlsxReadInvoked: 'No' as const,
  };
  workerScope.postMessage({ type: 'worker-received', context: base });
  if (!(buffer instanceof ArrayBuffer) || buffer.byteLength === 0) {
    postError(new Error('Worker received an empty or invalid ArrayBuffer.'), {
      ...base,
      sheetJsReadAttempt: 'not attempted',
      retryAttemptNumber: 0,
    });
    return;
  }
  for (const attempt of [1, 2]) {
    try {
      const data = await parseArrayBuffer(buffer, filename, fileSize, {
        fileReadMs,
        sheetJsReadAttempt: attempt === 1 ? 'default' : 'conservative',
        retryAttemptNumber: attempt,
        readOptions: attempt === 1 ? undefined : CONSERVATIVE_SHEETJS_READ_OPTIONS,
        onBeforeSheetJsRead: () =>
          workerScope.postMessage({
            type: 'sheetjs-read-start',
            context: {
              ...base,
              workerEvent: 'sheetjs-read',
              sheetJsReadAttempt: attempt === 1 ? 'default' : 'conservative',
              retryAttemptNumber: attempt,
              xlsxReadInvoked: 'Yes',
            },
          }),
        onProgress: (p) =>
          workerScope.postMessage({
            type: 'progress',
            progress: p,
            context: {
              ...base,
              sheetJsReadAttempt: attempt === 1 ? 'default' : 'conservative',
              retryAttemptNumber: attempt,
              xlsxReadInvoked: p.stage === 'Opening workbook' ? 'Yes' : 'Yes',
            },
          }),
      });
      data.diagnostics.memory = { ...(data.diagnostics.memory ?? {}), workerTerminated: 'No' };
      buffer = null;
      if (payload) payload.buffer = null;
      workerScope.postMessage({
        type: 'complete',
        data,
        context: {
          ...base,
          processingStage: 'Complete',
          workerEvent: 'complete',
          sheetJsReadAttempt: attempt === 1 ? 'default' : 'conservative',
          retryAttemptNumber: attempt,
          xlsxReadInvoked: 'Yes',
        },
      });
      return;
    } catch (error) {
      if (attempt === 2) {
        buffer = null;
        if (payload) payload.buffer = null;
        postError(error, {
          ...base,
          workerEvent: 'error',
          sheetJsReadAttempt: 'conservative',
          retryAttemptNumber: 2,
        });
        return;
      }
    }
  }
};
