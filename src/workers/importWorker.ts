import {
  serialiseThrown,
  technicalDetails,
  type ImportDebugContext,
} from '../lib/importDiagnostics';
import type { ParseOptions } from '../lib/parser';

const workerScope = self as unknown as Worker;

type ParserModule = typeof import('../lib/parser');
let parserModule: ParserModule | null = null;
let startupComplete = false;

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

function baseContext(overrides: Partial<ImportDebugContext> = {}): ImportDebugContext {
  return {
    processingStage: 'Opening workbook',
    workerEvent: 'worker-started',
    workbookFilename: 'Unknown',
    fileSize: 0,
    workerStarted: startupComplete ? 'Yes' : 'No',
    arrayBufferTransferSucceeded: 'Unknown',
    workerPayloadReceived: 'Unknown',
    sheetJsImported: parserModule?.sheetJsModuleLoaded ? 'Yes' : 'Unknown',
    xlsxReadInvoked: 'No',
    ...overrides,
  };
}

async function startup() {
  try {
    parserModule = await import('../lib/parser');
    startupComplete = true;
    workerScope.postMessage({
      type: 'worker-ready',
      context: baseContext({
        workerEvent: 'worker-ready',
        workerStarted: 'Yes',
        sheetJsImported: parserModule.sheetJsModuleLoaded ? 'Yes' : 'No',
      }),
    });
  } catch (error) {
    postError(error, baseContext({ workerEvent: 'worker-startup-error', sheetJsImported: 'No' }));
  }
}

workerScope.onmessageerror = () => {
  postError(
    new Error('Worker message could not be deserialised.'),
    baseContext({ workerEvent: 'worker-messageerror' }),
  );
};

workerScope.onmessage = async (e: MessageEvent) => {
  try {
    if (!parserModule) {
      postError(
        new Error('Worker received a workbook before startup completed.'),
        baseContext({ workerEvent: 'parse-start', workerPayloadReceived: 'No' }),
      );
      return;
    }
    const payload = e.data;
    let { buffer } = payload ?? {};
    const { filename, fileSize, arrayBufferSizeBeforeTransfer, fileReadMs } = payload ?? {};
    const transferOk =
      buffer instanceof ArrayBuffer &&
      buffer.byteLength === Number(arrayBufferSizeBeforeTransfer ?? buffer.byteLength);
    const base = baseContext({
      workerEvent: 'parse-start',
      workbookFilename: filename ?? 'Unknown',
      fileSize: Number(fileSize ?? 0),
      arrayBufferSizeBeforeTransfer: Number(arrayBufferSizeBeforeTransfer ?? 0),
      arrayBufferSizeInsideWorker: buffer instanceof ArrayBuffer ? buffer.byteLength : 0,
      workerStarted: 'Yes',
      arrayBufferTransferSucceeded: transferOk ? 'Yes' : 'No',
      workerPayloadReceived:
        payload && buffer instanceof ArrayBuffer && typeof filename === 'string' ? 'Yes' : 'No',
      sheetJsImported: parserModule.sheetJsModuleLoaded ? 'Yes' : 'No',
      xlsxReadInvoked: 'No',
    });
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
        const readOptions: ParseOptions['readOptions'] =
          attempt === 1 ? undefined : parserModule.CONSERVATIVE_SHEETJS_READ_OPTIONS;
        const data = await parserModule.parseArrayBuffer(buffer, filename, fileSize, {
          fileReadMs,
          sheetJsReadAttempt: attempt === 1 ? 'default' : 'conservative',
          retryAttemptNumber: attempt,
          readOptions,
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
                xlsxReadInvoked: 'Yes',
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
  } catch (error) {
    postError(error, baseContext({ workerEvent: 'worker-message-exception' }));
  }
};

void startup();
