import { CONSERVATIVE_SHEETJS_READ_OPTIONS, parseArrayBuffer } from './lib/parser';
import { serialiseThrown, technicalDetails, type ImportDebugContext } from './lib/importDiagnostics';

function postError(error: unknown, context: ImportDebugContext) {
  const serialised = serialiseThrown(error);
  (self as unknown as Worker).postMessage({
    type: 'error',
    error: {
      name: serialised.name,
      message: serialised.message || 'This statement could not be fully processed.',
      stack: serialised.stack,
      details: technicalDetails(context, error),
      context,
    },
  });
}

self.onmessage = async (e: MessageEvent) => {
  const payload = e.data;
  let { buffer } = payload;
  const { filename, fileSize, arrayBufferSizeBeforeTransfer, fileReadMs } = payload;
  const base = {
    processingStage: 'Opening workbook',
    workerEvent: 'parse-start',
    workbookFilename: filename ?? 'Unknown',
    fileSize: Number(fileSize ?? 0),
    arrayBufferSizeBeforeTransfer: Number(arrayBufferSizeBeforeTransfer ?? 0),
    arrayBufferSizeInsideWorker: buffer?.byteLength ?? 0,
  };
  if (!(buffer instanceof ArrayBuffer) || buffer.byteLength === 0) {
    postError(new Error('Worker received an empty or invalid ArrayBuffer.'), { ...base, sheetJsReadAttempt: 'not attempted', retryAttemptNumber: 0 });
    return;
  }
  for (const attempt of [1, 2]) {
    try {
      const data = await parseArrayBuffer(buffer, filename, fileSize, {
        fileReadMs,
        sheetJsReadAttempt: attempt === 1 ? 'default' : 'conservative',
        retryAttemptNumber: attempt,
        readOptions: attempt === 1 ? undefined : CONSERVATIVE_SHEETJS_READ_OPTIONS,
        onProgress: p => (self as unknown as Worker).postMessage({ type: 'progress', progress: p, context: { ...base, sheetJsReadAttempt: attempt === 1 ? 'default' : 'conservative', retryAttemptNumber: attempt } }),
      });
      data.diagnostics.memory={...(data.diagnostics.memory??{}),workerTerminated:'No'};
      buffer = null;
      if (payload) payload.buffer = null;
      (self as unknown as Worker).postMessage({ type: 'complete', data, context: { ...base, processingStage: 'Complete', workerEvent: 'complete', sheetJsReadAttempt: attempt === 1 ? 'default' : 'conservative', retryAttemptNumber: attempt } });
      return;
    } catch (error) {
      if (attempt === 2) {
        buffer = null;
        if (payload) payload.buffer = null;
        postError(error, { ...base, workerEvent: 'error', sheetJsReadAttempt: 'conservative', retryAttemptNumber: 2 });
        return;
      }
    }
  }
};
