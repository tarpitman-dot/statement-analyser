import type { ImportProgress, ImportStage } from './importProgress';

function formatBytes(size: number) {
  return size > 1024 * 1024
    ? `${(size / 1024 / 1024).toFixed(1)} MB`
    : `${(size / 1024).toFixed(1)} KB`;
}

export interface ActiveXhrConversion {
  xhr: XMLHttpRequest | null;
}

export type ConverterUploadProgress = (
  stage: ImportStage,
  fraction: number,
  message: string,
  extra?: Partial<ImportProgress>,
) => void;

export function uploadWorkbookWithProgress(options: {
  file: File;
  converterUrl: string;
  activeConversion: ActiveXhrConversion;
  diagnostics: Record<string, unknown>;
  onProgress: ConverterUploadProgress;
  uploadInactivityMs?: number;
  responseInactivityMs?: number;
}) {
  const {
    file: f,
    converterUrl,
    activeConversion,
    diagnostics,
    onProgress,
    uploadInactivityMs = 30000,
    responseInactivityMs = 120000,
  } = options;

  return new Promise<void>((resolve, reject) => {
    const xhr = new XMLHttpRequest();
    activeConversion.xhr = xhr;
    const started = Date.now();
    let lastUploaded = 0;
    let lastResponseLoaded = 0;
    let settled = false;
    let inactivityTimer = 0;

    const settle = (callback: () => void) => {
      if (settled) return;
      settled = true;
      window.clearTimeout(inactivityTimer);
      callback();
    };

    const resetInactivity = (
      reason: 'upload-inactivity-timeout' | 'converter-response-timeout',
    ) => {
      window.clearTimeout(inactivityTimer);
      inactivityTimer = window.setTimeout(
        () => {
          diagnostics.abortReason = reason;
          settle(() => {
            xhr.abort();
            reject(
              reason === 'upload-inactivity-timeout'
                ? new Error(
                    'No upload progress was reported for 30 seconds. The request was aborted so you can retry.',
                  )
                : new Error('The converter response timed out.'),
            );
          });
        },
        reason === 'upload-inactivity-timeout' ? uploadInactivityMs : responseInactivityMs,
      );
    };

    xhr.upload.onprogress = (e) => {
      const uploaded = e.loaded || lastUploaded;
      if (uploaded > lastUploaded) resetInactivity('upload-inactivity-timeout');
      lastUploaded = uploaded;
      diagnostics.bytesUploaded = uploaded;
      diagnostics.elapsedUploadMs = Date.now() - started;
      onProgress(
        'Uploading workbook',
        f.size ? uploaded / f.size : 1,
        `Uploading workbook (${formatBytes(uploaded)} of ${formatBytes(f.size)})`,
        { bytesUploaded: uploaded },
      );
    };
    xhr.upload.onload = () => {
      diagnostics.bytesUploaded = f.size;
      diagnostics.elapsedUploadMs = Date.now() - started;
      onProgress('Server received workbook', 1, 'Server received workbook', {
        bytesUploaded: f.size,
      });
      resetInactivity('converter-response-timeout');
    };
    xhr.onreadystatechange = () => {
      if (
        xhr.readyState >= XMLHttpRequest.HEADERS_RECEIVED &&
        !diagnostics.conversionStartReceived
      ) {
        diagnostics.httpStatus = xhr.status;
        diagnostics.conversionStartReceived =
          xhr.getResponseHeader('x-conversion-started') === 'true';
        resetInactivity('converter-response-timeout');
        onProgress('Converting Digital Sales worksheet', 0.2, 'Converting Digital Sales worksheet');
      }
    };
    xhr.onprogress = (e) => {
      if (xhr.status) diagnostics.httpStatus = xhr.status;
      const loaded = e.loaded || 0;
      if (loaded > lastResponseLoaded) resetInactivity('converter-response-timeout');
      lastResponseLoaded = loaded;
      diagnostics.responseSize = loaded;
      onProgress(
        'Receiving converted data',
        e.lengthComputable ? loaded / e.total : 0.5,
        `Receiving converted data (${formatBytes(loaded)})`,
        { responseBytes: loaded },
      );
    };
    xhr.onload = () => {
      diagnostics.httpStatus = xhr.status;
      diagnostics.responseSize = xhr.response?.size ?? 0;
      if (xhr.status < 200 || xhr.status >= 300) {
        settle(() => {
          if (xhr.response?.size) {
            xhr.response
              .text()
              .then((text: string) => {
                try {
                  const payload = JSON.parse(text) as {
                    error?: string;
                    stage?: string;
                    details?: string;
                  };
                  reject(
                    new Error(
                      `${payload.error || `Converter HTTP ${xhr.status}`} (stage: ${payload.stage || 'unknown'})${payload.details ? ` Details: ${payload.details}` : ''}`,
                    ),
                  );
                } catch {
                  reject(new Error(`Converter HTTP ${xhr.status}: ${text.slice(0, 500)}`));
                }
              })
              .catch(() => reject(new Error(`Converter HTTP ${xhr.status}`)));
          } else reject(new Error(`Converter HTTP ${xhr.status}`));
        });
        return;
      }
      diagnostics.csvBlob = xhr.response;
      onProgress(
        'Receiving converted data',
        1,
        `Receiving converted data (${formatBytes(xhr.response?.size ?? 0)})`,
        {
          responseBytes: xhr.response?.size ?? 0,
        },
      );
      settle(resolve);
    };
    xhr.onerror = () => {
      diagnostics.httpStatus = xhr.status || 0;
      diagnostics.abortReason = 'network-failure';
      settle(() =>
        reject(
          new Error(
            'Network failure: the browser could not reach the converter. Check the API route, proxy, CORS and converter health endpoint.',
          ),
        ),
      );
    };
    xhr.ontimeout = () => {
      diagnostics.abortReason = 'converter-response-timeout';
      settle(() => reject(new Error('The converter response timed out.')));
    };
    xhr.onabort = () =>
      settle(() => reject(new Error(String(diagnostics.abortReason || 'conversion-cancelled'))));
    xhr.open('POST', converterUrl);
    xhr.responseType = 'blob';
    xhr.setRequestHeader(
      'content-type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    );
    resetInactivity('upload-inactivity-timeout');
    xhr.send(f);
  });
}
