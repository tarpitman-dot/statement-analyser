export const LARGE_XLSX_CONVERTER_URL =
  import.meta.env.VITE_CONVERTER_API_URL ||
  import.meta.env.VITE_LARGE_XLSX_CONVERTER_URL ||
  '/api/convert-xlsx';
export const LARGE_XLSX_CONVERTER_HEALTH_URL =
  import.meta.env.VITE_CONVERTER_HEALTH_URL || deriveHealthUrl(LARGE_XLSX_CONVERTER_URL);
export const LARGE_XLSX_BROWSER_SIZE_LIMIT = Number(
  import.meta.env.VITE_LARGE_XLSX_BROWSER_SIZE_LIMIT || 10 * 1024 * 1024,
);
export const LARGE_XLSX_BROWSER_ROW_LIMIT = Number(
  import.meta.env.VITE_LARGE_XLSX_BROWSER_ROW_LIMIT || 50000,
);

export type ConverterHealth = {
  status: 'ok';
  maxUploadBytes: number;
  converterVersion: string;
  supportedWorksheet: string;
};

function deriveHealthUrl(convertUrl: string) {
  if (/\/convert-xlsx(?:$|[?#])/.test(convertUrl))
    return convertUrl.replace(/\/convert-xlsx(?=$|[?#])/, '/converter-health');
  return '/api/converter-health';
}

export async function checkConverterHealth(signal?: AbortSignal): Promise<ConverterHealth> {
  const response = await fetch(LARGE_XLSX_CONVERTER_HEALTH_URL, {
    method: 'GET',
    headers: { accept: 'application/json' },
    cache: 'no-store',
    signal,
  });
  if (!response.ok) throw new Error(`Converter health check failed with HTTP ${response.status}.`);
  const health = (await response.json()) as Partial<ConverterHealth>;
  if (health.status !== 'ok' || typeof health.maxUploadBytes !== 'number')
    throw new Error('Converter health check returned an invalid response.');
  return health as ConverterHealth;
}

export type LargeXlsxDecision = {
  shouldConvert: boolean;
  reason?: 'size' | 'rows';
  estimatedRows?: number;
};

export function isXlsxFilename(name: string) {
  return /\.xlsx$/i.test(name);
}

export function shouldOfferLargeXlsxConversion(
  file: { name: string; size: number },
  estimatedRows?: number,
): LargeXlsxDecision {
  if (!isXlsxFilename(file.name)) return { shouldConvert: false, estimatedRows };
  if ((estimatedRows ?? 0) > LARGE_XLSX_BROWSER_ROW_LIMIT)
    return { shouldConvert: true, reason: 'rows', estimatedRows };
  if (file.size > LARGE_XLSX_BROWSER_SIZE_LIMIT)
    return { shouldConvert: true, reason: 'size', estimatedRows };
  return { shouldConvert: false, estimatedRows };
}
