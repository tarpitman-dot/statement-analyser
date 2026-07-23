export const LARGE_XLSX_CONVERTER_URL =
  import.meta.env.VITE_LARGE_XLSX_CONVERTER_URL || '/api/convert-xlsx';
export const LARGE_XLSX_BROWSER_SIZE_LIMIT = Number(
  import.meta.env.VITE_LARGE_XLSX_BROWSER_SIZE_LIMIT || 10 * 1024 * 1024,
);
export const LARGE_XLSX_BROWSER_ROW_LIMIT = Number(
  import.meta.env.VITE_LARGE_XLSX_BROWSER_ROW_LIMIT || 50000,
);

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
