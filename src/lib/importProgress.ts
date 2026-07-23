export const IMPORT_STAGES = [
  'Reading file',
  'Uploading workbook',
  'Server received workbook',
  'Converting Digital Sales worksheet',
  'Receiving converted data',
  'Importing CSV',
  'Opening workbook',
  'Detecting worksheets and headers',
  'Importing transactions',
  'Calculating initial summaries',
  'Preparing dashboard',
  'Complete',
] as const;
export type ImportStage = (typeof IMPORT_STAGES)[number];
export const STAGE_RANGES: Record<ImportStage, [number, number]> = {
  'Reading file': [0, 5],
  'Uploading workbook': [5, 55],
  'Server received workbook': [55, 58],
  'Converting Digital Sales worksheet': [58, 72],
  'Receiving converted data': [72, 80],
  'Importing CSV': [80, 99],
  'Opening workbook': [15, 25],
  'Detecting worksheets and headers': [25, 35],
  'Importing transactions': [35, 80],
  'Calculating initial summaries': [80, 93],
  'Preparing dashboard': [93, 99],
  Complete: [100, 100],
};
export interface ImportProgress {
  stage: ImportStage;
  percent: number;
  rowsExamined: number;
  rowsImported: number;
  rowsSkipped: number;
  currentWorksheet?: string;
  message?: string;
  largeFileMode?: boolean;
  bytesRead?: number;
  bytesUploaded?: number;
  totalBytes?: number;
  responseBytes?: number;
}
export function clampProgress(previous: number, stage: ImportStage, fraction = 0) {
  const [min, max] = STAGE_RANGES[stage];
  return Math.max(
    previous,
    Math.min(100, Math.round(min + (max - min) * Math.max(0, Math.min(1, fraction)))),
  );
}
export function isLargeFile(fileSize = 0, estimatedRows = 0) {
  return fileSize > 10 * 1024 * 1024 || estimatedRows > 50000;
}
