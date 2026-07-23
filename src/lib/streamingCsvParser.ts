import * as XLSX from 'xlsx';
import { parseWorkbook, type ParseOptions } from './parser';

export async function parseCsvFileStreaming(file: File, opts: ParseOptions = {}) {
  const rows: string[][] = [];
  let current: string[] = [];
  let field = '';
  let quoted = false;
  let examined = 0;
  const decoder = new TextDecoder('utf-8');
  const reader =
    typeof file.stream === 'function'
      ? file.stream().getReader()
      : new Response(await file.arrayBuffer()).body!.getReader();
  const pushField = () => {
    current.push(field);
    field = '';
  };
  const pushRow = () => {
    if (current.length || field) {
      pushField();
      rows.push(current);
      examined++;
      if (examined % 1000 === 0)
        opts.onProgress?.({
          stage: 'Importing transactions',
          percent: Math.min(79, 35 + Math.round(examined / 1000)),
          rowsExamined: examined,
          rowsImported: Math.max(0, examined - 1),
          rowsSkipped: 0,
          largeFileMode: true,
          message: 'Streaming CSV rows',
        });
    }
    current = [];
  };
  let carry = '';
  while (true) {
    const { value, done } = await reader.read();
    const chunk = carry + (value ? decoder.decode(value, { stream: !done }) : '');
    carry = '';
    for (let i = 0; i < chunk.length; i++) {
      const ch = chunk[i];
      if (quoted) {
        if (ch === '"') {
          if (chunk[i + 1] === '"') {
            field += '"';
            i++;
          } else quoted = false;
        } else field += ch;
      } else if (ch === '"') quoted = true;
      else if (ch === ',') pushField();
      else if (ch === '\n') pushRow();
      else if (ch !== '\r') field += ch;
    }
    if (done) break;
  }
  if (field || current.length) pushRow();
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), 'Digital Sales');
  return parseWorkbook(wb, file.name, file.size, {
    ...opts,
    largeFileMode: true,
    estimatedRows: rows.length,
  });
}
