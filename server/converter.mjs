import http from 'node:http';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import * as XLSX from 'xlsx';

const PORT = Number(process.env.CONVERTER_PORT || 8787);
const MAX_UPLOAD_BYTES = Number(process.env.CONVERTER_MAX_UPLOAD_BYTES || 75 * 1024 * 1024);
const TIMEOUT_MS = Number(process.env.CONVERTER_TIMEOUT_MS || 120000);
const SHEET_NAME = 'Digital Sales';

export function selectDigitalSalesSheet(workbook) {
  return workbook.SheetNames.find((name) => name.toLowerCase() === SHEET_NAME.toLowerCase());
}

export function workbookToDigitalSalesCsv(buffer) {
  const workbook = XLSX.read(buffer, {
    type: 'buffer',
    raw: false,
    cellDates: false,
    cellStyles: false,
    cellHTML: false,
  });
  const sheetName = selectDigitalSalesSheet(workbook);
  if (!sheetName) {
    const available = workbook.SheetNames.join(', ') || 'none';
    throw Object.assign(
      new Error(`Digital Sales worksheet not found. Available worksheets: ${available}.`),
      { statusCode: 422 },
    );
  }
  const csv = XLSX.utils.sheet_to_csv(workbook.Sheets[sheetName], {
    FS: ',',
    RS: '\n',
    blankrows: false,
  });
  for (const name of workbook.SheetNames) delete workbook.Sheets[name];
  workbook.SheetNames.length = 0;
  return csv;
}

async function readBody(req) {
  const chunks = [];
  let total = 0;
  for await (const chunk of req) {
    total += chunk.length;
    if (total > MAX_UPLOAD_BYTES)
      throw Object.assign(new Error('Upload is larger than the configured converter limit.'), {
        statusCode: 413,
      });
    chunks.push(chunk);
  }
  return Buffer.concat(chunks);
}

export async function handleConvert(req, res) {
  let dir;
  const timer = setTimeout(() => req.destroy(new Error('Conversion timed out.')), TIMEOUT_MS);
  try {
    dir = await mkdtemp(join(tmpdir(), 'cargo-xlsx-'));
    const body = await readBody(req);
    const uploadPath = join(dir, 'upload.xlsx');
    await writeFile(uploadPath, body);
    const csv = workbookToDigitalSalesCsv(body);
    res.writeHead(200, {
      'content-type': 'text/csv; charset=utf-8',
      'cache-control': 'no-store',
      'x-converter-cleanup': 'temporary-files-deleted',
      'x-conversion-started': 'true',
      'x-upload-bytes-received': String(body.length),
    });
    res.end(csv);
  } catch (error) {
    const status = error.statusCode || (/timed out/i.test(error.message) ? 504 : 400);
    res.writeHead(status, {
      'content-type': 'application/json; charset=utf-8',
      'cache-control': 'no-store',
    });
    res.end(JSON.stringify({ error: error.message || 'The workbook could not be converted.' }));
  } finally {
    clearTimeout(timer);
    if (dir) await rm(dir, { recursive: true, force: true });
  }
}

export function createServer() {
  return http.createServer((req, res) => {
    if (req.method === 'POST' && req.url === '/convert-xlsx') void handleConvert(req, res);
    else if (req.url === '/health') {
      res.writeHead(200);
      res.end('ok');
    } else {
      res.writeHead(404);
      res.end('not found');
    }
  });
}

if (import.meta.url === `file://${process.argv[1]}`)
  createServer().listen(PORT, () => console.log(`converter listening on ${PORT}`));
