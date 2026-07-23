import http from 'node:http';
import { mkdtemp, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import * as XLSX from 'xlsx';

const PORT = Number(process.env.CONVERTER_PORT || 8787);
const MAX_UPLOAD_BYTES = Number(process.env.CONVERTER_MAX_UPLOAD_BYTES || 100 * 1024 * 1024);
const TIMEOUT_MS = Number(process.env.CONVERTER_TIMEOUT_MS || 300000);
const SHEET_NAME = 'Digital Sales';
const VERSION = process.env.CONVERTER_VERSION || 'node-xlsx-2026-07-23';
const ALLOWED_ORIGINS = (process.env.CONVERTER_ALLOWED_ORIGINS || '')
  .split(',')
  .map((origin) => origin.trim())
  .filter(Boolean);

function corsHeaders(req) {
  const origin = req.headers.origin;
  const allowOrigin = !origin
    ? '*'
    : ALLOWED_ORIGINS.length === 0 || ALLOWED_ORIGINS.includes(origin)
      ? origin
      : '';
  return {
    ...(allowOrigin ? { 'access-control-allow-origin': allowOrigin } : {}),
    'access-control-allow-methods': 'GET,POST,OPTIONS',
    'access-control-allow-headers': 'content-type,accept',
    'access-control-expose-headers':
      'x-conversion-started,x-source-worksheet,x-converted-row-count,x-upload-bytes-received,x-converter-cleanup',
    vary: 'Origin',
  };
}

function json(res, status, payload, req) {
  res.writeHead(status, {
    ...corsHeaders(req),
    'content-type': 'application/json; charset=utf-8',
    'cache-control': 'no-store',
  });
  res.end(JSON.stringify(payload));
}

export function healthPayload() {
  return {
    status: 'ok',
    maxUploadBytes: MAX_UPLOAD_BYTES,
    converterVersion: VERSION,
    supportedWorksheet: SHEET_NAME,
  };
}

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
      { statusCode: 422, stage: 'locating-worksheet' },
    );
  }
  const csv = XLSX.utils.sheet_to_csv(workbook.Sheets[sheetName], {
    FS: ',',
    RS: '\n',
    blankrows: false,
  });
  const rowCount = Math.max(0, csv.split('\n').filter(Boolean).length - 1);
  for (const name of workbook.SheetNames) delete workbook.Sheets[name];
  workbook.SheetNames.length = 0;
  return { csv, sheetName, rowCount };
}

async function readBody(req) {
  const contentLength = Number(req.headers['content-length'] || 0);
  if (contentLength > MAX_UPLOAD_BYTES)
    throw Object.assign(new Error('Upload is larger than the configured converter limit.'), {
      statusCode: 413,
      stage: 'receiving-upload',
    });
  const chunks = [];
  let total = 0;
  for await (const chunk of req) {
    total += chunk.length;
    if (total > MAX_UPLOAD_BYTES)
      throw Object.assign(new Error('Upload is larger than the configured converter limit.'), {
        statusCode: 413,
        stage: 'receiving-upload',
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
    const { csv, sheetName, rowCount } = workbookToDigitalSalesCsv(body);
    res.writeHead(200, {
      ...corsHeaders(req),
      'content-type': 'text/csv; charset=utf-8',
      'cache-control': 'no-store',
      'x-converter-cleanup': 'temporary-files-deleted',
      'x-conversion-started': 'true',
      'x-source-worksheet': sheetName,
      'x-converted-row-count': String(rowCount),
      'x-upload-bytes-received': String(body.length),
    });
    res.end(csv);
  } catch (error) {
    const status = error.statusCode || (/timed out/i.test(error.message) ? 504 : 500);
    json(
      res,
      status,
      {
        error: error.message || 'The workbook could not be converted.',
        stage: error.stage || (status === 504 ? 'timeout' : 'conversion'),
        details: status === 413 ? `maxUploadBytes=${MAX_UPLOAD_BYTES}` : String(error.stack || ''),
      },
      req,
    );
  } finally {
    clearTimeout(timer);
    if (dir) await rm(dir, { recursive: true, force: true });
  }
}

export function createServer() {
  return http.createServer((req, res) => {
    const path = new URL(req.url || '/', 'http://converter.local').pathname;
    if (req.method === 'OPTIONS') {
      res.writeHead(204, corsHeaders(req));
      res.end();
    } else if (req.method === 'POST' && (path === '/convert-xlsx' || path === '/api/convert-xlsx'))
      void handleConvert(req, res);
    else if (req.method === 'GET' && (path === '/health' || path === '/api/converter-health'))
      json(res, 200, healthPayload(), req);
    else json(res, 404, { error: 'Route not found.', stage: 'routing', details: path }, req);
  });
}

if (import.meta.url === `file://${process.argv[1]}`)
  createServer().listen(PORT, () => console.log(`converter listening on ${PORT}`));
