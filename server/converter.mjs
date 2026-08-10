import http from 'node:http';
import { createReadStream, createWriteStream } from 'node:fs';
import { mkdtemp, open, rm, stat } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { basename, join, normalize, posix } from 'node:path';
import { pipeline } from 'node:stream/promises';
import { Transform } from 'node:stream';
import { createInflateRaw } from 'node:zlib';

const PORT = Number(process.env.CONVERTER_PORT || 8787);
const MAX_UPLOAD_BYTES = Number(process.env.CONVERTER_MAX_UPLOAD_BYTES || 100 * 1024 * 1024);
const TIMEOUT_MS = Number(process.env.CONVERTER_TIMEOUT_MS || 300000);
const SHEET_NAME = 'Digital Sales';
const VERSION = process.env.CONVERTER_VERSION || 'node-xlsx-2026-07-23';
const ALLOWED_ORIGINS = (process.env.CONVERTER_ALLOWED_ORIGINS || '')
  .split(',')
  .map((origin) => origin.trim())
  .filter(Boolean);
const EXPOSED_HEADERS =
  'x-conversion-started,x-source-worksheet,x-converted-row-count,x-upload-bytes-received,x-converter-cleanup';

function log(event, details = {}) {
  console.log(JSON.stringify({ time: new Date().toISOString(), event, ...details }));
}

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
    'access-control-expose-headers': EXPOSED_HEADERS,
    vary: 'Origin',
  };
}

function json(res, status, payload, req, extraHeaders = {}) {
  if (res.headersSent || res.destroyed) return;
  res.writeHead(status, {
    ...corsHeaders(req),
    ...extraHeaders,
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

function decodeXml(text) {
  return text
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&')
    .replace(/&#x([0-9A-Fa-f]+);/g, (_, hex) => String.fromCodePoint(parseInt(hex, 16)))
    .replace(/&#(\d+);/g, (_, dec) => String.fromCodePoint(Number(dec)))
    .replace(/_x([0-9A-Fa-f]{4})_/g, (_, hex) => String.fromCharCode(parseInt(hex, 16)));
}

function attrs(tag) {
  const out = {};
  for (const match of tag.matchAll(/([A-Za-z_:][\w:.-]*)\s*=\s*"([^"]*)"/g))
    out[match[1]] = decodeXml(match[2]);
  return out;
}

function csvEscape(value) {
  const text = value == null ? '' : String(value);
  return /[",\r\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

// Cell <v> contents are already the workbook's underlying value. Keep them as text: converting
// through JavaScript numbers (or using formatted <w> display values) corrupts identifiers.
function exactCellValue(valueXml) {
  return decodeXml(valueXml);
}

function columnIndex(ref = '') {
  const letters = (ref.match(/^[A-Z]+/i)?.[0] || '').toUpperCase();
  let n = 0;
  for (const ch of letters) n = n * 26 + ch.charCodeAt(0) - 64;
  return Math.max(0, n - 1);
}

function resolveTarget(base, target) {
  return normalize(posix.join(posix.dirname(base), target))
    .replace(/^\\?/, '')
    .replace(/\\/g, '/');
}

async function readTail(path, maxBytes) {
  const size = (await stat(path)).size;
  const length = Math.min(size, maxBytes);
  const fh = await open(path, 'r');
  try {
    const buffer = Buffer.alloc(length);
    await fh.read(buffer, 0, length, size - length);
    return { buffer, offset: size - length, size };
  } finally {
    await fh.close();
  }
}

async function openZip(path) {
  const { buffer, offset } = await readTail(path, 1024 * 1024);
  const eocd = buffer.lastIndexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06]));
  if (eocd < 0)
    throw Object.assign(new Error('Invalid XLSX ZIP: end record not found.'), {
      statusCode: 422,
      stage: 'opening-zip',
    });
  const total = buffer.readUInt16LE(eocd + 10);
  const cdSize = buffer.readUInt32LE(eocd + 12);
  const cdOffset = buffer.readUInt32LE(eocd + 16);
  const fh = await open(path, 'r');
  const cd = Buffer.alloc(cdSize);
  await fh.read(cd, 0, cdSize, cdOffset);
  const entries = new Map();
  let p = 0;
  for (let i = 0; i < total; i += 1) {
    if (cd.readUInt32LE(p) !== 0x02014b50) break;
    const method = cd.readUInt16LE(p + 10);
    const compressedSize = cd.readUInt32LE(p + 20);
    const uncompressedSize = cd.readUInt32LE(p + 24);
    const nameLen = cd.readUInt16LE(p + 28);
    const extraLen = cd.readUInt16LE(p + 30);
    const commentLen = cd.readUInt16LE(p + 32);
    const localOffset = cd.readUInt32LE(p + 42);
    const name = cd.subarray(p + 46, p + 46 + nameLen).toString('utf8');
    entries.set(name, { name, method, compressedSize, uncompressedSize, localOffset });
    p += 46 + nameLen + extraLen + commentLen;
  }
  async function stream(name) {
    const entry = entries.get(name);
    if (!entry) return null;
    const header = Buffer.alloc(30);
    await fh.read(header, 0, 30, entry.localOffset);
    const nameLen = header.readUInt16LE(26);
    const extraLen = header.readUInt16LE(28);
    const start = entry.localOffset + 30 + nameLen + extraLen;
    const raw = createReadStream(path, { start, end: start + entry.compressedSize - 1 });
    return entry.method === 8 ? raw.pipe(createInflateRaw()) : raw;
  }
  return { entries, stream, close: () => fh.close() };
}

async function streamToString(zip, name, max = 8 * 1024 * 1024) {
  const s = await zip.stream(name);
  if (!s) return '';
  let total = 0;
  const chunks = [];
  for await (const chunk of s) {
    total += chunk.length;
    if (total > max)
      throw Object.assign(new Error(`${name} is larger than supported metadata limit.`), {
        statusCode: 422,
        stage: 'reading-metadata',
      });
    chunks.push(chunk);
  }
  return Buffer.concat(chunks).toString('utf8');
}

async function findWorksheet(zip) {
  const workbookXml = await streamToString(zip, 'xl/workbook.xml');
  const relsXml = await streamToString(zip, 'xl/_rels/workbook.xml.rels');
  const rels = new Map();
  for (const m of relsXml.matchAll(/<Relationship\b[^>]*>/g)) {
    const a = attrs(m[0]);
    if (a.Id && a.Target) rels.set(a.Id, resolveTarget('xl/workbook.xml', a.Target));
  }
  for (const m of workbookXml.matchAll(/<sheet\b[^>]*>/g)) {
    const a = attrs(m[0]);
    if (a.name?.toLowerCase() === SHEET_NAME.toLowerCase()) {
      const target = rels.get(a['r:id']);
      if (!target) break;
      log('worksheet found', { sheetName: a.name, target });
      return { sheetName: a.name, target };
    }
  }
  const available =
    [...workbookXml.matchAll(/<sheet\b[^>]*>/g)]
      .map((m) => attrs(m[0]).name)
      .filter(Boolean)
      .join(', ') || 'none';
  throw Object.assign(
    new Error(`Digital Sales worksheet not found. Available worksheets: ${available}.`),
    { statusCode: 422, stage: 'locating-worksheet' },
  );
}

async function buildSharedStrings(zip, dir) {
  if (!zip.entries.has('xl/sharedStrings.xml')) return null;
  const path = join(dir, 'shared-strings.jsonl');
  const offsets = [];
  let pos = 0;
  const out = createWriteStream(path);
  let carry = '';
  const s = await zip.stream('xl/sharedStrings.xml');
  for await (const chunk of s) {
    carry += chunk.toString('utf8');
    let end;
    while ((end = carry.indexOf('</si>')) >= 0) {
      const block = carry.slice(0, end + 5);
      carry = carry.slice(end + 5);
      const parts = [...block.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map((m) => decodeXml(m[1]));
      const line = JSON.stringify(parts.join('')) + '\n';
      offsets.push(pos);
      pos += Buffer.byteLength(line);
      if (!out.write(line)) await new Promise((r) => out.once('drain', r));
    }
    if (carry.length > 1024 * 1024) carry = carry.slice(-1024 * 1024);
  }
  await new Promise((r) => out.end(r));
  return { path, offsets };
}

async function sharedAt(shared, index) {
  if (!shared) return '';
  const offset = shared.offsets[index];
  if (offset == null) return '';
  const next = shared.offsets[index + 1];
  const fh = await open(shared.path, 'r');
  try {
    const len = (next ?? (await stat(shared.path)).size) - offset;
    const buf = Buffer.alloc(len);
    await fh.read(buf, 0, len, offset);
    return JSON.parse(buf.toString('utf8'));
  } finally {
    await fh.close();
  }
}

async function writeResponse(res, text) {
  if (!res.write(text))
    await new Promise((resolve, reject) => {
      res.once('drain', resolve);
      res.once('error', reject);
    });
}

async function streamWorksheetCsv(zip, worksheetPath, shared, res) {
  let carry = '';
  let row = [];
  let emitted = 0;
  const s = await zip.stream(worksheetPath);
  if (!s)
    throw Object.assign(new Error('Target worksheet XML was not found in the XLSX archive.'), {
      statusCode: 422,
      stage: 'locating-worksheet',
    });
  for await (const chunk of s) {
    carry += chunk.toString('utf8');
    let end;
    while ((end = carry.indexOf('</row>')) >= 0) {
      const block = carry.slice(0, end + 6);
      carry = carry.slice(end + 6);
      row = [];
      for (const m of block.matchAll(/<c\b([^>]*)>([\s\S]*?)<\/c>|<c\b([^>]*)\/>/g)) {
        const a = attrs(`<c ${m[1] || m[3] || ''}>`);
        const col = columnIndex(a.r);
        while (row.length < col) row.push('');
        const body = m[2] || '';
        const v = body.match(/<v\b[^>]*>([\s\S]*?)<\/v>/)?.[1] ?? '';
        const inline = [...body.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)]
          .map((x) => decodeXml(x[1]))
          .join('');
        let value = '';
        if (a.t === 's') value = await sharedAt(shared, Number(v));
        else if (a.t === 'inlineStr') value = inline;
        else if (a.t === 'b') value = v === '1' ? 'TRUE' : v === '0' ? 'FALSE' : v;
        else value = exactCellValue(v);
        row[col] = value;
      }
      await writeResponse(res, `${row.map(csvEscape).join(',')}\n`);
      emitted += 1;
      if (emitted % 10000 === 0)
        log('rows emitted', { rows: emitted, memory: process.memoryUsage().rss });
    }
    if (carry.length > 1024 * 1024) carry = carry.slice(-1024 * 1024);
  }
  log('rows emitted', { rows: emitted, memory: process.memoryUsage().rss });
  return emitted > 0 ? emitted - 1 : 0;
}

async function writeUploadToTemp(req, uploadPath) {
  const contentLength = Number(req.headers['content-length'] || 0);
  if (contentLength > MAX_UPLOAD_BYTES)
    throw Object.assign(new Error('Upload is larger than the configured converter limit.'), {
      statusCode: 413,
      stage: 'receiving-upload',
    });
  let total = 0;
  log('upload started', { contentLength });
  const counter = new Transform({
    transform(chunk, _enc, cb) {
      total += chunk.length;
      if (total > MAX_UPLOAD_BYTES)
        cb(
          Object.assign(new Error('Upload is larger than the configured converter limit.'), {
            statusCode: 413,
            stage: 'receiving-upload',
          }),
        );
      else cb(null, chunk);
    },
  });
  await pipeline(req, counter, createWriteStream(uploadPath));
  log('upload completed', { bytes: total });
  log('total bytes received', { bytes: total });
  return total;
}

export async function convertUploadedXlsxToCsv(uploadPath, res, req, dir, bytesReceived = 0) {
  const zip = await openZip(uploadPath);
  try {
    const { sheetName, target } = await findWorksheet(zip);
    const shared = await buildSharedStrings(zip, dir);
    res.writeHead(200, {
      ...corsHeaders(req),
      'content-type': 'text/csv; charset=utf-8',
      'cache-control': 'no-store',
      'x-converter-cleanup': 'temporary-files-deleted',
      'x-conversion-started': 'true',
      'x-source-worksheet': sheetName,
      'x-upload-bytes-received': String(bytesReceived),
    });
    log('conversion started', { sheetName, target, memory: process.memoryUsage().rss });
    const rowCount = await streamWorksheetCsv(zip, target, shared, res);
    res.addTrailers?.({ 'x-converted-row-count': String(rowCount) });
    res.end();
    return { sheetName, rowCount };
  } finally {
    await zip.close();
  }
}

export async function handleConvert(req, res) {
  let dir;
  let bytesReceived = 0;
  let disconnected = false;
  const timer = setTimeout(() => req.destroy(new Error('Conversion timed out.')), TIMEOUT_MS);
  req.on?.('close', () => {
    if (!res.writableEnded) {
      disconnected = true;
      log('client disconnected', { bytesReceived });
    }
  });
  try {
    dir = await mkdtemp(join(tmpdir(), 'cargo-xlsx-'));
    const uploadPath = join(dir, 'upload.xlsx');
    bytesReceived = await writeUploadToTemp(req, uploadPath);
    await convertUploadedXlsxToCsv(uploadPath, res, req, dir, bytesReceived);
  } catch (error) {
    log('conversion error', { message: error.message, stage: error.stage, stack: error.stack });
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
      {
        'x-converter-cleanup': 'temporary-files-deleted',
        'x-conversion-started': res.headersSent ? 'true' : 'false',
        'x-upload-bytes-received': String(bytesReceived),
      },
    );
    if (!res.writableEnded && !disconnected) res.end();
  } finally {
    clearTimeout(timer);
    if (dir) await rm(dir, { recursive: true, force: true });
    log('cleanup completed', { dir: dir ? basename(dir) : null });
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
