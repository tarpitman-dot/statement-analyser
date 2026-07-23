import { describe, it, expect, vi } from 'vitest';
import * as XLSX from 'xlsx';
import { Readable, Writable } from 'node:stream';
import { mkdtemp, writeFile, access } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import {
  createServer,
  convertUploadedXlsxToCsv,
  handleConvert,
  healthPayload,
} from './converter.mjs';

function bytes(
  sheetName = 'Digital Sales',
  rows = [
    ['Barcode', 'Catalog Number', 'Release Code', 'ISRC', 'Artist', 'Album Title'],
    ['0012345678901', 'CAT-001', 'REL-1', 'ISRC001', 'Artist', 'Album'],
  ],
) {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), sheetName);
  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
}

class CaptureRes extends Writable {
  constructor(origin) {
    super();
    this.headers = {};
    this.statusCode = 0;
    this.body = '';
    this.destroyed = false;
    this.req = { headers: origin ? { origin } : {} };
  }
  writeHead(status, headers) {
    this.statusCode = status;
    this.headers = { ...this.headers, ...headers };
    this.headersSent = true;
  }
  setHeader(k, v) {
    this.headers[k.toLowerCase()] = v;
  }
  addTrailers(headers) {
    this.trailers = headers;
  }
  _write(chunk, _enc, cb) {
    this.body += chunk.toString();
    cb();
  }
}

async function convert(buffer, origin) {
  const dir = await mkdtemp(join(tmpdir(), 'converter-test-'));
  const path = join(dir, 'upload.xlsx');
  await writeFile(path, buffer);
  const res = new CaptureRes(origin);
  await convertUploadedXlsxToCsv(path, res, res.req, dir, buffer.length);
  return res;
}

describe('converter service', () => {
  it('reports converter health metadata', () => {
    expect(healthPayload()).toMatchObject({
      status: 'ok',
      maxUploadBytes: expect.any(Number),
      converterVersion: expect.any(String),
      supportedWorksheet: 'Digital Sales',
    });
  });

  it('handles OPTIONS preflight and allowed origin CORS response', async () => {
    process.env.CONVERTER_ALLOWED_ORIGINS = 'https://statement-analyser-aux.pages.dev';
    const server = createServer().listen(0);
    const url = `http://127.0.0.1:${server.address().port}/api/convert-xlsx`;
    const response = await fetch(url, {
      method: 'OPTIONS',
      headers: { origin: 'https://statement-analyser-aux.pages.dev' },
    });
    server.close();
    expect(response.status).toBe(204);
    expect(response.headers.get('access-control-allow-origin')).toBe(
      'https://statement-analyser-aux.pages.dev',
    );
  });

  it('extracts Digital Sales and preserves shared strings', async () => {
    const res = await convert(bytes('digital sales'));
    expect(res.statusCode).toBe(200);
    expect(res.headers['x-source-worksheet']).toBe('digital sales');
    expect(res.body).toContain('0012345678901,CAT-001');
  });

  it('returns a CORS JSON error when Digital Sales is missing', async () => {
    const other = bytes('Other');
    const req = Readable.from([other]);
    req.headers = { 'content-length': String(other.length), origin: 'https://example.test' };
    const res = new CaptureRes('https://example.test');
    await handleConvert(req, res);
    expect(res.statusCode).toBe(422);
    expect(res.headers['access-control-allow-origin']).toBe('https://example.test');
    expect(res.body).toMatch(/Digital Sales worksheet not found/);
  });

  it('escapes CSV commas, quotes and embedded line breaks', async () => {
    const res = await convert(
      bytes('Digital Sales', [
        ['A', 'B', 'C'],
        ['comma,value', 'quote " value', 'line\nbreak'],
      ]),
    );
    expect(res.body).toContain('"comma,value","quote "" value","line\nbreak"');
  });

  it('preserves blank and missing cells with column alignment', async () => {
    const res = await convert(
      bytes('Digital Sales', [
        ['A', 'B', 'C', 'D'],
        ['left', undefined, 'right', 'end'],
      ]),
    );
    expect(res.body).toContain('left,,right,end');
  });

  it('preserves inline strings', async () => {
    const wb = XLSX.utils.book_new();
    const ws = {
      '!ref': 'A1:A2',
      A1: { t: 'inlineStr', v: 'Header' },
      A2: { t: 'inlineStr', v: 'Inline text' },
    };
    XLSX.utils.book_append_sheet(wb, ws, 'Digital Sales');
    const res = await convert(XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }));
    expect(res.body).toContain('Inline text');
  });

  it('streams date cell serials without crashing', async () => {
    const res = await convert(
      bytes('Digital Sales', [['Date'], [new Date('2026-07-23T00:00:00Z')]]),
    );
    expect(res.body).toMatch(/Date\n\d+/);
  });

  it('cleans up after client disconnect', async () => {
    const buffer = bytes();
    const req = Readable.from([buffer]);
    req.headers = { 'content-length': String(buffer.length) };
    const dir = await mkdtemp(join(tmpdir(), 'cargo-xlsx-'));
    const res = new CaptureRes();
    res.write = () => {
      res.destroyed = true;
      throw new Error('client gone');
    };
    await expect(handleConvert(req, res)).resolves.toBeUndefined();
    await expect(access(dir)).resolves.toBeUndefined();
  });

  it('keeps memory bounded for a generated large worksheet', async () => {
    const rows = [['A', 'B', 'C']];
    for (let i = 0; i < 20000; i += 1) rows.push([`id-${i}`, '', `value-${i}`]);
    const before = process.memoryUsage().rss;
    const res = await convert(bytes('Digital Sales', rows));
    const after = process.memoryUsage().rss;
    expect(res.body.split('\n').length).toBeGreaterThan(20000);
    expect(after - before).toBeLessThan(80 * 1024 * 1024);
  });
});
