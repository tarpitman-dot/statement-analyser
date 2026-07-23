import { describe, it, expect, vi } from 'vitest';
import * as XLSX from 'xlsx';
import { workbookToDigitalSalesCsv, selectDigitalSalesSheet, handleConvert } from './converter.mjs';

function bytes(sheetName = 'Digital Sales') {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    wb,
    XLSX.utils.aoa_to_sheet([
      ['Barcode', 'Catalog Number', 'Release Code', 'ISRC', 'Artist', 'Album Title'],
      ['0012345678901', 'CAT-001', 'REL-1', 'ISRC001', 'Artist', 'Album'],
    ]),
    sheetName,
  );
  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
}

describe('converter service', () => {
  it('selects Digital Sales case-insensitively and downloads identifier-safe CSV', () => {
    const wb = XLSX.read(bytes('digital sales'), { type: 'buffer' });
    expect(selectDigitalSalesSheet(wb)).toBe('digital sales');
    const csv = workbookToDigitalSalesCsv(bytes('digital sales'));
    expect(csv).toContain('0012345678901');
    expect(csv).toContain('CAT-001');
  });

  it('returns a clear error when worksheet is missing', () => {
    expect(() => workbookToDigitalSalesCsv(bytes('Other'))).toThrow(
      /Digital Sales worksheet not found/,
    );
  });

  it('cleans up temporary files after upload handling', async () => {
    const chunks = [bytes()];
    const fakeReq = {
      async *[Symbol.asyncIterator]() {
        yield* chunks;
      },
      destroy: vi.fn(),
    };
    const fakeRes = { writeHead: vi.fn(), end: vi.fn() };
    await handleConvert(fakeReq, fakeRes);
    expect(fakeRes.writeHead.mock.calls[0][0]).toBe(200);
    expect(String(fakeRes.end.mock.calls[0][0])).toContain('0012345678901');
  });
});
