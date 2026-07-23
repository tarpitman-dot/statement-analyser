import React from 'react';
import { describe, it, expect } from 'vitest';
import { render, screen, within, waitFor } from '@testing-library/react';
import '@testing-library/jest-dom/vitest';
import { Overview } from './main';
import type { Transaction } from './lib/types';

const base: Transaction = {
  sourceSheet: 'Digital Sales',
  sourceRow: 2,
  contractId: 'C1',
  contractName: 'Contract',
  assetType: 'Track',
  releaseCode: 'R1',
  albumTitle: 'A Very Long Release Title That Should Wrap Fully Instead Of Being Truncated',
  catalogNumber: 'CAT1',
  barcode: '1234567890123',
  isrc: 'ISRC1',
  artist: 'A Very Long Artist Name That Should Be Readable In Full',
  trackTitle: 'Track',
  usageType: 'Track Stream',
  country: 'United Kingdom',
  shop: 'A Very Long Shop Name That Should Not Be Truncated',
  salesPeriod: '2026-06',
  sales: '10',
  returns: '0',
  royaltyRate: '0.85',
  amount: '10',
  deduction1: '',
  deduction2: '',
  deduction3: '',
  contractDeductions: '',
  deduction4: '',
  deduction5: '',
  lineCharges: '',
  shareContract: '',
  ppu: '',
  share: '',
  rata1: '',
  rata2: '',
  royaltyAmount: '8',
};

function tx(overrides: Partial<Transaction>): Transaction {
  return { ...base, ...overrides };
}

describe('compact overview', () => {
  it('shows full names and expected ranked rows in compact overview tables', async () => {
    render(
      <Overview
        rows={[
          tx({ sourceRow: 2, barcode: '1000000000001', albumTitle: 'Long Release Winner Name Visible In Full', artist: 'Long Artist Winner Name Visible In Full', shop: 'Shop Winner Full Name', country: 'United Kingdom', usageType: 'Download', royaltyAmount: '30', sales: '3' }),
          tx({ sourceRow: 3, barcode: '1000000000002', albumTitle: 'Runner Up Release Full Name', artist: 'Runner Up Artist Full Name', shop: 'Shop Runner Full Name', country: 'United States', usageType: 'Stream', royaltyAmount: '20', sales: '2' }),
          tx({ sourceRow: 4, barcode: '1000000000003', albumTitle: 'Third Release Full Name', artist: 'Third Artist Full Name', shop: 'Shop Third Full Name', country: 'Germany', usageType: 'Stream', royaltyAmount: '10', sales: '1' }),
        ]}
      />,
    );

    const releases = await screen.findByRole('region', { name: 'Top releases' });
    await waitFor(() => expect(releases).toHaveTextContent('Long Release Winner Name Visible In Full'));
    expect(releases).toHaveTextContent('Long Artist Winner Name Visible In Full');
    const releaseRows = within(releases).getAllByRole('row');
    expect(releaseRows[1]).toHaveTextContent('1');
    expect(releaseRows[1]).toHaveTextContent('£30.00');
    expect(releaseRows[2]).toHaveTextContent('Runner Up Release Full Name');
    expect(screen.getByRole('region', { name: 'Top artists' })).toHaveTextContent('Long Artist Winner Name Visible In Full');
    expect(screen.getByRole('region', { name: 'Top shops' })).toHaveTextContent('Shop Winner Full Name');
    expect(screen.getByRole('region', { name: 'Usage types' })).toHaveTextContent('Stream');
    expect(screen.getByRole('region', { name: 'Countries' })).toHaveTextContent('United Kingdom');
  });
});
