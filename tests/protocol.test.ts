import assert from 'node:assert/strict';
import test from 'node:test';
import { isSelectionContext } from '../packages/protocol/src/index.js';

const selection = {
  host: 'wps',
  connectionId: 'test-session',
  workbookName: 'book.xlsx',
  sheetName: 'Sheet1',
  address: 'A1:B2',
  rowCount: 2,
  columnCount: 2,
  sampleValues: [
    ['机构', '数量'],
    ['测试', 3],
  ],
  truncated: false,
  capturedAt: '2026-10-19T00:00:00.000Z',
};

test('accepts bounded selection data and requires host identity', () => {
  assert.equal(isSelectionContext(selection), true);
  assert.equal(isSelectionContext({ ...selection, connectionId: '' }), false);
  assert.equal(isSelectionContext({ ...selection, sheetName: '' }), false);
});

test('rejects malformed cells, mismatched dimensions and invalid timestamps', () => {
  assert.equal(isSelectionContext(null), false);
  assert.equal(
    isSelectionContext({ ...selection, sampleValues: [[{ value: 3 }]] }),
    false,
  );
  assert.equal(
    isSelectionContext({
      ...selection,
      sampleValues: [
        [1, Number.NaN],
        [2, 3],
      ],
    }),
    false,
  );
  assert.equal(isSelectionContext({ ...selection, rowCount: 3 }), false);
  assert.equal(
    isSelectionContext({ ...selection, capturedAt: 'invalid' }),
    false,
  );
});

test('large selections require explicitly truncated, bounded samples', () => {
  const sampleValues = Array.from({ length: 20 }, () =>
    Array.from({ length: 15 }, () => 1),
  );
  const large = {
    ...selection,
    rowCount: 1000,
    columnCount: 30,
    sampleValues,
    truncated: true,
  };
  assert.equal(isSelectionContext(large), true);
  assert.equal(isSelectionContext({ ...large, truncated: false }), false);
  assert.equal(
    isSelectionContext({
      ...large,
      sampleValues: [...sampleValues, sampleValues[0]],
    }),
    false,
  );
});
