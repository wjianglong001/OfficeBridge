// Synthetic host contract test, not WPS integration evidence.
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { createContext, runInContext } from 'node:vm';
import { loadWpsHost } from '../../tests/wps-fixture.js';
const model = await loadWpsHost({
  Probe: [
    ['unused'],
    [null, 'OfficeBridge probe', 12],
    [null, 'Local selection', 34],
  ],
});
model.book.Name = 'OfficeBridge-XLL-Probe.xlsx';
model.select(2, 2, 2, 2);
const elements: Record<string, any> = Object.fromEntries(
  ['status', 'read', 'range', 'values'].map((id) => [
    id,
    { textContent: '', disabled: false, addEventListener: () => {} },
  ]),
);
const window: any = { addEventListener: () => {} };
const reports: any[] = [];
const context = createContext({
  window,
  OfficeBridgeHost: model.host,
  document: { getElementById: (id: string) => elements[id] },
  setTimeout,
  clearTimeout,
  fetch: async (url: string, options: any) => {
    if (url === '/probe/session')
      return { json: async () => ({ token: 'synthetic-token' }) };
    assert.equal(url, '/probe/read-result');
    reports.push(JSON.parse(options.body));
    return { ok: true };
  },
});
runInContext(
  await readFile(
    new URL('../../apps/wps-addin/executor.js', import.meta.url),
    'utf8',
  ),
  context,
);
(context as any).OfficeBridgeExecutor = window.OfficeBridgeExecutor;
runInContext(
  await readFile(new URL('./probe.js', import.meta.url), 'utf8'),
  context,
);
for (let n = 0; n < 100 && elements.read.disabled; n++)
  await new Promise((resolve) => setTimeout(resolve, 10));
assert.equal(elements.status.textContent, '已连接 WPS，选区读取成功');
assert.equal(elements.range.textContent, 'B2:C3');
assert.deepEqual(JSON.parse(elements.values.textContent), [
  ['OfficeBridge probe', 12],
  ['Local selection', 34],
]);
assert.deepEqual(reports, [{ address: 'B2:C3', rows: 2, columns: 2 }]);
model.book.Name = 'not-the-probe.xlsx';
runInContext(
  await readFile(new URL('./probe.js', import.meta.url), 'utf8'),
  context,
);
assert.match(elements.status.textContent, /请切换到/);
assert.equal(reports.length, 1);
console.info(
  'Synthetic selection contract passed: real executor/host path reads B2:C3, reports metadata only, rejects other workbooks. Not a real WPS integration pass.',
);
