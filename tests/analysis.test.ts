import assert from 'node:assert/strict';
import { once } from 'node:events';
import { readFile } from 'node:fs/promises';
import { runInNewContext } from 'node:vm';
import test from 'node:test';
import { createBridgeServer } from '../apps/bridge/src/server.js';
import { isAnalysisSelection } from '../packages/protocol/src/index.js';
import { runSelectionTool } from '../packages/agents/src/selection-tools.js';
import { fakeAdapter, fullSelection } from './fixtures.js';
import type { AgentAdapter } from '../packages/agents/src/index.js';

test('full selection validation rejects samples, malformed dimensions and oversized data', () => {
  assert.ok(isAnalysisSelection(fullSelection));
  assert.ok(!isAnalysisSelection({ ...fullSelection, rowCount: 50 }));
  assert.ok(!isAnalysisSelection({ ...fullSelection, values: [[Number.NaN]] }));
  assert.ok(
    !isAnalysisSelection({ ...fullSelection, rowCount: 5001, columnCount: 1 }),
  );
  assert.ok(
    !isAnalysisSelection({
      ...fullSelection,
      values: fullSelection.values.map((row) =>
        row.map(() => 'a'.repeat(10001)),
      ),
    }),
  );
});

test('statistics include all rows, group statuses correctly and exclude blank durations from averages', () => {
  const summary = runSelectionTool(
    fullSelection,
    'officebridge_summarize_selection',
    { groupByColumns: [0, 1] },
  ) as any;
  assert.equal(summary.rowCount, 3);
  assert.equal(summary.groups.length, 3);
  assert.equal(summary.groups[0].columns[2].average, 10);
  assert.equal(summary.groups[1].columns[2].numericCount, 0);
  assert.equal(summary.groups[1].columns[2].average, null);
  const all = runSelectionTool(
    fullSelection,
    'officebridge_summarize_selection',
    {},
  ) as any;
  assert.equal(all.groups[0].columns[2].average, 20);
  assert.throws(() =>
    runSelectionTool(fullSelection, 'officebridge_summarize_selection', {
      groupByColumns: [9],
    }),
  );
  assert.throws(() => runSelectionTool(fullSelection, 'write_cells', {}));
});

test('WPS reader captures rows beyond the old sample and blocks oversized selections', async () => {
  const script = await readFile(
    new URL('../apps/wps-addin/host.js', import.meta.url),
    'utf8',
  );
  const values = Array.from({ length: 30 }, (_, index) => [index]);
  const selection = {
    Row: 1,
    Column: 1,
    Rows: { Count: 30 },
    Columns: { Count: 1 },
    Areas: { Count: 1 },
    Address: 'A1:A30',
  };
  const window: any = {};
  runInNewContext(script, {
    window,
    Blob,
    Application: {
      ActiveWorkbook: { Name: 'test.xlsx' },
      ActiveSheet: {
        Name: 'Sheet1',
        Cells: {
          Item: (row: number) => ({ Value2: values[row - 1]?.[0] ?? row }),
        },
      },
      Selection: selection,
    },
  });
  const result = JSON.parse(
    JSON.stringify(window.OfficeBridgeHost.readSelection()),
  );
  assert.equal(result.values.length, 30);
  assert.equal(result.values[29][0], 29);
  assert.ok(isAnalysisSelection(result));
  selection.Rows.Count = 5001;
  selection.Address = 'A1:A5001';
  const large = JSON.parse(
    JSON.stringify(window.OfficeBridgeHost.readSelection()),
  );
  assert.equal(large.values.length, 5001);
  assert.equal(large.values[5000][0], 5001);
  assert.ok(isAnalysisSelection(large));
});

test('legacy selection reader stops on UTF-8 byte size instead of cell count', async () => {
  const script = await readFile(
    new URL('../apps/wps-addin/host.js', import.meta.url),
    'utf8',
  );
  const window: any = {};
  let reads = 0;
  runInNewContext(script, {
    window,
    Blob,
    Application: {
      ActiveWorkbook: { Name: 'test.xlsx' },
      ActiveSheet: {
        Name: 'Sheet1',
        Cells: {
          Item: () => {
            reads++;
            return { Value2: '测'.repeat(1000) };
          },
        },
      },
      Selection: {
        Row: 1,
        Column: 1,
        Rows: { Count: 100000 },
        Columns: { Count: 10000 },
        Areas: { Count: 1 },
        Address: 'A1:NTP100000',
      },
    },
  });
  assert.throws(() => window.OfficeBridgeHost.readSelection(), /512 KB/);
  assert.ok(
    reads < 1000,
    'do not read the whole huge selection before reporting byte overflow',
  );
});

test('analysis API authenticates requests and streams the full snapshot', async (t) => {
  let received: any;
  const adapter: AgentAdapter = {
    ...fakeAdapter,
    async *run(request) {
      received = request.context;
      yield { type: 'text', text: '完整数据已收到' };
      yield { type: 'complete' };
    },
  };
  const server = createBridgeServer({ adapter });
  server.listen(0, '127.0.0.1');
  await once(server, 'listening');
  t.after(() => {
    server.closeAllConnections();
    server.close();
  });
  const address = server.address();
  assert.ok(address && typeof address !== 'string');
  const base = `http://127.0.0.1:${address.port}`;
  const session = (await (await fetch(`${base}/api/v1/session`)).json()) as any;
  const headers = {
    Origin: base,
    'Content-Type': 'application/json',
    'X-OfficeBridge-Token': session.token,
  };
  const body = JSON.stringify({ prompt: '统计时长', selection: fullSelection });
  assert.equal(
    (
      await fetch(`${base}/api/v1/analyze`, {
        method: 'POST',
        headers: { ...headers, 'X-OfficeBridge-Token': 'invalid' },
        body,
      })
    ).status,
    401,
  );
  assert.equal(
    (
      await fetch(`${base}/api/v1/analyze`, {
        method: 'POST',
        headers: { ...headers, Origin: 'https://evil.example' },
        body,
      })
    ).status,
    403,
  );
  assert.equal(
    (
      await fetch(`${base}/api/v1/analyze`, {
        method: 'POST',
        headers,
        body: '{',
      })
    ).status,
    400,
  );
  assert.equal(
    (
      await fetch(`${base}/api/v1/analyze`, {
        method: 'POST',
        headers,
        body: JSON.stringify({
          prompt: 'x',
          selection: { ...fullSelection, rowCount: 100 },
        }),
      })
    ).status,
    400,
  );
  const response = await fetch(`${base}/api/v1/analyze`, {
    method: 'POST',
    headers,
    body,
  });
  assert.equal(response.status, 200);
  const events = (await response.text())
    .trim()
    .split('\n')
    .map((line) => JSON.parse(line));
  assert.equal(events.at(-1).type, 'complete');
  assert.deepEqual(received.values, fullSelection.values);
});

test('timeout aborts the adapter and releases the analysis slot', async (t) => {
  let aborted = false;
  const adapter: AgentAdapter = {
    ...fakeAdapter,
    async *run(request) {
      await new Promise<void>((resolve) => {
        request.signal.addEventListener(
          'abort',
          () => {
            aborted = true;
            resolve();
          },
          { once: true },
        );
      });
      yield { type: 'error', code: 'cancelled', message: 'stopped' };
    },
  };
  const server = createBridgeServer({ adapter, analysisTimeoutMs: 25 });
  server.listen(0, '127.0.0.1');
  await once(server, 'listening');
  t.after(() => {
    server.closeAllConnections();
    server.close();
  });
  const address = server.address();
  assert.ok(address && typeof address !== 'string');
  const base = `http://127.0.0.1:${address.port}`;
  const { token } = (await (
    await fetch(`${base}/api/v1/session`)
  ).json()) as any;
  const options = {
    method: 'POST',
    headers: {
      Origin: base,
      'Content-Type': 'application/json',
      'X-OfficeBridge-Token': token,
    },
    body: JSON.stringify({ prompt: '分析', selection: fullSelection }),
  };
  const first = await fetch(`${base}/api/v1/analyze`, options);
  assert.match(await first.text(), /"code":"timeout"/);
  assert.ok(aborted);
  const second = await fetch(`${base}/api/v1/analyze`, options);
  assert.equal(second.status, 200);
  await second.text();
});
