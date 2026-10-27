import assert from 'node:assert/strict';
import test from 'node:test';
import { Worker } from 'node:worker_threads';
import { readFile } from 'node:fs/promises';
import { loadWpsHost } from './wps-fixture.js';
import { executeComputation } from '../packages/agents/src/execution.js';

const data = {
  Data: [
    ['Month', 'Amount'],
    ['June', '1.10'],
    ['June', '2.20'],
    ['July', '4.50'],
  ],
};
const chartArgs = {
  rangeId: 'r1',
  chartType: 'line',
  title: 'Bill',
  categoryColumn: 0,
  valueColumns: [1],
  aggregation: 'sum',
};
async function staged() {
  const m = await loadWpsHost(data);
  m.select(1, 1, 4, 2);
  const d = m.host.beginRead('context');
  const reply = m.host.executeQuery(
    d.snapshotId,
    'officebridge_create_chart',
    chartArgs,
  );
  return { m, d, reply };
}

test('the bundled execution worker computes and validates real chart plans on a different thread without host access', async () => {
  const { m, reply } = await staged();
  const bundle = await readFile(
    new URL(
      '../apps/wps-addin/.generated/execution-worker.js',
      import.meta.url,
    ),
    'utf8',
  );
  const worker = new Worker(
    'const {parentPort}=require("node:worker_threads");globalThis.postMessage=x=>parentPort.postMessage(x);parentPort.on("message",data=>globalThis.onmessage({data}));' +
      bundle,
    { eval: true },
  );
  try {
    const result = new Promise<any>((resolve, reject) => {
      worker.on('message', (m) => {
        if (m.type === 'result') resolve(m.result);
        else if (m.type === 'error') reject(new Error(m.message));
      });
      worker.on('error', reject);
    });
    assert.ok(worker.threadId > 0);
    worker.postMessage({
      operation: 'prepare',
      payload: { snapshot: reply.snapshot, plan: reply.plan },
    });
    const output = await result;
    assert.equal(output.valid, true);
    assert.deepEqual(output.summary.series[0].values, [3.3, 4.5]);
    assert.match(output.formulas[0], /3\.3,4\.5/);
    assert.equal(m.sheets[0].ChartObjects().Count, 0);
  } finally {
    await worker.terminate();
  }
});

test('all query computations share one pure dispatcher and staged edits survive separate worker jobs', async () => {
  const m = await loadWpsHost({
    Data: [
      ['Name', 'Amount'],
      [' a ', 10],
      ['b', 20],
    ],
  });
  m.select(1, 1, 3, 2);
  const operations: string[] = [];
  const compute = async (operation: string, payload: unknown) => {
    operations.push(operation);
    await Promise.resolve();
    return executeComputation(operation, m.clone(payload));
  };
  const d = await m.host.beginReadAsync('context', m.host.describe('context'), {
    compute,
  });
  await m.host.executeQueryAsync(
    d.snapshotId,
    'officebridge_inspect_range',
    { rangeId: 'r1' },
    { compute },
  );
  const stats = await m.host.executeQueryAsync(
    d.snapshotId,
    'officebridge_summarize_range',
    { rangeId: 'r1', columns: [1] },
    { compute },
  );
  assert.equal(stats.result.groups[0].columns[0].sum, 30);
  await m.host.executeQueryAsync(
    d.snapshotId,
    'officebridge_clean_range',
    { rangeId: 'r1' },
    { compute },
  );
  const second = await m.host.executeQueryAsync(
    d.snapshotId,
    'officebridge_propose_edits',
    { edits: [{ rangeId: 'r1', row: 2, column: 1, value: 50 }] },
    { compute },
  );
  assert.equal(second.plan.edits.length, 2);
  assert.ok(operations.includes('summary'));
  assert.ok(operations.includes('inspect'));
  assert.ok(operations.includes('tool'));
  assert.equal(m.sheets[0].Cells.Item(2, 1).Value2, ' a ');
  await m.host.applyAsync(second.plan, {
    prepare: async (snapshot: unknown, plan: unknown) =>
      executeComputation('prepare', { snapshot, plan }),
  });
  assert.equal(m.sheets[0].Cells.Item(2, 1).Value2, 'a');
  assert.equal(m.sheets[0].Cells.Item(3, 2).Value2, 50);
  await m.host.undoAsync();
  assert.equal(m.sheets[0].Cells.Item(2, 1).Value2, ' a ');
});

test('async range reads and structure baseline preparation yield, cancel without publishing, and release for a fresh task', async () => {
  const m = await loadWpsHost({
    Data: Array.from({ length: 2000 }, (_, i) => [i, i + 1]),
  });
  m.select(1, 1, 2000, 2);
  const controller = new AbortController();
  let heartbeats = 0;
  const timer = setInterval(() => {
    heartbeats++;
    controller.abort();
  }, 1);
  try {
    await assert.rejects(
      m.host.beginReadAsync('context', m.host.describe('context'), {
        signal: controller.signal,
      }),
      /停止/,
    );
  } finally {
    clearInterval(timer);
  }
  assert.ok(heartbeats > 0);
  m.select(1, 1);
  const d = await m.host.beginReadAsync('context', m.host.describe('context'));
  const stop = new AbortController();
  const compute = async (op: string, payload: any) => {
    if (op === 'snapshot') stop.abort();
    return executeComputation(op, m.clone(payload));
  };
  await assert.rejects(
    m.host.executeQueryAsync(
      d.snapshotId,
      'officebridge_insert_columns',
      { rangeId: 'r1', column: 0, count: 1 },
      { signal: stop.signal, compute },
    ),
    /停止/,
  );
  assert.equal(m.sheets[0].Cells.Item(1, 1).Value2, 0);
  await assert.rejects(
    m.host.executeQueryAsync(d.snapshotId, 'officebridge_read_range', {
      rangeId: 'r1',
    }),
    /失效/,
  );
});

test('chart staging/apply/undo never scan unrelated sheets or read source formats; later border changes are preserved', async () => {
  const m = await loadWpsHost({ ...data, Other: [[99]] });
  m.select(1, 1, 4, 2);
  const d = m.host.beginRead('context');
  const item = m.sheets[0].Cells.Item;
  m.sheets[0].Cells.Item = (r: number, c: number) =>
    new Proxy(item(r, c), {
      get(target, key) {
        if (
          [
            'Font',
            'Interior',
            'NumberFormat',
            'Borders',
            'RowHeight',
            'ColumnWidth',
          ].includes(String(key))
        )
          throw new Error('unnecessary format read');
        return Reflect.get(target, key);
      },
    });
  m.sheets[1].Cells.Item = () => {
    throw new Error('unrelated cell read');
  };
  const reply = m.host.executeQuery(
    d.snapshotId,
    'officebridge_create_chart',
    chartArgs,
  );
  await m.host.applyAsync(reply.plan);
  assert.equal(m.sheets[0].ChartObjects().Count, 1);
  assert.equal(m.host.queryMetrics().styleReads, 0);
  assert.ok(m.host.chartMetrics().yields > 0);
  m.host.undo();
  assert.equal(m.sheets[0].ChartObjects().Count, 0);
});

test('chart bulk verification uses matrix reads with a per-cell fallback and rejects same-result formula drift', async () => {
  const { m, reply } = await staged();
  let bulk = 0;
  const range = m.sheets[0].Range,
    item = m.sheets[0].Cells.Item;
  m.sheets[0].Range = (address: string) => {
    const target = range(address),
      match = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(address);
    if (!match) return target;
    const col = (s: string) =>
        [...s].reduce((v, c) => v * 26 + c.charCodeAt(0) - 64, 0),
      a = col(match[1]!),
      b = col(match[3]!);
    return Object.defineProperties(target, {
      Value2: {
        get() {
          bulk++;
          return Array.from({ length: +match[4]! - +match[2]! + 1 }, (_, r) =>
            Array.from(
              { length: b - a + 1 },
              (_, c) => item(+match[2]! + r, a + c).Value2,
            ),
          );
        },
      },
      HasFormula: {
        get() {
          return false;
        },
      },
    });
  };
  await m.host.applyAsync(reply.plan);
  assert.ok(bulk > 0);
  assert.ok(m.host.chartMetrics().bulkReads > 0);
  const n = await loadWpsHost({
    Data: [
      ['x', 'y'],
      [10, 1],
      [10, 2],
    ],
  });
  n.sheets[0].Cells.Item(2, 2).Formula = '=SUM(A2:A2)';
  n.select(1, 1, 3, 2);
  const d = n.host.beginRead('context');
  const { aggregation: _aggregation, ...rawChart } = chartArgs;
  const p = n.host.executeQuery(
    d.snapshotId,
    'officebridge_create_chart',
    rawChart,
  );
  n.sheets[0].Cells.Item(2, 2).Formula = '=SUM(A3:A3)';
  await assert.rejects(n.host.applyAsync(p.plan), /公式.*变化/);
  assert.equal(n.sheets[0].ChartObjects().Count, 0);
});

test('stop and edits between async checks and commit prevent chart creation; late-read references still protect application', async () => {
  for (const mode of ['stop', 'change', 'file']) {
    const { m, reply } = await staged();
    const controller = new AbortController();
    await assert.rejects(
      m.host.applyAsync(reply.plan, {
        signal: controller.signal,
        onProgress() {
          if (mode === 'stop') controller.abort();
          else if (mode === 'file') m.book.FullName = 'Other.xlsx';
          else m.sheets[0].Cells.Item(2, 2).Value2 = '999';
        },
      }),
      /停止|变化|切换/,
    );
    assert.equal(m.sheets[0].ChartObjects().Count, 0);
  }
  const m = await loadWpsHost({ ...data, Other: [[99]] });
  m.select(1, 1, 4, 2);
  const d = m.host.beginRead('context');
  const p = m.host.executeQuery(
    d.snapshotId,
    'officebridge_create_chart',
    chartArgs,
  );
  const r = m.host.executeQuery(d.snapshotId, 'officebridge_resolve_range', {
    target: 'sheet',
    sheetName: 'Other',
  });
  m.host.executeQuery(d.snapshotId, 'officebridge_read_range', {
    rangeId: r.result.rangeId,
  });
  m.sheets[1].Cells.Item(1, 1).Value2 = 100;
  await assert.rejects(m.host.applyAsync(p.plan), /变化/);
  assert.equal(m.sheets[0].ChartObjects().Count, 0);
});

test('async chart commits restore redraw settings after success and configuration failure without leaking new charts', async () => {
  for (const fail of [false, true])
    for (const updating of [false, true]) {
      const { m, reply } = await staged();
      m.app.ScreenUpdating = updating;
      m.sheets[0].failNextSeries = fail;
      if (fail) await assert.rejects(m.host.applyAsync(reply.plan), /已恢复/);
      else await m.host.applyAsync(reply.plan);
      assert.equal(m.app.ScreenUpdating, updating);
      assert.equal(m.sheets[0].ChartObjects().Count, fail ? 0 : 1);
    }
});

test('stopping a queued recovery preserves the applied chart and its recovery record', async () => {
  const { m, reply } = await staged();
  await m.host.applyAsync(reply.plan);
  const stop = new AbortController();
  stop.abort();
  await assert.rejects(m.host.undoAsync({ signal: stop.signal }), /停止/);
  assert.equal(m.sheets[0].ChartObjects().Count, 1);
  assert.equal(m.host.canUndo(), true);
});
