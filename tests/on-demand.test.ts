import assert from 'node:assert/strict';
import test from 'node:test';
import { once } from 'node:events';
import { loadWpsHost } from './wps-fixture.js';
import { createTableTask } from '../packages/agents/src/table-task.js';
import {
  isWorkbookDescriptor,
  ReadParameterError,
  ToolParameterError,
  projectRange,
} from '../packages/protocol/src/on-demand.js';
import { createBridgeServer } from '../apps/bridge/src/server.js';
import { fakeAdapter } from './fixtures.js';
import { HostBroker } from '../apps/bridge/src/host-broker.js';

test('a missing formula equals sign at blank E12 is corrected before host access, then preview/apply/undo preserve the source', async () => {
  const source = [
    ['Month', '', '', '', 'Amount'],
    ...Array.from({ length: 10 }, (_, i) => ['Month ' + i, '', '', '', i + 1]),
  ];
  const m = await loadWpsHost({ Summary: source });
  m.select(12, 5);
  const descriptor = m.clone(m.host.beginRead('context'));
  let calls = 0;
  const task = createTableTask(
    descriptor,
    async (name, args) => {
      calls++;
      return m.clone(
        await m.host.executeQueryAsync(descriptor.snapshotId, name, args),
      );
    },
    new AbortController().signal,
  );
  const edit = {
    rangeId: 'r1',
    row: 0,
    column: 0,
    formula: 'SUM(E2:E11)',
    format: { bold: true },
  };
  const metrics = m.clone(m.host.queryMetrics());
  await assert.rejects(
    Promise.resolve(task.run('officebridge_propose_edits', { edits: [edit] })),
    (error) =>
      error instanceof ToolParameterError &&
      error.field === 'edits[0].formula' &&
      error.message.includes('必须以等号 = 开头'),
  );
  assert.equal(calls, 0);
  assert.deepEqual(m.clone(m.host.queryMetrics()), metrics);
  assert.equal(task.proposal(), null);
  await task.run('officebridge_propose_edits', {
    edits: [{ ...edit, formula: '=SUM(E2:E11)' }],
  });
  assert.equal(calls, 1);
  const plan = task.proposal();
  assert.ok(plan);
  assert.equal(m.sheets[0].Cells.Item(12, 5).Value2, null);
  m.host.apply(plan);
  assert.equal(m.sheets[0].Cells.Item(12, 5).Formula, '=SUM(E2:E11)');
  assert.equal(m.sheets[0].Cells.Item(12, 5).Value2, 55);
  assert.equal(m.sheets[0].Cells.Item(12, 5).Font.Bold, true);
  m.host.undo();
  assert.equal(m.sheets[0].Cells.Item(12, 5).Value2, null);
  for (let r = 2; r <= 11; r++)
    assert.equal(m.sheets[0].Cells.Item(r, 5).Value2, r - 1);
});

test('edit preflight rejects unsafe formulas and invalid bounds/fields without host access and discards an older proposal', async () => {
  const m = await loadWpsHost({ Summary: [[10]] });
  m.select(1, 1);
  const descriptor = m.clone(m.host.beginRead('context'));
  let calls = 0;
  const task = createTableTask(
    descriptor,
    async (name, args) => {
      calls++;
      return m.clone(
        await m.host.executeQueryAsync(descriptor.snapshotId, name, args),
      );
    },
    new AbortController().signal,
  );
  const base = { rangeId: 'r1', row: 0, column: 0 };
  for (const edits of [
    [],
    [{ ...base, formula: '=SUM([book]Sheet1!A1)' }],
    [{ ...base, formula: '=WEBSERVICE("https://example.com")' }],
    [{ ...base, formula: 'SUM(A1:A1)' }],
    [{ ...base, formula: '=1', value: 1 }],
    [{ ...base, value: '=SUM(A1:A1)' }],
    [{ ...base, value: 1, row: 1 }],
    [{ ...base, value: 1, column: 1 }],
    [{ ...base, value: 1, rangeId: 'historical' }],
    [{ ...base, value: 1, script: 'x' }],
    [{ ...base, format: { bold: 'yes' } }],
    [
      { ...base, value: 1 },
      { ...base, value: 2 },
    ],
  ]) {
    await assert.rejects(
      Promise.resolve(task.run('officebridge_propose_edits', { edits })),
      ToolParameterError,
    );
  }
  assert.equal(calls, 0);
  await task.run('officebridge_propose_edits', {
    edits: [{ ...base, value: 20 }],
  });
  await task.run('officebridge_continue_after_apply', { nextStep: 'next' });
  assert.ok(task.proposal());
  assert.ok(task.continuation());
  const before = calls;
  await assert.rejects(
    Promise.resolve(
      task.run('officebridge_propose_edits', {
        edits: [{ ...base, formula: 'SUM(A1:A1)' }],
      }),
    ),
    ToolParameterError,
  );
  assert.equal(calls, before);
  assert.equal(task.proposal(), null);
  assert.equal(task.continuation(), null);
  assert.equal(m.sheets[0].Cells.Item(1, 1).Value2, 10);
});

test('chart ordering parameters can be corrected after applying a result sheet without poisoning its snapshot', async () => {
  for (const correction of ['omit-order', 'aggregate']) {
    const m = await loadWpsHost({ Source: [['Bill'], ['0.10']] });
    const first = m.host.beginRead('context');
    const staged = m.host.executeQuery(
      first.snapshotId,
      'officebridge_propose_result_sheet',
      {
        name: '6-9月测试',
        values: [
          ['Month', 'Bill'],
          ['2026-06', 2.6],
          ['2026-07', 2.6],
          ['2026-08', 2.6],
          ['2026-09', 2.6],
        ],
      },
    );
    m.host.apply(staged.plan);
    const d = m.clone(m.host.beginRead('context'));
    let calls = 0;
    const task = createTableTask(
      d,
      async (name, args) => {
        calls++;
        return m.clone(
          await m.host.executeQueryAsync(d.snapshotId, name, args),
        );
      },
      new AbortController().signal,
    );
    const range: any = await task.run('officebridge_resolve_range', {
      target: 'sheet',
      sheetName: '6-9月测试',
    });
    const args = {
      rangeId: range.rangeId,
      chartType: 'line',
      title: 'Monthly bill',
      categoryColumn: 0,
      valueColumns: [1],
    };
    const before = calls;
    const metrics = m.clone(m.host.queryMetrics());
    await assert.rejects(
      Promise.resolve(
        task.run('officebridge_create_chart', {
          ...args,
          categoryOrder: 'ascending',
        }),
      ),
      (error) =>
        error instanceof ToolParameterError &&
        error.field === 'categoryOrder' &&
        error.message.includes('aggregation="sum"'),
    );
    assert.equal(calls, before);
    assert.deepEqual(m.clone(m.host.queryMetrics()), metrics);
    assert.equal(task.proposal(), null);
    await task.run('officebridge_create_chart', {
      ...args,
      ...(correction === 'aggregate'
        ? { aggregation: 'sum', categoryOrder: 'ascending' }
        : {}),
    });
    const plan = task.proposal();
    assert.ok(plan);
    assert.equal(m.book.Worksheets.Item('6-9月测试').ChartObjects().Count, 0);
    m.host.apply(plan);
    assert.equal(m.book.Worksheets.Item('6-9月测试').ChartObjects().Count, 1);
    assert.equal(m.sheets[0].Cells.Item(2, 1).Value2, '0.10');
  }
});

test('read parameter preflight permits correction on the same 105 by 23 source without poisoning the host or guessing columns', async () => {
  const data = [
    Array.from({ length: 23 }, (_, c) =>
      c === 2 ? 'Month' : c === 19 ? 'Bill' : 'Other',
    ),
    ...Array.from({ length: 104 }, (_, i) =>
      Array.from({ length: 23 }, (_, c) =>
        c === 2 ? '2026-0' + (6 + (i % 4)) : c === 19 ? '0.10' : 'private',
      ),
    ),
  ];
  const m = await loadWpsHost({ Data: data });
  m.select(15, 8);
  const d = m.clone(m.host.beginRead('context'));
  let calls = 0;
  const task = createTableTask(
    d,
    async (name, args) => {
      calls++;
      return m.clone(m.host.executeQuery(d.snapshotId, name, args));
    },
    new AbortController().signal,
  );
  const range: any = await task.run('officebridge_resolve_range', {
    target: 'sheet',
  });
  assert.equal(range.address, 'A1:W105');
  const header: any = await task.run('officebridge_read_range', {
    rangeId: range.rangeId,
    columns: [2, 19],
    rowCount: 1,
  });
  assert.deepEqual(header.values, [['Month', 'Bill']]);
  assert.equal(header.columnCount, 2);
  assert.equal(header.rowCount, 1);
  assert.equal(header.sourceRange.columnCount, 23);
  assert.equal(header.sourceRange.rowCount, 105);
  for (const [params, field] of [
    [{ rowOffset: 1, rowCount: 105 }, 'rowCount'],
    [{ columns: ['C', 'T'] }, 'columns'],
    [{ columns: [23] }, 'columns'],
    [{ columns: [2, 2] }, 'columns'],
    [{ rowOffset: 105 }, 'rowOffset'],
    [{ fields: ['value'] }, 'fields'],
    [{ fields: [] }, 'fields'],
    [{ rangeId: 'r1', columns: [2, 19] }, 'columns'],
    [{ rangeId: 'previous-task' }, 'rangeId'],
  ] as const) {
    const before = calls,
      metrics = m.clone(m.host.queryMetrics());
    await assert.rejects(
      Promise.resolve(
        task.run('officebridge_read_range', {
          rangeId: range.rangeId,
          ...params,
        }),
      ),
      (error) =>
        error instanceof ReadParameterError &&
        error.field === field &&
        error.message.includes('本次调用未进入宿主'),
    );
    assert.equal(calls, before);
    assert.deepEqual(m.clone(m.host.queryMetrics()), metrics);
    const corrected: any = await task.run('officebridge_read_range', {
      rangeId: range.rangeId,
      columns: [2, 19],
      rowOffset: 1,
    });
    assert.equal(corrected.rowCount, 104);
    assert.equal(corrected.values.length, 104);
    assert.deepEqual(corrected.columnIndices, [2, 19]);
    assert.equal(corrected.values[0][1], '0.10');
    assert.doesNotMatch(JSON.stringify(corrected), /private/);
  }
  const beforeStats = calls;
  await assert.rejects(
    Promise.resolve(
      task.run('officebridge_summarize_range', {
        rangeId: range.rangeId,
        columns: [23],
        groupByColumns: [2],
      }),
    ),
    /columns/,
  );
  assert.equal(calls, beforeStats);
  const stats: any = await task.run('officebridge_summarize_range', {
    rangeId: range.rangeId,
    columns: [19],
    groupByColumns: [2],
    numericText: true,
  });
  assert.equal(stats.groups.length, 4);
  assert.ok(stats.groups.every((g: any) => g.columns[0].sum === 2.6));
  assert.equal(m.host.queryMetrics().styleReads, 0);
  assert.equal(m.host.queryMetrics().baselineCaptures, 0);
  assert.equal(task.proposal(), null);
});

test('preflight correction does not bypass freshness, and projected metadata always retains its full source dimensions', async () => {
  const m = await loadWpsHost({
    Data: [
      ['Month', 'Bill'],
      ['June', 10],
      ['July', 20],
    ],
  });
  m.select(1, 1, 3, 2);
  const d = m.clone(m.host.beginRead('selection'));
  const task = createTableTask(
    d,
    async (name, args) =>
      m.clone(m.host.executeQuery(d.snapshotId, name, args)),
    new AbortController().signal,
  );
  await task.run('officebridge_read_range', { rangeId: 'r1', columns: [1] });
  await assert.rejects(
    Promise.resolve(
      task.run('officebridge_read_range', {
        rangeId: 'r1',
        rowOffset: 1,
        rowCount: 3,
      }),
    ),
    /最多读取 2 行/,
  );
  m.sheets[0].Cells.Item(2, 2).Value2 = 99;
  await assert.rejects(
    Promise.resolve(
      task.run('officebridge_read_range', { rangeId: 'r1', columns: [1] }),
    ),
    /已变化/,
  );
  assert.equal(task.proposal(), null);
  const snapshot = m.clone(m.host.capture('sheet'));
  const projected: any = projectRange(snapshot.ranges[0], {
    columns: [1],
    rowOffset: 1,
    rowCount: 1,
  });
  assert.equal(projected.rowCount, 1);
  assert.equal(projected.columnCount, 1);
  assert.equal(projected.sourceRange.rowCount, 3);
  assert.equal(projected.sourceRange.columnCount, 2);
});

test('one selected cell never reads UsedRange, unselected data, formats or a recovery baseline before sending', async () => {
  const m = await loadWpsHost({
    Data: [
      [1, 2],
      [3, 4],
    ],
    Huge: [['not selected']],
  });
  m.select(2, 1, 1, 1);
  let reads = 0;
  for (const sheet of m.sheets) {
    const item = sheet.Cells.Item;
    sheet.Cells.Item = (r: number, c: number) => {
      assert.equal(sheet.Name, 'Data');
      assert.equal(r, 2);
      assert.equal(c, 1);
      reads++;
      return new Proxy(item(r, c), {
        get(target, key) {
          if (
            ['Font', 'Interior', 'NumberFormat', 'Formula'].includes(
              String(key),
            )
          )
            throw new Error('unrequested fields');
          return Reflect.get(target, key);
        },
      });
    };
  }
  const guard = (sheet: any) =>
    new Proxy(sheet, {
      get(target, key) {
        if (key === 'UsedRange')
          throw new Error('UsedRange must not be scanned');
        return Reflect.get(target, key);
      },
    });
  m.app.ActiveSheet = guard(m.sheets[0]);
  const originalSheet = m.book.Worksheets.Item;
  m.book.Worksheets.Item = (key: any) => guard(originalSheet(key));
  const descriptor = m.host.beginRead(
    'selection',
    m.host.describe('selection'),
  );
  assert.ok(isWorkbookDescriptor(descriptor));
  assert.equal(reads, 1);
  assert.equal(descriptor.inline![0]!.values[0]![0], 3);
  assert.equal(m.host.queryMetrics().baselineCaptures, 0);
  assert.equal(m.host.queryMetrics().styleReads, 0);
  assert.equal(m.host.queryMetrics().formulaReads, 0);
});

test('large scope is metadata only; projected reads fetch only requested values and statistics return exact grouped numeric-text totals', async () => {
  const rows = [
    ['Month', 'Secret', 'Bill'],
    ...Array.from({ length: 104 }, (_, i) => [
      i % 2 ? '2026-01' : '2026-02',
      'private',
      '0.100000000000',
    ]),
  ];
  const m = await loadWpsHost({ Data: rows });
  m.select(1, 1, 105, 3);
  const descriptor = m.host.beginRead('selection');
  assert.equal(descriptor.inline, undefined);
  assert.equal(m.host.queryMetrics().valueReads, 0);
  const item = m.sheets[0].Cells.Item;
  m.sheets[0].Cells.Item = (r: number, c: number) => {
    assert.notEqual(c, 2, 'secret column must not be read');
    return new Proxy(item(r, c), {
      get(target, key) {
        if (
          ['Font', 'Interior', 'NumberFormat', 'Formula'].includes(String(key))
        )
          throw new Error('unrequested field');
        return Reflect.get(target, key);
      },
    });
  };
  const task = createTableTask(
    descriptor,
    async (name, args) =>
      m.clone(m.host.executeQuery(descriptor.snapshotId, name, args)),
    new AbortController().signal,
  );
  const first: any = await task.run('officebridge_read_range', {
    rangeId: 'r1',
    columns: [0, 2],
    rowCount: 1,
  });
  assert.deepEqual(first.values, [['Month', 'Bill']]);
  assert.equal(first.styles, undefined);
  assert.equal(first.formulas, undefined);
  const stats: any = await task.run('officebridge_summarize_range', {
    rangeId: 'r1',
    columns: [2],
    groupByColumns: [0],
    numericText: true,
  });
  assert.equal(stats.groups.length, 2);
  assert.deepEqual(
    stats.groups.map((g: any) => g.columns[0].sum),
    [5.2, 5.2],
  );
  assert.deepEqual(
    stats.groups[0].columns.map((c: any) => c.index),
    [2],
  );
  assert.ok(!JSON.stringify(stats).includes('private'));
  assert.equal(m.host.queryMetrics().baselineCaptures, 0);
  assert.equal(task.proposal(), null);
  // Exact repeated query reuses calculations but checks source values for freshness.
  assert.deepEqual(
    await task.run('officebridge_summarize_range', {
      rangeId: 'r1',
      columns: [2],
      groupByColumns: [0],
      numericText: true,
    }),
    stats,
  );
  item(2, 3).Value2 = '9';
  await assert.rejects(
    async () =>
      task.run('officebridge_read_range', {
        rangeId: 'r1',
        columns: [2],
        rowCount: 1,
      }),
    /已变化/,
  );
  assert.equal(task.proposal(), null);
});

test('on-demand writes stage real baselines only when needed, preserve data until apply and protect preview/undo against later edits', async () => {
  for (const tool of [
    'officebridge_propose_edits',
    'officebridge_create_chart',
    'officebridge_insert_columns',
    'officebridge_propose_result_sheet',
  ]) {
    const m = await loadWpsHost({
      Data: [
        ['Month', 'Bill'],
        ['Jan', 1],
        ['Feb', 2],
      ],
      Other: [[99]],
    });
    m.select(1, 1, 3, 2);
    const d = m.host.beginRead('selection');
    const task = createTableTask(
      d,
      async (name, args) =>
        m.clone(m.host.executeQuery(d.snapshotId, name, args)),
      new AbortController().signal,
    );
    const args: any =
      tool === 'officebridge_propose_edits'
        ? { edits: [{ rangeId: 'r1', row: 1, column: 1, value: 7 }] }
        : tool === 'officebridge_create_chart'
          ? {
              rangeId: 'r1',
              chartType: 'line',
              title: 'Bill',
              categoryColumn: 0,
              valueColumns: [1],
            }
          : tool === 'officebridge_insert_columns'
            ? { rangeId: 'r1', column: 1, count: 1 }
            : { name: 'Result', values: [['ok']] };
    await task.run(tool, args);
    assert.ok(task.proposal());
    assert.ok(task.snapshot);
    assert.equal(m.sheets[0].Cells.Item(2, 2).Value2, 1);
    assert.equal(
      m.host.queryMetrics().baselineCaptures,
      tool === 'officebridge_create_chart' ||
        tool === 'officebridge_insert_columns'
        ? 1
        : 0,
    );
    m.host.apply(task.proposal());
    m.host.undo();
    assert.equal(m.sheets.length, 2);
    assert.equal(m.sheets[0].Cells.Item(2, 2).Value2, 1);
  }
  const m = await loadWpsHost({
    Data: [
      ['Month', 'Bill'],
      ['Jan', 1],
    ],
    Other: [[99]],
  });
  m.select(1, 1, 2, 2);
  const d = m.host.beginRead('selection');
  const response = m.host.executeQuery(
    d.snapshotId,
    'officebridge_create_chart',
    {
      rangeId: 'r1',
      chartType: 'line',
      title: 'Bill',
      categoryColumn: 0,
      valueColumns: [1],
    },
  );
  m.sheets[1].Cells.Item(1, 1).Value2 = 100;
  m.host.apply(response.plan);
  m.host.undo();
  assert.equal(m.sheets[1].Cells.Item(1, 1).Value2, 100);
  assert.equal(m.sheets[0].ChartObjects().Count, 0);
});

test('on-demand boundaries reject invalid columns, stale run IDs, changed workbooks and byte overflow without publishing a plan', async () => {
  for (const args of [
    { rangeId: 'r1', columns: [9] },
    { rangeId: 'r1', rowOffset: 9 },
    { rangeId: 'r1', fields: ['script'] },
    { rangeId: 'unknown' },
  ]) {
    const m = await loadWpsHost();
    const d = m.host.beginRead('selection');
    assert.throws(() =>
      m.host.executeQuery(d.snapshotId, 'officebridge_read_range', args),
    );
  }
  const m = await loadWpsHost();
  const first = m.host.beginRead('selection');
  const next = m.host.beginRead('selection');
  assert.throws(
    () =>
      m.host.executeQuery(first.snapshotId, 'officebridge_read_range', {
        rangeId: 'r1',
      }),
    /失效/,
  );
  m.book.FullName = 'E:\\different.xlsx';
  assert.throws(
    () =>
      m.host.executeQuery(next.snapshotId, 'officebridge_read_range', {
        rangeId: 'r1',
      }),
    /切换/,
  );
  const huge = await loadWpsHost({
    Data: Array.from({ length: 100 }, () => ['x'.repeat(10000)]),
  });
  huge.select(1, 1, 100, 1);
  const h = huge.host.beginRead('selection');
  assert.throws(
    () =>
      huge.host.executeQuery(h.snapshotId, 'officebridge_read_range', {
        rangeId: 'r1',
      }),
    /512 KB/,
  );
});

test('host broker rejects mismatched and late replies and releases a pending read on cancellation', async () => {
  const broker = new HostBroker(),
    controller = new AbortController();
  let event: any;
  const run = broker.open('snapshot', controller.signal, (e) => {
    event = e;
  });
  const pending = run.call('officebridge_read_range', { rangeId: 'r1' });
  assert.equal(
    broker.reply({ ...event, runId: 'wrong', reply: { result: 1 } }),
    false,
  );
  controller.abort();
  await assert.rejects(pending, /取消/);
  assert.equal(broker.reply({ ...event, reply: { result: 1 } }), false);
  run.close();
});

test('host reads time out, reject late replies, and cannot restart after the run closes', async (t) => {
  t.mock.timers.enable({ apis: ['setTimeout'] });
  const broker = new HostBroker();
  let event: any;
  const run = broker.open('timeout', new AbortController().signal, (e) => {
    event = e;
  });
  const rejected = assert.rejects(
    run.call('officebridge_read_range', {}),
    /超时/,
  );
  t.mock.timers.tick(30000);
  await rejected;
  assert.equal(broker.reply({ ...event, reply: { result: 1 } }), false);
  run.close();
  await assert.rejects(run.call('officebridge_read_range', {}), /停止/);
});

test('cancelled and forged write replies never become applicable plans', async () => {
  const m = await loadWpsHost();
  const descriptor = m.clone(m.host.beginRead('selection'));
  const edits = [{ rangeId: 'r1', row: 0, column: 0, value: 'new' }];
  const staged = m.clone(
    m.host.executeQuery(descriptor.snapshotId, 'officebridge_propose_edits', {
      edits,
    }),
  );
  const controller = new AbortController();
  const task = createTableTask(
    descriptor,
    async () => {
      controller.abort();
      return staged;
    },
    controller.signal,
  );
  await assert.rejects(
    async () => task.run('officebridge_propose_edits', { edits }),
    /停止/,
  );
  assert.equal(task.proposal(), null);
  m.host.finishRead(descriptor.snapshotId, false);
  assert.throws(() => m.host.apply(staged.plan), /失效|重新|快照/);
  const tampered = structuredClone(staged);
  tampered.snapshot.ranges[0].startRow++;
  tampered.snapshot.ranges[0].address = 'A2:B4';
  const forged = createTableTask(
    descriptor,
    async () => tampered,
    new AbortController().signal,
  );
  await assert.rejects(
    async () => forged.run('officebridge_propose_edits', { edits }),
    /越界或无效/,
  );
  assert.equal(forged.proposal(), null);
});

test('bounded bulk reads use only requested column blocks and fall back safely on hosts without a matrix API', async () => {
  const m = await loadWpsHost({
    Data: Array.from({ length: 105 }, (_, i) => [i, 'secret', i + 1]),
  });
  m.select(1, 1, 105, 3);
  const range = m.sheets[0].Range,
    item = m.sheets[0].Cells.Item;
  const blocks: string[] = [];
  m.sheets[0].Range = (address: string) => {
    if (address === 'A1:A105' || address === 'C1:C105') {
      blocks.push(address);
      const column = address.startsWith('A') ? 1 : 3;
      return {
        get Value2() {
          return Array.from({ length: 105 }, (_, i) => [
            item(i + 1, column).Value2,
          ]);
        },
      };
    }
    return range(address);
  };
  const d = m.host.beginRead('selection');
  const result = m.host.executeQuery(d.snapshotId, 'officebridge_read_range', {
    rangeId: 'r1',
    columns: [2, 0],
  });
  assert.deepEqual(blocks, ['A1:A105', 'C1:C105']);
  assert.deepEqual(m.clone(result.result.values[104]), [105, 104]);
  assert.equal(m.host.queryMetrics().bulkReads, 2);
  assert.equal(m.host.queryMetrics().styleReads, 0);
});

test('descriptor validation rejects hidden data, malformed inline records and out-of-scope shapes without throwing', async () => {
  const m = await loadWpsHost();
  const d = m.clone(m.host.beginRead('selection'));
  for (const v of [
    null,
    [],
    { ...d, values: [[1]] },
    { ...d, inline: [null] },
    { ...d, inline: [{}] },
    { ...d, ranges: [{ ...d.ranges[0], values: [[1]] }] },
    { ...d, columnOperations: { available: true, script: 'bad' } },
  ])
    assert.equal(isWorkbookDescriptor(v), false);
});

test('real HTTP streams carry host read requests before a fake Agent response, preserve scope and refuse unauthenticated replies', async (t) => {
  const m = await loadWpsHost({ Data: [['Amount'], [12], [30]] });
  m.select(1, 1, 3, 1);
  const descriptor = m.host.beginRead('selection');
  const adapter = {
    ...fakeAdapter,
    async *run(request: any) {
      assert.equal(request.context, undefined);
      assert.equal(request.workbook.ranges[0].values, undefined);
      const task = createTableTask(
        request.workbook,
        request.hostTool,
        request.signal,
      );
      const result: any = await task.run('officebridge_summarize_range', {
        rangeId: 'r1',
        columns: [0],
      });
      assert.equal(result.groups[0].columns[0].sum, 42);
      yield { type: 'text' as const, text: '合计42' };
      yield { type: 'complete' as const };
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
  const base = 'http://127.0.0.1:' + address.port;
  const token = ((await (await fetch(base + '/api/v1/session')).json()) as any)
    .token;
  const headers = {
    Origin: base,
    'Content-Type': 'application/json',
    'X-OfficeBridge-Token': token,
  };
  const conversation: any = await (
    await fetch(base + '/api/v1/conversations', {
      method: 'POST',
      headers,
      body: JSON.stringify({ agentId: 'codex' }),
    })
  ).json();
  const response = await fetch(base + '/api/v1/run', {
    method: 'POST',
    headers,
    body: JSON.stringify({
      prompt: '合计',
      workbook: descriptor,
      conversationId: conversation.id,
    }),
  });
  assert.equal(response.status, 200);
  const events: any[] = [];
  let buffer = '';
  const reader = response.body!.getReader();
  const decoder = new TextDecoder();
  while (true) {
    const next = await reader.read();
    if (next.done) break;
    buffer += decoder.decode(next.value);
    const lines = buffer.split('\n');
    buffer = lines.pop()!;
    for (const line of lines) {
      if (!line) continue;
      const event = JSON.parse(line);
      events.push(event);
      if (event.type === 'host_request') {
        const reply = m.host.executeQuery(
          event.snapshotId,
          event.name,
          event.args,
        );
        assert.equal(
          (
            await fetch(base + '/api/v1/host-result', {
              method: 'POST',
              headers: { Origin: base, 'Content-Type': 'application/json' },
              body: JSON.stringify({ ...event, reply }),
            })
          ).status,
          401,
        );
        assert.equal(
          (
            await fetch(base + '/api/v1/host-result', {
              method: 'POST',
              headers,
              body: JSON.stringify({ ...event, reply }),
            })
          ).status,
          200,
        );
      }
    }
  }
  assert.deepEqual(
    events.map((e) => e.type),
    ['host_request', 'text', 'complete'],
  );
  assert.equal(
    (
      await fetch(base + '/api/v1/host-result', {
        method: 'POST',
        headers,
        body: JSON.stringify({ ...events[0], reply: { result: 0 } }),
      })
    ).status,
    410,
  );
});
