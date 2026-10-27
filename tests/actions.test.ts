import assert from 'node:assert/strict';
import test from 'node:test';
import { loadWpsHost } from './wps-fixture.js';
import {
  TableToolSession,
  tableTools,
} from '../packages/agents/src/table-tools.js';
import {
  isEditPlan,
  isWorkbookSnapshot,
  summarizeChart,
  summarySeriesFormula,
  type EditPlan,
  type WorkbookSnapshot,
  type SheetAction,
} from '../packages/protocol/src/table.js';

function base(snapshot: WorkbookSnapshot): EditPlan {
  return {
    snapshotId: snapshot.snapshotId,
    connectionId: snapshot.connectionId,
    workbookFullName: snapshot.workbookFullName,
    edits: [],
    newSheets: [],
  };
}
function action(snapshot: WorkbookSnapshot, change: SheetAction): EditPlan {
  return { ...base(snapshot), sheetActions: [change] };
}
function rows(
  snapshot: WorkbookSnapshot,
  kind: 'insert' | 'delete',
  row: number,
  count = 1,
  values?: any[][],
): EditPlan {
  return {
    ...base(snapshot),
    rowChanges: [
      {
        rangeId: 'r1',
        action: kind,
        row,
        count,
        ...(values ? { values } : {}),
      },
    ],
  };
}
function read(
  model: Awaited<ReturnType<typeof loadWpsHost>>,
  scope = 'sheet',
): WorkbookSnapshot {
  return model.clone(model.host.capture(scope));
}
const chart: SheetAction = {
  type: 'chart',
  rangeId: 'r1',
  chartType: 'line',
  title: 'Sales',
  categoryColumn: 0,
  valueColumns: [1],
};
const source = {
  Data: [
    ['Month', 'Sales'],
    ['Jan', 100],
    ['Feb', 150],
    ['Mar', 120],
  ],
};

test('whole-row insertion affects all columns, fills selected columns and restores formulas, row heights and hidden states', async () => {
  for (const offset of [0, 1, 3]) {
    const m = await loadWpsHost({
      Data: [
        ['a', 'b', 'c'],
        [1, '003', false],
        [2, 0, true],
      ],
      Other: [[null]],
    });
    m.sheets[1].Cells.Item(1, 1).Formula = '=Data!A3';
    m.sheets[0].Range('3:3').RowHeight = 24;
    m.sheets[0].Range('3:3').Hidden = true;
    m.select(1, 2, 3, 1);
    const snapshot = read(m, 'selection');
    assert.equal(snapshot.rowOperations?.available, true);
    const task = new TableToolSession(snapshot);
    const result: any = task.run('officebridge_insert_rows', {
      rangeId: 'r1',
      row: offset,
      count: 2,
      values: [['new'], [0]],
    });
    assert.equal(result.applied, false);
    assert.equal(m.sheets[0].Cells.Item(2, 1).Value2, 1);
    const output = m.host.apply(task.proposal());
    assert.equal(output.changedCells, 2);
    assert.equal(m.sheets[0].Cells.Item(offset + 1, 2).Value2, 'new');
    assert.equal(m.sheets[0].Cells.Item(offset + 2, 2).Value2, 0);
    const destination = offset <= 1 ? 4 : 2;
    assert.equal(m.sheets[0].Cells.Item(destination, 1).Value2, 1);
    assert.equal(m.sheets[0].Cells.Item(destination, 3).Value2, false);
    m.host.undo();
    assert.equal(m.sheets[0].Cells.Item(2, 2).Value2, '003');
    assert.equal(m.sheets[0].Cells.Item(3, 3).Value2, true);
    assert.equal(m.sheets[0].Range('3:3').RowHeight, 24);
    assert.equal(m.sheets[0].Range('3:3').Hidden, true);
    assert.equal(m.sheets[1].Cells.Item(1, 1).Formula, '=Data!A3');
  }
});

test('deleting partially selected rows removes every column, preserves surviving references and undo restores deleted formats', async () => {
  const m = await loadWpsHost({ ...source, Other: [[null]] });
  m.sheets[0].Cells.Item(2, 2).Font.Bold = true;
  m.sheets[0].Range('2:2').RowHeight = 27;
  m.sheets[1].Cells.Item(1, 1).Formula = '=Data!B4';
  m.select(2, 1, 2, 1);
  const snapshot = read(m, 'selection');
  const task = new TableToolSession(snapshot);
  task.run('officebridge_delete_rows', { rangeId: 'r1', row: 0, count: 2 });
  const output = m.host.apply(task.proposal());
  assert.equal(output.rowChanges[0].row, 2);
  assert.equal(m.sheets[0].Cells.Item(2, 1).Value2, 'Mar');
  assert.equal(m.sheets[0].Cells.Item(2, 2).Value2, 120);
  assert.equal(m.sheets[1].Cells.Item(1, 1).Formula, '=Data!B2');
  assert.equal(m.host.recovery().rowChanges[0].count, 2);
  m.host.undo();
  assert.equal(m.sheets[0].Cells.Item(2, 1).Value2, 'Jan');
  assert.equal(m.sheets[0].Cells.Item(2, 2).Font.Bold, true);
  assert.equal(m.sheets[0].Range('2:2').RowHeight, 27);
  assert.equal(m.sheets[1].Cells.Item(1, 1).Formula, '=Data!B4');
});

test('row failures roll back verified mutations; uncertain native exceptions disable undo; hidden zero-height deletion is rejected', async () => {
  for (const mode of ['write', 'formula', 'native', 'partial']) {
    const m = await loadWpsHost({ ...source, Other: [[null]] });
    if (mode === 'formula') m.sheets[1].Cells.Item(1, 1).Formula = '=Data!B2';
    const snapshot = read(m);
    if (mode === 'write') m.sheets[0].failNextCellWrite = true;
    if (mode === 'native') m.sheets[0].failNextRowOperation = true;
    if (mode === 'partial') m.sheets[0].failRowAfterMutation = true;
    const proposal =
      mode === 'write'
        ? rows(snapshot, 'insert', 1, 1, [['New', 0]])
        : rows(snapshot, 'delete', 1);
    assert.throws(
      () => m.host.apply(proposal),
      mode === 'partial' ? /无法确认.*自动恢复已禁用/ : /已恢复/,
    );
    if (mode !== 'partial')
      assert.equal(m.sheets[0].Cells.Item(2, 1).Value2, 'Jan');
    else assert.equal(m.host.recovery().incomplete, true);
    assert.equal(m.host.canUndo(), false);
  }
  const m = await loadWpsHost(source);
  m.app.hiddenHeightZero = true;
  m.sheets[0].Range('2:2').Hidden = true;
  const snapshot = read(m);
  assert.throws(
    () => m.host.apply(rows(snapshot, 'delete', 1)),
    /隐藏行.*行高/,
  );
  assert.equal(m.sheets[0].Cells.Item(2, 1).Value2, 'Jan');
});

test('new schemas reject out-of-scope actions, mixed operations, invalid names, filter arguments and chart sources before mutation', async () => {
  const m = await loadWpsHost(source),
    snapshot = read(m);
  const invalid: any[] = [
    rows(snapshot, 'delete', 4),
    rows(snapshot, 'delete', 3, 2),
    rows(snapshot, 'insert', 0, 101),
    rows(snapshot, 'insert', 0, 1, [['bad']]),
    {
      ...rows(snapshot, 'delete', 1),
      edits: [{ rangeId: 'r1', row: 1, column: 1, value: 7 }],
    },
    { ...rows(snapshot, 'delete', 1), sheetActions: [chart] },
    action(snapshot, { type: 'rename', rangeId: 'r1', name: 'DATA' }),
    action(snapshot, { type: 'rename', rangeId: 'r1', name: 'bad/name' }),
    action(snapshot, { type: 'rename', rangeId: 'r1', name: "'bad" }),
    action(snapshot, {
      type: 'filter',
      rangeId: 'r1',
      conditions: [{ column: 9, operator: 'eq', value: 1 }],
    }),
    action(snapshot, {
      type: 'filter',
      rangeId: 'r1',
      conditions: [{ column: 1, operator: 'gt', value: '100' }],
    }),
    action(snapshot, {
      type: 'filter',
      rangeId: 'r1',
      conditions: [
        { column: 1, operator: 'eq', value: 100 },
        { column: 1, operator: 'ne', value: 120 },
      ],
    }),
    action(snapshot, { ...chart, valueColumns: [0] } as SheetAction),
    action(snapshot, { ...chart, valueColumns: [1, 1] } as SheetAction),
    action(snapshot, {
      ...chart,
      chartType: 'script',
    } as unknown as SheetAction),
    {
      ...action(snapshot, chart),
      sheetActions: [chart, { type: 'rename', rangeId: 'r1', name: 'New' }],
    },
  ];
  for (const proposal of invalid) {
    assert.equal(isEditPlan(snapshot, proposal), false);
    assert.throws(() => m.host.apply(proposal));
  }
  const {
    rowOperations: _rows,
    sheetOperations: _sheets,
    ...noCaps
  } = snapshot;
  assert.equal(isEditPlan(noCaps, rows(snapshot, 'delete', 1)), false);
  assert.equal(isEditPlan(noCaps, action(snapshot, chart)), false);
  assert.equal(
    isWorkbookSnapshot({ ...snapshot, rowOperations: { available: 'true' } }),
    false,
  );
  const readOnly = new TableToolSession(snapshot, false);
  for (const [name, args] of [
    ['officebridge_delete_rows', { rangeId: 'r1', row: 1, count: 1 }],
    ['officebridge_create_chart', { ...chart, type: undefined }],
  ] as const)
    assert.throws(() => readOnly.run(name, args));
  const task = new TableToolSession(snapshot);
  task.run('officebridge_rename_sheet', { rangeId: 'r1', name: 'Renamed' });
  assert.throws(() =>
    task.run('officebridge_delete_rows', { rangeId: 'r1', row: 1, count: 1 }),
  );
  assert.throws(() => task.run('officebridge_clean_range', { rangeId: 'r1' }));
  assert.equal(m.sheets[0].Name, 'Data');
  assert.ok(
    tableTools.some((tool) => tool.name === 'officebridge_filter_range'),
  );
});

test('renaming uses the original sheet, updates cross-sheet references and restores names and references without copying data', async () => {
  const m = await loadWpsHost({ ...source, Other: [[null]] });
  m.sheets[1].Cells.Item(1, 1).Formula = '=Data!B2';
  const snapshot = read(m),
    task = new TableToolSession(snapshot);
  task.run('officebridge_rename_sheet', { rangeId: 'r1', name: '整理结果' });
  const original = m.sheets[0];
  m.host.apply(task.proposal());
  assert.equal(m.sheets[0], original);
  assert.equal(original.Name, '整理结果');
  assert.match(m.sheets[1].Cells.Item(1, 1).Formula, /整理结果/);
  assert.equal(m.sheets.length, 2);
  assert.equal(original.Cells.Item(2, 2).Value2, 100);
  m.host.undo();
  assert.equal(original.Name, 'Data');
  assert.equal(m.sheets[1].Cells.Item(1, 1).Formula, '=Data!B2');
});

test('native filter is non-destructive, handles literal wildcards/numbers/blanks and restores prior criteria and visible rows', async () => {
  for (const conditions of [
    [{ column: 0, operator: 'eq', value: 'Jan' }],
    [
      { column: 1, operator: 'gte', value: 120 },
      { column: 0, operator: 'ne', value: 'Mar' },
    ],
    [{ column: 0, operator: 'contains', value: 'a' }],
  ] as any[]) {
    const m = await loadWpsHost(source);
    m.app.hiddenHeightZero = true;
    const snapshot = read(m);
    m.host.apply(
      action(snapshot, { type: 'filter', rangeId: 'r1', conditions }),
    );
    assert.equal(m.sheets[0].AutoFilterMode, true);
    assert.equal(m.sheets[0].Cells.Item(3, 2).Value2, 150);
    assert.ok(
      [2, 3, 4].some((row) => m.sheets[0].Range(row + ':' + row).Hidden),
    );
    m.host.undo();
    assert.equal(m.sheets[0].AutoFilterMode, false);
    assert.ok(
      [2, 3, 4].every((row) => !m.sheets[0].Range(row + ':' + row).Hidden),
    );
  }
  const m = await loadWpsHost({
    Data: [
      ['Status', 'N'],
      ['*?', 1],
      ['abc', 2],
      [null, 3],
    ],
  });
  let snapshot = read(m);
  m.host.apply(
    action(snapshot, {
      type: 'filter',
      rangeId: 'r1',
      conditions: [{ column: 0, operator: 'eq', value: '*?' }],
    }),
  );
  assert.equal(m.sheets[0].Range('2:2').Hidden, false);
  assert.equal(m.sheets[0].Range('3:3').Hidden, true);
  snapshot = read(m);
  m.host.apply(
    action(snapshot, { type: 'filter', rangeId: 'r1', conditions: [] }),
  );
  assert.equal(m.sheets[0].AutoFilterMode, false);
  m.host.undo();
  assert.equal(m.sheets[0].AutoFilterMode, true);
  assert.equal(m.sheets[0].Range('3:3').Hidden, true);
  snapshot = read(m);
  m.host.apply(
    action(snapshot, {
      type: 'filter',
      rangeId: 'r1',
      conditions: [{ column: 0, operator: 'eq', value: null }],
    }),
  );
  assert.equal(m.sheets[0].Range('4:4').Hidden, false);
  assert.equal(m.sheets[0].Range('2:2').Hidden, true);
  m.host.undo();
  assert.equal(m.sheets[0].Range('2:2').Hidden, false);
});

test('four chart types create only bounded native series, preserve source data and existing charts, then undo only the new chart', async () => {
  for (const chartType of ['line', 'column', 'bar', 'pie'] as const) {
    const m = await loadWpsHost(source),
      snapshot = read(m),
      task = new TableToolSession(snapshot);
    task.run('officebridge_create_chart', {
      rangeId: 'r1',
      chartType,
      title: 'Sales',
      categoryColumn: 0,
      valueColumns: [1],
    });
    assert.equal(m.sheets[0].ChartObjects().Count, 0);
    m.host.apply(task.proposal());
    const collection = m.sheets[0].ChartObjects();
    assert.equal(collection.Count, 1);
    const series = collection.Item(1).Chart.SeriesCollection().Item(1);
    assert.match(series.Formula, /\$A\$2:\$A\$4/);
    assert.match(series.Formula, /\$B\$2:\$B\$4/);
    assert.equal(m.sheets[0].Cells.Item(3, 2).Value2, 150);
    // A second chart must preserve the first and recovery removes only its own.
    const second = read(m);
    assert.equal(second.columnOperations?.available, false);
    assert.equal(second.rowOperations?.available, false);
    assert.equal(second.sheetOperations?.available, true);
    m.host.apply(
      action(second, { ...chart, chartType, title: 'Second' } as SheetAction),
    );
    assert.equal(collection.Count, 2);
    m.host.undo();
    assert.equal(collection.Count, 1);
    assert.equal(collection.Item(1).Chart.ChartTitle.Text, 'Sales');
  }
});

test('monthly billing chart trims categories, sums numeric text exactly and sorts before staging, applying and undoing without source writes', async () => {
  const data = {
    Bills: [
      ['Month', 'Bill', 'Paid'],
      ['2026-10\t', '1,200.100000000000', '0.1'],
      ['2026-02\t', "'0.200000000000", 0.2],
      ['2026-10\t', '-0.100000000000', 0.2],
      ['2026-01\t', ' 0.100000000000 ', '0.1'],
      [null, null, null],
      ['2026-01\t', '0.200000000000', '0.2'],
    ],
  };
  for (const chartType of ['line', 'column', 'bar', 'pie'] as const) {
    const m = await loadWpsHost(data);
    const snapshot = read(m),
      task = new TableToolSession(snapshot);
    const { type: _type, ...args } = {
      ...chart,
      chartType,
      aggregation: 'sum',
      categoryOrder: 'ascending',
      valueColumns: chartType === 'pie' ? [1] : [1, 2],
    } as const;
    const result: any = task.run('officebridge_create_chart', args);
    assert.deepEqual(result.chartSummary.categories, [
      '2026-01',
      '2026-02',
      '2026-10',
    ]);
    assert.deepEqual(result.chartSummary.series[0].values, [0.3, 0.2, 1200]);
    if (chartType !== 'pie')
      assert.deepEqual(result.chartSummary.series[1].values, [0.3, 0.2, 0.3]);
    assert.equal(result.chartSummary.sourceRows, 5);
    assert.equal(result.chartSummary.ignoredBlankRows, 1);
    assert.equal(result.chartSummary.trimmedCategories, 5);
    assert.equal(
      result.chartSummary.convertedTextCells,
      chartType === 'pie' ? 5 : 8,
    );
    assert.equal(m.sheets[0].ChartObjects().Count, 0);
    assert.equal(isEditPlan(snapshot, task.proposal()), true);
    m.host.apply(task.proposal());
    const native = m.sheets[0].ChartObjects().Item(1).Chart;
    assert.equal(
      native.SeriesCollection().Item(1).Formula,
      '=SERIES("Bill",{"2026-01","2026-02","2026-10"},{0.3,0.2,1200},1)',
    );
    assert.equal(m.sheets[0].Cells.Item(2, 1).Value2, '2026-10\t');
    assert.equal(m.sheets[0].Cells.Item(2, 2).Value2, '1,200.100000000000');
    assert.equal(m.sheets.length, 1);
    m.host.undo();
    assert.equal(m.sheets[0].ChartObjects().Count, 0);
    assert.equal(m.sheets[0].Cells.Item(2, 2).Value2, '1,200.100000000000');
  }
});

test('aggregation rejects invalid or missing amounts, mixed categories, total rows and precision errors with row diagnostics and no staged or native edits', async () => {
  const invalid = [
    [['2026-01', '¥10'], /第 2 行.*十进制/],
    [['2026-01', '1,23'], /第 2 行.*十进制/],
    [['2026-01', '10%'], /第 2 行.*十进制/],
    [['2026-01', true], /第 2 行.*不是数值/],
    [['2026-01', null], /第 2 行.*缺失/],
    [['2026-01', '  '], /第 2 行.*十进制/],
    [[null, '10'], /第 2 行.*分类缺失/],
    [['合计', '10'], /第 2 行.*合计/],
    [['2026-01', '1.000000000000001'], /第 2 行.*精度/],
    [['2026-01', '9007199254740993'], /第 2 行.*精度/],
    [['2026-01', '0.0000000000000001'], /第 2 行.*精度/],
  ] as const;
  for (const [badRow, error] of invalid) {
    const m = await loadWpsHost({ Data: [['Month', 'Bill'], [...badRow]] });
    const snapshot = read(m),
      task = new TableToolSession(snapshot);
    const { type: _type, ...args } = { ...chart, aggregation: 'sum' };
    assert.throws(() => task.run('officebridge_create_chart', args), error);
    assert.equal(task.proposal(), null);
    assert.equal(m.sheets[0].ChartObjects().Count, 0);
    assert.equal(
      isEditPlan(snapshot, action(snapshot, { ...chart, aggregation: 'sum' })),
      false,
    );
  }
  for (const values of [
    [
      ['2026-01', 1],
      [1, 2],
    ],
    [
      ['2026-01', '999999999999999'],
      ['2026-01', '0.1'],
    ],
  ]) {
    const m = await loadWpsHost({ Data: [['Month', 'Bill'], ...values] });
    const task = new TableToolSession(read(m));
    assert.throws(
      () =>
        task.run('officebridge_create_chart', {
          rangeId: 'r1',
          chartType: 'line',
          title: 'Bill',
          categoryColumn: 0,
          valueColumns: [1],
          aggregation: 'sum',
        }),
      /类型混合|精度/,
    );
    assert.equal(task.proposal(), null);
  }
});

test('aggregation rules remain explicit and bounded, escape labels, and share native chart rollback and stale guards', async () => {
  const m = await loadWpsHost({
    Data: [
      ['Month', 'Bill'],
      ['2026-01\t', '0.1'],
      ['2026-01', '0.2'],
    ],
  });
  const snapshot = read(m);
  assert.equal(isEditPlan(snapshot, action(snapshot, chart)), false); // No implicit conversion.
  for (const extra of [
    { aggregation: 'avg' },
    { categoryOrder: 'ascending' },
    { aggregation: 'sum', categoryOrder: 'descending' },
    { aggregation: 'sum', data: [[1, 2]] },
    { aggregation: 'sum', valueColumns: [2] },
  ])
    assert.equal(
      isEditPlan(
        snapshot,
        action(snapshot, { ...chart, ...extra } as SheetAction),
      ),
      false,
    );
  const aggregated = { ...chart, aggregation: 'sum' } as SheetAction;
  m.sheets[0].failNextSeries = true;
  assert.throws(() => m.host.apply(action(snapshot, aggregated)), /已恢复/);
  assert.equal(m.sheets[0].ChartObjects().Count, 0);
  const fresh = read(m);
  m.sheets[0].Cells.Item(2, 2).Value2 = '4';
  assert.throws(() => m.host.apply(action(fresh, aggregated)), /已有变化/);
  const latest = read(m);
  m.host.apply(action(latest, aggregated));
  m.sheets[0].ChartObjects().Item(1).Chart.SeriesCollection().Item(1).Formula =
    '=SERIES("Bill",{"2026-01"},{99},1)';
  assert.throws(() => m.host.undo(), /已有变化/);
  const escaped = await loadWpsHost({
    Data: [
      ['Category', 'A"B'],
      ['x"},DDE!A1,{"', '0.1'],
    ],
  });
  const range = read(escaped).ranges[0]!;
  assert.equal(
    summarySeriesFormula(summarizeChart(range, aggregated as any), 0),
    '=SERIES("A""B",{"x""},DDE!A1,{"""},{0.1},1)',
  );
  const long = await loadWpsHost({
    Data: [
      ['Category', 'Bill'],
      ['x'.repeat(9000), '1'],
    ],
  });
  assert.throws(
    () => summarizeChart(read(long).ranges[0]!, aggregated as any),
    /8192/,
  );
});

test('grouped chart handles small native numbers, signed decimals, year boundaries and numeric category order without binary sum noise', async () => {
  for (const [data, expected] of [
    [
      [
        ['2026-02', 1e-7],
        ['2025-12\t', '+0.1000'],
        ['2025-12', '-0.2000'],
      ],
      { categories: ['2025-12', '2026-02'], values: [-0.1, 0.0000001] },
    ],
    [
      [
        [10, 1],
        [2, 2],
        [10, 3],
      ],
      { categories: [2, 10], values: [2, 4] },
    ],
  ] as const) {
    const m = await loadWpsHost({
      Data: [['Category', 'Amount'], ...data.map((row) => [...row])],
    });
    const snapshot = read(m);
    const change = {
      ...chart,
      aggregation: 'sum',
      categoryOrder: 'ascending',
    } as const;
    const summary = summarizeChart(snapshot.ranges[0]!, change);
    assert.deepEqual(summary.categories, expected.categories);
    assert.deepEqual(summary.series[0]!.values, expected.values);
    assert.ok(!summarySeriesFormula(summary, 0).includes('1e-7'));
    m.host.apply(action(snapshot, change));
    m.host.undo();
  }
});

test('action guards refuse stale data, later chart edits and changed filters; failed actions restore or report uncertainty', async () => {
  for (const mutate of [
    (m: any) => {
      m.sheets[0].ChartObjects().Item(1).Chart.ChartTitle.Text = 'Changed';
    },
    (m: any) => {
      m.sheets[0].ChartObjects().Item(1).Left = 99;
    },
    (m: any) => {
      m.sheets[0]
        .ChartObjects()
        .Item(1)
        .Chart.SeriesCollection()
        .Item(1).Formula = '=SERIES("Changed",Data!A2:A4,Data!B2:B4,1)';
    },
    (m: any) => {
      m.sheets[0].Cells.Item(2, 2).Value2 = 999;
    },
  ]) {
    const m = await loadWpsHost(source);
    m.host.apply(action(read(m), chart));
    mutate(m);
    assert.throws(() => m.host.undo(), /已有变化/);
    assert.equal(m.sheets[0].ChartObjects().Count, 1);
  }
  for (const mode of ['rename', 'filter', 'chart', 'partial-chart']) {
    const m = await loadWpsHost(source),
      snapshot = read(m);
    let proposal: EditPlan;
    if (mode === 'rename') {
      m.sheets[0].failRenameAfterMutation = true;
      proposal = action(snapshot, {
        type: 'rename',
        rangeId: 'r1',
        name: 'New',
      });
    } else if (mode === 'filter') {
      m.sheets[0].failNextFilter = true;
      proposal = action(snapshot, {
        type: 'filter',
        rangeId: 'r1',
        conditions: [{ column: 0, operator: 'eq', value: 'Jan' }],
      });
    } else {
      m.sheets[0][
        mode === 'chart' ? 'failNextSeries' : 'failChartAddAfterMutation'
      ] = true;
      proposal = action(snapshot, chart);
    }
    assert.throws(
      () => m.host.apply(proposal),
      mode === 'partial-chart' ? /无法确认完整恢复/ : /已恢复/,
    );
    assert.equal(m.sheets[0].Name, 'Data');
    assert.equal(m.sheets[0].AutoFilterMode, false);
    assert.equal(
      m.sheets[0].ChartObjects().Count,
      mode === 'partial-chart' ? 1 : 0,
    );
    if (mode === 'partial-chart')
      assert.equal(m.host.recovery().incomplete, true);
  }
  const m = await loadWpsHost(source),
    snapshot = read(m);
  m.sheets[0].Cells.Item(2, 2).Value2 = 999;
  assert.throws(() => m.host.apply(action(snapshot, chart)), /已有变化/);
  assert.equal(m.sheets[0].ChartObjects().Count, 0);
});

test('existing other-range filters and manual hidden rows are not silently replaced; protection and complex data expose unavailable capability', async () => {
  const m = await loadWpsHost(source);
  m.sheets[0].Range('A1:B4').AutoFilter(1, '=Jan');
  m.select(1, 1, 3, 2);
  let snapshot = read(m, 'selection');
  assert.throws(
    () =>
      m.host.apply(
        action(snapshot, { type: 'filter', rangeId: 'r1', conditions: [] }),
      ),
    /其他范围筛选/,
  );
  m.sheets[0].AutoFilterMode = false;
  m.sheets[0].Range('2:2').Hidden = true;
  snapshot = read(m);
  assert.throws(
    () =>
      m.host.apply(
        action(snapshot, {
          type: 'filter',
          rangeId: 'r1',
          conditions: [{ column: 0, operator: 'eq', value: 'Jan' }],
        }),
      ),
    /手动隐藏/,
  );
  m.sheets[0].ProtectContents = true;
  snapshot = read(m);
  assert.equal(snapshot.rowOperations?.available, false);
  assert.equal(snapshot.sheetOperations?.available, false);
});
