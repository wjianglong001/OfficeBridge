import {
  TABLE_LIMITS,
  rangeAddress,
  type WorkbookSnapshot,
  type TableRange,
  isCellValue,
} from './table.js';

export type RangeDescriptor = Omit<
  TableRange,
  'values' | 'formulas' | 'styles'
>;
export type WorkbookDescriptor = Omit<WorkbookSnapshot, 'ranges'> & {
  readMode: 'on-demand';
  ranges: RangeDescriptor[];
  inline?: { rangeId: string; values: TableRange['values'] }[];
  navigation?: {
    mode: 'auto' | 'locked' | 'context';
    selectionProvided?: boolean;
    activeSheet: string;
    allowedSheets: string[];
    anchor: RangeDescriptor;
    previousRange?: RangeDescriptor;
  };
};
export interface HostToolReply {
  result: unknown;
  plan?: unknown;
  snapshot?: unknown;
  range?: RangeDescriptor;
}
export function validRange(r: any): r is RangeDescriptor {
  return (
    !!r &&
    typeof r === 'object' &&
    !Array.isArray(r) &&
    typeof r.id === 'string' &&
    r.id.length > 0 &&
    r.id.length <= 1024 &&
    typeof r.sheetName === 'string' &&
    r.sheetName.length > 0 &&
    r.sheetName.length <= 31 &&
    Object.keys(r).every((k) =>
      [
        'id',
        'sheetName',
        'address',
        'startRow',
        'startColumn',
        'rowCount',
        'columnCount',
      ].includes(k),
    ) &&
    [r.startRow, r.startColumn, r.rowCount, r.columnCount].every(
      (n) => Number.isSafeInteger(n) && n > 0,
    ) &&
    r.startRow + r.rowCount - 1 <= 1048576 &&
    r.startColumn + r.columnCount - 1 <= 16384 &&
    r.address ===
      rangeAddress(r.startRow, r.startColumn, r.rowCount, r.columnCount)
  );
}
export function rangeWithin(r: RangeDescriptor, allowed: RangeDescriptor) {
  return (
    r.sheetName === allowed.sheetName &&
    r.startRow >= allowed.startRow &&
    r.startColumn >= allowed.startColumn &&
    r.startRow + r.rowCount <= allowed.startRow + allowed.rowCount &&
    r.startColumn + r.columnCount <= allowed.startColumn + allowed.columnCount
  );
}
export function authorizedRange(d: WorkbookDescriptor, r: RangeDescriptor) {
  return (
    validRange(r) &&
    (d.navigation?.mode === 'auto' || d.navigation?.mode === 'context'
      ? d.navigation.allowedSheets.includes(r.sheetName)
      : d.ranges.some((a) => rangeWithin(r, a)))
  );
}
export function validateQuestion(p: any): {
  question: string;
  options: string[];
} {
  if (
    !p ||
    typeof p.question !== 'string' ||
    !p.question.trim() ||
    p.question.length > 300 ||
    !Array.isArray(p.options) ||
    p.options.length < 2 ||
    p.options.length > 3 ||
    p.options.some(
      (o: any) => typeof o !== 'string' || !o.trim() || o.length > 100,
    ) ||
    new Set(p.options).size !== p.options.length
  )
    throw new Error('请提供一个简短问题和 2–3 个不同选项。');
  return { question: p.question, options: p.options };
}
export type HostToolCall = (
  name: string,
  args: unknown,
) => Promise<HostToolReply>;
/** Rejected before host access; the Agent may correct parameters in this task. */
export class ToolParameterError extends Error {
  override name = 'ToolParameterError';
  readonly code: string = 'invalid_tool_parameters';
  constructor(
    readonly field: string,
    message: string,
  ) {
    super(message);
  }
}
export class ReadParameterError extends ToolParameterError {
  override name = 'ReadParameterError';
  override readonly code = 'invalid_read_parameters';
}
export function readSourceRange(range: RangeDescriptor) {
  return {
    rangeId: range.id,
    address: range.address,
    startRow: range.startRow,
    startColumn: range.startColumn,
    rowCount: range.rowCount,
    columnCount: range.columnCount,
  };
}
export function readOptions(
  range: RangeDescriptor,
  params: Record<string, any>,
) {
  const columns: number[] =
    params.columns ?? Array.from({ length: range.columnCount }, (_, i) => i);
  const rowOffset = params.rowOffset ?? 0;
  const rowCount = params.rowCount ?? range.rowCount - rowOffset;
  const fields: string[] = params.fields ?? ['values'];
  if (
    !Array.isArray(columns) ||
    !columns.length ||
    new Set(columns).size !== columns.length ||
    columns.some(
      (c) => !Number.isSafeInteger(c) || c < 0 || c >= range.columnCount,
    )
  )
    throw new ReadParameterError(
      'columns',
      `columns 必须是不重复的整数下标，范围为 0–${range.columnCount - 1}（相对于当前完整范围 ${range.address}）。不能使用列字母、从 1 开始的列号或上一次投影结果的列位置；需要其他范围时先调用 resolve_range。`,
    );
  if (
    !Number.isSafeInteger(rowOffset) ||
    rowOffset < 0 ||
    rowOffset >= range.rowCount
  )
    throw new ReadParameterError(
      'rowOffset',
      `rowOffset 必须为 0–${range.rowCount - 1} 的整数，0 表示当前完整范围 ${range.address} 的首行。跳过表头可用 rowOffset=1，rowCount 省略时读取剩余行。`,
    );
  if (
    !Number.isSafeInteger(rowCount) ||
    rowCount < 1 ||
    rowOffset + rowCount > range.rowCount
  )
    throw new ReadParameterError(
      'rowCount',
      `当前完整范围 ${range.address} 共 ${range.rowCount} 行，rowOffset=${rowOffset} 后最多读取 ${range.rowCount - rowOffset} 行。rowCount 必须为 1–${range.rowCount - rowOffset} 的整数；省略 rowCount 可自动读取剩余行。`,
    );
  if (
    !Array.isArray(fields) ||
    !fields.length ||
    new Set(fields).size !== fields.length ||
    fields.some((f) => !['values', 'formulas', 'styles'].includes(f))
  )
    throw new ReadParameterError(
      'fields',
      'fields 必须是不重复的数组，只接受 "values"、"formulas"、"styles"；普通取值可省略 fields，或传 ["values"]。',
    );
  return { columns, rowOffset, rowCount, fields };
}
export function projectRange(range: TableRange, params: Record<string, any>) {
  const o = readOptions(range, params);
  const result: Record<string, unknown> = {
    id: range.id,
    rangeId: range.id,
    sheetName: range.sheetName,
    address: range.address,
    startRow: range.startRow + o.rowOffset,
    startColumn: range.startColumn,
    rowOffset: o.rowOffset,
    rowCount: o.rowCount,
    columnCount: o.columns.length,
    columnIndices: o.columns,
    sourceRange: readSourceRange(range),
  };
  for (const field of o.fields)
    result[field] = (range as any)[field]
      .slice(o.rowOffset, o.rowOffset + o.rowCount)
      .map((r: any[]) => o.columns.map((c) => r[c]));
  return result;
}
export function isWorkbookDescriptor(
  input: unknown,
): input is WorkbookDescriptor {
  const v = input as WorkbookDescriptor | undefined;
  if (
    !v ||
    typeof v !== 'object' ||
    Array.isArray(v) ||
    Object.keys(v).some(
      (k) =>
        ![
          'host',
          'readMode',
          'connectionId',
          'snapshotId',
          'workbookName',
          'workbookFullName',
          'scope',
          'capturedAt',
          'sheetNames',
          'ranges',
          'inline',
          'columnOperations',
          'rowOperations',
          'sheetOperations',
          'navigation',
        ].includes(k),
    ) ||
    v.readMode !== 'on-demand' ||
    v.host !== 'wps' ||
    !['connectionId', 'snapshotId', 'workbookName', 'workbookFullName'].every(
      (k) =>
        typeof (v as any)[k] === 'string' &&
        (v as any)[k].length > 0 &&
        (v as any)[k].length <= 1024,
    ) ||
    !['selection', 'rows', 'columns', 'sheet', 'workbook'].includes(v.scope) ||
    typeof v.capturedAt !== 'string' ||
    !Number.isFinite(Date.parse(v.capturedAt)) ||
    !Array.isArray(v.sheetNames) ||
    !v.sheetNames.length ||
    v.sheetNames.length > 500 ||
    v.sheetNames.some(
      (n) => typeof n !== 'string' || !n.length || n.length > 31,
    ) ||
    new Set(v.sheetNames).size !== v.sheetNames.length ||
    !Array.isArray(v.ranges) ||
    !v.ranges.length ||
    v.ranges.length > TABLE_LIMITS.ranges ||
    (v.scope !== 'workbook' && v.ranges.length !== 1)
  )
    return false;
  const ids = new Set<string>(),
    sheets = new Set<string>();
  for (const r of v.ranges) {
    if (
      !r ||
      typeof r.id !== 'string' ||
      !r.id.length ||
      ids.has(r.id) ||
      (sheets.has(r.sheetName) &&
        !['auto', 'context'].includes(v.navigation?.mode || '')) ||
      !v.sheetNames.includes(r.sheetName) ||
      Object.keys(r).some(
        (k) =>
          ![
            'id',
            'sheetName',
            'address',
            'startRow',
            'startColumn',
            'rowCount',
            'columnCount',
          ].includes(k),
      ) ||
      ![r.startRow, r.startColumn, r.rowCount, r.columnCount].every(
        (n) => Number.isSafeInteger(n) && n > 0,
      ) ||
      r.startRow + r.rowCount - 1 > 1048576 ||
      r.startColumn + r.columnCount - 1 > 16384 ||
      r.address !==
        rangeAddress(r.startRow, r.startColumn, r.rowCount, r.columnCount)
    )
      return false;
    ids.add(r.id);
    sheets.add(r.sheetName);
  }
  const nav = v.navigation;
  if (nav !== undefined) {
    if (
      !nav ||
      typeof nav !== 'object' ||
      Array.isArray(nav) ||
      Object.keys(nav).some(
        (k) =>
          ![
            'mode',
            'activeSheet',
            'allowedSheets',
            'anchor',
            'previousRange',
            'selectionProvided',
          ].includes(k),
      ) ||
      !['auto', 'locked', 'context'].includes(nav.mode) ||
      !v.sheetNames.includes(nav.activeSheet) ||
      !Array.isArray(nav.allowedSheets) ||
      !nav.allowedSheets.length ||
      nav.allowedSheets.some((s) => !v.sheetNames.includes(s)) ||
      new Set(nav.allowedSheets).size !== nav.allowedSheets.length ||
      !nav.allowedSheets.includes(nav.activeSheet) ||
      !validRange(nav.anchor) ||
      nav.anchor.sheetName !== nav.activeSheet ||
      (nav.previousRange !== undefined &&
        (!validRange(nav.previousRange) ||
          !nav.allowedSheets.includes(nav.previousRange.sheetName)))
    )
      return false;
    if (
      ['auto', 'context'].includes(nav.mode) &&
      (v.scope !== 'workbook' ||
        v.ranges.some((r) => !nav.allowedSheets.includes(r.sheetName)))
    )
      return false;
    if (
      nav.mode === 'locked' &&
      (nav.allowedSheets.length !== 1 ||
        v.ranges.length !== 1 ||
        JSON.stringify(v.ranges[0]) !== JSON.stringify(nav.anchor))
    )
      return false;
  }
  for (const key of [
    'columnOperations',
    'rowOperations',
    'sheetOperations',
  ] as const) {
    const c = v[key];
    if (
      c !== undefined &&
      (!c ||
        typeof c !== 'object' ||
        Array.isArray(c) ||
        Object.keys(c).some((k) => !['available', 'reason'].includes(k)) ||
        typeof c.available !== 'boolean' ||
        (c.reason !== undefined &&
          (typeof c.reason !== 'string' || c.reason.length > 1024)))
    )
      return false;
  }
  if (v.inline !== undefined) {
    if (
      !Array.isArray(v.inline) ||
      !v.inline.every(
        (r) =>
          r &&
          typeof r === 'object' &&
          !Array.isArray(r) &&
          Object.keys(r).every((k) => ['rangeId', 'values'].includes(k)),
      ) ||
      new Set(v.inline.map((r) => r.rangeId)).size !== v.inline.length
    )
      return false;
    let cells = 0;
    for (const item of v.inline) {
      const range = v.ranges.find((r) => r.id === item.rangeId);
      if (
        !range ||
        !Array.isArray(item.values) ||
        item.values.length !== range.rowCount ||
        item.values.some(
          (row) =>
            !Array.isArray(row) ||
            row.length !== range.columnCount ||
            !row.every(isCellValue),
        )
      )
        return false;
      cells += range.rowCount * range.columnCount;
    }
    if (cells > 16 && nav?.mode !== 'context') return false;
  }
  if (
    nav?.mode === 'context' &&
    (JSON.stringify(v.ranges.find((r) => r.id === nav.anchor.id)) !==
      JSON.stringify(nav.anchor) ||
      typeof nav.selectionProvided !== 'boolean' ||
      (nav.selectionProvided
        ? !v.inline ||
          v.inline.length !== 1 ||
          v.inline[0]!.rangeId !== nav.anchor.id
        : v.inline !== undefined))
  )
    return false;
  return (
    new TextEncoder().encode(JSON.stringify(v)).length <= TABLE_LIMITS.bytes
  );
}
export function authorizedSnapshot(
  descriptor: WorkbookDescriptor,
  snapshot: WorkbookSnapshot,
): boolean {
  return (
    snapshot.snapshotId === descriptor.snapshotId &&
    snapshot.connectionId === descriptor.connectionId &&
    snapshot.workbookFullName === descriptor.workbookFullName &&
    snapshot.scope === descriptor.scope &&
    JSON.stringify(snapshot.sheetNames) ===
      JSON.stringify(descriptor.sheetNames) &&
    snapshot.ranges.every((r) => {
      const allowed = descriptor.ranges.find((a) => a.id === r.id);
      return (
        allowed &&
        r.sheetName === allowed.sheetName &&
        r.startRow === allowed.startRow &&
        r.startColumn === allowed.startColumn &&
        r.rowCount <= allowed.rowCount &&
        r.columnCount <= allowed.columnCount
      );
    })
  );
}
