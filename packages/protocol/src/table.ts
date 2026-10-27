import type { AnalysisSelection, CellValue } from './index.js';
import { summarizeChart } from './chart.js';
export { summarizeChart, summarySeriesFormula } from './chart.js';

export const TABLE_LIMITS = {
  bytes: 512000,
  ranges: 50,
  newSheets: 3,
  columns: 20,
  rows: 100,
} as const;
export type Scope = 'selection' | 'rows' | 'columns' | 'sheet' | 'workbook';
export interface CellStyle {
  numberFormat: string;
  bold: boolean;
  fillColor: number;
  fillPattern: number;
}
export interface TableRange {
  id: string;
  sheetName: string;
  address: string;
  startRow: number;
  startColumn: number;
  rowCount: number;
  columnCount: number;
  values: CellValue[][];
  formulas: (string | null)[][];
  styles: (CellStyle | null)[][];
}
export interface WorkbookSnapshot {
  host: 'wps';
  connectionId: string;
  snapshotId: string;
  workbookName: string;
  workbookFullName: string;
  scope: Scope;
  capturedAt: string;
  sheetNames: string[];
  ranges: TableRange[];
  columnOperations?: { available: boolean; reason?: string };
  rowOperations?: { available: boolean; reason?: string };
  sheetOperations?: { available: boolean; reason?: string };
}
export interface CellEdit {
  rangeId: string;
  row: number;
  column: number;
  value?: CellValue;
  formula?: string;
  format?: { numberFormat?: string; bold?: boolean; fillColor?: number };
}
export interface ResultSheet {
  name: string;
  values: CellValue[][];
}
export interface ColumnChange {
  rangeId: string;
  action: 'insert' | 'delete';
  column: number;
  count: number;
  values?: CellValue[][];
}
export interface EditPlan {
  snapshotId: string;
  connectionId: string;
  workbookFullName: string;
  edits: CellEdit[];
  newSheets: ResultSheet[];
  columnChanges?: ColumnChange[];
  rowChanges?: RowChange[];
  sheetActions?: SheetAction[];
}
export interface RowChange {
  rangeId: string;
  action: 'insert' | 'delete';
  row: number;
  count: number;
  values?: CellValue[][];
}
export interface FilterCondition {
  column: number;
  operator: 'eq' | 'ne' | 'gt' | 'gte' | 'lt' | 'lte' | 'contains';
  value: CellValue;
}
export type SheetAction =
  | { type: 'rename'; rangeId: string; name: string }
  | { type: 'filter'; rangeId: string; conditions: FilterCondition[] }
  | {
      type: 'chart';
      rangeId: string;
      chartType: 'line' | 'column' | 'bar' | 'pie';
      title: string;
      categoryColumn: number;
      valueColumns: number[];
      aggregation?: 'sum';
      categoryOrder?: 'source' | 'ascending';
    };

const object = (value: unknown): value is Record<string, any> =>
  typeof value === 'object' && value !== null && !Array.isArray(value);
const owns = (value: object, key: PropertyKey) =>
  Object.prototype.hasOwnProperty.call(value, key);
const text = (value: unknown, max = 1024): value is string =>
  typeof value === 'string' && value.length > 0 && value.length <= max;
export const isCellValue = (value: unknown): value is CellValue =>
  value === null ||
  typeof value === 'boolean' ||
  (typeof value === 'string' && value.length <= 10000) ||
  (typeof value === 'number' && Number.isFinite(value));
const integer = (value: unknown, min: number, max: number) =>
  Number.isSafeInteger(value) &&
  (value as number) >= min &&
  (value as number) <= max;
const only = (value: Record<string, unknown>, keys: string[]) =>
  Object.keys(value).every((key) => keys.includes(key));
function matrix(
  value: unknown,
  rows: number,
  columns: number,
  valid: (cell: unknown) => boolean,
): boolean {
  return (
    Array.isArray(value) &&
    value.length === rows &&
    value.every(
      (row) => Array.isArray(row) && row.length === columns && row.every(valid),
    )
  );
}
function isStyle(value: unknown): boolean {
  return (
    value === null ||
    (object(value) &&
      typeof value.numberFormat === 'string' &&
      value.numberFormat.length <= 256 &&
      typeof value.bold === 'boolean' &&
      integer(value.fillColor, -2147483648, 2147483647) &&
      integer(value.fillPattern, -5000, 100))
  );
}
export function cellAddress(row: number, column: number): string {
  let letters = '';
  for (let index = column; index > 0; index = Math.floor((index - 1) / 26))
    letters = String.fromCharCode(65 + ((index - 1) % 26)) + letters;
  return letters + row;
}
export function rangeAddress(
  row: number,
  column: number,
  rows: number,
  columns: number,
): string {
  const first = cellAddress(row, column);
  const last = cellAddress(row + rows - 1, column + columns - 1);
  return first === last ? first : first + ':' + last;
}
export function isWorkbookSnapshot(value: unknown): value is WorkbookSnapshot {
  if (
    !object(value) ||
    value.host !== 'wps' ||
    !['connectionId', 'snapshotId', 'workbookName', 'workbookFullName'].every(
      (key) => text(value[key]),
    ) ||
    !['selection', 'rows', 'columns', 'sheet', 'workbook'].includes(
      value.scope,
    ) ||
    !text(value.capturedAt) ||
    !Number.isFinite(Date.parse(value.capturedAt)) ||
    !Array.isArray(value.sheetNames) ||
    !value.sheetNames.length ||
    value.sheetNames.length > 500 ||
    !value.sheetNames.every((name: unknown) => text(name, 31)) ||
    new Set(value.sheetNames).size !== value.sheetNames.length ||
    !Array.isArray(value.ranges) ||
    !value.ranges.length ||
    value.ranges.length > TABLE_LIMITS.ranges
  )
    return false;
  for (const key of ['columnOperations', 'rowOperations', 'sheetOperations']) {
    const capability = value[key];
    if (
      capability !== undefined &&
      (!object(capability) ||
        !only(capability, ['available', 'reason']) ||
        typeof capability.available !== 'boolean' ||
        (capability.reason !== undefined && !text(capability.reason)))
    )
      return false;
  }
  const ids = new Set<string>();
  const sheets = new Set<string>();
  for (const range of value.ranges) {
    if (
      !object(range) ||
      !text(range.id) ||
      ids.has(range.id) ||
      !value.sheetNames.includes(range.sheetName) ||
      sheets.has(range.sheetName) ||
      !integer(range.startRow, 1, 1048576) ||
      !integer(range.startColumn, 1, 16384) ||
      !integer(range.rowCount, 1, 1048576) ||
      !integer(range.columnCount, 1, 16384) ||
      range.startRow + range.rowCount - 1 > 1048576 ||
      range.startColumn + range.columnCount - 1 > 16384 ||
      range.address !==
        rangeAddress(
          range.startRow,
          range.startColumn,
          range.rowCount,
          range.columnCount,
        )
    )
      return false;
    if (
      !matrix(range.values, range.rowCount, range.columnCount, isCellValue) ||
      !matrix(
        range.formulas,
        range.rowCount,
        range.columnCount,
        (cell) => cell === null || (text(cell, 8192) && cell.startsWith('=')),
      ) ||
      !matrix(range.styles, range.rowCount, range.columnCount, isStyle)
    )
      return false;
    ids.add(range.id);
    sheets.add(range.sheetName);
  }
  return (
    (value.scope === 'workbook' || value.ranges.length === 1) &&
    new TextEncoder().encode(JSON.stringify(value)).length <= TABLE_LIMITS.bytes
  );
}

const functions = new Set(
  'SUM SUMIF SUMIFS SUMPRODUCT SUBTOTAL AGGREGATE AVERAGE AVERAGEIF AVERAGEIFS MIN MAX MEDIAN ROUND ROUNDUP ROUNDDOWN INT MOD ABS IF IFS IFERROR IFNA AND OR NOT TRUE FALSE COUNT COUNTA COUNTBLANK COUNTIF COUNTIFS LEN TRIM CLEAN UPPER LOWER PROPER LEFT RIGHT MID FIND SEARCH SUBSTITUTE REPLACE TEXT VALUE CONCAT CONCATENATE TEXTJOIN DATE YEAR MONTH DAY HOUR MINUTE SECOND TIME TODAY NOW DAYS DATEDIF EDATE EOMONTH WEEKDAY WEEKNUM NETWORKDAYS WORKDAY INDEX MATCH XMATCH XLOOKUP VLOOKUP HLOOKUP ISBLANK ISNUMBER ISTEXT ISERROR EXACT UNIQUE SORT FILTER SEQUENCE ROW COLUMN ROWS COLUMNS CHOOSE RANK RANK.EQ LARGE SMALL'.split(
    ' ',
  ),
);
export function validFormula(value: unknown): value is string {
  if (
    !text(value, 1000) ||
    !value.startsWith('=') ||
    /[\[\]!|{}\\\r\n]/.test(value) ||
    /(?:https?:|file:|ftp:)/i.test(value)
  )
    return false;
  const withoutStrings = value.replace(/"(?:[^"]|"")*"/g, '""');
  if (withoutStrings.includes('"') && withoutStrings.match(/"/g)!.length % 2)
    return false;
  for (const match of withoutStrings.matchAll(
    /([A-Za-z_][A-Za-z0-9_.]*)\s*\(/g,
  ))
    if (!functions.has(match[1]!.toUpperCase())) return false;
  // No named ranges, XLM names, implicit external data, or local-language functions.
  const tokens = withoutStrings
    .replace(/""/g, '')
    .replace(/[A-Za-z_][A-Za-z0-9_.]*\s*\(/g, '(')
    .replace(/\$?[A-Za-z]{1,3}\$?\d+/g, '')
    .replace(/\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}/g, '')
    .replace(/\b(?:TRUE|FALSE)\b/gi, '');
  return !/[A-Za-z_\u0080-\uffff]/.test(tokens);
}
export function validateEdits(
  snapshot: WorkbookSnapshot,
  value: unknown,
): value is CellEdit[] {
  return editParameterIssue(snapshot.ranges, value) === null;
}
/** Parameter-only checks shared by bridge preflight and final snapshot validation. */
export function editParameterIssue(
  ranges: readonly Pick<TableRange, 'id' | 'rowCount' | 'columnCount'>[],
  value: unknown,
): { field: string; message: string } | null {
  if (!Array.isArray(value))
    return { field: 'edits', message: 'edits 必须是修改项数组。' };
  const positions = new Set<string>();
  for (const [index, edit] of value.entries()) {
    const issue = (field: string, message: string) => ({
      field: `edits[${index}]${field ? '.' + field : ''}`,
      message,
    });
    if (
      !object(edit) ||
      !only(edit, ['rangeId', 'row', 'column', 'value', 'formula', 'format'])
    )
      return issue(
        '',
        '修改项必须是对象，且仅包含 rangeId、row、column、value、formula、format。',
      );
    const range = ranges.find((item) => item.id === edit.rangeId);
    if (!range)
      return issue(
        'rangeId',
        'rangeId 必须属于本轮已登记范围，请先按需定位目标。',
      );
    if (!integer(edit.row, 0, range.rowCount - 1))
      return issue(
        'row',
        `row 为范围内零基偏移，合法值为 0–${range.rowCount - 1}。`,
      );
    if (!integer(edit.column, 0, range.columnCount - 1))
      return issue(
        'column',
        `column 为范围内零基偏移，合法值为 0–${range.columnCount - 1}。`,
      );
    const key = `${edit.rangeId}:${edit.row}:${edit.column}`;
    if (positions.has(key))
      return issue('', '同一次调用不得重复修改同一范围中的同一单元格。');
    positions.add(key);
    const hasValue = owns(edit, 'value');
    const hasFormula = owns(edit, 'formula');
    if (hasValue && hasFormula)
      return issue(
        'formula',
        'value 与 formula 不能同时提供；写公式时仅使用 formula 字段。',
      );
    if (
      hasValue &&
      (!isCellValue(edit.value) ||
        (typeof edit.value === 'string' && /^\s*=/.test(edit.value)))
    )
      return issue(
        'value',
        'value 必须是有效单元格值；以等号开头的公式请使用 formula 字段。',
      );
    if (hasFormula && !validFormula(edit.formula)) {
      if (typeof edit.formula === 'string' && !edit.formula.startsWith('='))
        return issue(
          'formula',
          'formula 必须以等号 = 开头，例如 =SUM(E2:E11)；请修正后重新提交。',
        );
      return issue(
        'formula',
        'formula 必须是长度不超过 1000 的公式，使用允许的常见英文函数及本表引用，禁止外部或跨表引用。',
      );
    }
    if (edit.format !== undefined) {
      if (
        !object(edit.format) ||
        !Object.keys(edit.format).length ||
        !only(edit.format, ['numberFormat', 'bold', 'fillColor']) ||
        (edit.format.numberFormat !== undefined &&
          (!text(edit.format.numberFormat, 128) ||
            /[\r\n]/.test(edit.format.numberFormat))) ||
        (edit.format.bold !== undefined &&
          typeof edit.format.bold !== 'boolean') ||
        (edit.format.fillColor !== undefined &&
          !integer(edit.format.fillColor, 0, 16777215))
      )
        return issue(
          'format',
          'format 仅支持有效的 numberFormat、布尔 bold 和 0–16777215 的整数 fillColor。',
        );
    }
    if (!hasValue && !hasFormula && !edit.format)
      return issue('', '修改项至少需要 value、formula 或 format。');
  }
  return null;
}
export function validateNewSheets(
  snapshot: WorkbookSnapshot,
  value: unknown,
): value is ResultSheet[] {
  if (!Array.isArray(value) || value.length > TABLE_LIMITS.newSheets)
    return false;
  const names = new Set(snapshot.sheetNames.map((name) => name.toLowerCase()));
  for (const sheet of value) {
    if (
      !object(sheet) ||
      !only(sheet, ['name', 'values']) ||
      !text(sheet.name, 31) ||
      /[\[\]:*?\/\\\u0000-\u001f]/.test(sheet.name) ||
      /^'|'$/.test(sheet.name) ||
      names.has(sheet.name.toLowerCase()) ||
      !Array.isArray(sheet.values) ||
      !sheet.values.length ||
      !Array.isArray(sheet.values[0]) ||
      !sheet.values[0].length ||
      sheet.values.length > 1048576 ||
      sheet.values[0].length > 16384
    )
      return false;
    const columns = sheet.values[0].length;
    if (
      !matrix(
        sheet.values,
        sheet.values.length,
        columns,
        (cell) =>
          isCellValue(cell) &&
          !(typeof cell === 'string' && /^\s*=/.test(cell)),
      )
    )
      return false;
    names.add(sheet.name.toLowerCase());
  }
  return true;
}
export function isEditPlan(
  snapshot: WorkbookSnapshot,
  value: unknown,
): value is EditPlan {
  return (
    object(value) &&
    value.snapshotId === snapshot.snapshotId &&
    value.connectionId === snapshot.connectionId &&
    value.workbookFullName === snapshot.workbookFullName &&
    validateEdits(snapshot, value.edits) &&
    validateNewSheets(snapshot, value.newSheets) &&
    validateColumnChanges(snapshot, value.columnChanges) &&
    validateRowChanges(snapshot, value.rowChanges) &&
    validateSheetActions(snapshot, value.sheetActions) &&
    [value.columnChanges, value.rowChanges, value.sheetActions].filter(
      (items) => items?.length,
    ).length <= 1 &&
    (!(value.rowChanges?.length || value.sheetActions?.length) ||
      (value.edits.length === 0 && value.newSheets.length === 0)) &&
    (!value.columnChanges?.length ||
      (value.edits.length === 0 && value.newSheets.length === 0)) &&
    new TextEncoder().encode(JSON.stringify(value)).length <= TABLE_LIMITS.bytes
  );
}
export function validateColumnChanges(
  snapshot: WorkbookSnapshot,
  value: unknown,
): value is ColumnChange[] | undefined {
  if (value === undefined) return true;
  if (!Array.isArray(value) || value.length > 1) return false;
  for (const change of value) {
    if (
      !object(change) ||
      !only(change, ['rangeId', 'action', 'column', 'count', 'values']) ||
      snapshot.columnOperations?.available !== true
    )
      return false;
    const range = snapshot.ranges.find((item) => item.id === change.rangeId);
    if (
      !range ||
      !['insert', 'delete'].includes(change.action) ||
      !integer(change.count, 1, TABLE_LIMITS.columns) ||
      !integer(
        change.column,
        0,
        range.columnCount - (change.action === 'delete' ? 1 : 0),
      )
    )
      return false;
    const absolute = range.startColumn + change.column;
    if (
      absolute + change.count - 1 > 16384 ||
      (change.action === 'delete' &&
        change.column + change.count > range.columnCount)
    )
      return false;
    if (
      change.values !== undefined &&
      (change.action !== 'insert' ||
        !matrix(
          change.values,
          range.rowCount,
          change.count,
          (cell) =>
            isCellValue(cell) &&
            !(typeof cell === 'string' && /^\s*=/.test(cell)),
        ))
    )
      return false;
  }
  return true;
}
export function validateRowChanges(
  snapshot: WorkbookSnapshot,
  value: unknown,
): value is RowChange[] | undefined {
  if (value === undefined) return true;
  if (!Array.isArray(value) || value.length > 1) return false;
  return value.every((change) => {
    if (
      !object(change) ||
      !only(change, ['rangeId', 'action', 'row', 'count', 'values']) ||
      snapshot.rowOperations?.available !== true
    )
      return false;
    const range = snapshot.ranges.find((item) => item.id === change.rangeId);
    if (
      !range ||
      !['insert', 'delete'].includes(change.action) ||
      !integer(change.count, 1, TABLE_LIMITS.rows) ||
      !integer(
        change.row,
        0,
        range.rowCount - (change.action === 'delete' ? 1 : 0),
      ) ||
      range.startRow + change.row + change.count - 1 > 1048576 ||
      (change.action === 'delete' && change.row + change.count > range.rowCount)
    )
      return false;
    return (
      change.values === undefined ||
      (change.action === 'insert' &&
        matrix(
          change.values,
          change.count,
          range.columnCount,
          (cell) =>
            isCellValue(cell) &&
            !(typeof cell === 'string' && /^\s*=/.test(cell)),
        ))
    );
  });
}
export function validSheetName(name: unknown): name is string {
  return (
    text(name, 31) &&
    name.trim() === name &&
    !/[\\/?*\[\]:\x00-\x1f]/.test(name) &&
    !name.startsWith("'") &&
    !name.endsWith("'")
  );
}
export function validateSheetActions(
  snapshot: WorkbookSnapshot,
  value: unknown,
): value is SheetAction[] | undefined {
  if (value === undefined) return true;
  if (!Array.isArray(value) || value.length > 1) return false;
  return value.every((action) => {
    if (!object(action) || snapshot.sheetOperations?.available !== true)
      return false;
    const range = snapshot.ranges.find((item) => item.id === action.rangeId);
    if (!range) return false;
    if (action.type === 'rename')
      return (
        only(action, ['type', 'rangeId', 'name']) &&
        validSheetName(action.name) &&
        !snapshot.sheetNames.some(
          (name) => name.toLowerCase() === action.name.toLowerCase(),
        )
      );
    if (action.type === 'filter') {
      if (
        !only(action, ['type', 'rangeId', 'conditions']) ||
        range.rowCount < 2 ||
        !Array.isArray(action.conditions) ||
        action.conditions.length > 3
      )
        return false;
      const columns = new Set<number>();
      return action.conditions.every((condition: unknown) => {
        if (
          !object(condition) ||
          !only(condition, ['column', 'operator', 'value']) ||
          !integer(condition.column, 0, range.columnCount - 1) ||
          columns.has(condition.column) ||
          !['eq', 'ne', 'gt', 'gte', 'lt', 'lte', 'contains'].includes(
            condition.operator,
          ) ||
          !isCellValue(condition.value)
        )
          return false;
        columns.add(condition.column);
        if (['gt', 'gte', 'lt', 'lte'].includes(condition.operator))
          return typeof condition.value === 'number';
        return (
          condition.operator !== 'contains' ||
          (typeof condition.value === 'string' && condition.value.length > 0)
        );
      });
    }
    if (
      action.type !== 'chart' ||
      !only(action, [
        'type',
        'rangeId',
        'chartType',
        'title',
        'categoryColumn',
        'valueColumns',
        'aggregation',
        'categoryOrder',
      ]) ||
      !['line', 'column', 'bar', 'pie'].includes(action.chartType) ||
      !text(action.title, 100) ||
      range.rowCount < 2 ||
      !integer(action.categoryColumn, 0, range.columnCount - 1) ||
      !Array.isArray(action.valueColumns) ||
      action.valueColumns.length < 1 ||
      action.valueColumns.length > (action.chartType === 'pie' ? 1 : 3) ||
      new Set(action.valueColumns).size !== action.valueColumns.length
    )
      return false;
    if (
      (owns(action, 'aggregation') && action.aggregation !== 'sum') ||
      (owns(action, 'categoryOrder') &&
        (action.aggregation !== 'sum' ||
          !['source', 'ascending'].includes(action.categoryOrder)))
    )
      return false;
    if (action.aggregation === 'sum') {
      if (
        !action.valueColumns.every(
          (column: unknown) =>
            integer(column, 0, range.columnCount - 1) &&
            column !== action.categoryColumn,
        )
      )
        return false;
      try {
        summarizeChart(
          range,
          action as Extract<SheetAction, { type: 'chart' }>,
        );
        return true;
      } catch {
        return false;
      }
    }
    return action.valueColumns.every(
      (column: unknown) =>
        integer(column, 0, range.columnCount - 1) &&
        column !== action.categoryColumn &&
        range.values
          .slice(1)
          .some((row) => typeof row[column as number] === 'number') &&
        range.values
          .slice(1)
          .every(
            (row) =>
              row[column as number] === null ||
              row[column as number] === '' ||
              typeof row[column as number] === 'number',
          ),
    );
  });
}
export function asSelection(
  snapshot: WorkbookSnapshot,
  range: TableRange,
): AnalysisSelection {
  return {
    host: 'wps',
    connectionId: snapshot.connectionId,
    workbookName: snapshot.workbookName,
    sheetName: range.sheetName,
    address: range.address,
    rowCount: range.rowCount,
    columnCount: range.columnCount,
    values: range.values,
    capturedAt: snapshot.capturedAt,
  };
}
