export const PROTOCOL_VERSION = '0.6.0';
export const DEFAULT_PORT = 4318;
export const SELECTION_SAMPLE_LIMITS = { rows: 20, columns: 15 } as const;

export type CellValue = string | number | boolean | null;
export const MAX_ANALYSIS_BYTES = 512_000;

export interface AnalysisSelection {
  host: 'wps';
  connectionId: string;
  workbookName: string;
  sheetName: string;
  address: string;
  rowCount: number;
  columnCount: number;
  values: CellValue[][];
  capturedAt: string;
}

export function isAnalysisSelection(
  value: unknown,
): value is AnalysisSelection {
  if (typeof value !== 'object' || value === null) return false;
  const data = value as Record<string, unknown>;
  if (
    data.host !== 'wps' ||
    !['connectionId', 'workbookName', 'sheetName', 'address'].every(
      (key) =>
        typeof data[key] === 'string' &&
        data[key].length > 0 &&
        data[key].length <= 1024,
    )
  )
    return false;
  if (
    !Number.isSafeInteger(data.rowCount) ||
    !Number.isSafeInteger(data.columnCount) ||
    (data.rowCount as number) < 1 ||
    (data.columnCount as number) < 1 ||
    (data.rowCount as number) > 1048576 ||
    (data.columnCount as number) > 16384 ||
    typeof data.capturedAt !== 'string' ||
    !Number.isFinite(Date.parse(data.capturedAt)) ||
    !Array.isArray(data.values) ||
    data.values.length !== data.rowCount
  )
    return false;
  return (
    data.values.every(
      (row: unknown) =>
        Array.isArray(row) &&
        row.length === data.columnCount &&
        row.every(
          (cell: unknown) =>
            cell === null ||
            typeof cell === 'boolean' ||
            (typeof cell === 'string' && cell.length <= 10000) ||
            (typeof cell === 'number' && Number.isFinite(cell)),
        ),
    ) && Buffer.byteLength(JSON.stringify(value), 'utf8') <= MAX_ANALYSIS_BYTES
  );
}

/** Addresses are scoped to this host connection and must never be replayed elsewhere. */
export interface SelectionContext {
  host: 'wps';
  connectionId: string;
  workbookName: string;
  sheetName: string;
  address: string;
  rowCount: number;
  columnCount: number;
  sampleValues: CellValue[][];
  truncated: boolean;
  capturedAt: string;
}

export interface HealthResponse {
  service: 'office-bridge';
  status: 'ok';
  protocolVersion: string;
}

/** Validate unknown host data before it enters the Agent boundary. */
export function isSelectionContext(value: unknown): value is SelectionContext {
  if (typeof value !== 'object' || value === null) return false;
  const data = value as Record<string, unknown>;
  const strings = ['connectionId', 'workbookName', 'sheetName', 'address'];
  if (data.host !== 'wps') return false;
  if (
    !strings.every(
      (key) => typeof data[key] === 'string' && data[key].length > 0,
    )
  ) {
    return false;
  }
  if (
    !Number.isSafeInteger(data.rowCount) ||
    !Number.isSafeInteger(data.columnCount) ||
    (data.rowCount as number) < 1 ||
    (data.columnCount as number) < 1 ||
    typeof data.truncated !== 'boolean' ||
    typeof data.capturedAt !== 'string' ||
    !Number.isFinite(Date.parse(data.capturedAt)) ||
    !Array.isArray(data.sampleValues)
  )
    return false;

  const expectedRows = Math.min(
    data.rowCount as number,
    SELECTION_SAMPLE_LIMITS.rows,
  );
  const expectedColumns = Math.min(
    data.columnCount as number,
    SELECTION_SAMPLE_LIMITS.columns,
  );
  const truncated =
    data.rowCount !== expectedRows || data.columnCount !== expectedColumns;
  return (
    data.truncated === truncated &&
    data.sampleValues.length === expectedRows &&
    data.sampleValues.every(
      (row: unknown) =>
        Array.isArray(row) &&
        row.length === expectedColumns &&
        row.every(
          (cell: unknown) =>
            cell === null ||
            typeof cell === 'string' ||
            typeof cell === 'boolean' ||
            (typeof cell === 'number' && Number.isFinite(cell)),
        ),
    )
  );
}
