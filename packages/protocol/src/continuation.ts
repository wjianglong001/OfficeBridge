export interface TaskContinuation {
  nextStep: string;
}
export interface ResumeOffer extends TaskContinuation {
  token: string | null;
  reason?: string;
}
const object = (v: any) => v && typeof v === 'object' && !Array.isArray(v);
export function isTaskContinuation(v: any): v is TaskContinuation {
  return (
    object(v) &&
    typeof v.nextStep === 'string' &&
    v.nextStep.trim().length > 0 &&
    v.nextStep.length <= 600
  );
}
export function isResumeOffer(v: any): v is ResumeOffer {
  return (
    isTaskContinuation(v) &&
    (v as any).token !== undefined &&
    ((v as any).token === null ||
      (typeof (v as any).token === 'string' &&
        /^[\w-]{36}$/.test((v as any).token))) &&
    ((v as any).reason === undefined ||
      (typeof (v as any).reason === 'string' &&
        (v as any).reason.length <= 300))
  );
}
export interface ApplicationReceipt {
  status: 'succeeded';
  snapshotId: string;
  workbookFullName: string;
  result: {
    changedCells: number;
    newSheets: string[];
    recoverable?: boolean;
    columnChanges?: unknown[];
    rowChanges?: unknown[];
    sheetActions?: unknown[];
  };
}
export function isApplicationReceipt(v: any): v is ApplicationReceipt {
  if (
    !object(v) ||
    v.status !== 'succeeded' ||
    typeof v.snapshotId !== 'string' ||
    !v.snapshotId ||
    v.snapshotId.length > 1024 ||
    typeof v.workbookFullName !== 'string' ||
    !v.workbookFullName ||
    v.workbookFullName.length > 1024 ||
    !object(v.result)
  )
    return false;
  const r = v.result;
  return (
    Object.keys(v).every((k) =>
      ['status', 'snapshotId', 'workbookFullName', 'result'].includes(k),
    ) &&
    Object.keys(r).every((k) =>
      [
        'changedCells',
        'newSheets',
        'recoverable',
        'columnChanges',
        'rowChanges',
        'sheetActions',
      ].includes(k),
    ) &&
    Number.isSafeInteger(r.changedCells) &&
    r.changedCells >= 0 &&
    Array.isArray(r.newSheets) &&
    r.newSheets.length <= 3 &&
    r.newSheets.every(
      (n: any) => typeof n === 'string' && n.length > 0 && n.length <= 31,
    ) &&
    (r.recoverable === undefined || typeof r.recoverable === 'boolean') &&
    ['columnChanges', 'rowChanges', 'sheetActions'].every(
      (k) => r[k] === undefined || (Array.isArray(r[k]) && r[k].length <= 1),
    ) &&
    new TextEncoder().encode(JSON.stringify(v)).length <= 16000
  );
}
