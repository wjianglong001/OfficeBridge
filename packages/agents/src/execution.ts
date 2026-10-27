import { prepareExecution } from '../../protocol/src/execution.js';
import { isWorkbookSnapshot } from '../../protocol/src/table.js';
import { TableToolSession, inspectData } from './table-tools.js';
import { summarizeData } from './selection-tools.js';

export function executeComputation(operation: string, payload: any): unknown {
  if (operation === 'prepare')
    return prepareExecution(payload.snapshot, payload.plan);
  if (operation === 'summary')
    return summarizeData(
      payload.range,
      payload.params,
      payload.values,
      payload.columns,
    );
  if (operation === 'inspect')
    return inspectData(payload.range, payload.headerRow);
  if (operation === 'snapshot') {
    if (!isWorkbookSnapshot(payload))
      throw new Error('修改范围无效或超过 512 KB。');
    return true;
  }
  if (operation === 'tool') {
    if (!isWorkbookSnapshot(payload.snapshot))
      throw new Error('修改范围无效或超过 512 KB。');
    const task = new TableToolSession(payload.snapshot, true, payload.previous);
    const result = task.run(payload.name, payload.params);
    return { result, plan: task.proposal() };
  }
  throw new Error('后台计算类型不受支持。');
}
