import {
  isWorkbookSnapshot,
  isEditPlan,
  summarizeChart,
  summarySeriesFormula,
} from './table.js';

// Pure data only: executed in a Web Worker, with no workbook or DOM access.
export function prepareExecution(snapshot: unknown, plan: unknown) {
  if (!isWorkbookSnapshot(snapshot) || !isEditPlan(snapshot, plan))
    throw new Error('修改方案无效，未执行任何修改。请重新发送。');
  const action = plan.sheetActions?.[0];
  const summary =
    action?.type === 'chart' && action.aggregation === 'sum'
      ? summarizeChart(
          snapshot.ranges.find((r) => r.id === action.rangeId)!,
          action,
        )
      : null;
  return {
    valid: true,
    planKey: JSON.stringify(plan),
    snapshotId: snapshot.snapshotId,
    summary,
    formulas: summary
      ? summary.series.map((_, i) => summarySeriesFormula(summary, i))
      : [],
  };
}
