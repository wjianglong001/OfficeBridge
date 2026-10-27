import { TableToolSession, tableTools } from './table-tools.js';
import {
  isTaskContinuation,
  type TaskContinuation,
} from '../../protocol/src/continuation.js';
import {
  isWorkbookSnapshot,
  isEditPlan,
  editParameterIssue,
  type WorkbookSnapshot,
  type EditPlan,
} from '../../protocol/src/table.js';
import {
  isWorkbookDescriptor,
  authorizedSnapshot,
  authorizedRange,
  readOptions,
  ReadParameterError,
  ToolParameterError,
  type WorkbookDescriptor,
  type HostToolCall,
} from '../../protocol/src/on-demand.js';
import { summaryColumns } from './selection-tools.js';

export interface TableTask {
  readonly snapshot?: WorkbookSnapshot | undefined;
  run(name: string, args: unknown): unknown;
  proposal(): EditPlan | null;
  continuation(): TaskContinuation | null;
}
class RemoteTableTask implements TableTask {
  snapshot: WorkbookSnapshot | undefined;
  private plan: EditPlan | null = null;
  private next: TaskContinuation | null = null;
  private queue = Promise.resolve();
  constructor(
    private descriptor: WorkbookDescriptor,
    private call: HostToolCall,
    private signal: AbortSignal,
  ) {}
  proposal() {
    return this.plan;
  }
  continuation() {
    return !this.signal.aborted && this.plan ? this.next : null;
  }
  run(name: string, args: unknown): unknown {
    const definition = tableTools.find((t) => t.name === name);
    if (
      !definition ||
      !args ||
      typeof args !== 'object' ||
      Array.isArray(args) ||
      Object.keys(args).some(
        (k) => !Object.hasOwn(definition.inputSchema.properties, k),
      ) ||
      definition.inputSchema.required.some((k) => !Object.hasOwn(args, k))
    ) {
      this.next = null;
      this.plan = null;
      this.snapshot = undefined;
      throw new Error('工具或参数不受支持。');
    }
    if (this.signal.aborted) throw new Error('任务已停止。');
    if (name === 'officebridge_list_ranges')
      return {
        ...this.descriptor,
        unsupportedOperations: ['native_sort'],
        note: 'inline 是随本轮消息提供的真实选区值。先判断是否足够；未提供的数据按需读取，不能从历史结果推断当前值。',
      };
    const execute = async () => {
      if (this.signal.aborted) throw new Error('任务已停止。');
      if (name === 'officebridge_continue_after_apply') {
        if (!this.plan || !isTaskContinuation(args))
          throw new Error('先生成有效修改方案，再登记应用后的下一步。');
        this.next = { nextStep: args.nextStep };
        return { awaitingApply: true, nextStep: args.nextStep };
      }
      if (name === 'officebridge_propose_edits') {
        const params = args as Record<string, unknown>;
        const issue =
          Array.isArray(params.edits) && params.edits.length === 0
            ? { field: 'edits', message: '至少提供一项修改。' }
            : editParameterIssue(this.descriptor.ranges, params.edits);
        if (issue)
          throw new ToolParameterError(
            issue.field,
            issue.field +
              '：' +
              issue.message +
              ' 本次调用未进入宿主，未读取或修改表格，可修正参数后在本轮重试。',
          );
      }
      if (name === 'officebridge_create_chart') {
        const params = args as Record<string, unknown>;
        if (
          Object.hasOwn(params, 'categoryOrder') &&
          (params.aggregation !== 'sum' ||
            !['source', 'ascending'].includes(params.categoryOrder as string))
        )
          throw new ToolParameterError(
            'categoryOrder',
            'categoryOrder 仅可与 aggregation="sum" 一起使用，值为 source 或 ascending。已有汇总数据按当前行序画普通图时请省略 categoryOrder；需要分组排序时请显式指定 aggregation="sum"。本次调用未进入宿主，未读取或修改表格，可修正参数后在本轮重试。',
          );
      }
      // Validate against the REGISTERED full range before opening a host query.
      // A malformed read must not poison the host's stable snapshot.
      if (
        name === 'officebridge_read_range' ||
        name === 'officebridge_summarize_range'
      ) {
        try {
          const params = args as Record<string, any>;
          const range = this.descriptor.ranges.find(
            (r) => r.id === params.rangeId,
          );
          if (!range)
            throw new ReadParameterError(
              'rangeId',
              'rangeId 不属于本轮已登记范围；请使用本轮 list_ranges 或 resolve_range 返回的 rangeId。',
            );
          if (name === 'officebridge_read_range') readOptions(range, params);
          else summaryColumns(range, params);
        } catch (error) {
          throw new ReadParameterError(
            error instanceof ReadParameterError ? error.field : 'summary',
            (error as Error).message +
              ' 本次调用未进入宿主，未读取或修改表格，可修正参数后在本轮重试。',
          );
        }
      }
      const reply = await this.call(name, args);
      if (this.signal.aborted) throw new Error('任务已停止。');
      if (reply.range !== undefined) {
        if (
          name !== 'officebridge_resolve_range' ||
          !authorizedRange(this.descriptor, reply.range)
        )
          throw new Error('宿主返回了未授权的定位范围。');
        const existing = this.descriptor.ranges.find(
          (r) => r.id === reply.range!.id,
        );
        if (
          existing &&
          JSON.stringify(existing) !== JSON.stringify(reply.range)
        )
          throw new Error('定位范围 ID 不得重定向。');
        if (!existing) {
          if (this.descriptor.ranges.length >= 50)
            throw new Error('本轮定位范围超过 50 个。');
          this.descriptor.ranges.push(reply.range);
        }
      }
      if (name === 'officebridge_ask_user') {
        this.plan = null;
        this.snapshot = undefined;
        this.next = null;
      }
      if (reply.plan !== undefined) {
        if (
          !isWorkbookSnapshot(reply.snapshot) ||
          !authorizedSnapshot(this.descriptor, reply.snapshot) ||
          !isEditPlan(reply.snapshot, reply.plan)
        )
          throw new Error('宿主返回了越界或无效的修改方案。');
        this.snapshot = reply.snapshot;
        this.plan = reply.plan;
      }
      return reply.result;
    };
    const next = this.queue.then(execute);
    this.queue = next.then(
      () => {},
      () => {},
    );
    return next.catch((error) => {
      this.plan = null;
      this.snapshot = undefined;
      this.next = null;
      throw error;
    });
  }
}
export function createTableTask(
  workbook: WorkbookSnapshot | WorkbookDescriptor,
  call: HostToolCall | undefined,
  signal: AbortSignal,
): TableTask {
  if (isWorkbookDescriptor(workbook)) {
    if (!call) throw new Error('按需读取通道未连接。');
    return new RemoteTableTask(workbook, call, signal);
  }
  return new TableToolSession(workbook as WorkbookSnapshot);
}
