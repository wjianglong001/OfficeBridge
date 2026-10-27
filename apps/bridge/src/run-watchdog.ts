import { AGENT_RUN_LIMITS } from '../../../packages/agents/src/run-limits.js';

export type RunPhase = 'agent' | 'thinking' | 'tool' | 'table';
export class RunWatchdog {
  private idle: ReturnType<typeof setTimeout> | undefined;
  private total: ReturnType<typeof setTimeout>;
  private stopped = false;
  phase: RunPhase = 'agent';
  reason: 'idle' | 'total' | undefined;
  private idleMs: number;
  constructor(
    private options: {
      idleMs?: number | undefined;
      maxMs?: number | undefined;
      onTimeout: () => void;
    },
  ) {
    this.idleMs = options.idleMs ?? AGENT_RUN_LIMITS.idleMs;
    this.total = setTimeout(
      () => this.expire('total'),
      options.maxMs ?? AGENT_RUN_LIMITS.maxMs,
    );
    this.touch('agent');
  }
  private expire(reason: 'idle' | 'total') {
    if (this.stopped) return;
    this.reason = reason;
    this.stop();
    this.options.onTimeout();
  }
  /** Only actual new native output or completed/started host work counts. */
  touch(phase: RunPhase) {
    if (this.stopped) return;
    this.phase = phase;
    clearTimeout(this.idle);
    this.idle = setTimeout(() => this.expire('idle'), this.idleMs);
  }
  message(agent: string) {
    const phase = {
      agent: '等待回复',
      thinking: '思考',
      tool: '准备工具参数',
      table: '处理表格',
    }[this.phase];
    return this.reason === 'total'
      ? `${agent} 本轮处理达到总时长上限，已停止。本轮未应用修改。`
      : `${agent} 在${phase}阶段未返回新的内容或有效进度，等待超时，已停止。本轮未应用修改。`;
  }
  stop() {
    this.stopped = true;
    clearTimeout(this.idle);
    clearTimeout(this.total);
  }
}
