import { randomUUID } from 'node:crypto';
import type {
  HostToolCall,
  HostToolReply,
} from '../../../packages/protocol/src/on-demand.js';

export class HostBroker {
  private pending = new Map<
    string,
    {
      runId: string;
      snapshotId: string;
      resolve: (reply: HostToolReply) => void;
      reject: (error: Error) => void;
      timer: ReturnType<typeof setTimeout>;
    }
  >();
  open(
    snapshotId: string,
    signal: AbortSignal,
    send: (event: unknown) => void,
  ) {
    const runId = randomUUID();
    let closed = false;
    const close = () => {
      closed = true;
      for (const [id, p] of this.pending)
        if (p.runId === runId) {
          clearTimeout(p.timer);
          this.pending.delete(id);
          p.reject(new Error('宿主读取任务已结束或取消。'));
        }
    };
    signal.addEventListener('abort', close, { once: true });
    const call: HostToolCall = (name, args) =>
      new Promise((resolve, reject) => {
        if (signal.aborted || closed) {
          reject(new Error('任务已停止。'));
          return;
        }
        const id = randomUUID();
        const timer = setTimeout(() => {
          this.pending.delete(id);
          reject(new Error('宿主读取超时，请检查 WPS 是否仍连接。'));
        }, 30000);
        this.pending.set(id, { runId, snapshotId, resolve, reject, timer });
        try {
          send({
            type: 'host_request',
            runId,
            callId: id,
            snapshotId,
            name,
            args,
          });
        } catch (error) {
          clearTimeout(timer);
          this.pending.delete(id);
          reject(error);
        }
      });
    return {
      call,
      close: () => {
        signal.removeEventListener('abort', close);
        close();
      },
    };
  }
  reply(input: any): boolean {
    if (
      !input ||
      typeof input.callId !== 'string' ||
      typeof input.runId !== 'string' ||
      typeof input.snapshotId !== 'string'
    )
      return false;
    const p = this.pending.get(input.callId);
    if (!p || p.runId !== input.runId || p.snapshotId !== input.snapshotId)
      return false;
    this.pending.delete(input.callId);
    clearTimeout(p.timer);
    if (typeof input.error === 'string')
      p.reject(new Error(input.error.slice(0, 1000)));
    else if (
      input.reply &&
      typeof input.reply === 'object' &&
      !Array.isArray(input.reply) &&
      Object.hasOwn(input.reply, 'result')
    )
      p.resolve(input.reply);
    else p.reject(new Error('宿主读取返回无效数据。'));
    return true;
  }
}
