import {
  spawn,
  execFile,
  type ChildProcessWithoutNullStreams,
} from 'node:child_process';
import { createInterface } from 'node:readline';

export type RpcMessage = {
  id?: number | string;
  method?: string;
  params?: any;
  result?: any;
  error?: { message: string };
};
export interface RpcTransport {
  request(method: string, params: unknown): Promise<any>;
  notify(method: string, params: unknown): void;
  onMessage(callback: (message: RpcMessage) => void): void;
  onFailure(callback: (error: Error) => void): void;
  reply(id: number | string, result: unknown): void;
  reject(id: number | string): void;
  close(): void;
}

/** Minimal stdio JSON-RPC client. Never uses a shell or forwards stderr to the UI. */
export class CodexRpc implements RpcTransport {
  private process: ChildProcessWithoutNullStreams;
  private nextId = 0;
  private pending = new Map<
    number,
    {
      resolve: (value: any) => void;
      reject: (error: Error) => void;
      timer: NodeJS.Timeout;
    }
  >();
  private messageHandler: (message: RpcMessage) => void = () => {};
  private failureHandler: (error: Error) => void = () => {};
  private closed = false;
  private stopReading: () => void;

  constructor(executable: string, cwd: string) {
    this.process = spawn(
      executable,
      [
        'app-server',
        '--stdio',
        ...[
          'features.shell_tool=false',
          'features.apps=false',
          'features.plugins=false',
          'features.hooks=false',
          'features.multi_agent=false',
          'features.browser_use=false',
          'features.computer_use=false',
          'features.skill_search=false',
          'features.shell_snapshot=false',
          'features.skip_host_skill_discovery=true',
          'mcp_servers={}',
          'web_search="disabled"',
        ].flatMap((setting) => ['-c', setting]),
      ],
      { cwd, windowsHide: true, stdio: ['pipe', 'pipe', 'pipe'] },
    );
    const lines = createInterface({
      input: this.process.stdout,
      crlfDelay: Infinity,
    });
    this.stopReading = () => lines.close();
    lines.on('line', (line) => {
      if (!line.trim()) return;
      try {
        const message = JSON.parse(line) as RpcMessage;
        if (typeof message.id === 'number' && !message.method) {
          const waiter = this.pending.get(message.id);
          if (waiter) {
            clearTimeout(waiter.timer);
            this.pending.delete(message.id);
            if (message.error) waiter.reject(new Error(message.error.message));
            else waiter.resolve(message.result);
          }
        } else this.messageHandler(message);
      } catch {
        this.fail(new Error('Codex 返回了无法解析的协议消息。'));
      }
    });
    this.process.stderr.resume();
    this.process.stdin.on('error', () =>
      this.fail(new Error('Codex 通信管道已关闭。')),
    );
    this.process.on('error', () =>
      this.fail(new Error('无法启动 Codex，请检查可执行文件路径。')),
    );
    this.process.on('exit', (code) => {
      if (!this.closed)
        this.fail(new Error(`Codex 进程提前退出（${code ?? 'signal'}）。`));
    });
    this.process.once('close', () => lines.close());
  }

  onMessage(callback: (message: RpcMessage) => void) {
    this.messageHandler = callback;
  }
  onFailure(callback: (error: Error) => void) {
    this.failureHandler = callback;
  }
  private send(message: unknown) {
    if (this.closed) throw new Error('Codex 连接已关闭。');
    this.process.stdin.write(JSON.stringify(message) + '\n');
  }
  request(method: string, params: unknown): Promise<any> {
    if (this.closed) return Promise.reject(new Error('Codex 连接已关闭。'));
    const id = ++this.nextId;
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        this.pending.delete(id);
        reject(new Error(`Codex ${method} 请求超时。`));
      }, 30000);
      this.pending.set(id, { resolve, reject, timer });
      try {
        this.send({ id, method, params });
      } catch (error) {
        clearTimeout(timer);
        this.pending.delete(id);
        reject(error);
      }
    });
  }
  reply(id: number | string, result: unknown) {
    this.send({ id, result });
  }
  notify(method: string, params: unknown) {
    this.send({ method, params });
  }
  reject(id: number | string) {
    this.send({
      id,
      error: {
        code: -32601,
        message: 'OfficeBridge only supports read-only selection tools.',
      },
    });
  }
  private fail(error: Error) {
    if (this.closed) return;
    this.failureHandler(error);
    this.close();
  }
  close() {
    if (this.closed) return;
    this.closed = true;
    for (const waiter of this.pending.values()) {
      clearTimeout(waiter.timer);
      waiter.reject(new Error('Codex 连接已关闭。'));
    }
    this.pending.clear();
    this.stopReading();
    this.process.stdin.destroy();
    this.process.stdout.destroy();
    this.process.stderr.destroy();
    if (process.platform === 'win32' && this.process.pid) {
      // Kill only the app-server process tree created by this transport.
      execFile(
        'taskkill.exe',
        ['/PID', String(this.process.pid), '/T', '/F'],
        { windowsHide: true },
        () => {},
      );
    } else this.process.kill();
  }
}
