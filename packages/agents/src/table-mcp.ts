import { createServer, type Server } from 'node:http';
import { randomBytes } from 'node:crypto';
import { once } from 'node:events';
import { tableTools } from './table-tools.js';
import type { TableTask } from './table-task.js';
import {
  ReadParameterError,
  ToolParameterError,
} from '../../protocol/src/on-demand.js';

/** A private, stateless MCP endpoint owned by one native Agent conversation. */
export class TableMcp {
  private server: Server;
  private token = randomBytes(32).toString('hex');
  private active:
    | {
        task: TableTask;
        signal: AbortSignal;
        parameterErrors: number;
        blocked: boolean;
        onParameterLimit?:
          ((message: string, code: string) => void) | undefined;
      }
    | undefined;
  private port = 0;
  initialized = false;
  listed = false;
  successfulCalls = 0;
  constructor() {
    this.server = createServer(async (request, response) => {
      response.setHeader('Cache-Control', 'no-store');
      const reply = (status: number, data?: unknown) => {
        response.writeHead(status, { 'Content-Type': 'application/json' });
        response.end(data === undefined ? undefined : JSON.stringify(data));
      };
      if (
        request.headers.host !== `127.0.0.1:${this.port}` ||
        request.headers.origin ||
        request.url !== '/mcp' ||
        request.headers.authorization !== `Bearer ${this.token}`
      ) {
        request.resume();
        reply(403);
        return;
      }
      if (request.method !== 'POST') {
        request.resume();
        reply(405);
        return;
      }
      if (!request.headers['content-type']?.startsWith('application/json')) {
        request.resume();
        reply(415);
        return;
      }
      // Capture identity BEFORE reading the body; a delayed old request cannot use a new snapshot.
      const active = this.active;
      let message: any;
      try {
        const chunks: Buffer[] = [];
        let length = 0;
        for await (const chunk of request) {
          length += chunk.length;
          if (length > 1_048_576) {
            reply(413);
            return;
          }
          chunks.push(chunk);
        }
        message = JSON.parse(Buffer.concat(chunks).toString('utf8'));
      } catch {
        if (!response.destroyed) reply(400);
        return;
      }
      if (
        !message ||
        Array.isArray(message) ||
        message.jsonrpc !== '2.0' ||
        typeof message.method !== 'string'
      ) {
        reply(400);
        return;
      }
      if (message.id === undefined) {
        reply(202);
        return;
      }
      const result = (value: unknown) =>
        reply(200, { jsonrpc: '2.0', id: message.id, result: value });
      if (message.method === 'initialize') {
        this.initialized = true;
        result({
          protocolVersion: '2025-03-26',
          capabilities: { tools: {} },
          serverInfo: { name: 'officebridge', version: '0.10.0' },
        });
      } else if (message.method === 'ping') result({});
      else if (message.method === 'tools/list') {
        this.listed = true;
        result({
          tools: tableTools.map(({ name, description, inputSchema }) => ({
            name,
            description,
            inputSchema,
          })),
        });
      } else if (message.method === 'tools/call') {
        try {
          if (!active || active !== this.active || active.signal.aborted)
            throw new Error(
              '没有正在处理的表格任务；旧快照和已停止的任务不可访问。',
            );
          if (active.blocked)
            throw new Error('本轮工具参数修正已达上限，任务已停止。');
          const output = await active.task.run(
            message.params?.name,
            message.params?.arguments ?? {},
          );
          if (active !== this.active || active.signal.aborted)
            throw new Error('任务已停止。');
          this.successfulCalls++;
          result({ content: [{ type: 'text', text: JSON.stringify(output) }] });
        } catch (error) {
          const limited =
            error instanceof ToolParameterError &&
            active === this.active &&
            active &&
            ++active.parameterErrors >= 3;
          if (limited) active.blocked = true;
          result({
            isError: true,
            content: [
              {
                type: 'text',
                text: error instanceof Error ? error.message : '工具执行失败。',
              },
            ],
          });
          if (limited)
            active.onParameterLimit?.(
              (error instanceof ReadParameterError ? '读取' : '工具') +
                '参数修正仍未通过，本轮已停止：' +
                (error as Error).message,
              error instanceof ReadParameterError
                ? 'read_parameters_failed'
                : 'tool_parameters_failed',
            );
        }
      } else
        reply(200, {
          jsonrpc: '2.0',
          id: message.id,
          error: { code: -32601, message: 'Unsupported MCP operation.' },
        });
    });
  }
  async start() {
    this.server.listen(0, '127.0.0.1');
    await once(this.server, 'listening');
    const address = this.server.address();
    if (!address || typeof address === 'string')
      throw new Error('无法启动表格工具接口。');
    this.port = address.port;
    return {
      type: 'http',
      name: 'officebridge',
      url: `http://127.0.0.1:${this.port}/mcp`,
      headers: [{ name: 'Authorization', value: `Bearer ${this.token}` }],
    };
  }
  bind(
    task: TableTask,
    signal: AbortSignal,
    onParameterLimit?: (message: string, code: string) => void,
  ) {
    this.successfulCalls = 0;
    this.active = {
      task,
      signal,
      parameterErrors: 0,
      blocked: false,
      onParameterLimit,
    };
  }
  async waitUntilReady(timeoutMs = 10000) {
    const deadline = Date.now() + timeoutMs;
    while (!this.initialized || !this.listed) {
      if (Date.now() >= deadline)
        throw new Error(
          'WorkBuddy 未完成表格工具连接。请开启新对话重试；未发送模型请求。',
        );
      await new Promise((resolve) => setTimeout(resolve, 25));
    }
  }
  clear() {
    this.active = undefined;
  }
  close() {
    this.clear();
    this.server.closeAllConnections();
    this.server.close();
  }
}
