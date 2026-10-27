import {
  createServer,
  type IncomingMessage,
  type ServerResponse,
} from 'node:http';
import { randomBytes, randomUUID } from 'node:crypto';
import { readWpsAsset } from './assets.js';
import {
  asSelection,
  isWorkbookSnapshot,
  isEditPlan,
} from '../../../packages/protocol/src/table.js';
import { plannedAgents } from '../../../packages/agents/src/index.js';
import type { AgentAdapter } from '../../../packages/agents/src/index.js';
import type { AgentEvent } from '../../../packages/agents/src/index.js';
import { CodexAdapter } from '../../../packages/agents/src/codex.js';
import { WorkBuddyAdapter } from '../../../packages/agents/src/workbuddy.js';
import { DoubaoAdapter } from '../../../packages/agents/src/doubao.js';
import {
  createDoubaoConnection,
  DoubaoConnectionChangedError,
} from '../../../packages/agents/src/doubao-connection.js';
import {
  PROTOCOL_VERSION,
  type HealthResponse,
  isAnalysisSelection,
} from '../../../packages/protocol/src/index.js';
import { isWorkbookDescriptor } from '../../../packages/protocol/src/on-demand.js';
import { HostBroker } from './host-broker.js';
import { RunWatchdog } from './run-watchdog.js';
import { Continuations } from './continuations.js';
import {
  isTaskContinuation,
  isApplicationReceipt,
} from '../../../packages/protocol/src/continuation.js';

const assets = new Map([
  ['/wps/index.html', ['.generated/index.html', 'text/html; charset=utf-8']],
  ['/wps/taskpane.html', ['taskpane.html', 'text/html; charset=utf-8']],
  ['/wps/taskpane.js', ['taskpane.js', 'text/javascript; charset=utf-8']],
  ['/wps/executor.js', ['executor.js', 'text/javascript; charset=utf-8']],
  [
    '/wps/execution-worker.js',
    ['.generated/execution-worker.js', 'text/javascript; charset=utf-8'],
  ],
  ['/wps/host.js', ['host.js', 'text/javascript; charset=utf-8']],
  [
    '/wps/host-operations.js',
    ['host-operations.js', 'text/javascript; charset=utf-8'],
  ],
  [
    '/wps/contract.js',
    ['.generated/contract.js', 'text/javascript; charset=utf-8'],
  ],
  ['/wps/styles.css', ['styles.css', 'text/css; charset=utf-8']],
  ['/wps/main.js', ['main.js', 'text/javascript; charset=utf-8']],
  ['/wps/ribbon.xml', ['ribbon.xml', 'application/xml; charset=utf-8']],
]);

function json(response: ServerResponse, status: number, body: unknown) {
  response.writeHead(status, {
    'Content-Type': 'application/json; charset=utf-8',
  });
  response.end(JSON.stringify(body));
}

function readJson(request: IncomingMessage): Promise<unknown> {
  return new Promise((resolve, reject) => {
    const chunks: Buffer[] = [];
    let length = 0;
    let failed = false;
    request.on('data', (chunk: Buffer) => {
      if (failed) return;
      length += chunk.length;
      if (length > 1_048_576) {
        failed = true;
        chunks.length = 0;
        reject(new Error('请求超过 1 MB。'));
        return;
      }
      chunks.push(chunk);
    });
    request.on('end', () => {
      if (failed) return;
      try {
        resolve(JSON.parse(Buffer.concat(chunks).toString('utf8')));
      } catch {
        reject(new Error('请求 JSON 无效。'));
      }
    });
    request.on('error', () => reject(new Error('读取请求失败。')));
    request.on('aborted', () => reject(new Error('请求已取消。')));
  });
}

export function createBridgeServer(
  options: {
    adapter?: AgentAdapter;
    adapters?: AgentAdapter[];
    analysisTimeoutMs?: number;
    analysisIdleTimeoutMs?: number;
    analysisMaxDurationMs?: number;
    doubaoConnection?: ReturnType<typeof createDoubaoConnection>;
  } = {},
) {
  const hostBroker = new HostBroker();
  const continuations = new Continuations();
  const doubaoConnection = options.doubaoConnection ?? createDoubaoConnection();
  const adapters = new Map(
    (
      options.adapters ??
      (options.adapter
        ? [options.adapter]
        : [new CodexAdapter(), new WorkBuddyAdapter(), new DoubaoAdapter()])
    ).map((adapter) => [adapter.id, adapter]),
  );
  const token = randomBytes(32).toString('hex');
  let busy = false;
  const conversations = new Map<string, AgentAdapter>();
  async function invalidateDoubaoConversations() {
    for (const [id, owner] of conversations) {
      if (owner.id !== 'doubao') continue;
      conversations.delete(id);
      continuations.clear(id);
      await owner.closeConversation(id).catch(() => {});
    }
  }
  const server = createServer(async (request, response) => {
    response.setHeader('Cache-Control', 'no-store');
    response.setHeader('X-Content-Type-Options', 'nosniff');
    response.setHeader(
      'Content-Security-Policy',
      "default-src 'self'; connect-src 'self'; frame-ancestors 'self'",
    );

    const address = request.socket.localAddress;
    if (address !== '127.0.0.1' && address !== '::ffff:127.0.0.1') {
      json(response, 403, { error: 'loopback_only' });
      return;
    }

    const expectedHost = `127.0.0.1:${request.socket.localPort}`;
    if (request.headers.host !== expectedHost) {
      json(response, 403, { error: 'invalid_host' });
      return;
    }
    if (
      request.headers.origin &&
      request.headers.origin !== `http://${expectedHost}`
    ) {
      json(response, 403, { error: 'invalid_origin' });
      return;
    }
    const analyzePath =
      request.url === '/api/v1/analyze' || request.url === '/api/v1/run';
    const conversationPath =
      request.url === '/api/v1/conversations' ||
      request.url === '/api/v1/conversations/close';
    const hostResultPath = request.url === '/api/v1/host-result';
    const doubaoConnectionPath =
      request.url === '/api/v1/agents/doubao/connect';
    const cancelContinuationPath =
      request.url === '/api/v1/continuations/cancel';
    if (
      request.method !== 'GET' &&
      !(
        request.method === 'POST' &&
        (analyzePath ||
          conversationPath ||
          hostResultPath ||
          doubaoConnectionPath ||
          cancelContinuationPath)
      )
    ) {
      response.setHeader('Allow', 'GET');
      json(response, 405, { error: 'method_not_allowed' });
      return;
    }

    let path: string;
    try {
      path = new URL(request.url ?? '/', `http://${expectedHost}`).pathname;
    } catch {
      json(response, 400, { error: 'invalid_url' });
      return;
    }

    if (path === '/') {
      response.writeHead(302, { Location: '/wps/taskpane.html' });
      response.end();
      return;
    }
    if (path === '/favicon.ico') {
      response.writeHead(204);
      response.end();
      return;
    }
    if (path === '/api/v1/health') {
      const health: HealthResponse = {
        service: 'office-bridge',
        status: 'ok',
        protocolVersion: PROTOCOL_VERSION,
      };
      json(response, 200, health);
      return;
    }
    if (path === '/api/v1/agents') {
      try {
        const statuses = new Map(
          await Promise.all(
            [...adapters].map(async ([id, adapter]) => {
              try {
                return [id, await adapter.probe()] as const;
              } catch {
                return [
                  id,
                  { available: false, reason: 'Agent 状态检测失败。' },
                ] as const;
              }
            }),
          ),
        );
        json(response, 200, {
          agents: plannedAgents.map((agent) =>
            statuses.has(agent.id)
              ? {
                  ...agent,
                  status: statuses.get(agent.id)!.available
                    ? 'ready'
                    : 'unavailable',
                  reason: statuses.get(agent.id)!.reason ?? null,
                }
              : agent,
          ),
        });
      } catch {
        json(response, 503, {
          error: 'probe_failed',
          message: 'Codex 状态检测失败。',
        });
      }
      return;
    }
    if (path === '/api/v1/session') {
      json(response, 200, { token });
      return;
    }
    if (
      analyzePath ||
      conversationPath ||
      hostResultPath ||
      doubaoConnectionPath ||
      cancelContinuationPath
    ) {
      if (request.method !== 'POST') {
        json(response, 405, { error: 'method_not_allowed' });
        return;
      }
      if (
        request.headers.origin !== `http://${expectedHost}` ||
        request.headers['x-officebridge-token'] !== token
      ) {
        request.resume();
        json(response, 401, {
          error: 'unauthorized',
          message: '连接已失效，请刷新侧栏。',
        });
        return;
      }
      if (!request.headers['content-type']?.startsWith('application/json')) {
        request.resume();
        json(response, 415, { error: 'json_required' });
        return;
      }
      if (
        busy &&
        path !== '/api/v1/conversations/close' &&
        !hostResultPath &&
        !cancelContinuationPath
      ) {
        request.resume();
        json(response, 409, {
          error: 'busy',
          message: '已有分析正在运行，请稍后再试。',
        });
        return;
      }
      let input: any;
      try {
        input = await readJson(request);
      } catch (error) {
        json(response, 400, {
          error: 'invalid_request',
          message: error instanceof Error ? error.message : '请求无效。',
        });
        return;
      }
      if (hostResultPath) {
        json(response, hostBroker.reply(input) ? 200 : 410, {
          status: 'received_or_expired',
        });
        return;
      }
      if (doubaoConnectionPath) {
        if (
          !input ||
          Array.isArray(input) ||
          typeof input !== 'object' ||
          Object.keys(input).some((key) => key !== 'restartToken') ||
          (input.restartToken !== undefined &&
            (typeof input.restartToken !== 'string' ||
              !/^[a-f0-9]{48}$/.test(input.restartToken)))
        ) {
          json(response, 400, { error: 'invalid_request' });
          return;
        }
        if (!adapters.has('doubao')) {
          json(response, 400, { error: 'unsupported_agent' });
          return;
        }
        if (busy) {
          json(response, 409, {
            error: 'busy',
            message: '正在处理任务，请稍后重试。',
          });
          return;
        }
        busy = true;
        try {
          const result = await doubaoConnection.prepare(input.restartToken);
          if (result.status === 'ready' && result.changed) {
            await invalidateDoubaoConversations();
          }
          json(response, 200, result);
        } catch (error) {
          const changed = error instanceof DoubaoConnectionChangedError;
          if (changed) await invalidateDoubaoConversations();
          json(response, 503, {
            error: 'doubao_connection_failed',
            changed,
            message:
              error instanceof Error ? error.message : '连接豆包失败，请重试。',
          });
        } finally {
          busy = false;
        }
        return;
      }
      if (cancelContinuationPath) {
        continuations.cancel(input?.token, input?.conversationId);
        json(response, 200, { cancelled: true });
        return;
      }
      if (conversationPath) {
        if (
          !input ||
          typeof input !== 'object' ||
          Array.isArray(input) ||
          (input.id !== undefined &&
            (typeof input.id !== 'string' ||
              !/^[\w-]{1,100}$/.test(input.id))) ||
          (input.agentId !== undefined && typeof input.agentId !== 'string') ||
          Object.keys(input).some((key) => key !== 'id' && key !== 'agentId')
        ) {
          json(response, 400, { error: 'invalid_request' });
          return;
        }
        if (path.endsWith('/close')) {
          if (typeof input.id !== 'string') {
            json(response, 400, { error: 'invalid_request' });
            return;
          }
          const owner = conversations.get(input.id);
          conversations.delete(input.id);
          continuations.clear(input.id);
          if (owner) await owner.closeConversation(input.id);
          json(response, 200, { closed: true });
          return;
        }
        if (busy) {
          json(response, 409, {
            error: 'busy',
            message: '正在处理任务，请稍后重试。',
          });
          return;
        }
        busy = true;
        try {
          const adapter = adapters.get(input.agentId ?? 'codex');
          if (!adapter) {
            json(response, 400, {
              error: 'unsupported_agent',
              message: '尚未接入这个 Agent。',
            });
            return;
          }
          const previous = conversations.get(input.id);
          const conversation = await adapter.openConversation(
            previous === adapter ? input.id : undefined,
          );
          if (response.destroyed) {
            await adapter.closeConversation(conversation.id);
            return;
          }
          conversations.set(conversation.id, adapter);
          if (previous) {
            conversations.delete(input.id);
            continuations.clear(input.id);
            await previous.closeConversation(input.id);
          }
          json(response, 200, { ...conversation, agentId: adapter.id });
        } catch (error) {
          json(response, 503, {
            error: 'conversation_failed',
            message:
              error instanceof Error ? error.message : '无法创建 Agent 会话。',
          });
        } finally {
          busy = false;
        }
        return;
      }
      if (
        !input ||
        typeof input.prompt !== 'string' ||
        !input.prompt.trim() ||
        input.prompt.length > 4000 ||
        !(path === '/api/v1/run'
          ? isWorkbookSnapshot(input.workbook) ||
            isWorkbookDescriptor(input.workbook)
          : isAnalysisSelection(input.selection)) ||
        (input.resume !== undefined &&
          (path !== '/api/v1/run' ||
            !input.resume ||
            typeof input.resume !== 'object' ||
            Array.isArray(input.resume) ||
            typeof input.resume.token !== 'string' ||
            !isApplicationReceipt(input.resume.application) ||
            Object.keys(input.resume).some(
              (k) => !['token', 'application'].includes(k),
            ))) ||
        (path === '/api/v1/run' &&
          (typeof input.conversationId !== 'string' ||
            input.history !== undefined ||
            (input.operationFeedback !== undefined &&
              (typeof input.operationFeedback !== 'string' ||
                input.operationFeedback.length > 4000))))
      ) {
        json(response, 400, {
          error: 'invalid_request',
          message:
            '指令或完整范围无效。请检查范围、数据类型及 512 KB 数据体积上限。',
        });
        return;
      }
      if (path === '/api/v1/run' && !conversations.has(input.conversationId)) {
        json(response, 410, {
          error: 'conversation_lost',
          message: 'Agent 会话已失效，请从右上角开启新对话。原对话未被替换。',
        });
        return;
      }
      if (busy) {
        json(response, 409, {
          error: 'busy',
          message: '已有分析正在运行，请稍后再试。',
        });
        return;
      }
      const resume = input.resume
        ? continuations.consume(
            input.resume.token,
            input.conversationId,
            input.workbook,
            input.resume.application,
          )
        : null;
      if (input.resume && !resume) {
        json(response, 410, {
          error: 'continuation_expired',
          message:
            '自动续办凭据已失效、已使用或文件已变化；本次未继续，请手动说明后续要求。',
        });
        return;
      }
      if (path === '/api/v1/run' && !input.resume)
        continuations.clear(input.conversationId);
      busy = true;
      const adapter =
        path === '/api/v1/run'
          ? conversations.get(input.conversationId)!
          : adapters.get('codex');
      if (!adapter) {
        busy = false;
        json(response, 503, { error: 'agent_unavailable' });
        return;
      }
      const controller = new AbortController();
      const agentLabel =
        adapter.id === 'workbuddy'
          ? 'WorkBuddy'
          : adapter.id === 'doubao'
            ? '豆包'
            : 'Codex';
      const watchdog = new RunWatchdog({
        maxMs: options.analysisTimeoutMs ?? options.analysisMaxDurationMs,
        idleMs: options.analysisIdleTimeoutMs,
        onTimeout: () => controller.abort(),
      });
      response.on('close', () => {
        if (!response.writableEnded) controller.abort();
      });
      response.writeHead(200, {
        'Content-Type': 'application/x-ndjson; charset=utf-8',
      });
      response.flushHeaders();
      const send = (event: unknown) => {
        if (!response.destroyed) response.write(JSON.stringify(event) + '\n');
      };
      const lazy =
        path === '/api/v1/run' && isWorkbookDescriptor(input.workbook);
      const host = lazy
        ? hostBroker.open(input.workbook.snapshotId, controller.signal, send)
        : undefined;
      let terminal = false;
      let staged: Extract<AgentEvent, { type: 'proposal' }> | undefined;
      try {
        for await (const event of adapter.run({
          requestId: randomUUID(),
          prompt: resume
            ? `上一方案已由用户点击应用并成功执行。请继续本会话中用户原始任务尚未完成的步骤，不扩展任务、不重复已应用操作。待办步骤：${JSON.stringify(resume.nextStep)}。请按需重新读取最新工作簿；后续修改仍只生成待应用方案，等待用户确认。若原任务已全部完成，请直接汇报，不登记续办。`
            : input.prompt.trim(),
          ...(!lazy
            ? {
                context:
                  path === '/api/v1/run'
                    ? asSelection(input.workbook, input.workbook.ranges[0])
                    : input.selection,
              }
            : {}),
          ...(host
            ? {
                hostTool: async (name, args) => {
                  watchdog.touch('table');
                  const reply = await host.call(name, args);
                  watchdog.touch('agent');
                  return reply;
                },
              }
            : {}),
          ...(path === '/api/v1/run'
            ? {
                workbook: input.workbook,
                conversationId: input.conversationId,
                operationFeedback: resume
                  ? JSON.stringify(resume.application)
                  : input.operationFeedback,
              }
            : {}),
          signal: controller.signal,
        })) {
          if (watchdog.reason) {
            send({
              type: 'error',
              code: 'timeout',
              message: watchdog.message(agentLabel),
            });
            terminal = true;
            break;
          }
          if (event.type === 'text' && event.text.trim())
            watchdog.touch('agent');
          if (event.type === 'activity') watchdog.touch(event.phase);
          if (event.type === 'proposal') {
            const snapshot =
              event.snapshot ?? (!lazy ? input.workbook : undefined);
            if (
              !isWorkbookSnapshot(snapshot) ||
              !isEditPlan(snapshot, event.plan)
            )
              throw new Error('Invalid proposal');
            staged = { ...event, snapshot };
            continue;
          }
          if (
            event.type === 'complete' &&
            staged &&
            !controller.signal.aborted
          ) {
            const offer = isTaskContinuation(staged.continuation)
              ? continuations.issue({
                  conversationId: input.conversationId,
                  workbookFullName: staged.plan.workbookFullName,
                  snapshotId: staged.plan.snapshotId,
                  nextStep: staged.continuation.nextStep,
                  step: resume ? resume.step + 1 : 1,
                })
              : undefined;
            send({ ...staged, continuation: offer });
          }
          send(event);
          if (event.type === 'complete' || event.type === 'error') {
            terminal = true;
            break;
          }
        }
        if (!terminal && !response.destroyed)
          send({
            type: 'error',
            code: watchdog.reason ? 'timeout' : 'incomplete',
            message: watchdog.reason
              ? watchdog.message(agentLabel)
              : '分析未完成，请重试。',
          });
      } catch {
        send({
          type: 'error',
          code: watchdog.reason ? 'timeout' : 'analysis_failed',
          message: watchdog.reason
            ? watchdog.message(agentLabel)
            : `${agentLabel} 处理失败，请检查客户端状态后重试。`,
        });
      } finally {
        host?.close();
        watchdog.stop();
        busy = false;
        response.end();
      }
      return;
    }

    if (path === '/wps/install.html') {
      try {
        const template = (
          await readWpsAsset('.generated/publish.template.html')
        ).toString('utf8');
        const body = template.replaceAll(
          '__OFFICEBRIDGE_ADDON_BASE_URL__',
          `http://${expectedHost}/wps/`,
        );
        // Official publisher uses inline handlers and the WPS loopback relay.
        // Keep this exception scoped to the installer, not the application API.
        response.setHeader(
          'Content-Security-Policy',
          "default-src 'self'; script-src 'self' 'unsafe-inline'; style-src 'self' 'unsafe-inline'; img-src 'self' data:; connect-src 'self' http://127.0.0.1:58890 http://localhost:58890; frame-ancestors 'none'",
        );
        response.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
        response.end(body);
      } catch {
        json(response, 503, {
          error: 'installer_not_prepared',
          message: 'Run npm run wps:prepare and restart the server.',
        });
      }
      return;
    }

    const asset = assets.get(path);
    if (!asset) {
      json(response, 404, { error: 'not_found' });
      return;
    }
    try {
      const body = await readWpsAsset(asset[0]!);
      response.writeHead(200, { 'Content-Type': asset[1]! });
      response.end(body);
    } catch {
      json(response, 500, { error: 'asset_unavailable' });
    }
  });
  server.on('close', () => {
    continuations.clearAll();
    for (const adapter of adapters.values()) adapter.dispose();
    conversations.clear();
  });
  return server;
}
