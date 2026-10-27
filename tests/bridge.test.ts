import assert from 'node:assert/strict';
import { request } from 'node:http';
import { once } from 'node:events';
import test from 'node:test';
import { createBridgeServer } from '../apps/bridge/src/server.js';
import { fakeAdapter } from './fixtures.js';
import { loadWpsHost } from './wps-fixture.js';

test('bridge serves health and assets without exposing execution or cross-origin access', async (t) => {
  const server = createBridgeServer({ adapter: fakeAdapter });
  server.listen(0, '127.0.0.1');
  await once(server, 'listening');
  t.after(
    () =>
      new Promise<void>((resolve, reject) => {
        server.close((error) => (error ? reject(error) : resolve()));
        server.closeAllConnections();
      }),
  );
  const address = server.address();
  assert.ok(address && typeof address !== 'string');
  const base = `http://127.0.0.1:${address.port}`;

  const health = await fetch(`${base}/api/v1/health`);
  assert.equal(health.status, 200);
  assert.deepEqual(await health.json(), {
    service: 'office-bridge',
    status: 'ok',
    protocolVersion: '0.6.0',
  });
  const agents = await fetch(`${base}/api/v1/agents`);
  const body = (await agents.json()) as { agents: { status: string }[] };
  assert.equal(body.agents.length, 5);
  assert.equal(body.agents[0]?.status, 'ready');
  assert.ok(
    body.agents.slice(1).every((agent) => agent.status === 'not_implemented'),
  );

  const page = await fetch(`${base}/wps/taskpane.html`);
  assert.match(await page.text(), /OfficeBridge/);
  assert.match(page.headers.get('content-type') ?? '', /text\/html/);
  const bootstrap = await fetch(`${base}/wps/index.html`);
  assert.equal(bootstrap.status, 200);
  assert.match(await bootstrap.text(), /src="\.\/main.js"/);
  const installer = await fetch(`${base}/wps/install.html`);
  assert.equal(installer.status, 200);
  const installerHtml = await installer.text();
  assert.ok(installerHtml.includes(`${base}/wps/`));
  assert.ok(installerHtml.includes('"name":"OfficeBridge"'));
  assert.ok(!installerHtml.includes('PUBLISH_REPLACE_STRING'));
  assert.ok(!installerHtml.includes('SERVERID_REPLEASE_STRING'));
  assert.ok(!installerHtml.includes('onclick="ClearAll()"'));
  assert.match(
    installer.headers.get('content-security-policy') ?? '',
    /127\.0\.0\.1:58890/,
  );
  assert.ok(
    !page.headers.get('content-security-policy')?.includes('unsafe-inline'),
  );
  assert.equal(
    (
      await fetch(`${base}/api/v1/health`, {
        headers: { Origin: 'https://example.com' },
      })
    ).status,
    403,
  );
  assert.equal(
    (await fetch(`${base}/api/v1/health`, { method: 'POST', body: '{}' }))
      .status,
    405,
  );
  assert.equal((await fetch(`${base}/wps/%2e%2e/package.json`)).status, 404);
  assert.equal((await fetch(`${base}/api/v1/execute`)).status, 404);

  const invalidHostStatus = await new Promise<number | undefined>(
    (resolve, reject) => {
      const req = request(
        `${base}/api/v1/health`,
        { headers: { Host: 'example.com' } },
        (res) => {
          res.resume();
          res.on('end', () => resolve(res.statusCode));
        },
      );
      req.on('error', reject);
      req.end();
    },
  );
  assert.equal(invalidHostStatus, 403);
});

test('each conversation routes to its chosen Agent and a failed switch preserves its previous owner', async (t) => {
  const seen: string[] = [],
    closed: string[] = [];
  let failWorkBuddy = false,
    count = 0;
  const make = (id: 'codex' | 'workbuddy') => ({
    ...fakeAdapter,
    id,
    async openConversation() {
      if (id === 'workbuddy' && failWorkBuddy) throw new Error('unavailable');
      return { id: `${id}-${++count}` };
    },
    async closeConversation(conversationId: string) {
      closed.push(conversationId);
    },
    async *run(request: any) {
      seen.push(`${id}:${request.conversationId}`);
      yield { type: 'complete' as const };
    },
  });
  const server = createBridgeServer({
    adapters: [make('codex'), make('workbuddy')],
  });
  server.listen(0, '127.0.0.1');
  await once(server, 'listening');
  t.after(() => {
    server.closeAllConnections();
    server.close();
  });
  const address = server.address();
  assert.ok(address && typeof address !== 'string');
  const base = `http://127.0.0.1:${address.port}`;
  const { token } = (await (
    await fetch(`${base}/api/v1/session`)
  ).json()) as any;
  const post = (path: string, body: unknown) =>
    fetch(base + path, {
      method: 'POST',
      headers: {
        Origin: base,
        'Content-Type': 'application/json',
        'X-OfficeBridge-Token': token,
      },
      body: JSON.stringify(body),
    });
  const agents = (await (await fetch(`${base}/api/v1/agents`)).json()) as any;
  assert.equal(
    agents.agents.find((a: any) => a.id === 'workbuddy').status,
    'ready',
  );
  const first = (await (
    await post('/api/v1/conversations', { agentId: 'codex' })
  ).json()) as any;
  const model = await loadWpsHost();
  const workbook = model.clone(model.host.capture('sheet'));
  const run = async (id: string) => {
    const response = await post('/api/v1/run', {
      conversationId: id,
      prompt: 'test',
      workbook,
    });
    await response.text();
    return response.status;
  };
  failWorkBuddy = true;
  assert.equal(
    (
      await post('/api/v1/conversations', {
        id: first.id,
        agentId: 'workbuddy',
      })
    ).status,
    503,
  );
  assert.equal(await run(first.id), 200);
  assert.deepEqual(closed, []);
  failWorkBuddy = false;
  const second = (await (
    await post('/api/v1/conversations', { id: first.id, agentId: 'workbuddy' })
  ).json()) as any;
  assert.equal(second.agentId, 'workbuddy');
  assert.deepEqual(closed, [first.id]);
  assert.equal(await run(second.id), 200);
  assert.equal(await run(second.id), 200);
  assert.equal(await run(first.id), 410);
  assert.deepEqual(seen, [
    `codex:${first.id}`,
    `workbuddy:${second.id}`,
    `workbuddy:${second.id}`,
  ]);
  assert.equal(
    (
      await post('/api/v1/conversations', {
        id: second.id,
        agentId: 'claude-code',
      })
    ).status,
    400,
  );
  assert.equal(await run(second.id), 200);
  await post('/api/v1/conversations/close', { id: second.id });
  assert.equal(closed.at(-1), second.id);
});

test('conversation lifecycle is authenticated, replacement failures preserve the old conversation and closed IDs cannot run', async (t) => {
  const live = new Set<string>(),
    closed: string[] = [],
    requests: any[] = [];
  let count = 0,
    failOpen = false,
    disposed = false;
  const adapter = {
    ...fakeAdapter,
    async openConversation() {
      if (failOpen) throw new Error('agent unavailable');
      const id = 'conversation-' + ++count;
      live.add(id);
      return { id };
    },
    async closeConversation(id: string) {
      live.delete(id);
      closed.push(id);
    },
    dispose() {
      disposed = true;
      live.clear();
    },
    async *run(request: any) {
      requests.push(request);
      assert.ok(live.has(request.conversationId));
      yield { type: 'complete' as const };
    },
  };
  const server = createBridgeServer({ adapter });
  server.listen(0, '127.0.0.1');
  await once(server, 'listening');
  t.after(() => {
    server.closeAllConnections();
    server.close();
  });
  const address = server.address();
  assert.ok(address && typeof address !== 'string');
  const base = `http://127.0.0.1:${address.port}`;
  const { token } = (await (
    await fetch(`${base}/api/v1/session`)
  ).json()) as any;
  const headers = {
    Origin: base,
    'Content-Type': 'application/json',
    'X-OfficeBridge-Token': token,
  };
  const post = (path: string, body: unknown, custom = headers) =>
    fetch(base + path, {
      method: 'POST',
      headers: custom,
      body: JSON.stringify(body),
    });
  assert.equal(
    (
      await post(
        '/api/v1/conversations',
        {},
        { ...headers, 'X-OfficeBridge-Token': 'wrong' },
      )
    ).status,
    401,
  );
  assert.equal(count, 0);
  const first = (await (await post('/api/v1/conversations', {})).json()) as any;
  const model = await loadWpsHost();
  const workbook = model.clone(model.host.capture('sheet'));
  const run = async (id: string, feedback = '') => {
    const response = await post('/api/v1/run', {
      prompt: 'continue',
      workbook,
      conversationId: id,
      operationFeedback: feedback,
    });
    await response.text();
    return response.status;
  };
  assert.equal(await run(first.id), 200);
  assert.equal(await run(first.id, 'applied'), 200);
  assert.equal(count, 1);
  assert.equal(requests[0].conversationId, requests[1].conversationId);
  assert.equal(requests[1].operationFeedback, 'applied');
  assert.equal(requests[1].history, undefined);
  failOpen = true;
  assert.equal(
    (await post('/api/v1/conversations', { id: first.id })).status,
    503,
  );
  assert.equal(await run(first.id), 200);
  assert.deepEqual(closed, []);
  failOpen = false;
  const second = (await (
    await post('/api/v1/conversations', { id: first.id })
  ).json()) as any;
  assert.notEqual(first.id, second.id);
  assert.deepEqual(closed, [first.id]);
  assert.equal(await run(first.id), 410);
  assert.equal(await run(second.id), 200);
  assert.equal(
    (await post('/api/v1/conversations/close', { id: second.id })).status,
    200,
  );
  assert.equal(await run(second.id), 410);
  server.closeAllConnections();
  await new Promise<void>((resolve) => server.close(() => resolve()));
  assert.ok(disposed);
});
