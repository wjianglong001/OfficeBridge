import assert from 'node:assert/strict';
import test from 'node:test';
import { once } from 'node:events';
import {
  createDoubaoConnection,
  DoubaoConnectionChangedError,
} from '../packages/agents/src/doubao-connection.js';
import { DoubaoDiscoveryError } from '../packages/agents/src/doubao-cdp.js';
import { createBridgeServer } from '../apps/bridge/src/server.js';
import { fakeAdapter } from './fixtures.js';

const unavailable = () =>
  new DoubaoDiscoveryError('endpoint_unavailable', '接口未开启');
function fixture() {
  let enabled = false,
    running = true,
    stopped = 0,
    started = 0,
    clock = 1000;
  let processes = [
    {
      pid: 100,
      created: 'test-start',
      executable: 'D:/Doubao.exe',
      main: true,
    },
  ];
  const connection = createDoubaoConnection({
    probe: async () => {
      if (!enabled) throw unavailable();
      return 'ws://127.0.0.1:9225/test';
    },
    discover: async () => ({ executable: 'D:/Doubao.exe', running, processes }),
    stop: async () => {
      stopped++;
      running = false;
    },
    start: async () => {
      started++;
      running = enabled = true;
      return { started: true, message: 'ready' };
    },
    now: () => clock,
  });
  return {
    connection,
    setEnabled: (v: boolean) => {
      enabled = v;
    },
    setRunning: (v: boolean) => {
      running = v;
    },
    setProcesses: (v: typeof processes) => {
      processes = v;
    },
    expire: () => {
      clock += 120001;
    },
    counts: () => ({ stopped, started }),
  };
}

test('Doubao integrated connect reuses enabled clients, starts stopped clients, and only restarts after confirmation', async () => {
  const f = fixture();
  f.setEnabled(true);
  assert.equal((await f.connection.prepare()).status, 'ready');
  assert.deepEqual(f.counts(), { stopped: 0, started: 0 });
  f.setEnabled(false);
  const prompt = await f.connection.prepare();
  assert.equal(prompt.status, 'confirmation_required');
  assert.deepEqual(f.counts(), { stopped: 0, started: 0 }); // cancel = no mutation
  if (prompt.status !== 'confirmation_required') throw new Error();
  assert.doesNotMatch(JSON.stringify(prompt), /D:|test-start|"pid"/);
  assert.equal(
    (await f.connection.prepare(prompt.restartToken)).status,
    'ready',
  );
  assert.deepEqual(f.counts(), { stopped: 1, started: 1 });
  await assert.rejects(f.connection.prepare(prompt.restartToken), /过期/);
  f.setRunning(false);
  f.setEnabled(false);
  assert.equal((await f.connection.prepare()).status, 'ready');
  assert.deepEqual(f.counts(), { stopped: 1, started: 2 });
});

test('Doubao confirmation expires, is bound to the main process, and permits ordinary child-process churn', async () => {
  const f = fixture();
  const a = await f.connection.prepare();
  if (a.status !== 'confirmation_required') throw new Error();
  f.expire();
  await assert.rejects(f.connection.prepare(a.restartToken), /过期/);
  assert.deepEqual(f.counts(), { stopped: 0, started: 0 });
  const b = await f.connection.prepare();
  if (b.status !== 'confirmation_required') throw new Error();
  f.setProcesses([
    { pid: 100, created: 'new-start', executable: 'D:/Doubao.exe', main: true },
  ]);
  const changed = await f.connection.prepare(b.restartToken);
  assert.equal(changed.status, 'confirmation_required');
  assert.deepEqual(f.counts(), { stopped: 0, started: 0 });
  if (changed.status !== 'confirmation_required') throw new Error();
  f.setProcesses([
    { pid: 100, created: 'new-start', executable: 'D:/Doubao.exe', main: true },
    {
      pid: 101,
      created: 'child-start',
      executable: 'D:/Doubao.exe',
      main: false,
    },
  ]);
  assert.equal(
    (await f.connection.prepare(changed.restartToken)).status,
    'ready',
  );
  assert.deepEqual(f.counts(), { stopped: 1, started: 1 });
});

test('Doubao never restarts a repaired connection, another endpoint, or an ambiguous installation', async () => {
  const f = fixture();
  const a = await f.connection.prepare();
  if (a.status !== 'confirmation_required') throw new Error();
  f.setEnabled(true);
  assert.equal((await f.connection.prepare(a.restartToken)).status, 'ready');
  assert.deepEqual(f.counts(), { stopped: 0, started: 0 });
  f.setEnabled(false);
  f.setProcesses([
    {
      pid: 100,
      created: 'test-start',
      executable: 'E:/other/Doubao.exe',
      main: true,
    },
  ]);
  await assert.rejects(f.connection.prepare(), /唯一确认/);
  const wrong = createDoubaoConnection({
    probe: async () => {
      throw new DoubaoDiscoveryError('unexpected_endpoint', 'wrong port');
    },
    stop: async () => {
      throw new Error('must not stop');
    },
    start: async () => {
      throw new Error('must not start');
    },
  });
  await assert.rejects(wrong.prepare(), /wrong port/);
});

test('failed restart is not retried, does not launch, and consumes its confirmation', async () => {
  let stops = 0,
    starts = 0;
  const connection = createDoubaoConnection({
    probe: async () => {
      throw unavailable();
    },
    discover: async () => ({
      executable: 'D:/Doubao.exe',
      running: true,
      processes: [
        { pid: 100, created: 'test', executable: 'D:/Doubao.exe', main: true },
      ],
    }),
    stop: async () => {
      stops++;
      throw new Error('exit failed');
    },
    start: async () => {
      starts++;
      return { started: true, message: '' };
    },
  });
  const a = await connection.prepare();
  if (a.status !== 'confirmation_required') throw new Error();
  await assert.rejects(
    connection.prepare(a.restartToken),
    (error) =>
      error instanceof DoubaoConnectionChangedError &&
      error.message === 'exit failed',
  );
  await assert.rejects(connection.prepare(a.restartToken), /过期/);
  assert.equal(stops, 1);
  assert.equal(starts, 0);
});

test('integrated connection endpoint requires same-origin POST and session token, validates input and serializes setup', async (t) => {
  let preparations = 0;
  let release: (() => void) | undefined;
  const server = createBridgeServer({
    adapters: [{ ...fakeAdapter, id: 'doubao' }],
    doubaoConnection: {
      prepare: async () => {
        preparations++;
        await new Promise<void>((resolve) => {
          release = resolve;
        });
        return { status: 'ready' as const, changed: false };
      },
    },
  });
  server.listen(0, '127.0.0.1');
  await once(server, 'listening');
  t.after(() => {
    release?.();
    server.closeAllConnections();
    server.close();
  });
  const address = server.address();
  assert.ok(address && typeof address !== 'string');
  const base = `http://127.0.0.1:${address.port}`,
    url = base + '/api/v1/agents/doubao/connect';
  const { token } = (await (await fetch(base + '/api/v1/session')).json()) as {
    token: string;
  };
  const headers = {
    Origin: base,
    'Content-Type': 'application/json',
    'X-OfficeBridge-Token': token,
  };
  assert.equal((await fetch(url)).status, 405);
  assert.equal((await fetch(url, { method: 'POST', body: '{}' })).status, 401);
  assert.equal(
    (
      await fetch(url, {
        method: 'POST',
        headers: { ...headers, Origin: 'https://foreign.example' },
        body: '{}',
      })
    ).status,
    403,
  );
  for (const body of [
    { restartToken: true },
    { executable: 'other.exe' },
    null,
  ])
    assert.equal(
      (
        await fetch(url, {
          method: 'POST',
          headers,
          body: JSON.stringify(body),
        })
      ).status,
      400,
    );
  await fetch(base + '/api/v1/agents');
  assert.equal(preparations, 0);
  const pending = fetch(url, { method: 'POST', headers, body: '{}' });
  while (!release) await new Promise((resolve) => setTimeout(resolve, 5));
  assert.equal(
    (await fetch(url, { method: 'POST', headers, body: '{}' })).status,
    409,
  );
  assert.equal(
    (
      await fetch(base + '/api/v1/conversations', {
        method: 'POST',
        headers,
        body: JSON.stringify({ agentId: 'doubao' }),
      })
    ).status,
    409,
  );
  release();
  assert.equal((await pending).status, 200);
  assert.equal(preparations, 1);
});

test('successful client restart invalidates old Doubao bindings but preserves other Agent conversations', async (t) => {
  const closed: string[] = [];
  let changed = false;
  let fail = false;
  const make = (id: 'codex' | 'doubao') => ({
    ...fakeAdapter,
    id,
    async openConversation() {
      return { id: id + '-conversation' };
    },
    async closeConversation(id: string) {
      closed.push(id);
    },
  });
  const server = createBridgeServer({
    adapters: [make('codex'), make('doubao')],
    doubaoConnection: {
      prepare: async () => {
        if (fail) throw new DoubaoConnectionChangedError('restart failed');
        return { status: 'ready', changed };
      },
    },
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
  const { token } = (await (await fetch(base + '/api/v1/session')).json()) as {
    token: string;
  };
  const headers = {
    Origin: base,
    'Content-Type': 'application/json',
    'X-OfficeBridge-Token': token,
  };
  const post = (path: string, body: unknown) =>
    fetch(base + path, { method: 'POST', headers, body: JSON.stringify(body) });
  for (const agentId of ['codex', 'doubao'])
    assert.equal(
      (await post('/api/v1/conversations', { agentId })).status,
      200,
    );
  await post('/api/v1/agents/doubao/connect', {});
  assert.deepEqual(closed, []);
  changed = true;
  await post('/api/v1/agents/doubao/connect', {});
  assert.deepEqual(closed, ['doubao-conversation']);
  await post('/api/v1/conversations', { agentId: 'doubao' });
  fail = true;
  const failure = await post('/api/v1/agents/doubao/connect', {});
  assert.equal(failure.status, 503);
  assert.equal(((await failure.json()) as { changed: boolean }).changed, true);
  assert.deepEqual(closed, ['doubao-conversation', 'doubao-conversation']);
});
