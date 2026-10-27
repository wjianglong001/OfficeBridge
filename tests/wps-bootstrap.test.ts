import assert from 'node:assert/strict';
import test from 'node:test';
import { readFile } from 'node:fs/promises';
import { createContext, runInContext } from 'node:vm';

const source = await readFile(
  new URL('../apps/wps-bootstrap/launcher.js', import.meta.url),
  'utf8',
);
const ENTRY = 'http://127.0.0.1:4318/wps/taskpane.html';
const READY = { state: 'ready', entryUrl: ENTRY };
const OFFLINE = { state: 'offline', entryUrl: ENTRY };
const STARTING = { state: 'starting', entryUrl: ENTRY };
type Reply = { status?: number; body?: unknown; json?: () => Promise<unknown> };
type Request = { path: string; options: any };
class Element {
  textContent = '';
  hidden = false;
  attributes: Record<string, string> = {};
  listeners: Record<string, () => void> = {};
  setAttribute(name: string, value: string) {
    this.attributes[name] = value;
  }
  addEventListener(name: string, listener: () => void) {
    this.listeners[name] = listener;
  }
}
function deferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>((yes) => {
    resolve = yes;
  });
  return { promise, resolve };
}
async function flush() {
  for (let n = 0; n < 40; n++) await Promise.resolve();
}
function fixture(
  options: {
    statuses?: unknown[];
    startReply?: Reply;
    host?: boolean;
    origin?: string;
    request?: (request: Request) => Reply | Promise<Reply> | undefined;
  } = {},
) {
  let now = 1000,
    nextTimer = 0,
    statusIndex = 0,
    legacyCalls = 0;
  const timers = new Map<number, { due: number; callback: () => void }>();
  const elements = Object.fromEntries(
    ['launcher', 'status', 'hint', 'retry', 'cancel'].map((id) => [
      id,
      new Element(),
    ]),
  ) as Record<string, Element>;
  const requests: Request[] = [];
  const redirects: string[] = [];
  const events: Record<string, () => void> = {};
  const statuses = options.statuses ?? [READY];
  const context: any = {
    document: { getElementById: (id: string) => elements[id] },
    window: {
      location: {
        origin: options.origin ?? 'http://127.0.0.1:43192',
        replace: (url: string) => redirects.push(url),
      },
      addEventListener: (name: string, listener: () => void) => {
        events[name] = listener;
      },
    },
    AbortController,
    Date: { now: () => now },
    setTimeout: (callback: () => void, milliseconds: number) => {
      const id = ++nextTimer;
      timers.set(id, { callback, due: now + milliseconds });
      return id;
    },
    clearTimeout: (id: number) => timers.delete(id),
    fetch: async (path: string, requestOptions: any) => {
      const request = { path, options: requestOptions };
      requests.push(request);
      let reply = await options.request?.(request);
      if (!reply) {
        if (path === '/native/bridge-status')
          reply = {
            body: statuses[Math.min(statusIndex++, statuses.length - 1)],
          };
        else if (path === '/probe/session')
          reply = { body: { token: 'synthetic-session-token' } };
        else if (path === '/native/start')
          reply = options.startReply ?? {
            status: 202,
            body: { state: 'starting' },
          };
        else throw new Error('Unexpected request: ' + path);
      }
      const status = reply.status ?? 200;
      return {
        status,
        ok: status >= 200 && status < 300,
        json: reply.json ?? (async () => reply!.body),
      };
    },
  };
  if (options.host !== false) {
    context.wps = {};
    Object.defineProperty(context.wps, 'OAAssist', {
      get() {
        legacyCalls++;
        throw new Error('ShellExecute must not be used');
      },
    });
  }
  runInContext(source, createContext(context));
  const element = (id: string) => elements[id]!;
  return {
    requests,
    redirects,
    timers,
    element,
    legacyCalls: () => legacyCalls,
    click: (id: string) => element(id).listeners.click?.(),
    event: (name: string) => events[name]?.(),
    state: () => element('launcher').attributes['data-state'],
    startBodies: () =>
      requests
        .filter((request) => request.path === '/native/start')
        .map((request) => JSON.parse(request.options.body)),
    async tick(milliseconds: number) {
      now += milliseconds;
      const due = [...timers].filter(([, timer]) => timer.due <= now);
      for (const [id, timer] of due) if (timers.delete(id)) timer.callback();
      await flush();
    },
  };
}

test('bootstrap reuses only the fixed healthy entry with no startup request', async () => {
  const h = fixture({ statuses: [{ ...READY, errorCode: 5, exitCode: 1 }] });
  await flush();
  assert.deepEqual(h.redirects, [ENTRY]);
  assert.deepEqual(h.startBodies(), []);
  assert.equal(h.requests.length, 1);
  assert.equal(h.legacyCalls(), 0);
  assert.equal(h.timers.size, 0);
  const wrong = fixture({
    statuses: [{ state: 'ready', entryUrl: 'https://example.invalid/' }],
  });
  await flush();
  assert.equal(wrong.state(), 'error');
  assert.equal(wrong.redirects.length, 0);
  assert.match(wrong.element('hint').textContent, /入口未通过核对/);
});

test('offline requests one native startup with token and retry only, then waits for health', async () => {
  const h = fixture({ statuses: [OFFLINE, OFFLINE, READY] });
  await flush();
  assert.deepEqual(h.startBodies(), [{ retry: false }]);
  const post = h.requests.find((request) => request.path === '/native/start')!;
  assert.equal(
    post.options.headers['X-OfficeBridge-Probe-Token'],
    'synthetic-session-token',
  );
  assert.equal(post.options.credentials, 'same-origin');
  assert.equal(post.options.redirect, 'error');
  assert.equal(h.redirects.length, 0);
  await h.tick(300);
  assert.deepEqual(h.redirects, [ENTRY]);
  assert.equal(h.startBodies().length, 1);
  assert.equal(h.legacyCalls(), 0);
  assert.equal(h.timers.size, 0);
});

test('a 200 ready or failed start response always requires separate health verification', async () => {
  const failed = { state: 'failed', errorCode: 5 };
  for (const body of [READY, failed]) {
    const h = fixture({
      statuses: [OFFLINE, OFFLINE, READY],
      startReply: { status: 200, body },
    });
    await flush();
    assert.equal(h.redirects.length, 0);
    assert.deepEqual(h.startBodies(), [{ retry: false }]);
    await h.tick(300);
    assert.deepEqual(h.redirects, [ENTRY]);
    assert.equal(
      h.requests.filter((request) => request.path === '/native/bridge-status')
        .length,
      3,
    );
  }
  const threadFailure = fixture({
    statuses: [OFFLINE, failed],
    startReply: { status: 200, body: failed },
  });
  await flush();
  assert.equal(threadFailure.state(), 'error');
  assert.match(threadFailure.element('hint').textContent, /系统错误码 5/);
  assert.equal(threadFailure.redirects.length, 0);
  assert.deepEqual(threadFailure.startBodies(), [{ retry: false }]);
  assert.equal(
    threadFailure.requests.filter(
      (request) => request.path === '/native/bridge-status',
    ).length,
    2,
  );
  assert.equal(threadFailure.timers.size, 0);
});
test('an already starting health state only waits and never requests startup again', async () => {
  const h = fixture({ statuses: [STARTING, OFFLINE, READY] });
  await flush();
  assert.deepEqual(h.startBodies(), []);
  assert.equal(
    h.requests.some((request) => request.path === '/probe/session'),
    false,
  );
  assert.equal(h.redirects.length, 0);
  await h.tick(300);
  assert.deepEqual(h.redirects, [ENTRY]);
  assert.equal(h.legacyCalls(), 0);
});

test('only already_requested 409 waits; runtime, conflict and all other rejections stop', async () => {
  const waiting = fixture({
    statuses: [OFFLINE, READY],
    startReply: { status: 409, body: { error: 'already_requested' } },
  });
  await flush();
  assert.deepEqual(waiting.redirects, [ENTRY]);
  assert.deepEqual(waiting.startBodies(), [{ retry: false }]);
  for (const [error, expected] of [
    ['runtime_unavailable', /安装文件缺失或不匹配/],
    ['port_conflict', /端口被其他程序占用/],
    ['already_running', /启动请求未确认/],
    ['unknown', /启动请求未确认/],
  ] as const) {
    const h = fixture({
      statuses: [OFFLINE, READY],
      startReply: { status: 409, body: { error } },
    });
    await flush();
    assert.equal(h.state(), 'error');
    assert.match(h.element('hint').textContent, expected);
    assert.equal(
      h.requests.filter((request) => request.path === '/native/bridge-status')
        .length,
      1,
    );
    assert.equal(h.redirects.length, 0);
    assert.equal(h.legacyCalls(), 0);
    assert.equal(h.timers.size, 0);
  }
});

test('native failure is immediate and preserves system errors and confirmed zero exit codes', async () => {
  for (const [failed, expected] of [
    [{ state: 'failed', errorCode: 5, exitCode: null }, /系统错误码 5/],
    [{ state: 'failed', errorCode: 0, exitCode: 0 }, /进程退出码 0/],
    [
      { state: 'failed', errorCode: 0, exitCode: 3221225781 },
      /进程退出码 3221225781/,
    ],
  ] as const) {
    const h = fixture({ statuses: [OFFLINE, failed] });
    await flush();
    assert.equal(h.state(), 'error');
    assert.match(h.element('hint').textContent, expected);
    assert.equal(h.redirects.length, 0);
    assert.equal(h.startBodies().length, 1);
    assert.equal(h.timers.size, 0);
  }
  const existing = fixture({
    statuses: [{ state: 'failed', errorCode: 5, exitCode: 259 }],
  });
  await flush();
  assert.deepEqual(existing.startBodies(), []);
  assert.match(existing.element('hint').textContent, /系统错误码 5/);
  assert.doesNotMatch(existing.element('hint').textContent, /退出码 259/);
  const malformed = fixture({
    statuses: [
      { state: 'failed', errorCode: '<script>', exitCode: 'private internals' },
    ],
  });
  await flush();
  assert.equal(malformed.state(), 'error');
  assert.doesNotMatch(
    malformed.element('hint').textContent,
    /script|private internals/,
  );
});

test('failed native startup never retries automatically; explicit retry can replace the failed state', async () => {
  const failed = { state: 'failed', errorCode: 5, exitCode: null };
  const h = fixture({ statuses: [OFFLINE, failed, failed, READY] });
  await flush();
  assert.equal(h.state(), 'error');
  assert.deepEqual(h.startBodies(), [{ retry: false }]);
  h.click('retry');
  h.click('retry');
  await flush();
  assert.deepEqual(h.startBodies(), [{ retry: false }, { retry: true }]);
  assert.deepEqual(h.redirects, [ENTRY]);
  assert.equal(h.legacyCalls(), 0);
});

test('cancel aborts fetches and late results cannot navigate or replace a newer attempt', async () => {
  const pending = deferred<Reply>();
  let statusCount = 0;
  const h = fixture({
    request: (request) =>
      request.path === '/native/bridge-status' && ++statusCount === 1
        ? pending.promise
        : undefined,
  });
  await flush();
  const signal = h.requests[0]!.options.signal as AbortSignal;
  h.click('cancel');
  assert.equal(signal.aborted, true);
  assert.equal(h.state(), 'cancelled');
  h.click('retry');
  await flush();
  assert.deepEqual(h.redirects, [ENTRY]);
  pending.resolve({ body: READY });
  await flush();
  assert.deepEqual(h.redirects, [ENTRY]);
  assert.equal(h.startBodies().length, 0);
  assert.equal(h.timers.size, 0);
});

test('waiting has a 20 second bound; late healthy data and page hide cannot resume or restart', async () => {
  const h = fixture({ statuses: [OFFLINE] });
  await flush();
  await h.tick(20001);
  assert.equal(h.state(), 'error');
  assert.match(h.element('hint').textContent, /等待超时/);
  assert.equal(h.startBodies().length, 1);
  assert.equal(h.redirects.length, 0);
  assert.equal(h.timers.size, 0);
  let count = 0;
  const pending = deferred<unknown>();
  const late = fixture({
    statuses: [OFFLINE],
    request: (request) =>
      request.path === '/native/bridge-status' && ++count === 2
        ? { json: () => pending.promise }
        : undefined,
  });
  await flush();
  await late.tick(20001);
  pending.resolve(READY);
  await flush();
  assert.equal(late.state(), 'error');
  assert.equal(late.redirects.length, 0);
  assert.equal(late.startBodies().length, 1);
  const hidden = fixture({ statuses: [OFFLINE] });
  await flush();
  hidden.event('pagehide');
  await flush();
  await hidden.tick(20001);
  assert.equal(hidden.startBodies().length, 1);
  assert.equal(hidden.redirects.length, 0);
  assert.equal(hidden.timers.size, 0);
  assert.ok(
    hidden.requests.every((request) =>
      ['/native/bridge-status', '/probe/session', '/native/start'].includes(
        request.path,
      ),
    ),
  );
});

test('missing host, conflicts, network loss and malformed startup responses safely stop', async () => {
  const noHost = fixture({ host: false });
  await flush();
  assert.equal(noHost.requests.length, 0);
  assert.match(noHost.element('hint').textContent, /请在 WPS 表格中/);
  const wrongOrigin = fixture({ origin: 'http://localhost:43192' });
  await flush();
  assert.equal(wrongOrigin.requests.length, 0);
  const conflict = fixture({ statuses: [{ state: 'conflict' }] });
  await flush();
  assert.equal(conflict.state(), 'error');
  assert.equal(conflict.startBodies().length, 0);
  const lost = fixture({
    request: () => {
      throw new TypeError('network internals');
    },
  });
  await flush();
  assert.equal(lost.state(), 'error');
  assert.equal(lost.requests.length, 1);
  assert.doesNotMatch(lost.element('hint').textContent, /network internals/);
  for (const startReply of [
    { status: 202, body: READY },
    { status: 200, body: { state: 'starting' } },
    { status: 200, body: { executable: 'ignored', arguments: 'ignored' } },
    { status: 202, body: [] },
  ]) {
    const h = fixture({ statuses: [OFFLINE], startReply });
    await flush();
    assert.equal(h.state(), 'error');
    assert.equal(h.redirects.length, 0);
    assert.equal(h.legacyCalls(), 0);
  }
});
