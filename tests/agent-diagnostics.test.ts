import assert from 'node:assert/strict';
import test from 'node:test';
import { workBuddyFailure } from '../packages/agents/src/workbuddy-errors.js';
import { startDoubaoConnection } from '../packages/agents/src/doubao-launch.js';
import {
  findDoubaoTarget,
  DoubaoDiscoveryError,
} from '../packages/agents/src/doubao-cdp.js';
import { connectDoubao } from '../packages/agents/src/doubao-driver.js';

const unavailable = () =>
  new DoubaoDiscoveryError('endpoint_unavailable', 'endpoint unavailable');

test('WorkBuddy quota, rate, auth, refusal and limits use specific safe messages without exposing native metadata', () => {
  const secret = 'Bearer private-secret';
  const quota = {
    stopReason: 'refusal',
    _meta: {
      'codebuddy.ai/errorMessage': JSON.stringify({
        code: -32003,
        message: '429 too many requests ' + secret,
        data: {
          statusCode: 429,
          code: 14003,
          category: 'quota',
          details: 'https://provider/private?api_key=secret',
        },
      }),
    },
  };
  assert.equal(workBuddyFailure(quota).code, 'agent_quota');
  assert.match(workBuddyFailure(quota).message, /429/);
  assert.match(workBuddyFailure(quota).message, /不足以判断积分余额/);
  assert.doesNotMatch(
    workBuddyFailure(quota).message,
    /积分(?:已用完|耗尽|不足)|请.*查看额度/,
  );
  assert.equal(
    workBuddyFailure({
      stopReason: 'refusal',
      errorMessage: JSON.stringify({
        data: { httpStatus: 429, bizCode: 14003 },
      }),
    }).code,
    'agent_quota',
  );
  assert.match(
    workBuddyFailure({
      stopReason: 'refusal',
      errorMessage: JSON.stringify({
        data: { httpStatus: 429, bizCode: 14003 },
      }),
    }).message,
    /服务码 14003/,
  );
  assert.equal(
    workBuddyFailure({
      stopReason: 'refusal',
      errorMessage: JSON.stringify({ data: { statusCode: 429 } }),
    }).code,
    'agent_rate_limited',
  );
  assert.equal(
    workBuddyFailure({
      stopReason: 'refusal',
      errorMessage: '429 too many requests',
    }).code,
    'agent_rate_limited',
  );
  assert.equal(
    workBuddyFailure({
      stopReason: 'refusal',
      errorMessage: JSON.stringify({ data: { statusCode: 401 } }),
    }).code,
    'agent_auth',
  );
  assert.equal(
    workBuddyFailure({
      stopReason: 'refusal',
      errorMessage: 'Permission denied for tool(s): x',
    }).code,
    'agent_permission',
  );
  assert.equal(
    workBuddyFailure({ stopReason: 'refusal', finishReason: 'content_filter' })
      .code,
    'agent_refused',
  );
  for (const [reason, code] of [
    ['max_tokens', 'agent_token_limit'],
    ['max_turn_requests', 'agent_turn_limit'],
    ['cancelled', 'cancelled'],
    ['end_turn', 'agent_empty'],
  ])
    assert.equal(workBuddyFailure({ stopReason: reason }).code, code);
  assert.equal(
    workBuddyFailure({ ...quota, stopReason: 'cancelled' }).code,
    'cancelled',
  );
  for (const input of [
    quota,
    { stopReason: 'refusal', errorMessage: secret },
    {
      stopReason: 'strange',
      errorMessage: 'https://api.example/?token=secret',
    },
  ]) {
    assert.doesNotMatch(
      JSON.stringify(workBuddyFailure(input)),
      /private-secret|api_key|https?:|Bearer/,
    );
  }
});

test('Doubao unavailable endpoint retains an actionable setup reason instead of being called version incompatible', async (t) => {
  t.mock.method(globalThis, 'fetch', async () => {
    throw new Error('socket failure with private URL');
  });
  await assert.rejects(
    connectDoubao(),
    (error) =>
      error instanceof Error &&
      error.name === 'DoubaoSetupError' &&
      error.message.includes('9225') &&
      error.message.includes('接口未开启') &&
      !error.message.includes('private'),
  );
});

test('Doubao discovery distinguishes wrong endpoint, missing and ambiguous chat windows without reading conversations', async (t) => {
  let body: any = [],
    status = 200;
  t.mock.method(
    globalThis,
    'fetch',
    async () => new Response(JSON.stringify(body), { status }),
  );
  await assert.rejects(
    findDoubaoTarget('http://127.0.0.1:9225'),
    (error) =>
      error instanceof DoubaoDiscoveryError && error.code === 'no_chat',
  );
  body = [1, 2].map(() => ({
    type: 'page',
    url: 'doubao://doubao-chat/chat/',
    webSocketDebuggerUrl: 'ws://127.0.0.1:9225/x',
  }));
  await assert.rejects(
    findDoubaoTarget('http://127.0.0.1:9225'),
    (error) =>
      error instanceof DoubaoDiscoveryError && error.code === 'ambiguous_chat',
  );
  status = 404;
  await assert.rejects(
    findDoubaoTarget('http://127.0.0.1:9225'),
    (error) =>
      error instanceof DoubaoDiscoveryError &&
      error.code === 'unexpected_endpoint',
  );
});

test('Doubao launch helper reuses an enabled interface and refuses to stop a running client', async () => {
  let spawned = 0;
  const ready = await startDoubaoConnection({
    probe: async () => true,
    discover: async () => {
      throw new Error('must not discover');
    },
  });
  assert.equal(ready.started, false);
  await assert.rejects(
    startDoubaoConnection({
      probe: async () => {
        throw unavailable();
      },
      discover: async () => ({ executable: 'D:/Doubao.exe', running: true }),
      launch: async () => {
        spawned++;
      },
    }),
    /不会终止/,
  );
  assert.equal(spawned, 0);
});

test('explicit Doubao helper starts once with loopback-only parameters, waits for metadata and never changes config', async () => {
  let probes = 0,
    spawned = 0;
  const calls: unknown[] = [];
  const result = await startDoubaoConnection({
    probe: async () => {
      if (++probes < 3) throw unavailable();
      return true;
    },
    discover: async () => ({ executable: 'D:/Doubao.exe', running: false }),
    launch: async (exe, args) => {
      spawned++;
      calls.push({ exe, args });
    },
    wait: async () => {},
  });
  assert.equal(result.started, true);
  assert.equal(spawned, 1);
  assert.deepEqual(calls, [
    {
      exe: 'D:/Doubao.exe',
      args: [
        '--remote-debugging-address=127.0.0.1',
        '--remote-debugging-port=9225',
      ],
    },
  ]);
  assert.doesNotMatch(
    JSON.stringify(calls),
    /kill|taskkill|disable|allow-origins|profile|credential|config/,
  );
});

test('Doubao launch timeout does not restart, and an occupied non-Doubao endpoint is never taken over', async () => {
  let spawned = 0;
  const options = {
    probe: async () => {
      throw unavailable();
    },
    discover: async () => ({ executable: 'D:/Doubao.exe', running: false }),
    launch: async () => {
      spawned++;
    },
    wait: async () => {},
  };
  await assert.rejects(startDoubaoConnection(options), /未再次启动/);
  assert.equal(spawned, 1);
  await assert.rejects(
    startDoubaoConnection({
      ...options,
      probe: async () => {
        throw new DoubaoDiscoveryError('unexpected_endpoint', 'wrong port');
      },
    }),
    /wrong port/,
  );
  assert.equal(spawned, 1);
  await assert.rejects(
    startDoubaoConnection({ ...options, endpoint: 'http://example.com' }),
    /只允许/,
  );
  assert.equal(spawned, 1);
});
