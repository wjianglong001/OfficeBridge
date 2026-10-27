import assert from 'node:assert/strict';
import test from 'node:test';
import { once } from 'node:events';
import { RunWatchdog } from '../apps/bridge/src/run-watchdog.js';
import { createBridgeServer } from '../apps/bridge/src/server.js';
import { fakeAdapter, fullSelection } from './fixtures.js';
import type { AgentEvent, AgentRequest } from '../packages/agents/src/index.js';

test('actual progress renews the idle deadline, while the absolute cap and stop remain bounded', (t) => {
  t.mock.timers.enable({ apis: ['setTimeout', 'Date'] });
  let stops = 0;
  const w = new RunWatchdog({
    idleMs: 100,
    maxMs: 500,
    onTimeout: () => stops++,
  });
  t.mock.timers.tick(90);
  w.touch('thinking');
  t.mock.timers.tick(90);
  assert.equal(stops, 0);
  w.touch('tool');
  t.mock.timers.tick(99);
  assert.equal(stops, 0);
  t.mock.timers.tick(1);
  assert.equal(w.reason, 'idle');
  assert.equal(stops, 1);
  assert.match(w.message('WorkBuddy'), /准备工具参数/);
  w.touch('agent');
  t.mock.timers.tick(1000);
  assert.equal(stops, 1);
  const cap = new RunWatchdog({
    idleMs: 100,
    maxMs: 250,
    onTimeout: () => stops++,
  });
  for (let i = 0; i < 3; i++) {
    t.mock.timers.tick(80);
    cap.touch('thinking');
  }
  t.mock.timers.tick(10);
  assert.equal(cap.reason, 'total');
  assert.equal(stops, 2);
  const done = new RunWatchdog({
    idleMs: 100,
    maxMs: 250,
    onTimeout: () => stops++,
  });
  done.stop();
  t.mock.timers.tick(1000);
  assert.equal(stops, 2);
});

test('HTTP tasks survive genuine progress, ignore repeated status labels, and stop active work at the total cap', async (t) => {
  for (const mode of ['active', 'status', 'cap'] as const) {
    let aborted = false;
    const adapter = {
      ...fakeAdapter,
      async *run(request: AgentRequest): AsyncIterable<AgentEvent> {
        request.signal.addEventListener(
          'abort',
          () => {
            aborted = true;
          },
          { once: true },
        );
        for (let i = 0; i < 8 && !request.signal.aborted; i++) {
          await new Promise((r) => setTimeout(r, 40));
          yield mode === 'status'
            ? { type: 'status', message: 'unchanged heartbeat' }
            : { type: 'activity', phase: 'thinking' };
        }
        if (request.signal.aborted)
          yield { type: 'error', code: 'cancelled', message: 'stopped' };
        else {
          yield { type: 'text', text: 'completed' };
          yield { type: 'complete' };
        }
      },
    };
    const server = createBridgeServer({
      adapter,
      analysisIdleTimeoutMs: 150,
      analysisMaxDurationMs: mode === 'cap' ? 200 : 2000,
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
      await fetch(base + '/api/v1/session')
    ).json()) as { token: string };
    const response = await fetch(base + '/api/v1/analyze', {
      method: 'POST',
      headers: {
        Origin: base,
        'Content-Type': 'application/json',
        'X-OfficeBridge-Token': token,
      },
      body: JSON.stringify({ prompt: 'test', selection: fullSelection }),
    });
    const events = (await response.text())
      .trim()
      .split('\n')
      .map((s) => JSON.parse(s));
    if (mode === 'active') {
      assert.equal(events.at(-1).type, 'complete');
      assert.equal(aborted, false);
    } else {
      assert.equal(events.at(-1).code, 'timeout');
      assert.equal(aborted, true);
      assert.equal(events.filter((e) => e.code === 'timeout').length, 1);
      assert.doesNotMatch(events.at(-1).message, /缩小选区/);
      if (mode === 'cap') assert.match(events.at(-1).message, /总时长/);
    }
    server.closeAllConnections();
    await new Promise<void>((r) => server.close(() => r()));
  }
});
