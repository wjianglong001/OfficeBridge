import assert from 'node:assert/strict';
import test from 'node:test';
import { once } from 'node:events';
import { Continuations } from '../apps/bridge/src/continuations.js';
import { isApplicationReceipt } from '../packages/protocol/src/continuation.js';
import { createTableTask } from '../packages/agents/src/table-task.js';
import { createBridgeServer } from '../apps/bridge/src/server.js';
import { fakeAdapter } from './fixtures.js';
import { loadWpsHost } from './wps-fixture.js';
import type { AgentRequest, AgentEvent } from '../packages/agents/src/index.js';

const binding = {
  conversationId: 'conversation',
  workbookFullName: 'E:\\test.xlsx',
  snapshotId: 'before',
  nextStep: '在结果表创建折线图',
  step: 1,
};
const receipt = {
  status: 'succeeded' as const,
  snapshotId: 'before',
  workbookFullName: binding.workbookFullName,
  result: { changedCells: 0, newSheets: ['Result'], recoverable: true },
};
const fresh = {
  snapshotId: 'after',
  workbookFullName: binding.workbookFullName,
};

test('continuation receipts are bounded; tokens bind the same conversation, file and applied snapshot and can be used only once', () => {
  const store = new Continuations(),
    offer = store.issue(binding);
  assert.ok(isApplicationReceipt(receipt));
  for (const bad of [
    { ...receipt, status: 'failed' },
    { ...receipt, status: 'cancelled' },
    { ...receipt, result: { changedCells: -1, newSheets: [] } },
    { ...receipt, result: { changedCells: 0, newSheets: [], script: 'x' } },
  ])
    assert.equal(isApplicationReceipt(bad), false);
  assert.equal(store.consume(offer.token, 'other', fresh, receipt), null);
  assert.equal(
    store.consume(
      offer.token,
      binding.conversationId,
      { ...fresh, workbookFullName: 'Other' },
      receipt,
    ),
    null,
  );
  assert.equal(
    store.consume(
      offer.token,
      binding.conversationId,
      { ...fresh, snapshotId: 'before' },
      receipt,
    ),
    null,
  );
  assert.equal(
    store.consume(offer.token, binding.conversationId, fresh, {
      ...receipt,
      snapshotId: 'wrong',
    }),
    null,
  );
  assert.equal(
    store.consume(offer.token, binding.conversationId, fresh, {
      ...receipt,
      status: 'failed',
    }),
    null,
  );
  assert.deepEqual(
    store.consume(offer.token, binding.conversationId, fresh, receipt),
    { nextStep: binding.nextStep, step: 1, application: receipt },
  );
  assert.equal(
    store.consume(offer.token, binding.conversationId, fresh, receipt),
    null,
  );
});

test('cancellation, superseding tasks, expiry and the eight-round bound terminate automatic continuation', (t) => {
  const store = new Continuations();
  let offer = store.issue(binding);
  assert.equal(store.cancel(offer.token, 'other'), false);
  assert.equal(store.cancel(offer.token, binding.conversationId), true);
  assert.equal(
    store.consume(offer.token, binding.conversationId, fresh, receipt),
    null,
  );
  offer = store.issue(binding);
  store.issue(binding);
  assert.equal(
    store.consume(offer.token, binding.conversationId, fresh, receipt),
    null,
  );
  offer = store.issue(binding);
  store.clear(binding.conversationId);
  assert.equal(
    store.consume(offer.token, binding.conversationId, fresh, receipt),
    null,
  );
  offer = store.issue(binding);
  t.mock.method(Date, 'now', () => Number.MAX_SAFE_INTEGER);
  assert.equal(
    store.consume(offer.token, binding.conversationId, fresh, receipt),
    null,
  );
  t.mock.restoreAll();
  assert.ok(store.issue({ ...binding, step: 8 }).token);
  assert.equal(store.issue({ ...binding, step: 9 }).token, null);
});

test('an explicit next step requires a staged plan, awaits queued writes and disappears after errors, questions or cancellation', async () => {
  for (const stop of ['error', 'question', 'abort', 'invalid']) {
    const m = await loadWpsHost(),
      d = m.clone(m.host.beginRead('context')),
      controller = new AbortController();
    const task = createTableTask(
      d,
      async (name, args) =>
        m.clone(m.host.executeQuery(d.snapshotId, name, args)),
      controller.signal,
    );
    await assert.rejects(
      async () =>
        task.run('officebridge_continue_after_apply', { nextStep: 'draw' }),
      /先生成/,
    );
    const write = task.run('officebridge_propose_result_sheet', {
      name: 'Result',
      values: [
        ['Month', 'Amount'],
        ['Jan', 1],
      ],
    });
    const marker = task.run('officebridge_continue_after_apply', {
      nextStep: 'draw',
    });
    await Promise.all([write, marker]);
    assert.deepEqual(task.continuation(), { nextStep: 'draw' });
    assert.equal(m.sheets.length, 1);
    if (stop === 'abort') controller.abort();
    else if (stop === 'question')
      await task.run('officebridge_ask_user', {
        question: '哪个口径？',
        options: ['A', 'B'],
      });
    else
      await assert.rejects(async () =>
        task.run(stop === 'invalid' ? 'unknown' : 'officebridge_read_range', {
          rangeId: 'missing',
        }),
      );
    assert.equal(task.continuation(), null);
  }
});

test('automatic capture sends fresh metadata without reading incidental selection values', async () => {
  const m = await loadWpsHost({ Data: [['header'], ['old']] }),
    d = m.host.beginRead('context');
  assert.ok(d.inline);
  m.host.finishRead(d.snapshotId, false);
  let reads = 0;
  const original = m.sheets[0].Cells.Item;
  m.sheets[0].Cells.Item = (r: number, c: number) => {
    reads++;
    return original(r, c);
  };
  const next = m.host.beginRead('context', m.host.describe('context'), {
    omitSelection: true,
  });
  assert.equal(next.inline, undefined);
  assert.equal(next.navigation.selectionProvided, false);
  assert.notEqual(next.snapshotId, d.snapshotId);
  assert.equal(reads, 0);
});

test('successful application resumes the same Agent with actual results; failed turns, cancellation and manual tasks cannot resume or replay a token', async (t) => {
  const m = await loadWpsHost(),
    calls: AgentRequest[] = [];
  const adapter = {
    ...fakeAdapter,
    async *run(req: AgentRequest): AsyncIterable<AgentEvent> {
      calls.push(req);
      const task = createTableTask(req.workbook!, req.hostTool, req.signal);
      await task.run('officebridge_propose_result_sheet', {
        name: 'Result',
        values: [['ok']],
      });
      await task.run('officebridge_continue_after_apply', {
        nextStep: '在 Result 上创建折线图',
      });
      yield {
        type: 'proposal',
        plan: task.proposal()!,
        snapshot: task.snapshot,
        continuation: task.continuation()!,
      };
      if (req.prompt === 'fail')
        yield { type: 'error', code: 'test', message: 'failed' };
      else if (req.prompt !== 'incomplete') yield { type: 'complete' };
    },
  };
  const server = createBridgeServer({ adapter });
  server.listen(0, '127.0.0.1');
  await once(server, 'listening');
  t.after(() => {
    server.closeAllConnections();
    return new Promise<void>((resolve) => server.close(() => resolve()));
  });
  const address = server.address();
  assert.ok(address && typeof address !== 'string');
  const base = `http://127.0.0.1:${address.port}`;
  const { token } = (await (
    await fetch(base + '/api/v1/session')
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
  const { id } = (await (
    await post('/api/v1/conversations', {})
  ).json()) as any;
  const snapshot = m.clone(m.host.capture('sheet'));
  const run = async (prompt: string) => {
    const response = await post('/api/v1/run', {
      prompt,
      workbook: snapshot,
      conversationId: id,
    });
    assert.equal(response.status, 200);
    return (await response.text())
      .trim()
      .split('\n')
      .map((l) => JSON.parse(l));
  };
  for (const mode of ['fail', 'incomplete']) {
    const events = await run(mode);
    assert.equal(
      events.some((e) => e.type === 'proposal'),
      false,
    );
    assert.equal(events.at(-1).type, 'error');
  }
  const events = await run('multi-step'),
    proposal = events.find((e) => e.type === 'proposal');
  assert.ok(proposal.continuation.token);
  assert.equal(events.at(-1).type, 'complete');
  const body = {
    prompt: 'continue',
    workbook: { ...snapshot, snapshotId: 'fresh' },
    conversationId: id,
    resume: {
      token: proposal.continuation.token,
      application: {
        status: 'succeeded',
        snapshotId: snapshot.snapshotId,
        workbookFullName: snapshot.workbookFullName,
        result: { changedCells: 0, newSheets: ['Result'], recoverable: true },
      },
    },
  };
  const resumed = await post('/api/v1/run', body);
  assert.equal(resumed.status, 200);
  await resumed.text();
  assert.equal(calls.at(-1)!.conversationId, id);
  assert.match(calls.at(-1)!.prompt, /在 Result 上创建折线图/);
  assert.deepEqual(
    JSON.parse(calls.at(-1)!.operationFeedback!),
    body.resume.application,
  );
  const count = calls.length;
  assert.equal((await post('/api/v1/run', body)).status, 410);
  assert.equal(calls.length, count);
  const cancelled = (await run('multi-step')).find(
    (e) => e.type === 'proposal',
  );
  await post('/api/v1/continuations/cancel', {
    token: cancelled.continuation.token,
    conversationId: id,
  });
  assert.equal(
    (
      await post('/api/v1/run', {
        ...body,
        resume: { ...body.resume, token: cancelled.continuation.token },
      })
    ).status,
    410,
  );
  const superseded = (await run('multi-step')).find(
    (e) => e.type === 'proposal',
  );
  await run('another manual task');
  assert.equal(
    (
      await post('/api/v1/run', {
        ...body,
        resume: { ...body.resume, token: superseded.continuation.token },
      })
    ).status,
    410,
  );
});
