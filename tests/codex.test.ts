import assert from 'node:assert/strict';
import test from 'node:test';
import { CodexAdapter } from '../packages/agents/src/codex.js';
import type {
  RpcMessage,
  RpcTransport,
} from '../packages/agents/src/codex-rpc.js';
import { fullSelection } from './fixtures.js';
import { loadWpsHost } from './wps-fixture.js';

test('Codex awaits on-demand host tools and publishes the verified write snapshot without eagerly reading workbook data', async () => {
  const model = await loadWpsHost();
  model.select(1, 1);
  const workbook = model.clone(model.host.beginRead('context'));
  const rpc = new FakeRpc();
  rpc.toolCalls = [
    { tool: 'officebridge_resolve_range', arguments: { target: 'sheet' } },
    {
      tool: 'officebridge_propose_edits',
      arguments: { edits: [{ rangeId: 'q1', row: 1, column: 1, value: 50 }] },
    },
    {
      tool: 'officebridge_continue_after_apply',
      arguments: { nextStep: '创建图表' },
    },
  ];
  let calls = 0;
  const adapter = new CodexAdapter({ transport: async () => rpc });
  const events = await collect(
    adapter.run({
      ...request(),
      workbook,
      hostTool: async (name, args) => {
        calls++;
        await new Promise((r) => setTimeout(r, 5));
        return model.clone(
          model.host.executeQuery(workbook.snapshotId, name, args),
        );
      },
    }),
  );
  assert.equal(calls, 2);
  const proposal = events.find((e) => e.type === 'proposal');
  assert.ok(proposal?.type === 'proposal');
  assert.deepEqual(proposal.continuation, { nextStep: '创建图表' });
  assert.ok(proposal.snapshot);
  assert.equal(model.host.queryMetrics().baselineCaptures, 0);
  model.host.apply(proposal.plan);
  assert.equal(model.sheets[0].Cells.Item(2, 2).Value2, 50);
  model.host.undo();
  adapter.dispose();
});

async function collect<T>(source: AsyncIterable<T>): Promise<T[]> {
  const result: T[] = [];
  for await (const item of source) result.push(item);
  return result;
}

test('Codex receives the full selection before any tool call and can answer directly from that context', async () => {
  const m = await loadWpsHost({
    Data: Array.from({ length: 40 }, (_, i) => ['record ' + i]),
  });
  m.select(1, 1, 40, 1);
  const workbook = m.clone(m.host.beginRead('context'));
  const rpc = new FakeRpc();
  rpc.toolCalls = [];
  const adapter = new CodexAdapter({ transport: async () => rpc });
  try {
    const events = await collect(
      adapter.run({
        ...request(),
        workbook,
        hostTool: async () => {
          throw new Error('No host read needed');
        },
      }),
    );
    assert.equal(events.at(-1)?.type, 'complete');
    const turn = rpc.messages.find((m) => m.method === 'turn/start')!;
    assert.match(turn.params.input[0].text, /record 39/);
    assert.match(turn.params.input[0].text, /selectionProvided/);
    assert.equal(rpc.replies.length, 0);
  } finally {
    adapter.dispose();
  }
});

class FakeRpc implements RpcTransport {
  closed = false;
  messages: { method: string; params: any }[] = [];
  replies: any[] = [];
  rejected: (number | string)[] = [];
  handler: (message: RpcMessage) => void = () => {};
  failure: (error: Error) => void = () => {};
  authenticated = true;
  outcome = 'completed';
  sendText = true;
  hang = false;
  turnCount = 0;
  toolCalls: { tool: string; arguments: unknown }[] | undefined;
  onMessage(callback: (message: RpcMessage) => void) {
    this.handler = callback;
  }
  onFailure(callback: (error: Error) => void) {
    this.failure = callback;
  }
  notify(method: string, params: unknown) {
    this.messages.push({ method, params });
  }
  reply(id: number | string, result: unknown) {
    this.replies.push({ id, result });
  }
  reject(id: number | string) {
    this.rejected.push(id);
  }
  close() {
    this.closed = true;
  }
  async request(method: string, params: any): Promise<any> {
    this.messages.push({ method, params });
    if (method === 'account/read')
      return {
        account: this.authenticated ? { type: 'chatgpt' } : null,
        requiresOpenaiAuth: true,
      };
    if (method === 'thread/start') return { thread: { id: 'own-thread' } };
    if (method === 'turn/interrupt') {
      this.handler({
        method: 'turn/completed',
        params: {
          threadId: params.threadId,
          turn: { id: params.turnId, status: 'interrupted' },
        },
      });
      return {};
    }
    if (method === 'turn/start') {
      const turnId = 'turn-' + ++this.turnCount;
      if (!this.hang)
        setTimeout(() => {
          this.handler({
            id: 90,
            method: 'item/tool/call',
            params: {
              threadId: 'other-thread',
              turnId,
              tool: 'officebridge_read_selection',
              arguments: {},
            },
          });
          if (!this.toolCalls)
            this.handler({
              id: 91,
              method: 'item/tool/call',
              params: {
                threadId: 'own-thread',
                turnId,
                tool: 'officebridge_summarize_selection',
                arguments: {},
              },
            });
          this.toolCalls?.forEach((call, index) =>
            this.handler({
              id: 100 + index,
              method: 'item/tool/call',
              params: { threadId: 'own-thread', turnId, ...call },
            }),
          );
          if (this.sendText) {
            this.handler({
              method: 'item/agentMessage/delta',
              params: {
                threadId: 'own-thread',
                turnId,
                itemId: 'item',
                delta: '平均时长 20 秒',
              },
            });
            this.handler({
              method: 'item/completed',
              params: {
                threadId: 'own-thread',
                turnId,
                item: {
                  type: 'agentMessage',
                  id: 'item',
                  text: '平均时长 20 秒',
                },
              },
            });
          }
          this.handler({
            method: 'turn/completed',
            params: {
              threadId: 'own-thread',
              turn: { id: turnId, status: this.outcome },
            },
          });
        }, 1);
      return { turn: { id: turnId } };
    }
    return {};
  }
}

function request(signal = new AbortController().signal) {
  return {
    requestId: 'test',
    prompt: '统计时长',
    context: fullSelection,
    signal,
  };
}

test('Codex uses an isolated read-only thread, routes scoped tools and streams without duplicate messages', async () => {
  const rpc = new FakeRpc();
  const adapter = new CodexAdapter({ transport: async () => rpc });
  const events = await collect(adapter.run(request()));
  assert.equal(events.filter((event) => event.type === 'text').length, 1);
  assert.equal(events.at(-1)?.type, 'complete');
  const start = rpc.messages.find(
    (message) => message.method === 'thread/start',
  )!;
  assert.equal(start.params.sandbox, 'read-only');
  assert.equal(start.params.ephemeral, true);
  assert.equal(start.params.dynamicTools.length, 2);
  assert.ok(rpc.messages.some((message) => message.method === 'initialized'));
  assert.ok(rpc.rejected.includes(90));
  assert.equal(rpc.replies[0].result.success, true);
  const stats = JSON.parse(rpc.replies[0].result.contentItems[0].text);
  assert.equal(stats.groups[0].columns[2].average, 20);
  assert.ok(rpc.closed);
});

test('failed, interrupted and empty completed turns are never reported as success', async () => {
  for (const [outcome, sendText] of [
    ['failed', true],
    ['interrupted', true],
    ['completed', false],
  ] as const) {
    const rpc = new FakeRpc();
    rpc.outcome = outcome;
    rpc.sendText = sendText;
    const events = await collect(
      new CodexAdapter({ transport: async () => rpc }).run(request()),
    );
    assert.equal(events.at(-1)?.type, 'error');
    assert.ok(rpc.closed);
  }
});

test('missing authentication fails before any model turn starts', async () => {
  const rpc = new FakeRpc();
  rpc.authenticated = false;
  const adapter = new CodexAdapter({ transport: async () => rpc });
  const events = await collect(adapter.run(request()));
  assert.equal(events.at(-1)?.type, 'error');
  assert.ok(!rpc.messages.some((message) => message.method === 'turn/start'));
  assert.ok(rpc.closed);
});

test('cancellation closes the owning transport and unblocks waiting consumers', async () => {
  const rpc = new FakeRpc();
  rpc.hang = true;
  const controller = new AbortController();
  const events: unknown[] = [];
  for await (const event of new CodexAdapter({
    transport: async () => rpc,
  }).run(request(controller.signal))) {
    events.push(event);
    if (rpc.messages.some((message) => message.method === 'turn/start'))
      controller.abort();
  }
  assert.equal((events.at(-1) as any).code, 'cancelled');
  assert.ok(rpc.closed);
});

test('Codex publishes validated proposals only after a completed turn, and staged tool calls never claim application', async () => {
  const model = await loadWpsHost();
  const workbook = model.clone(model.host.capture('sheet'));
  for (const outcome of ['completed', 'interrupted', 'failed']) {
    const rpc = new FakeRpc();
    rpc.outcome = outcome;
    rpc.toolCalls = [
      { tool: 'officebridge_clean_range', arguments: { rangeId: 'r1' } },
    ];
    const events = await collect(
      new CodexAdapter({ transport: async () => rpc }).run({
        ...request(),
        workbook,
        operationFeedback: '已应用上次修改。',
      }),
    );
    const output = JSON.parse(rpc.replies[0].result.contentItems[0].text);
    assert.equal(output.staged, true);
    assert.equal(output.applied, false);
    assert.equal(
      events.some((event) => event.type === 'proposal'),
      outcome === 'completed',
    );
    if (outcome === 'completed') {
      assert.equal(events.at(-2)?.type, 'proposal');
      const proposal = events.find((event) => event.type === 'proposal')!;
      assert.equal(proposal.plan.edits[0]?.value, 'a');
    }
    const start = rpc.messages.find(
      (message) => message.method === 'thread/start',
    )!;
    assert.equal(start.params.dynamicTools.length, 17);
    assert.equal(start.params.sandbox, 'read-only');
    const turn = rpc.messages.find(
      (message) => message.method === 'turn/start',
    )!;
    assert.ok(turn.params.input[0].text.includes('已应用上次修改。'));
    assert.ok(turn.params.input[0].text.includes('本次范围'));
    assert.equal(model.sheets[0].Cells.Item(2, 1).Value2, ' a ');
  }
});

test('Codex routes the new native actions through scoped tools and publishes only after successful completion', async () => {
  for (const call of [
    {
      tool: 'officebridge_delete_rows',
      arguments: { rangeId: 'r1', row: 1, count: 1 },
    },
    {
      tool: 'officebridge_rename_sheet',
      arguments: { rangeId: 'r1', name: '整理结果' },
    },
    {
      tool: 'officebridge_filter_range',
      arguments: {
        rangeId: 'r1',
        conditions: [{ column: 1, operator: 'gte', value: 15 }],
      },
    },
    {
      tool: 'officebridge_create_chart',
      arguments: {
        rangeId: 'r1',
        chartType: 'line',
        title: 'Amount',
        categoryColumn: 0,
        valueColumns: [1],
      },
    },
    {
      tool: 'officebridge_create_chart',
      arguments: {
        rangeId: 'r1',
        chartType: 'line',
        title: 'Totals',
        categoryColumn: 0,
        valueColumns: [1],
        aggregation: 'sum',
        categoryOrder: 'ascending',
      },
    },
  ]) {
    const model = await loadWpsHost();
    const rpc = new FakeRpc();
    rpc.toolCalls = [call];
    const events = await collect(
      new CodexAdapter({ transport: async () => rpc }).run({
        ...request(),
        workbook: model.clone(model.host.capture('sheet')),
      }),
    );
    const proposal = events.find((event) => event.type === 'proposal');
    assert.ok(proposal?.type === 'proposal', call.tool);
    assert.equal(events.at(-1)?.type, 'complete');
    assert.equal(model.sheets[0].Name, 'Data');
    assert.equal(model.sheets[0].Cells.Item(2, 1).Value2, ' a ');
    model.host.apply(proposal.plan);
    model.host.undo();
    assert.equal(model.sheets[0].Name, 'Data');
    assert.equal(model.sheets[0].Cells.Item(2, 1).Value2, ' a ');
  }
});
test('invalid model edits fail at the tool boundary and leave no proposal', async () => {
  const model = await loadWpsHost();
  const rpc = new FakeRpc();
  rpc.toolCalls = [
    {
      tool: 'officebridge_propose_edits',
      arguments: { edits: [{ rangeId: 'r1', row: 80, column: 0, value: 'x' }] },
    },
  ];
  const events = await collect(
    new CodexAdapter({ transport: async () => rpc }).run({
      ...request(),
      workbook: model.clone(model.host.capture('sheet')),
    }),
  );
  assert.equal(rpc.replies[0].result.success, false);
  assert.ok(!events.some((event) => event.type === 'proposal'));
});

test('one conversation reuses its native thread across turns, refreshes snapshots and never replays chat history', async () => {
  const model = await loadWpsHost();
  const rpc = new FakeRpc();
  rpc.toolCalls = [
    { tool: 'officebridge_read_range', arguments: { rangeId: 'r1' } },
  ];
  let connections = 0;
  const adapter = new CodexAdapter({
    transport: async () => {
      connections++;
      return rpc;
    },
  });
  const conversation = await adapter.openConversation();
  const run = (workbook: any, prompt: string) =>
    collect(
      adapter.run({
        ...request(),
        prompt,
        workbook,
        conversationId: conversation.id,
      }),
    );
  let workbook = model.clone(model.host.capture('selection'));
  assert.equal(
    (await run(workbook, '记住我想统计时长')).at(-1)?.type,
    'complete',
  );
  model.select(2, 1);
  workbook = model.clone(model.host.capture('selection'));
  assert.equal(
    (await run(workbook, '现在处理这一格')).at(-1)?.type,
    'complete',
  );
  assert.equal(connections, 1);
  assert.equal(
    rpc.messages.filter((m) => m.method === 'thread/start').length,
    1,
  );
  const turns = rpc.messages.filter((m) => m.method === 'turn/start');
  assert.equal(turns.length, 2);
  assert.equal(turns[0]!.params.threadId, turns[1]!.params.threadId);
  assert.ok(!turns[1]!.params.input[0].text.includes('记住我想统计时长'));
  assert.ok(turns[1]!.params.input[0].text.includes('A2'));
  assert.ok(!rpc.closed);
  assert.equal(
    JSON.parse(rpc.replies.at(-1).result.contentItems[0].text).address,
    'A2',
  );
  // Late calls outside a running turn cannot access a stale snapshot.
  rpc.handler({
    id: 999,
    method: 'item/tool/call',
    params: {
      threadId: 'own-thread',
      turnId: 'turn-1',
      tool: 'officebridge_read_range',
      arguments: { rangeId: 'r1' },
    },
  });
  assert.ok(rpc.rejected.includes(999));
  await adapter.closeConversation(conversation.id);
  assert.ok(rpc.closed);
  const lost = await run(workbook, '继续');
  assert.equal((lost.at(-1) as any).code, 'conversation_lost');
  assert.equal(connections, 1);
});

test('new conversations isolate transports and cancel only the active turn, with no staged plan leakage', async () => {
  const model = await loadWpsHost();
  const workbook = model.clone(model.host.capture('sheet'));
  const transports: FakeRpc[] = [];
  const adapter = new CodexAdapter({
    transport: async () => {
      const rpc = new FakeRpc();
      transports.push(rpc);
      return rpc;
    },
  });
  const first = await adapter.openConversation();
  const rpc = transports[0]!;
  rpc.hang = true;
  rpc.toolCalls = [];
  const controller = new AbortController();
  const events = [];
  for await (const event of adapter.run({
    ...request(controller.signal),
    conversationId: first.id,
    workbook,
  })) {
    events.push(event);
    if (rpc.turnCount) {
      rpc.handler({
        id: 400,
        method: 'item/tool/call',
        params: {
          threadId: 'own-thread',
          turnId: 'turn-1',
          tool: 'officebridge_clean_range',
          arguments: { rangeId: 'r1' },
        },
      });
      controller.abort();
    }
  }
  assert.ok(!events.some((e) => e.type === 'proposal'));
  assert.ok(!rpc.closed);
  assert.ok(rpc.messages.some((m) => m.method === 'turn/interrupt'));
  rpc.hang = false;
  const resumed = await collect(
    adapter.run({ ...request(), conversationId: first.id, workbook }),
  );
  assert.equal(resumed.at(-1)?.type, 'complete');
  assert.ok(!resumed.some((e) => e.type === 'proposal'));
  assert.equal(
    rpc.messages.filter((m) => m.method === 'thread/start').length,
    1,
  );
  const second = await adapter.openConversation();
  assert.notEqual(first.id, second.id);
  assert.equal(transports.length, 2);
  await adapter.closeConversation(first.id);
  assert.ok(rpc.closed);
  assert.ok(!transports[1]!.closed);
  adapter.dispose();
  assert.ok(transports[1]!.closed);
});

test('a disconnected native conversation fails explicitly and never silently creates a replacement', async () => {
  const rpc = new FakeRpc();
  let connections = 0;
  const adapter = new CodexAdapter({
    transport: async () => {
      connections++;
      return rpc;
    },
  });
  const conversation = await adapter.openConversation();
  rpc.failure(new Error('connection dropped'));
  const events = await collect(
    adapter.run({ ...request(), conversationId: conversation.id }),
  );
  assert.equal((events.at(-1) as any).code, 'conversation_lost');
  assert.equal(connections, 1);
  adapter.dispose();
});

test('concurrent turns cannot share tools, and an unacknowledged interrupt invalidates rather than replaces the thread', async () => {
  const rpc = new FakeRpc();
  const original = rpc.request.bind(rpc);
  rpc.request = async (method, params) =>
    method === 'turn/interrupt' ? {} : original(method, params);
  rpc.hang = true;
  const adapter = new CodexAdapter({ transport: async () => rpc });
  const conversation = await adapter.openConversation();
  const controller = new AbortController();
  const stream = adapter.run({
    ...request(controller.signal),
    conversationId: conversation.id,
  });
  const iterator = stream[Symbol.asyncIterator]();
  await iterator.next(); // connecting status
  await iterator.next(); // turn started
  const concurrent = await collect(
    adapter.run({ ...request(), conversationId: conversation.id }),
  );
  assert.equal((concurrent.at(-1) as any).code, 'conversation_busy');
  controller.abort();
  await collect(stream);
  assert.ok(rpc.closed);
  const lost = await collect(
    adapter.run({ ...request(), conversationId: conversation.id }),
  );
  assert.equal((lost.at(-1) as any).code, 'conversation_lost');
  assert.equal(
    rpc.messages.filter((m) => m.method === 'thread/start').length,
    1,
  );
  adapter.dispose();
});

test('closing a conversation during generation wakes consumers without waiting for the turn timeout', async () => {
  const rpc = new FakeRpc();
  rpc.hang = true;
  const adapter = new CodexAdapter({ transport: async () => rpc });
  const conversation = await adapter.openConversation();
  const stream = adapter.run({ ...request(), conversationId: conversation.id });
  const iterator = stream[Symbol.asyncIterator]();
  await iterator.next();
  await iterator.next();
  await adapter.closeConversation(conversation.id);
  const events = await collect(stream);
  assert.equal((events.at(-1) as any).code, 'conversation_lost');
  assert.ok(rpc.closed);
});
