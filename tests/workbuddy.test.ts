import assert from 'node:assert/strict';
import test from 'node:test';
import {
  WorkBuddyAdapter,
  workBuddyArgs,
} from '../packages/agents/src/workbuddy.js';
import type {
  RpcMessage,
  RpcTransport,
} from '../packages/agents/src/codex-rpc.js';
import { loadWpsHost } from './wps-fixture.js';
import { fullSelection } from './fixtures.js';
import { TableMcp } from '../packages/agents/src/table-mcp.js';
import { TableToolSession } from '../packages/agents/src/table-tools.js';

test('WorkBuddy private MCP awaits on-demand reads and writes in the existing native conversation', async () => {
  const model = await loadWpsHost();
  model.select(1, 1);
  const workbook = model.clone(model.host.beginRead('context'));
  const rpc = new FakeAcp();
  rpc.actionCalls = [
    { name: 'officebridge_resolve_range', arguments: { target: 'sheet' } },
    {
      name: 'officebridge_propose_edits',
      arguments: { edits: [{ rangeId: 'q1', row: 1, column: 1, value: 50 }] },
    },
    {
      name: 'officebridge_continue_after_apply',
      arguments: { nextStep: '创建图表' },
    },
  ];
  const adapter = new WorkBuddyAdapter({ transport: async () => rpc });
  let calls = 0;
  try {
    const { id } = await adapter.openConversation();
    const events = await collect(
      adapter.run({
        requestId: 'lazy',
        prompt: '修改金额',
        workbook,
        conversationId: id,
        signal: new AbortController().signal,
        hostTool: async (name, args) => {
          calls++;
          return model.clone(
            model.host.executeQuery(workbook.snapshotId, name, args),
          );
        },
      }),
    );
    const proposal = events.find((e) => e.type === 'proposal');
    assert.ok(proposal?.type === 'proposal');
    assert.deepEqual(proposal.continuation, { nextStep: '创建图表' });
    assert.ok(proposal.snapshot);
    assert.equal(calls, 3);
    model.host.apply(proposal.plan);
    assert.equal(model.sheets[0].Cells.Item(2, 2).Value2, 50);
    model.host.undo();
  } finally {
    adapter.dispose();
  }
});

async function collect<T>(stream: AsyncIterable<T>) {
  const events: T[] = [];
  for await (const e of stream) events.push(e);
  return events;
}

test('WorkBuddy corrects chart ordering in the same MCP task, with a shared bounded parameter correction budget', async () => {
  for (const invalidCalls of [1, 3]) {
    const m = await loadWpsHost({
      Summary: [
        ['Month', 'Bill'],
        ['2026-06', 2.6],
        ['2026-07', 2.6],
      ],
    });
    m.select(1, 1, 3, 2);
    const rpc = new FakeAcp();
    const original = rpc.request.bind(rpc);
    rpc.request = async (method, params) => {
      if (method !== 'session/prompt') return original(method, params);
      rpc.messages.push({ method, params });
      const cancelled = new Promise((resolve) => {
        rpc.finish = resolve;
      });
      const args = {
        rangeId: 'r1',
        chartType: 'line',
        title: 'Bill',
        categoryColumn: 0,
        valueColumns: [1],
        categoryOrder: 'ascending',
      };
      for (let i = 0; i < invalidCalls; i++) {
        const mixedRead = invalidCalls === 3 && i === 1;
        const mixedFormula = invalidCalls === 3 && i === 2;
        const result = await rpc.tool(
          mixedRead
            ? 'officebridge_read_range'
            : mixedFormula
              ? 'officebridge_propose_edits'
              : 'officebridge_create_chart',
          mixedRead
            ? { rangeId: 'r1', rowCount: 4 }
            : mixedFormula
              ? {
                  edits: [
                    { rangeId: 'r1', row: 0, column: 0, formula: 'SUM(B2:B3)' },
                  ],
                }
              : args,
        );
        assert.equal(result.isError, true);
        assert.match(
          result.content[0].text,
          mixedRead
            ? /rowCount/
            : mixedFormula
              ? /必须以等号/
              : /aggregation="sum"/,
        );
      }
      if (invalidCalls === 3) return cancelled;
      const corrected = await rpc.tool('officebridge_create_chart', {
        ...args,
        aggregation: 'sum',
      });
      assert.equal(corrected.isError, undefined);
      return { stopReason: 'end_turn' };
    };
    const adapter = new WorkBuddyAdapter({ transport: async () => rpc });
    const hostCalls: string[] = [];
    try {
      const { id } = await adapter.openConversation();
      const workbook = m.clone(m.host.beginRead('context'));
      const events = await collect(
        adapter.run({
          requestId: 'chart-correction',
          conversationId: id,
          prompt: '画折线图',
          workbook,
          signal: new AbortController().signal,
          hostTool: async (name, args) => {
            hostCalls.push(name);
            return m.clone(
              await m.host.executeQueryAsync(workbook.snapshotId, name, args),
            );
          },
        }),
      );
      if (invalidCalls === 1) {
        assert.equal(events.at(-1)?.type, 'complete');
        assert.ok(events.some((e) => e.type === 'proposal'));
        assert.deepEqual(hostCalls, ['officebridge_create_chart']);
      } else {
        assert.equal((events.at(-1) as any).code, 'tool_parameters_failed');
        assert.ok(!events.some((e) => e.type === 'proposal'));
        assert.deepEqual(hostCalls, []);
        assert.ok(rpc.messages.some((m) => m.method === 'session/cancel'));
      }
      assert.equal(
        rpc.messages.filter((m) => m.method === 'session/new').length,
        1,
      );
      assert.equal(m.sheets[0].ChartObjects().Count, 0);
    } finally {
      adapter.dispose();
    }
  }
});

test('WorkBuddy receives a precise formula error then stages corrected E12 in the same native turn without a host failure', async () => {
  const m = await loadWpsHost({
    Summary: [
      ['Month', '', '', '', 'Amount'],
      ...Array.from({ length: 10 }, (_, i) => [String(i), '', '', '', i + 1]),
    ],
  });
  m.select(12, 5);
  const workbook = m.clone(m.host.beginRead('context'));
  const edit = {
    rangeId: 'r1',
    row: 0,
    column: 0,
    formula: 'SUM(E2:E11)',
    format: { bold: true },
  };
  const rpc = new FakeAcp();
  rpc.actionCalls = [
    { name: 'officebridge_propose_edits', arguments: { edits: [edit] } },
    {
      name: 'officebridge_propose_edits',
      arguments: { edits: [{ ...edit, formula: '=SUM(E2:E11)' }] },
    },
  ];
  const adapter = new WorkBuddyAdapter({ transport: async () => rpc });
  const calls: string[] = [];
  try {
    const { id } = await adapter.openConversation();
    const events = await collect(
      adapter.run({
        requestId: 'formula-correction',
        conversationId: id,
        prompt: '在这格合计金额',
        workbook,
        signal: new AbortController().signal,
        hostTool: async (name, args) => {
          calls.push(name);
          return m.clone(
            await m.host.executeQueryAsync(workbook.snapshotId, name, args),
          );
        },
      }),
    );
    assert.equal(rpc.results[1].isError, true);
    assert.match(
      rpc.results[1].content[0].text,
      /edits\[0\]\.formula.*必须以等号/,
    );
    assert.equal(rpc.results[2].isError, undefined);
    assert.deepEqual(calls, [
      'officebridge_read_range',
      'officebridge_propose_edits',
    ]);
    assert.equal(events.at(-1)?.type, 'complete');
    const proposal = events.find((e) => e.type === 'proposal');
    assert.ok(proposal?.type === 'proposal');
    assert.equal(m.sheets[0].Cells.Item(12, 5).Value2, null);
    m.host.apply(proposal.plan);
    assert.equal(m.sheets[0].Cells.Item(12, 5).Value2, 55);
    assert.equal(
      rpc.messages.filter((m) => m.method === 'session/new').length,
      1,
    );
    assert.equal(
      rpc.messages.filter((m) => m.method === 'session/prompt').length,
      1,
    );
  } finally {
    adapter.dispose();
  }
});

test('WorkBuddy can answer from a full attached selection without requesting data again or showing a false warning', async () => {
  const m = await loadWpsHost({
    Data: Array.from({ length: 40 }, (_, i) => ['record ' + i]),
  });
  m.select(1, 1, 40, 1);
  const workbook = m.clone(m.host.beginRead('context'));
  const rpc = new FakeAcp(),
    baseRequest = rpc.request.bind(rpc);
  rpc.request = async (method, params) => {
    if (method !== 'session/prompt') return baseRequest(method, params);
    assert.match(params.prompt[0].text, /record 39/);
    rpc.handler({
      method: 'session/update',
      params: {
        sessionId: params.sessionId,
        update: {
          sessionUpdate: 'agent_message_chunk',
          content: { type: 'text', text: '已收到完整数据。' },
        },
      },
    });
    return { stopReason: 'end_turn' };
  };
  const adapter = new WorkBuddyAdapter({ transport: async () => rpc });
  try {
    const { id } = await adapter.openConversation();
    const events = await collect(
      adapter.run({
        requestId: 'provided',
        prompt: '分析这些',
        workbook,
        conversationId: id,
        signal: new AbortController().signal,
        hostTool: async () => {
          throw new Error('No extra read needed');
        },
      }),
    );
    const end = events.at(-1);
    assert.equal(end?.type, 'complete');
    assert.equal((end as any).notice, undefined);
    assert.equal(rpc.results.length, 0);
  } finally {
    adapter.dispose();
  }
});
class FakeAcp implements RpcTransport {
  messages: { method: string; params: any }[] = [];
  handler: (message: RpcMessage) => void = () => {};
  failure: (error: Error) => void = () => {};
  rejected: (string | number)[] = [];
  closed = false;
  mcp: any;
  mode = 'clean';
  outcome: any;
  actionCall: { name: string; arguments: unknown } | undefined;
  actionCalls: { name: string; arguments: unknown }[] | undefined;
  reads: any[] = [];
  results: any[] = [];
  finish: ((result: any) => void) | undefined;
  async mcpRequest(method: string, params = {}) {
    return (
      (await (
        await fetch(this.mcp.url, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            Authorization: this.mcp.headers[0].value,
          },
          body: JSON.stringify({ jsonrpc: '2.0', id: 1, method, params }),
        })
      ).json()) as any
    ).result;
  }
  async tool(name: string, args: unknown = {}) {
    const response = await fetch(this.mcp.url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: this.mcp.headers[0].value,
      },
      body: JSON.stringify({
        jsonrpc: '2.0',
        id: 1,
        method: 'tools/call',
        params: { name, arguments: args },
      }),
    });
    const result = ((await response.json()) as any).result;
    this.results.push(result);
    return result;
  }
  async request(method: string, params: any): Promise<any> {
    this.messages.push({ method, params });
    if (method === 'initialize')
      return {
        protocolVersion: 1,
        agentCapabilities: { mcpCapabilities: { http: true } },
      };
    if (method === 'session/new') {
      this.mcp = params.mcpServers[0];
      await this.mcpRequest('initialize', {
        protocolVersion: '2025-03-26',
        capabilities: {},
        clientInfo: { name: 'test', version: '1' },
      });
      await this.mcpRequest('tools/list');
      return { sessionId: 'native-session' };
    }
    if (method === 'session/prompt') {
      if (this.mode === 'text') {
        this.handler({
          method: 'session/update',
          params: {
            sessionId: params.sessionId,
            update: {
              sessionUpdate: 'agent_message_chunk',
              content: {
                type: 'text',
                text: '<tool_call>readRange。方案已暂存，请点击应用。',
              },
            },
          },
        });
        return { stopReason: 'end_turn' };
      }
      const range = await this.tool('officebridge_read_range', {
        rangeId: 'r1',
      });
      this.reads.push(JSON.parse(range.content[0].text));
      if (this.actionCalls)
        for (const call of this.actionCalls)
          await this.tool(call.name, call.arguments);
      else if (this.actionCall)
        await this.tool(this.actionCall.name, this.actionCall.arguments);
      else if (this.mode !== 'empty')
        await this.tool('officebridge_clean_range', {
          rangeId: 'r1',
          headerRow: this.reads.at(-1).rowCount > 1,
        });
      this.handler({
        method: 'session/update',
        params: {
          sessionId: 'other-session',
          update: {
            sessionUpdate: 'agent_message_chunk',
            content: { type: 'text', text: 'wrong session' },
          },
        },
      });
      this.handler({ id: 99, method: 'fs/write_text_file', params: {} });
      if (this.mode === 'hang')
        return new Promise((resolve) => {
          this.finish = resolve;
        });
      if (this.mode === 'fail') throw new Error('private provider failure');
      if (this.mode !== 'empty')
        this.handler({
          method: 'session/update',
          params: {
            sessionId: params.sessionId,
            update: {
              sessionUpdate: 'agent_message_chunk',
              content: { type: 'text', text: '方案已准备好，等待应用。' },
            },
          },
        });
      return {
        ...(this.outcome ?? {}),
        stopReason:
          this.outcome?.stopReason ??
          (this.mode === 'limited' ? 'max_tokens' : 'end_turn'),
      };
    }
    return {};
  }
  notify(method: string, params: unknown) {
    this.messages.push({ method, params });
    if (method === 'session/cancel') this.finish?.({ stopReason: 'cancelled' });
  }
  onMessage(handler: (message: RpcMessage) => void) {
    this.handler = handler;
  }
  onFailure(handler: (error: Error) => void) {
    this.failure = handler;
  }
  reply() {}
  reject(id: string | number) {
    this.rejected.push(id);
  }
  close() {
    this.closed = true;
    this.finish?.({ stopReason: 'cancelled' });
  }
}

test('native thinking and argument generation expose only throttled activity, with no private content or wrong-session renewal', async (t) => {
  const rpc = new FakeAcp();
  const original = rpc.request.bind(rpc);
  let clock = 1000;
  t.mock.method(Date, 'now', () => clock);
  rpc.request = async (method, p) => {
    if (method !== 'session/prompt') return original(method, p);
    const push = (sessionId: string, update: unknown) =>
      rpc.handler({ method: 'session/update', params: { sessionId, update } });
    push('other-session', {
      sessionUpdate: 'agent_thought_chunk',
      content: { type: 'text', text: 'PRIVATE_WRONG_SESSION' },
    });
    push('native-session', {
      sessionUpdate: 'agent_thought_chunk',
      content: { type: 'text', text: 'PRIVATE_REASONING' },
    });
    push('native-session', {
      sessionUpdate: 'agent_thought_chunk',
      content: { type: 'text', text: 'PRIVATE_DUPLICATE' },
    });
    clock += 600;
    push('native-session', {
      sessionUpdate: 'tool_call_update',
      toolCallId: 'x',
      status: 'in_progress',
      rawInput: { secret: 'PRIVATE_ARGS' },
      content: [
        { type: 'content', content: { type: 'text', text: 'PRIVATE_ARGS' } },
      ],
    });
    clock += 600;
    push('native-session', {
      sessionUpdate: 'tool_call_update',
      toolCallId: 'x',
      status: 'in_progress',
      rawInput: { secret: 'PRIVATE_ARGS' },
      content: [
        { type: 'content', content: { type: 'text', text: 'PRIVATE_ARGS' } },
      ],
    });
    clock += 600;
    push('native-session', {
      sessionUpdate: 'tool_call',
      toolCallId: 'x',
      status: 'in_progress',
    });
    push('native-session', {
      sessionUpdate: 'agent_message_chunk',
      content: { type: 'text', text: '完成' },
    });
    return { stopReason: 'end_turn' };
  };
  const adapter = new WorkBuddyAdapter({ transport: async () => rpc });
  try {
    const { id } = await adapter.openConversation();
    const m = await loadWpsHost();
    const events = await collect(
      adapter.run({
        requestId: 'progress',
        conversationId: id,
        prompt: 'test',
        workbook: m.clone(m.host.capture('sheet')),
        signal: new AbortController().signal,
      }),
    );
    assert.deepEqual(
      events.filter((e) => e.type === 'activity').map((e) => (e as any).phase),
      ['thinking', 'tool'],
    );
    assert.doesNotMatch(JSON.stringify(events), /PRIVATE|secret/);
    assert.equal(events.at(-1)?.type, 'complete');
  } finally {
    adapter.dispose();
  }
});

test('WorkBuddy exposes and executes new scoped native actions through the private MCP without applying before confirmation', async () => {
  for (const call of [
    {
      name: 'officebridge_delete_rows',
      arguments: { rangeId: 'r1', row: 1, count: 1 },
    },
    {
      name: 'officebridge_rename_sheet',
      arguments: { rangeId: 'r1', name: '整理结果' },
    },
    {
      name: 'officebridge_filter_range',
      arguments: {
        rangeId: 'r1',
        conditions: [{ column: 1, operator: 'gte', value: 15 }],
      },
    },
    {
      name: 'officebridge_create_chart',
      arguments: {
        rangeId: 'r1',
        chartType: 'line',
        title: 'Amount',
        categoryColumn: 0,
        valueColumns: [1],
      },
    },
    {
      name: 'officebridge_create_chart',
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
    const model = await loadWpsHost(),
      rpc = new FakeAcp();
    rpc.actionCall = call;
    const adapter = new WorkBuddyAdapter({ transport: async () => rpc });
    try {
      const { id } = await adapter.openConversation();
      const events = await collect(
        adapter.run({
          requestId: 'native',
          conversationId: id,
          prompt: 'native operation',
          context: fullSelection,
          workbook: model.clone(model.host.capture('sheet')),
          signal: new AbortController().signal,
        }),
      );
      const proposal = events.find((event) => event.type === 'proposal');
      assert.ok(proposal?.type === 'proposal', call.name);
      assert.equal(model.sheets[0].Name, 'Data');
      assert.equal(model.sheets[0].Cells.Item(2, 1).Value2, ' a ');
      model.host.apply(proposal.plan);
      model.host.undo();
      assert.equal(model.sheets[0].Name, 'Data');
    } finally {
      adapter.dispose();
    }
  }
});
test('WorkBuddy reuses a native ACP session and scoped MCP tools across refreshed snapshots without replaying history', async () => {
  const model = await loadWpsHost();
  const transports: FakeAcp[] = [];
  const adapter = new WorkBuddyAdapter({
    transport: async () => {
      const rpc = new FakeAcp();
      transports.push(rpc);
      return rpc;
    },
  });
  const { id } = await adapter.openConversation();
  const rpc = transports[0]!;
  const run = () =>
    collect(
      adapter.run({
        requestId: 'test',
        conversationId: id,
        prompt: '去空格',
        context: fullSelection,
        workbook: model.clone(model.host.capture('selection')),
        signal: new AbortController().signal,
      }),
    );
  const first = await run();
  assert.equal(first.at(-1)?.type, 'complete');
  assert.equal(first.at(-2)?.type, 'proposal');
  assert.ok(
    !first.some((e) => e.type === 'text' && e.text.includes('wrong session')),
  );
  assert.ok(rpc.rejected.includes(99));
  assert.equal(model.sheets[0].Cells.Item(2, 1).Value2, ' a ');
  assert.equal(JSON.parse(rpc.results[1].content[0].text).applied, false);
  assert.equal(
    (await rpc.tool('officebridge_read_range', { rangeId: 'r1' })).isError,
    true,
  );
  model.select(2, 1);
  const second = await run();
  assert.equal(second.at(-1)?.type, 'complete');
  assert.equal(rpc.reads.at(-1).address, 'A2');
  assert.equal(
    rpc.messages.filter((m) => m.method === 'session/new').length,
    1,
  );
  const prompts = rpc.messages.filter((m) => m.method === 'session/prompt');
  assert.equal(prompts.length, 2);
  assert.equal(prompts[0]!.params.sessionId, prompts[1]!.params.sessionId);
  assert.ok(!prompts[1]!.params.prompt[0].text.includes('方案已准备好'));
  assert.ok(!rpc.closed);
  const replacement = await adapter.openConversation(id);
  assert.notEqual(replacement.id, id);
  await adapter.closeConversation(id);
  assert.ok(rpc.closed);
  assert.ok(!transports[1]!.closed);
  adapter.dispose();
  assert.ok(transports[1]!.closed);
});

test('WorkBuddy never publishes proposals after failure, empty completion, token limits or interruption', async () => {
  const model = await loadWpsHost();
  for (const mode of ['fail', 'empty', 'limited', 'hang']) {
    const rpc = new FakeAcp();
    rpc.mode = mode;
    const adapter = new WorkBuddyAdapter({ transport: async () => rpc });
    const { id } = await adapter.openConversation();
    const controller = new AbortController();
    const events = [];
    for await (const event of adapter.run({
      requestId: 'x',
      conversationId: id,
      prompt: 'test',
      context: fullSelection,
      workbook: model.clone(model.host.capture('sheet')),
      signal: controller.signal,
    })) {
      events.push(event);
      if (mode === 'hang' && event.type === 'status')
        setTimeout(() => controller.abort(), 20);
    }
    assert.equal(events.at(-1)?.type, 'error');
    assert.ok(!events.some((e) => e.type === 'proposal'));
    assert.ok(!JSON.stringify(events).includes('private provider failure'));
    if (mode === 'hang') {
      assert.ok(!rpc.closed);
      rpc.mode = 'clean';
      const resumed = await collect(
        adapter.run({
          requestId: 'next',
          conversationId: id,
          prompt: 'continue',
          context: fullSelection,
          workbook: model.clone(model.host.capture('sheet')),
          signal: new AbortController().signal,
        }),
      );
      assert.equal(resumed.at(-1)?.type, 'complete');
      assert.equal(
        rpc.messages.filter((m) => m.method === 'session/new').length,
        1,
      );
    }
    adapter.dispose();
  }
});

test('WorkBuddy connection loss is explicit and a new native session is never silently created', async () => {
  const rpc = new FakeAcp();
  const adapter = new WorkBuddyAdapter({ transport: async () => rpc });
  const { id } = await adapter.openConversation();
  rpc.failure(new Error('lost'));
  const events = await collect(
    adapter.run({
      requestId: 'x',
      conversationId: id,
      prompt: 'continue',
      context: fullSelection,
      signal: new AbortController().signal,
    }),
  );
  assert.equal((events.at(-1) as any).code, 'conversation_lost');
  assert.equal(
    rpc.messages.filter((m) => m.method === 'session/new').length,
    1,
  );
  adapter.dispose();
});

test('a quota refusal discards staged output without retry and the same native conversation can continue after recovery', async () => {
  const m = await loadWpsHost(),
    rpc = new FakeAcp(),
    adapter = new WorkBuddyAdapter({ transport: async () => rpc });
  try {
    const { id } = await adapter.openConversation();
    rpc.outcome = {
      stopReason: 'refusal',
      _meta: {
        'codebuddy.ai/errorMessage': JSON.stringify({
          data: { statusCode: 429, category: 'quota', code: 14003 },
          message: 'private-provider-response',
        }),
      },
    };
    const run = () =>
      collect(
        adapter.run({
          requestId: 'quota',
          conversationId: id,
          prompt: 'clean',
          workbook: m.clone(m.host.capture('sheet')),
          signal: new AbortController().signal,
        }),
      );
    const events = await run();
    assert.equal((events.at(-1) as any).code, 'agent_quota');
    assert.ok(!events.some((e) => e.type === 'proposal'));
    assert.doesNotMatch(JSON.stringify(events), /private-provider-response/);
    assert.equal(rpc.closed, false);
    assert.equal(
      rpc.messages.filter((m) => m.method === 'session/prompt').length,
      1,
    );
    rpc.outcome = undefined;
    const recovered = await run();
    assert.equal(recovered.at(-1)?.type, 'complete');
    assert.ok(recovered.some((e) => e.type === 'proposal'));
    assert.equal(
      rpc.messages.filter((m) => m.method === 'session/new').length,
      1,
    );
  } finally {
    adapter.dispose();
  }
});

test('WorkBuddy corrects read parameters through the same MCP/session, and repeated invalid reads stop without host access or proposals', async () => {
  for (const invalidReads of [1, 3]) {
    const m = await loadWpsHost({
      Data: [
        ['Month', 'Bill'],
        ['June', 10],
        ['July', 20],
      ],
    });
    m.select(2, 2);
    const rpc = new FakeAcp();
    const original = rpc.request.bind(rpc);
    let bad = invalidReads;
    rpc.request = async (method, params) => {
      if (method !== 'session/prompt') return original(method, params);
      rpc.messages.push({ method, params });
      const cancelled = new Promise((resolve) => {
        rpc.finish = resolve;
      });
      const resolved = await rpc.tool('officebridge_resolve_range', {
        target: 'sheet',
      });
      const q = JSON.parse(resolved.content[0].text);
      const head = await rpc.tool('officebridge_read_range', {
        rangeId: q.rangeId,
        rowCount: 1,
      });
      assert.equal(JSON.parse(head.content[0].text).sourceRange.rowCount, 3);
      for (let i = 0; i < bad; i++) {
        const result = await rpc.tool('officebridge_read_range', {
          rangeId: q.rangeId,
          rowOffset: 1,
          rowCount: 3,
        });
        assert.equal(result.isError, true);
        assert.match(result.content[0].text, /最多读取 2 行/);
      }
      if (bad >= 3) return cancelled;
      const data = await rpc.tool('officebridge_read_range', {
        rangeId: q.rangeId,
        rowOffset: 1,
      });
      assert.equal(data.isError, undefined);
      assert.equal(JSON.parse(data.content[0].text).rowCount, 2);
      const chart = await rpc.tool('officebridge_create_chart', {
        rangeId: q.rangeId,
        chartType: 'line',
        title: 'Bill',
        categoryColumn: 0,
        valueColumns: [1],
      });
      assert.equal(chart.isError, undefined);
      rpc.handler({
        method: 'session/update',
        params: {
          sessionId: 'native-session',
          update: {
            sessionUpdate: 'agent_message_chunk',
            content: { type: 'text', text: '图表方案已准备。' },
          },
        },
      });
      return { stopReason: 'end_turn' };
    };
    const adapter = new WorkBuddyAdapter({ transport: async () => rpc });
    try {
      const { id } = await adapter.openConversation();
      const calls: string[] = [];
      const run = () => {
        const d = m.clone(m.host.beginRead('context'));
        return collect(
          adapter.run({
            requestId: 'read-correction',
            conversationId: id,
            prompt: '画折线图',
            workbook: d,
            hostTool: async (name, args) => {
              calls.push(name);
              return m.clone(m.host.executeQuery(d.snapshotId, name, args));
            },
            signal: new AbortController().signal,
          }),
        );
      };
      const events = await run();
      if (invalidReads === 1) {
        assert.equal(events.at(-1)?.type, 'complete');
        assert.ok(events.some((e) => e.type === 'proposal'));
        assert.equal(
          calls.filter((x) => x === 'officebridge_read_range').length,
          2,
        );
      } else {
        assert.equal((events.at(-1) as any).code, 'read_parameters_failed');
        assert.ok(!events.some((e) => e.type === 'proposal'));
        assert.deepEqual(calls, [
          'officebridge_resolve_range',
          'officebridge_read_range',
        ]);
        assert.ok(rpc.messages.some((m) => m.method === 'session/cancel'));
        bad = 0;
        const recovered = await run();
        assert.equal(recovered.at(-1)?.type, 'complete');
        assert.ok(recovered.some((e) => e.type === 'proposal'));
      }
      assert.equal(
        rpc.messages.filter((m) => m.method === 'session/new').length,
        1,
      );
      assert.equal(m.sheets[0].ChartObjects().Count, 0);
    } finally {
      adapter.dispose();
    }
  }
});

test('private MCP denies unauthenticated, idle, aborted, unknown and out-of-range calls without changing the workbook', async () => {
  const model = await loadWpsHost();
  const gateway = new TableMcp();
  const config = await gateway.start();
  try {
    const controller = new AbortController();
    const task = new TableToolSession(
      model.clone(model.host.capture('selection')),
    );
    const call = async (
      method: string,
      params: unknown = {},
      auth = config.headers[0]!.value,
    ) => {
      const response = await fetch(config.url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', Authorization: auth },
        body: JSON.stringify({ jsonrpc: '2.0', id: 1, method, params }),
      });
      return {
        status: response.status,
        body: response.status === 200 ? ((await response.json()) as any) : null,
      };
    };
    assert.equal((await call('tools/list', {}, 'wrong')).status, 403);
    assert.equal((await call('tools/list')).body.result.tools.length, 17);
    assert.equal(
      (await call('tools/call', { name: 'officebridge_list_ranges' })).body
        .result.isError,
      true,
    );
    gateway.bind(task, controller.signal);
    assert.equal(
      (
        await call('tools/call', {
          name: 'officebridge_propose_edits',
          arguments: {
            edits: [{ rangeId: 'r1', row: 999, column: 0, value: 'x' }],
          },
        })
      ).body.result.isError,
      true,
    );
    assert.equal(
      (await call('tools/call', { name: 'fs_write', arguments: {} })).body
        .result.isError,
      true,
    );
    controller.abort();
    assert.equal(
      (
        await call('tools/call', {
          name: 'officebridge_read_range',
          arguments: { rangeId: 'r1' },
        })
      ).body.result.isError,
      true,
    );
    assert.equal(task.proposal(), null);
    assert.equal(model.sheets[0].Cells.Item(2, 1).Value2, ' a ');
  } finally {
    gateway.close();
  }
});

test('WorkBuddy launch policy disables built-ins, unrelated MCP, hooks and transcript persistence', () => {
  const args = workBuddyArgs('cli.js');
  const tools = args[args.indexOf('--tools') + 1]!.split(',');
  assert.equal(tools.length, 17);
  assert.ok(
    tools.every((tool) =>
      /^NoDefer\(mcp__officebridge__officebridge_[a-z_]+\)$/.test(tool),
    ),
  );
  assert.equal(args[args.indexOf('--agent') + 1], 'cli');
  for (const flag of [
    '--acp',
    '--strict-mcp-config',
    '--no-session-persistence',
  ])
    assert.ok(args.includes(flag));
  assert.equal(args[args.indexOf('--permission-mode') + 1], 'dontAsk');
  assert.equal(
    args[args.indexOf('--settings') + 1],
    '{"disableAllHooks":true}',
  );
  assert.ok(!args.includes('--dangerously-skip-permissions'));
});

test('text that pretends to call tools cannot create a proposal and gets an explicit non-execution notice on every turn', async () => {
  const model = await loadWpsHost();
  const rpc = new FakeAcp();
  const adapter = new WorkBuddyAdapter({ transport: async () => rpc });
  try {
    const { id } = await adapter.openConversation();
    const run = () =>
      collect(
        adapter.run({
          requestId: 'x',
          conversationId: id,
          prompt: '新增测试工作表',
          workbook: model.clone(model.host.capture('selection')),
          context: fullSelection,
          signal: new AbortController().signal,
        }),
      );
    rpc.mode = 'clean';
    const real = await run();
    assert.equal(real.at(-2)?.type, 'proposal');
    for (let i = 0; i < 2; i++) {
      rpc.mode = 'text';
      const textOnly = await run();
      assert.ok(!textOnly.some((e) => e.type === 'proposal'));
      assert.match((textOnly.at(-1) as any).notice, /未实际调用/);
    }
  } finally {
    adapter.dispose();
  }
});

test('MCP readiness requires both initialization and listing, with no model request needed', async () => {
  const gateway = new TableMcp();
  const config = await gateway.start();
  try {
    await assert.rejects(gateway.waitUntilReady(30), /未完成表格工具连接/);
    const request = (method: string) =>
      fetch(config.url, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          Authorization: config.headers[0]!.value,
        },
        body: JSON.stringify({ jsonrpc: '2.0', id: 1, method }),
      });
    await (await request('initialize')).text();
    await assert.rejects(gateway.waitUntilReady(30), /未完成表格工具连接/);
    await (await request('tools/list')).text();
    await gateway.waitUntilReady(30);
  } finally {
    gateway.close();
  }
});
