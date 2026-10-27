import assert from 'node:assert/strict';
import test from 'node:test';
import { runInContext, createContext } from 'node:vm';
import { webcrypto } from 'node:crypto';
import {
  DoubaoAdapter,
  parseDoubaoFrame,
} from '../packages/agents/src/doubao.js';
import {
  connectDoubao,
  type DoubaoDriver,
  type DoubaoRpc,
} from '../packages/agents/src/doubao-driver.js';
import {
  doubaoEndpoint,
  isDoubaoChat,
  findDoubaoTarget,
} from '../packages/agents/src/doubao-cdp.js';
import { fullSelection } from './fixtures.js';
import { loadWpsHost } from './wps-fixture.js';

test('Doubao receives complete selection data on its first native message and can finish without redundant tools or warnings', async () => {
  const m = await loadWpsHost({
    Data: Array.from({ length: 40 }, (_, i) => ['record ' + i]),
  });
  m.select(1, 1, 40, 1);
  const workbook = m.clone(m.host.beginRead('context'));
  let turns = 0;
  const adapter = new DoubaoAdapter({
    connect: async () => ({
      conversationId: undefined,
      async send(text) {
        turns++;
        const frame = JSON.parse(text);
        assert.equal(frame.ranges.inline[0].values.length, 40);
        assert.equal(frame.ranges.inline[0].values[39][0], 'record 39');
        return {
          conversationId: '7333333333333',
          runId: '8333333333333',
          text: JSON.stringify({
            protocol: 'officebridge/1',
            requestId: frame.requestId,
            type: 'final',
            text: '已收到完整数据。',
          }),
        };
      },
      async cancel() {
        return true;
      },
      close() {},
    }),
  });
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
    assert.equal(events.at(-1)?.type, 'complete');
    assert.equal((events.at(-1) as any).notice, undefined);
    assert.equal(turns, 1);
  } finally {
    adapter.dispose();
  }
});

test('Doubao structured requests await on-demand host results without sending host write baselines to the model', async () => {
  const model = await loadWpsHost();
  model.select(1, 1);
  const workbook = model.clone(model.host.beginRead('context'));
  let count = 0,
    calls = 0;
  const adapter = new DoubaoAdapter({
    connect: async () => ({
      conversationId: undefined,
      async send(text) {
        const message = JSON.parse(text);
        count++;
        if (count === 1) {
          assert.equal(message.ranges.readMode, 'on-demand');
          assert.equal(message.ranges.ranges[0].values, undefined);
        } else if (count === 2) {
          assert.equal(message.results[0].result.rangeId, 'q1');
        } else {
          assert.equal(message.results[0].result.staged, true);
          assert.equal(message.results[0].snapshot, undefined);
        }
        return {
          conversationId: '7333333333333',
          runId: String(8333333333330 + count),
          text: JSON.stringify({
            protocol: 'officebridge/1',
            requestId: message.requestId,
            ...(count === 1
              ? {
                  type: 'tools',
                  calls: [
                    {
                      name: 'officebridge_resolve_range',
                      arguments: { target: 'sheet' },
                    },
                  ],
                }
              : count === 2
                ? {
                    type: 'tools',
                    calls: [
                      {
                        name: 'officebridge_propose_edits',
                        arguments: {
                          edits: [
                            { rangeId: 'q1', row: 1, column: 1, value: 50 },
                          ],
                        },
                      },
                      {
                        name: 'officebridge_continue_after_apply',
                        arguments: { nextStep: '创建图表' },
                      },
                    ],
                  }
                : { type: 'final', text: '方案已准备。' }),
          }),
        };
      },
      async cancel() {
        return true;
      },
      close() {},
    }),
  });
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
    assert.equal(calls, 2);
    model.host.apply(proposal.plan);
    assert.equal(model.sheets[0].Cells.Item(2, 2).Value2, 50);
    model.host.undo();
  } finally {
    adapter.dispose();
  }
});
import type { AgentEvent } from '../packages/agents/src/index.js';
import { reduceStreamEvent } from '../packages/agents/vendor/doubao-cli/protocol.mjs';

test('Doubao end-event brief cannot overwrite a complete streamed tool request', () => {
  const answer = JSON.stringify({
    protocol: 'officebridge/1',
    requestId: 'stream-test',
    type: 'tools',
    calls: [{ name: 'officebridge_read_range', arguments: { rangeId: 'r1' } }],
  });
  const state = { answer: '', completed: false };
  reduceStreamEvent(
    state,
    'STREAM_CHUNK',
    {
      patch_op: [
        {
          patch_value: {
            content_block: [
              { block_type: 10000, content: { text_block: { text: answer } } },
            ],
          },
        },
      ],
    },
    { waitForReply: true },
  );
  reduceStreamEvent(
    state,
    'SSE_REPLY_END',
    { end_type: 1, msg_finish_attr: { brief: answer.slice(0, 80) } },
    { waitForReply: true },
  );
  assert.equal(state.answer, answer);
  assert.equal(parseDoubaoFrame(state.answer, 'stream-test').type, 'tools');
  reduceStreamEvent(
    state,
    'SSE_REPLY_END',
    { end_type: 3 },
    { waitForReply: true },
  );
  assert.equal(state.completed, true);
  const fallback = { answer: '', completed: false };
  reduceStreamEvent(
    fallback,
    'SSE_REPLY_END',
    { end_type: 1, msg_finish_attr: { brief: 'fallback reply' } },
    { waitForReply: true },
  );
  assert.equal(fallback.answer, 'fallback reply');
});

async function collect(stream: AsyncIterable<AgentEvent>) {
  const events: AgentEvent[] = [];
  for await (const event of stream) events.push(event);
  return events;
}

test('Doubao extracts one complete frame with brief narration, while rejecting stale IDs, unknown tools and ambiguous or quoted payloads', () => {
  const frame = JSON.stringify({
    protocol: 'officebridge/1',
    requestId: 'fenced',
    type: 'tools',
    calls: [
      {
        name: 'officebridge_propose_result_sheet',
        arguments: { name: 'Test', values: [['']] },
      },
    ],
  });
  const wrapped = '```json\n' + frame + '\n```';
  assert.equal(
    parseDoubaoFrame(wrapped, 'fenced').calls[0].name,
    'officebridge_propose_result_sheet',
  );
  assert.throws(() => parseDoubaoFrame(wrapped, 'stale'));
  assert.throws(() =>
    parseDoubaoFrame(
      wrapped.replace('officebridge_propose_result_sheet', 'exec'),
      'fenced',
    ),
  );
  assert.throws(() => parseDoubaoFrame('Example: ' + wrapped, 'fenced'));
  assert.equal(
    parseDoubaoFrame('现在生成新增工作表方案。\n' + frame, 'fenced').type,
    'tools',
  );
  assert.equal(
    parseDoubaoFrame(
      '已获取分组统计。现在生成"机构汇总"新工作表方案。\n' + frame,
      'fenced',
    ).type,
    'tools',
  );
  assert.equal(
    parseDoubaoFrame('已读取数据。\n' + wrapped + '\n等待应用。', 'fenced')
      .type,
    'tools',
  );
  for (const ambiguous of [
    frame + '\n' + frame,
    '[' + frame + ']',
    '引用：' + frame,
    '“' + frame + '”',
    'x'.repeat(201) + frame,
    '```json\n' + frame + '\n```\n```json\n' + frame + '\n```',
    JSON.stringify(frame),
    frame.slice(0, -1),
  ])
    assert.throws(() => parseDoubaoFrame(ambiguous, 'fenced'));
  const escaped = frame.replace(
    '"Test"',
    JSON.stringify('brace } { and "quote"'),
  );
  assert.equal(
    parseDoubaoFrame('准备方案。' + escaped, 'fenced').calls[0].arguments.name,
    'brace } { and "quote"',
  );
  assert.throws(() =>
    parseDoubaoFrame(
      frame.replace('"values":[[""]]', '"values":[[""]]]'),
      'fenced',
    ),
  );
  assert.equal(parseDoubaoFrame('普通文字答复', 'fenced'), null);
});

test('Doubao bounds format corrections in its original conversation and only publishes an actually staged result sheet', async () => {
  for (const mode of ['corrected', 'still-invalid', 'invalid-after-staging']) {
    const model = await loadWpsHost();
    const messages: any[] = [];
    const nativeId = '7333333333333';
    const adapter = new DoubaoAdapter({
      connect: async () => ({
        conversationId: undefined,
        async send(text, conversationId) {
          assert.equal(conversationId, messages.length ? nativeId : undefined);
          const message = JSON.parse(text);
          messages.push(message);
          const frame = {
            protocol: 'officebridge/1',
            requestId: message.requestId,
            type: 'tools',
            calls: [
              {
                name: 'officebridge_propose_result_sheet',
                arguments: { name: 'Test', values: [['']] },
              },
            ],
          };
          let reply = JSON.stringify(frame);
          if (
            mode === 'still-invalid' ||
            (mode === 'corrected' && messages.length === 1) ||
            (mode === 'invalid-after-staging' && messages.length > 1)
          )
            reply = reply.replace('"values":[[""]]', '"values":[[""]]]');
          else if (message.type === 'tool_results')
            reply = JSON.stringify({
              protocol: 'officebridge/1',
              requestId: message.requestId,
              type: 'final',
              text: '新增工作表方案已准备，等待应用。',
            });
          return {
            conversationId: nativeId,
            runId: String(8333333333330 + messages.length),
            text: reply,
          };
        },
        async cancel() {
          return true;
        },
        close() {},
      }),
    });
    const { id } = await adapter.openConversation();
    const events = await collect(
      adapter.run({
        requestId: 'format-check',
        conversationId: id,
        prompt: 'Create a blank sheet named Test',
        context: fullSelection,
        workbook: model.clone(model.host.capture('selection')),
        signal: new AbortController().signal,
      }),
    );
    assert.ok(!JSON.stringify(events).includes('"protocol"'));
    assert.ok(
      messages.every(
        (message) => message.instructions && message.history === undefined,
      ),
    );
    assert.equal(
      messages.filter((message) => message.type === 'format_error').length,
      mode === 'corrected' ? 1 : 2,
    );
    if (mode === 'corrected') {
      assert.equal(events.at(-1)?.type, 'complete');
      const proposal = events.find((event) => event.type === 'proposal');
      assert.ok(proposal?.type === 'proposal');
      assert.deepEqual(proposal.plan.newSheets, [
        { name: 'Test', values: [['']] },
      ]);
      assert.equal(model.sheets.length, 1);
      model.host.apply(proposal.plan);
      assert.equal(model.sheets.at(-1).Name, 'Test');
      model.host.undo();
      assert.equal(model.sheets.length, 1);
    } else {
      assert.equal(events.at(-1)?.type, 'error');
      assert.ok(!events.some((event) => event.type === 'proposal'));
      assert.ok(messages.length <= 4);
      assert.equal(model.sheets.length, 1);
    }
    adapter.dispose();
  }
});

let serial = 0;
class FakeDoubao implements DoubaoDriver {
  conversationId: string | undefined;
  nativeId = String(7000000000000 + ++serial);
  calls: { message: any; conversationId?: string | undefined }[] = [];
  mode = 'clean';
  cancelCalls = 0;
  stopConfirmed = true;
  closed = false;
  step = 0;
  finish: (() => void) | undefined;
  async send(text: string, conversationId?: string) {
    assert.equal(conversationId, this.conversationId);
    this.conversationId = this.nativeId;
    const message = JSON.parse(text);
    this.calls.push({ message, conversationId });
    if (message.type === 'user') this.step = 0;
    this.step++;
    const frame = (data: unknown) =>
      JSON.stringify({
        protocol: 'officebridge/1',
        requestId: message.requestId,
        ...(data as object),
      });
    let result = '';
    if (
      ['hang', 'unconfirmed'].includes(this.mode) ||
      (this.mode === 'hang-after-plan' && this.step === 3)
    ) {
      await new Promise<void>((resolve) => {
        this.finish = resolve;
      });
      throw new Error('cancelled in native client');
    }
    if (this.mode === 'fail')
      throw new Error('private-provider-token-should-not-leak');
    if (this.mode === 'disconnect') throw new Error('豆包客户端连接已断开。');
    if (this.mode === 'text')
      result = '<tool_call>readRange 已暂存方案，点击应用。';
    else if (this.mode === 'wrong-id')
      result = JSON.stringify({
        protocol: 'officebridge/1',
        requestId: 'old',
        type: 'final',
        text: 'old',
      });
    else if (this.mode === 'out-of-range')
      result = frame({
        type: 'tools',
        calls: [
          {
            name: 'officebridge_propose_edits',
            arguments: {
              edits: [{ rangeId: 'r1', row: 9999, column: 0, value: 'bad' }],
            },
          },
        ],
      });
    else if (this.mode === 'unknown-tool')
      result = frame({
        type: 'tools',
        calls: [{ name: 'exec', arguments: { command: 'anything' } }],
      });
    else if (this.mode === 'limit')
      result = frame({
        type: 'tools',
        calls: [{ name: 'officebridge_list_ranges', arguments: {} }],
      });
    else if (this.step === 1)
      result = frame({
        type: 'tools',
        calls: [
          { name: 'officebridge_read_range', arguments: { rangeId: 'r1' } },
        ],
      });
    else if (this.step === 2)
      result = frame({
        type: 'tools',
        calls: [
          {
            name: 'officebridge_clean_range',
            arguments: { rangeId: 'r1', headerRow: false },
          },
        ],
      });
    else if (['fail-after-plan', 'timeout-after-plan'].includes(this.mode))
      throw Object.assign(new Error('private provider detail'), {
        code: this.mode === 'timeout-after-plan' ? 'timeout' : 'connection',
      });
    else if (this.mode === 'text-after-plan') result = '完成了，请应用。';
    else result = frame({ type: 'final', text: '方案已准备，等待应用。' });
    return {
      conversationId:
        this.mode === 'different-native' && conversationId
          ? '8888888888888'
          : this.nativeId,
      runId: String(
        8000000000000 + (this.mode === 'duplicate-run' ? 1 : this.calls.length),
      ),
      text: result,
    };
  }
  async cancel() {
    this.cancelCalls++;
    this.finish?.();
    return this.stopConfirmed && this.mode !== 'unconfirmed';
  }
  close() {
    this.closed = true;
    this.finish?.();
  }
}

test('Doubao reuses its native conversation, refreshes selection, and stages real bounded tool results without history replay', async () => {
  const model = await loadWpsHost();
  const drivers: FakeDoubao[] = [];
  const adapter = new DoubaoAdapter({
    connect: async () => {
      const driver = new FakeDoubao();
      drivers.push(driver);
      return driver;
    },
  });
  const { id } = await adapter.openConversation();
  const run = () =>
    collect(
      adapter.run({
        requestId: 'test',
        conversationId: id,
        prompt: 'trim-now',
        context: fullSelection,
        workbook: model.clone(model.host.capture('selection')),
        signal: new AbortController().signal,
      }),
    );
  const first = await run();
  assert.equal(first.at(-1)?.type, 'complete');
  const proposal = first.find((event) => event.type === 'proposal');
  assert.ok(proposal && proposal.type === 'proposal');
  assert.equal(model.sheets[0].Cells.Item(2, 1).Value2, ' a ');
  model.host.apply(proposal.plan);
  assert.equal(model.sheets[0].Cells.Item(2, 1).Value2, 'a');
  model.host.undo();
  model.select(2, 1);
  const second = await run();
  assert.equal(second.at(-1)?.type, 'complete');
  assert.equal(drivers.length, 1);
  assert.equal(drivers[0]!.calls[0]!.conversationId, undefined);
  assert.ok(
    drivers[0]!.calls
      .slice(1)
      .every((call) => call.conversationId === drivers[0]!.nativeId),
  );
  const users = drivers[0]!.calls.filter(
    (call) => call.message.type === 'user',
  );
  assert.equal(users.length, 2);
  const taskIds = users.map((call) => call.message.requestId);
  assert.notEqual(taskIds[0], taskIds[1]);
  assert.ok(
    drivers[0]!.calls
      .slice(0, 3)
      .every((call) => call.message.requestId === taskIds[0]),
  );
  assert.ok(
    drivers[0]!.calls
      .slice(3)
      .every((call) => call.message.requestId === taskIds[1]),
  );
  assert.equal(users[1]!.message.ranges.ranges[0].address, 'A2');
  assert.ok(
    drivers[0]!.calls.every((call) => call.message.history === undefined),
  );
  await adapter.closeConversation(id);
  assert.ok(drivers[0]!.closed);
});

test('Doubao advertises and stages native row, rename, filter and chart tools; only host confirmation changes the workbook', async () => {
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
    const model = await loadWpsHost();
    let count = 0;
    const adapter = new DoubaoAdapter({
      connect: async () => ({
        conversationId: undefined,
        async send(text) {
          const message = JSON.parse(text);
          count++;
          if (count === 1)
            assert.ok(
              message.tools.some((tool: any) => tool.name === call.name),
            );
          return {
            conversationId: '7333333333333',
            runId: String(8333333333330 + count),
            text: JSON.stringify({
              protocol: 'officebridge/1',
              requestId: message.requestId,
              ...(count === 1
                ? { type: 'tools', calls: [call] }
                : { type: 'final', text: '方案已准备，等待应用。' }),
            }),
          };
        },
        async cancel() {
          return true;
        },
        close() {},
      }),
    });
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
    assert.equal(events.at(-1)?.type, 'complete');
    assert.equal(model.sheets[0].Name, 'Data');
    model.host.apply(proposal.plan);
    model.host.undo();
    assert.equal(model.sheets[0].Name, 'Data');
    assert.equal(model.sheets[0].Cells.Item(2, 1).Value2, ' a ');
    adapter.dispose();
  }
});
test('Doubao never publishes a plan for unstructured text, invalid/stale tools, incomplete final replies, transport failures or loop limits', async () => {
  for (const mode of [
    'text',
    'wrong-id',
    'out-of-range',
    'unknown-tool',
    'limit',
    'fail',
    'fail-after-plan',
    'text-after-plan',
    'different-native',
    'duplicate-run',
  ]) {
    const model = await loadWpsHost();
    const driver = new FakeDoubao();
    driver.mode = mode;
    const adapter = new DoubaoAdapter({ connect: async () => driver });
    const { id } = await adapter.openConversation();
    const events = await collect(
      adapter.run({
        requestId: mode,
        conversationId: id,
        prompt: 'trim',
        context: fullSelection,
        workbook: model.clone(model.host.capture('selection')),
        signal: new AbortController().signal,
      }),
    );
    assert.ok(!events.some((event) => event.type === 'proposal'), mode);
    assert.ok(!JSON.stringify(events).includes('private'), mode);
    assert.equal(model.sheets[0].Cells.Item(2, 1).Value2, ' a ');
    assert.ok(driver.calls.length <= 8);
    if (mode.startsWith('text'))
      assert.ok(
        events.some((event) => event.type === 'complete' && event.notice),
      );
    else assert.equal(events.at(-1)?.type, 'error');
    adapter.dispose();
  }
});

test('Doubao cancellation confirms the owning native turn, permits continued conversation, and closing a chat releases an active run', async () => {
  const model = await loadWpsHost();
  const driver = new FakeDoubao();
  driver.mode = 'hang';
  const adapter = new DoubaoAdapter({ connect: async () => driver });
  const { id } = await adapter.openConversation();
  const controller = new AbortController();
  const request = {
    requestId: 'cancel',
    conversationId: id,
    prompt: 'wait',
    context: fullSelection,
    workbook: model.clone(model.host.capture('selection')),
    signal: controller.signal,
  };
  const waiting = collect(adapter.run(request));
  while (!driver.finish) await new Promise((resolve) => setTimeout(resolve, 1));
  const busy = await collect(
    adapter.run({ ...request, signal: new AbortController().signal }),
  );
  assert.ok(
    busy.some(
      (event) => event.type === 'error' && event.code === 'conversation_busy',
    ),
  );
  controller.abort();
  assert.ok(
    (await waiting).some(
      (event) => event.type === 'error' && event.code === 'cancelled',
    ),
  );
  driver.mode = 'clean';
  assert.equal(
    (
      await collect(
        adapter.run({ ...request, signal: new AbortController().signal }),
      )
    ).at(-1)?.type,
    'complete',
  );
  driver.mode = 'hang';
  driver.finish = undefined;
  const closing = collect(
    adapter.run({ ...request, signal: new AbortController().signal }),
  );
  while (!driver.finish) await new Promise((resolve) => setTimeout(resolve, 1));
  await adapter.closeConversation(id);
  assert.equal((await closing).at(-1)?.type, 'error');
  assert.ok(driver.closed);
});

test('Doubao new chat uses a separate binding and transport loss is never repaired by replaying or silently creating a session', async () => {
  const model = await loadWpsHost();
  const drivers: FakeDoubao[] = [];
  const adapter = new DoubaoAdapter({
    connect: async () => {
      const driver = new FakeDoubao();
      drivers.push(driver);
      return driver;
    },
  });
  const first = await adapter.openConversation();
  const second = await adapter.openConversation();
  assert.notEqual(first.id, second.id);
  drivers[0]!.mode = 'disconnect';
  const request = {
    requestId: 'loss',
    conversationId: first.id,
    prompt: 'test',
    context: fullSelection,
    workbook: model.clone(model.host.capture('selection')),
    signal: new AbortController().signal,
  };
  assert.ok(
    (await collect(adapter.run(request))).some(
      (event) => event.type === 'error' && event.code === 'conversation_lost',
    ),
  );
  await collect(adapter.run(request));
  assert.equal(drivers[0]!.calls.length, 1);
  assert.equal(drivers.length, 2);
  adapter.dispose();
});

test('Doubao recovers a mismatched request only through a new valid reply; prose never leaks and invalid tools never stage edits', async () => {
  const model = await loadWpsHost();
  const messages: any[] = [];
  const adapter = new DoubaoAdapter({
    connect: async () => ({
      conversationId: undefined,
      async send(text) {
        const message = JSON.parse(text);
        messages.push(message);
        const frame = {
          protocol: 'officebridge/1',
          requestId: message.requestId,
          ...(messages.length === 4
            ? { type: 'final', text: '方案已准备。' }
            : {
                type: 'tools',
                calls: [
                  {
                    name:
                      messages.length === 1
                        ? 'officebridge_read_range'
                        : 'officebridge_propose_edits',
                    arguments:
                      messages.length === 1
                        ? { rangeId: 'r1' }
                        : {
                            edits: [
                              {
                                rangeId: 'r1',
                                row: 1,
                                column: 0,
                                value:
                                  messages.length === 2
                                    ? 'must never stage'
                                    : 'a',
                              },
                            ],
                          },
                  },
                ],
              }),
        };
        if (messages.length === 2) frame.requestId = 'previous-user-task';
        return {
          conversationId: '7333333333333',
          runId: String(8333333333330 + messages.length),
          text: '现在继续处理。\n' + JSON.stringify(frame),
        };
      },
      async cancel() {
        return true;
      },
      close() {},
    }),
  });
  const { id } = await adapter.openConversation();
  const events = await collect(
    adapter.run({
      requestId: 'mismatch',
      conversationId: id,
      prompt: 'trim',
      context: fullSelection,
      workbook: model.clone(model.host.capture('selection')),
      signal: new AbortController().signal,
    }),
  );
  assert.equal(events.at(-1)?.type, 'complete');
  assert.equal(messages[2].type, 'format_error');
  assert.equal(new Set(messages.map((m) => m.requestId)).size, 1);
  const proposal = events.find((e) => e.type === 'proposal');
  assert.ok(proposal?.type === 'proposal');
  assert.equal(proposal.plan.edits.length, 1);
  assert.equal(proposal.plan.edits[0]?.value, 'a');
  assert.ok(!JSON.stringify(events).includes('must never stage'));
  assert.ok(!JSON.stringify(events).includes('officebridge/1'));
  assert.equal(model.sheets[0].Cells.Item(2, 1).Value2, ' a ');
  adapter.dispose();
});

test('Doubao transport timeout cancels before closing, reports stop confirmation honestly and discards staged plans without resending', async () => {
  for (const confirmed of [true, false]) {
    const model = await loadWpsHost();
    const driver = new FakeDoubao();
    driver.mode = 'timeout-after-plan';
    driver.stopConfirmed = confirmed;
    const close = driver.close.bind(driver);
    driver.close = () => {
      assert.ok(
        driver.cancelCalls >= 1,
        'must cancel before closing transport',
      );
      close();
    };
    const adapter = new DoubaoAdapter({ connect: async () => driver });
    const { id } = await adapter.openConversation();
    const request = {
      requestId: 'timeout',
      conversationId: id,
      prompt: 'trim',
      context: fullSelection,
      workbook: model.clone(model.host.capture('selection')),
      signal: new AbortController().signal,
    };
    const events = await collect(adapter.run(request));
    const error = events.at(-1);
    assert.ok(error?.type === 'error');
    assert.equal(error.code, 'conversation_lost');
    assert.match(error.message, /超时/);
    assert.ok(
      error.message.includes(
        confirmed ? '已确认本轮任务结束' : '未确认豆包任务已停止',
      ),
    );
    assert.ok(!events.some((e) => e.type === 'proposal'));
    assert.equal(driver.cancelCalls, 1);
    assert.ok(driver.closed);
    await collect(adapter.run(request));
    assert.equal(driver.calls.length, 3);
    assert.equal(model.sheets[0].Cells.Item(2, 1).Value2, ' a ');
    adapter.dispose();
  }
});

test('stopping after a staged edit discards the proposal; an unconfirmed stop invalidates the native binding', async () => {
  for (const mode of ['hang-after-plan', 'unconfirmed']) {
    const model = await loadWpsHost();
    const driver = new FakeDoubao();
    driver.mode = mode;
    const adapter = new DoubaoAdapter({ connect: async () => driver });
    const { id } = await adapter.openConversation();
    const controller = new AbortController();
    const request = {
      requestId: mode,
      conversationId: id,
      prompt: 'trim',
      context: fullSelection,
      workbook: model.clone(model.host.capture('selection')),
      signal: controller.signal,
    };
    const pending = collect(adapter.run(request));
    while (!driver.finish)
      await new Promise((resolve) => setTimeout(resolve, 1));
    controller.abort();
    const events = await pending;
    assert.ok(!events.some((event) => event.type === 'proposal'));
    assert.ok(
      events.some(
        (event) =>
          event.type === 'error' &&
          event.code ===
            (mode === 'unconfirmed' ? 'conversation_lost' : 'cancelled'),
      ),
    );
    assert.equal(model.sheets[0].Cells.Item(2, 1).Value2, ' a ');
    adapter.dispose();
  }
});

test('Doubao discovery rejects network endpoints, browser/web targets, stale frames and unsupported commands', async () => {
  assert.equal(doubaoEndpoint(), 'http://127.0.0.1:9225');
  for (const endpoint of [
    'http://localhost:9225',
    'http://example.com',
    'https://127.0.0.1:9225',
    'http://user@127.0.0.1:9225',
    'http://127.0.0.1:9225/x',
  ])
    assert.throws(() => doubaoEndpoint(endpoint));
  assert.ok(isDoubaoChat('doubao://doubao-chat/chat/'));
  assert.ok(!isDoubaoChat('https://www.doubao.com/chat/'));
  assert.ok(!isDoubaoChat('chrome://settings/'));
  assert.throws(() =>
    parseDoubaoFrame(
      '{"protocol":"officebridge/1","requestId":"old","type":"final","text":"x"}',
      'new',
    ),
  );
  assert.equal(parseDoubaoFrame('ordinary text', 'new'), null);
  const original = globalThis.fetch;
  try {
    globalThis.fetch = async () =>
      new Response(
        JSON.stringify([
          {
            type: 'page',
            url: 'https://www.doubao.com/chat/',
            webSocketDebuggerUrl: 'ws://127.0.0.1:9225/x',
          },
        ]),
      );
    await assert.rejects(findDoubaoTarget('http://127.0.0.1:9225'));
    globalThis.fetch = async () =>
      new Response(
        JSON.stringify([
          {
            type: 'page',
            url: 'doubao://doubao-chat/chat/',
            webSocketDebuggerUrl: 'ws://example.com/x',
          },
        ]),
      );
    await assert.rejects(findDoubaoTarget('http://127.0.0.1:9225'));
  } finally {
    globalThis.fetch = original;
  }
});

class Renderer implements DoubaoRpc {
  bodies: any[] = [];
  lookups: any[] = [];
  persistedReply = 'native reply';
  storedRunMismatch = false;
  streamError = '';
  storedRunning = false;
  closed = false;
  account = 'fixture-user';
  listeners = new Map<string, (params: any) => void>();
  sandbox: any;
  context: ReturnType<typeof createContext>;
  constructor() {
    this.sandbox = {
      location: {
        protocol: 'doubao:',
        hostname: 'doubao-chat',
        pathname: '/chat/',
      },
      localStorage: { getItem: () => this.account },
      performance: {
        getEntriesByType: () => [
          {
            name: 'https://www.doubao.com/im/test?aid=582478&device_id=fixture-device',
          },
        ],
      },
      neotix: {
        taskMode: {
          runtime: {
            queryRuntimeInfo: async () => ({
              env: { environmentId: 'fixture-env' },
            }),
          },
        },
      },
      '@flow-web/desktop:stable': [],
      crypto: webcrypto,
      URL,
      URLSearchParams,
      setTimeout,
      clearTimeout,
      AbortController,
      AbortSignal,
      TextDecoder,
      fetch: async (_url: string, options: any) => {
        const body = JSON.parse(options.body);
        if (_url.includes('/im/conversation/batch_get')) {
          this.lookups.push(body);
          const runId = String(8111111111110 + this.bodies.length);
          return Response.json({
            downlink_body: {
              batch_get_conv_info_downlink_body: {
                conversation_info_list: [
                  {
                    conversation_id: '7111111111111',
                    messages: [
                      {
                        user_type: 1,
                        message_id: this.storedRunMismatch
                          ? '8999999999999'
                          : runId,
                      },
                      {
                        user_type: 2,
                        message_id: String(9111111111110 + this.bodies.length),
                        bot_reply_message_id: runId,
                        ext: { is_finish: this.storedRunning ? '0' : '1' },
                        content_status: this.storedRunning ? 100 : undefined,
                        content_block: [
                          {
                            block_type: 10000,
                            content: {
                              text_block: { text: this.persistedReply },
                            },
                          },
                        ],
                      },
                    ],
                  },
                ],
              },
            },
          });
        }
        this.bodies.push(body);
        const events = [
          [
            'SSE_ACK',
            {
              ack_client_meta: { conversation_id: '7111111111111' },
              query_list: [
                { question_id: String(8111111111110 + this.bodies.length) },
              ],
            },
          ],
          [
            'STREAM_CHUNK',
            {
              patch_op: [
                {
                  patch_value: {
                    content_block: [
                      {
                        block_type: 10000,
                        content: { text_block: { text: 'native reply' } },
                      },
                    ],
                  },
                },
              ],
            },
          ],
          this.streamError
            ? ['STREAM_ERROR', { error_code: this.streamError }]
            : ['SSE_REPLY_END', { end_type: 3 }],
        ];
        return new Response(
          events
            .map(
              ([event, data]) =>
                `event: ${event}\ndata: ${JSON.stringify(data)}\n\n`,
            )
            .join(''),
        );
      },
    };
    this.sandbox.window = this.sandbox;
    this.context = createContext(this.sandbox);
  }
  async evaluate(expression: string) {
    return await runInContext(expression, this.context);
  }
  async send(method: string, params: any) {
    if (method === 'Runtime.addBinding')
      this.sandbox[params.name] = (payload: string) =>
        this.listeners.get('Runtime.bindingCalled')?.({
          name: params.name,
          payload,
        });
    if (method === 'Runtime.removeBinding') delete this.sandbox[params.name];
    return {};
  }
  subscribe(method: string, callback: (params: any) => void) {
    this.listeners.set(method, callback);
    return () => {
      this.listeners.delete(method);
    };
  }
  close() {
    this.closed = true;
  }
}

test('native Doubao driver probes without inference, uses cloud context with no local sandbox or shared folders, and pins account/native conversation', async () => {
  const renderer = new Renderer();
  const driver = await connectDoubao({ client: renderer });
  assert.equal(renderer.bodies.length, 0);
  assert.equal(renderer.lookups.length, 0);
  const first = await driver.send('first');
  assert.equal(first.text, 'native reply');
  const second = await driver.send('second', first.conversationId);
  assert.equal(second.conversationId, first.conversationId);
  assert.equal(renderer.bodies.length, 2);
  assert.equal(renderer.lookups.length, 2);
  for (const lookup of renderer.lookups)
    assert.deepEqual(
      lookup.uplink_body.batch_get_conv_info_uplink_body.conversation_id,
      ['7111111111111'],
    );
  for (const body of renderer.bodies) {
    const context = body.option.general_task_param;
    assert.equal(context.runtime_type, 1);
    assert.equal(context.client_option, undefined);
    assert.deepEqual(JSON.parse(context.task_input_json), {
      schema_version: 1,
      project_context: {},
    });
    assert.equal(body.user_context.length, 0);
    assert.ok(!JSON.stringify(body).includes('fixture-user'));
  }
  renderer.account = 'changed-account';
  await assert.rejects(driver.send('not sent', second.conversationId));
  assert.equal(renderer.bodies.length, 2);
  driver.close();
  assert.ok(renderer.closed);
});

test('native Doubao uses the exact completed stored reply when the stream is incomplete, and rejects a missing requested run', async () => {
  const renderer = new Renderer();
  renderer.persistedReply = JSON.stringify({
    protocol: 'officebridge/1',
    requestId: 'stored',
    type: 'final',
    text: '完整回复，超过流式预览内容。',
  });
  const driver = await connectDoubao({ client: renderer });
  const reply = await driver.send('read full reply');
  assert.equal(reply.text, renderer.persistedReply);
  assert.equal(parseDoubaoFrame(reply.text, 'stored').type, 'final');
  renderer.storedRunMismatch = true;
  await assert.rejects(
    driver.send('exact turn required', reply.conversationId),
  );
  assert.equal(renderer.bodies.length, 2);
  driver.close();
});

test('native timeout observes only the acknowledged run once, recovers completed replies and rejects running or missing runs', async () => {
  for (const mode of ['completed', 'running', 'missing']) {
    const renderer = new Renderer();
    renderer.streamError = 'timeout';
    renderer.storedRunning = mode === 'running';
    renderer.storedRunMismatch = mode === 'missing';
    renderer.persistedReply = 'exact completed reply';
    const driver = await connectDoubao({ client: renderer });
    if (mode === 'completed') {
      const result = await driver.send('one submission only');
      assert.equal(result.text, renderer.persistedReply);
      assert.equal(result.runId, '8111111111111');
    } else {
      await assert.rejects(driver.send('one submission only'), {
        code: 'timeout',
      });
    }
    assert.equal(renderer.bodies.length, 1);
    assert.equal(renderer.lookups.length, 1);
    driver.close();
  }
});

test('bundling preserves the native protocol functions injected into the renderer', async () => {
  const { build } = await import('esbuild');
  const output = await build({
    entryPoints: ['packages/agents/src/doubao-driver.ts'],
    bundle: true,
    write: false,
    platform: 'node',
    format: 'esm',
    target: 'node24',
  });
  const bundled = await import(
    'data:text/javascript;base64,' +
      Buffer.from(output.outputFiles[0]!.text).toString('base64')
  );
  const renderer = new Renderer();
  const driver = await bundled.connectDoubao({ client: renderer });
  assert.equal((await driver.send('bundle verification')).text, 'native reply');
  driver.close();
});
