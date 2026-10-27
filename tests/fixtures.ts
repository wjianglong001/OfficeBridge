import type { AnalysisSelection } from '../packages/protocol/src/index.js';
import type {
  AgentAdapter,
  AgentEvent,
  AgentRequest,
} from '../packages/agents/src/index.js';

export const fullSelection: AnalysisSelection = {
  host: 'wps',
  connectionId: 'test-pane',
  workbookName: 'test.xlsx',
  sheetName: '视频',
  address: 'A1:C4',
  rowCount: 4,
  columnCount: 3,
  values: [
    ['机构', '状态', '时长'],
    ['甲', 'completed', 10],
    ['甲', 'failed', null],
    ['乙', 'completed', 30],
  ],
  capturedAt: '2026-10-19T00:00:00.000Z',
};

export const fakeAdapter: AgentAdapter = {
  id: 'codex',
  async probe() {
    return { available: true };
  },
  async openConversation() {
    return { id: crypto.randomUUID() };
  },
  async closeConversation(_id: string) {},
  dispose() {},
  async *run(_request: AgentRequest): AsyncIterable<AgentEvent> {
    yield { type: 'text', text: '测试分析结果' };
    yield { type: 'complete' };
  },
};
