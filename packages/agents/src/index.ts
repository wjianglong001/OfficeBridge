import type { AnalysisSelection } from '../../protocol/src/index.js';
import type { EditPlan, WorkbookSnapshot } from '../../protocol/src/table.js';
import type { TaskContinuation } from '../../protocol/src/continuation.js';
import type {
  WorkbookDescriptor,
  HostToolCall,
} from '../../protocol/src/on-demand.js';

export type AgentId = 'codex' | 'claude-code' | 'dsh' | 'workbuddy' | 'doubao';

export interface AgentRequest {
  requestId: string;
  prompt: string;
  context?: AnalysisSelection;
  workbook?: WorkbookSnapshot | WorkbookDescriptor;
  hostTool?: HostToolCall;
  conversationId?: string;
  operationFeedback?: string;
  signal: AbortSignal;
}

export type AgentEvent =
  | { type: 'activity'; phase: 'thinking' | 'tool' }
  | { type: 'status'; message: string }
  | { type: 'text'; text: string }
  | {
      type: 'proposal';
      plan: EditPlan;
      snapshot?: WorkbookSnapshot | undefined;
      continuation?: TaskContinuation | undefined;
    }
  | { type: 'complete'; notice?: string }
  | { type: 'error'; code: string; message: string };

/** Each implementation owns its supported invocation interface, not user credentials. */
export interface AgentAdapter {
  readonly id: AgentId;
  probe(): Promise<{ available: boolean; reason?: string }>;
  /** The agent owns conversation history. IDs only address conversations opened here. */
  openConversation(replaceId?: string): Promise<{ id: string }>;
  closeConversation(id: string): Promise<void>;
  dispose(): void;
  run(request: AgentRequest): AsyncIterable<AgentEvent>;
}

/** Roadmap metadata only; this is not installed-agent discovery. */
export const plannedAgents = [
  { id: 'codex', name: 'Codex', status: 'not_implemented' },
  { id: 'claude-code', name: 'Claude Code', status: 'not_implemented' },
  { id: 'dsh', name: 'DSH', status: 'not_implemented' },
  { id: 'workbuddy', name: 'WorkBuddy', status: 'not_implemented' },
  { id: 'doubao', name: '豆包（实验）', status: 'not_implemented' },
] as const;
