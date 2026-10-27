import { createHash } from 'node:crypto';
import {
  DoubaoCdp,
  doubaoEndpoint,
  findDoubaoTarget,
  DoubaoDiscoveryError,
} from './doubao-cdp.js';
import { withApp } from '../vendor/doubao-cli/app.mjs';
import {
  AUTOMATIC_MODEL_PROTOCOL,
  runtimeParameters,
  sendChatCompletion,
} from '../vendor/doubao-cli/protocol.mjs';
import { readTurn, stopTurn } from '../vendor/doubao-cli/turns.mjs';

export interface DoubaoDriver {
  readonly conversationId: string | undefined;
  send(
    message: string,
    conversationId?: string,
  ): Promise<{ conversationId: string; runId: string; text: string }>;
  cancel(): Promise<boolean>;
  close(): void;
}
export type DoubaoRpc = Pick<
  DoubaoCdp,
  'evaluate' | 'send' | 'subscribe' | 'close'
>;

const identityExpression = `(() => {
  if (!['doubao:', 'chrome:'].includes(location.protocol) || location.hostname !== 'doubao-chat' || !/^\\/chat(?:\\/|$)/.test(location.pathname)) return null;
  const account = localStorage.getItem('flow_tea_user_id');
  if (!account || account === '0' || !window.neotix?.taskMode?.runtime?.queryRuntimeInfo || typeof window['@flow-web/desktop:stable']?.push !== 'function') return null;
  return { account };
})()`;

export class DoubaoSetupError extends Error {
  override name = 'DoubaoSetupError';
}

// Explicit context: neither the upstream app scanner nor its CLI is loaded.
const app = {
  id: 'doubao',
  name: 'Doubao',
  scheme: 'doubao',
  aid: '582478',
  profile: 'Default',
  dataDir: '',
};
export const doubaoCloudContext = Object.freeze({
  runtime: 'cloud',
  projectId: '',
  previousProjectId: '',
  previousAgentParam: { runtime_type: 1 },
  projectContext: {},
});

async function fingerprint(client: DoubaoRpc) {
  const identity = await client.evaluate(identityExpression);
  if (!identity || typeof identity.account !== 'string' || !identity.account)
    throw new DoubaoSetupError(
      '请在豆包主窗口登录，并打开支持工作会话的客户端版本。',
    );
  return createHash('sha256').update(identity.account).digest('hex');
}

class NativeDoubaoDriver implements DoubaoDriver {
  private receipt: any;
  private nativeId: string | undefined;
  private closed = false;
  private active = false;
  private inFlight = false;
  private stopping: Promise<boolean> | undefined;
  constructor(
    private client: DoubaoRpc,
    private account: string,
  ) {}
  get conversationId() {
    return this.nativeId;
  }

  private async identity() {
    if (this.closed || (await fingerprint(this.client)) !== this.account)
      throw new Error('account_changed');
  }

  private async completedReply(runId: string) {
    await this.identity();
    if (!this.nativeId || !/^\d{12,24}$/.test(runId))
      throw new Error('native_reply_not_confirmed');
    const { result: stored } = await withApp(app, () =>
      readTurn(this.client, this.nativeId!, {
        runId,
        receipt: this.receipt,
        deadline: Date.now() + 10000,
      }),
    );
    await this.identity();
    if (
      stored?.conversationId !== this.nativeId ||
      stored?.runId !== runId ||
      stored?.status !== 'completed' ||
      typeof stored?.reply?.text !== 'string' ||
      !stored.reply.text.trim()
    )
      throw new Error('native_reply_not_confirmed');
    this.inFlight = false;
    return {
      conversationId: this.nativeId,
      runId,
      text: stored.reply.text,
    };
  }

  async send(message: string, conversationId?: string) {
    if (this.active || this.closed || this.nativeId !== conversationId)
      throw new Error('native_conversation_mismatch');
    await this.identity();
    this.active = true;
    this.inFlight = true;
    this.stopping = undefined;
    this.receipt = undefined;
    try {
      const result = await withApp(app, () =>
        sendChatCompletion(this.client, {
          conversationId,
          message,
          model: AUTOMATIC_MODEL_PROTOCOL,
          timeoutMs: 90000,
          waitForReply: true,
          workspace: '/officebridge',
          skillPaths: [],
          permission: 'AlwaysAsk',
          localConnectors: [],
          sharedFolderPath: [],
          taskContext: doubaoCloudContext,
          debug: false,
          onReceipt: (receipt: any) => {
            this.receipt = receipt;
            if (receipt.conversationId) {
              if (
                !/^\d{12,24}$/.test(receipt.conversationId) ||
                (this.nativeId && receipt.conversationId !== this.nativeId)
              )
                throw new Error('native_conversation_mismatch');
              this.nativeId = receipt.conversationId;
            }
          },
        }),
      ).catch(async (error: unknown) => {
        // A lost completion event does not necessarily mean inference is still
        // running. Observe this exact acknowledged run once; never resend it.
        if (
          error instanceof Error &&
          'code' in error &&
          error.code === 'timeout' &&
          this.receipt?.conversationId === this.nativeId &&
          /^\d{12,24}$/.test(this.receipt?.runId ?? '')
        ) {
          const recovered = await this.completedReply(
            String(this.receipt.runId),
          ).catch(() => undefined);
          if (recovered) return { recovered };
        }
        throw error;
      });
      if ('recovered' in result) return result.recovered;
      await this.identity();
      if (
        !result.completed ||
        result.waitingInput ||
        result.handoffs?.length ||
        !result.conversationId ||
        !result.runId ||
        typeof result.answer !== 'string'
      )
        throw new Error('native_turn_not_completed');
      if (
        !/^\d{12,24}$/.test(result.conversationId) ||
        !/^\d{12,24}$/.test(String(result.runId)) ||
        (this.nativeId && result.conversationId !== this.nativeId)
      )
        throw new Error('native_conversation_mismatch');
      this.nativeId = result.conversationId;
      this.inFlight = false;
      // The desktop stream can omit text deltas, and brief is only a preview.
      // Read the exact completed turn, never another chat or an inferred turn.
      return await this.completedReply(String(result.runId));
    } finally {
      this.active = false;
    }
  }

  cancel(): Promise<boolean> {
    if (this.stopping) return this.stopping;
    this.stopping = (async () => {
      if (!this.inFlight) return true;
      if (
        !this.receipt?.conversationId ||
        !this.receipt?.runId ||
        this.receipt.conversationId !== this.nativeId
      )
        return false;
      await this.identity();
      const result = await withApp(app, () =>
        stopTurn(this.client, this.receipt.conversationId, {
          runId: this.receipt.runId,
          receipt: this.receipt,
          timeoutMs: 5000,
        }),
      );
      if (result?.stopped === true) this.inFlight = false;
      return result?.stopped === true;
    })().catch(() => false);
    return this.stopping;
  }

  close() {
    this.closed = true;
    this.client.close();
  }
}

export async function connectDoubao(
  options: { endpoint?: string; client?: DoubaoRpc } = {},
): Promise<DoubaoDriver> {
  let client: DoubaoRpc | undefined;
  let endpoint: string;
  try {
    endpoint = doubaoEndpoint(options.endpoint);
  } catch {
    throw new DoubaoSetupError(
      '豆包接口地址配置无效，只允许 http://127.0.0.1:端口。',
    );
  }
  try {
    client =
      options.client ??
      (await DoubaoCdp.connect(await findDoubaoTarget(endpoint)));
    const account = await fingerprint(client);
    // Metadata only: validates the native protocol's non-secret app identity.
    await withApp(app, () => runtimeParameters(client));
    return new NativeDoubaoDriver(client, account);
  } catch (error) {
    client?.close();
    if (error instanceof DoubaoSetupError) throw error;
    if (error instanceof DoubaoDiscoveryError)
      throw new DoubaoSetupError(error.message);
    throw new DoubaoSetupError(
      '豆包本机接口未就绪或版本不兼容。请在豆包中登录并打开主窗口，再在侧栏重新连接。',
    );
  }
}
