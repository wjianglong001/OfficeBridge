import { randomUUID } from 'node:crypto';
import {
  isApplicationReceipt,
  type ApplicationReceipt,
  type ResumeOffer,
} from '../../../packages/protocol/src/continuation.js';

interface Pending {
  conversationId: string;
  workbookFullName: string;
  snapshotId: string;
  nextStep: string;
  step: number;
  expires: number;
}
export class Continuations {
  private pending = new Map<string, Pending>();
  clear(conversationId: string) {
    for (const [token, p] of this.pending)
      if (p.conversationId === conversationId) this.pending.delete(token);
  }
  clearAll() {
    this.pending.clear();
  }
  cancel(token: unknown, conversationId: unknown) {
    const p = typeof token === 'string' ? this.pending.get(token) : undefined;
    return (
      !!p &&
      p.conversationId === conversationId &&
      this.pending.delete(token as string)
    );
  }
  issue(input: Omit<Pending, 'expires'>): ResumeOffer {
    this.clear(input.conversationId);
    for (const [token, p] of this.pending)
      if (p.expires < Date.now()) this.pending.delete(token);
    if (input.step > 8)
      return {
        nextStep: input.nextStep,
        token: null,
        reason: '已达到本任务 8 次自动续办上限，应用后请手动说明是否继续。',
      };
    const token = randomUUID();
    this.pending.set(token, { ...input, expires: Date.now() + 30 * 60 * 1000 });
    return { nextStep: input.nextStep, token };
  }
  consume(
    token: unknown,
    conversationId: string,
    workbook: { snapshotId: string; workbookFullName: string },
    application: unknown,
  ): {
    nextStep: string;
    step: number;
    application: ApplicationReceipt;
  } | null {
    if (typeof token !== 'string') return null;
    const p = this.pending.get(token);
    if (
      !p ||
      p.expires < Date.now() ||
      p.conversationId !== conversationId ||
      p.workbookFullName !== workbook.workbookFullName ||
      workbook.snapshotId === p.snapshotId ||
      !isApplicationReceipt(application) ||
      application.snapshotId !== p.snapshotId ||
      application.workbookFullName !== p.workbookFullName
    )
      return null;
    this.pending.delete(token);
    return { nextStep: p.nextStep, step: p.step, application };
  }
}
