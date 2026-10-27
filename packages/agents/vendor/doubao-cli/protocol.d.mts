export const AUTOMATIC_MODEL_PROTOCOL: {
  key: string;
  ndt: number;
  provider: string;
};
export function runtimeParameters(client: any, options?: any): Promise<any>;
export function sendChatCompletion(client: any, options: any): Promise<any>;
export function reduceStreamEvent(
  state: any,
  event: string,
  data: any,
  options: any,
): 'stop' | null;
export function conversationExt(
  model: any,
  messageId: string,
  workspace: string,
  options?: any,
): any;
