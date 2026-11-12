export function readTurn(
  client: any,
  conversationId: string,
  options?: any,
): Promise<any>;
export function stopTurn(
  client: any,
  conversationId: string,
  options?: any,
): Promise<any>;
export function terminateNativeTask(
  client: any,
  threadId: string,
  deadline: number,
): Promise<any>;
