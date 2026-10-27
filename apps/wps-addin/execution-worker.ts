import { executeComputation } from '../../packages/agents/src/execution.js';
const scope = globalThis as unknown as {
  onmessage: (event: { data: unknown }) => void;
  postMessage: (message: unknown) => void;
};
scope.onmessage = (event) => {
  try {
    const data = event.data as { operation: string; payload: unknown };
    scope.postMessage({ type: 'progress', message: '正在后台校验和计算方案…' });
    scope.postMessage({
      type: 'result',
      result: executeComputation(data.operation, data.payload),
    });
  } catch (error) {
    scope.postMessage({
      type: 'error',
      message: error instanceof Error ? error.message : '后台方案计算失败。',
    });
  }
};
