type Pending = {
  resolve: (value: any) => void;
  reject: (error: Error) => void;
  timer: ReturnType<typeof setTimeout>;
};

/** Only used behind the local Agent adapter; no page-facing evaluate endpoint. */
export class DoubaoCdp {
  private socket: WebSocket;
  private nextId = 1;
  private pending = new Map<number, Pending>();
  private listeners = new Map<string, Set<(params: any) => void>>();
  private closed = false;

  private constructor(url: string) {
    this.socket = new WebSocket(url);
    this.socket.addEventListener('message', (event) => {
      try {
        const raw = String(event.data);
        if (Buffer.byteLength(raw) > 4_000_000) throw new Error();
        const message = JSON.parse(raw);
        if (typeof message.id === 'number') {
          const call = this.pending.get(message.id);
          if (!call) return;
          clearTimeout(call.timer);
          this.pending.delete(message.id);
          if (message.error) call.reject(new Error('豆包客户端接口调用失败。'));
          else call.resolve(message.result);
        } else if (typeof message.method === 'string') {
          for (const handler of this.listeners.get(message.method) ?? [])
            handler(message.params);
        }
      } catch {
        this.close();
      }
    });
    this.socket.addEventListener('close', () => this.close());
    this.socket.addEventListener('error', () => this.close());
  }

  static async connect(url: string) {
    const parsed = new URL(url);
    if (
      parsed.protocol !== 'ws:' ||
      parsed.hostname !== '127.0.0.1' ||
      parsed.username ||
      parsed.password
    )
      throw new Error('豆包调试连接只允许本机 WebSocket。');
    const client = new DoubaoCdp(url);
    try {
      await new Promise<void>((resolve, reject) => {
        const timer = setTimeout(
          () => reject(new Error('豆包客户端连接超时。')),
          5000,
        );
        const done = (error?: Error) => {
          clearTimeout(timer);
          client.socket.removeEventListener('open', open);
          client.socket.removeEventListener('error', failed);
          client.socket.removeEventListener('close', failed);
          error ? reject(error) : resolve();
        };
        const open = () => done();
        const failed = () => done(new Error('豆包客户端未连接。'));
        client.socket.addEventListener('open', open, { once: true });
        client.socket.addEventListener('error', failed, { once: true });
        client.socket.addEventListener('close', failed, { once: true });
      });
      return client;
    } catch (error) {
      client.close();
      throw error;
    }
  }

  send(method: string, params: unknown = {}): Promise<any> {
    if (this.closed || this.socket.readyState !== WebSocket.OPEN)
      return Promise.reject(new Error('豆包客户端连接已断开。'));
    const id = this.nextId++;
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        this.pending.delete(id);
        reject(new Error('豆包客户端接口响应超时。'));
      }, 125000);
      this.pending.set(id, { resolve, reject, timer });
      try {
        this.socket.send(JSON.stringify({ id, method, params }));
      } catch {
        this.close();
      }
    });
  }

  async evaluate(expression: string) {
    const result = await this.send('Runtime.evaluate', {
      expression,
      awaitPromise: true,
      returnByValue: true,
      userGesture: false,
    });
    if (result?.exceptionDetails) throw new Error('豆包客户端执行接口不兼容。');
    return result?.result?.value;
  }

  subscribe(method: string, handler: (params: any) => void) {
    const handlers = this.listeners.get(method) ?? new Set();
    handlers.add(handler);
    this.listeners.set(method, handlers);
    return () => {
      handlers.delete(handler);
      if (!handlers.size) this.listeners.delete(method);
    };
  }

  close() {
    if (this.closed) return;
    this.closed = true;
    for (const call of this.pending.values()) {
      clearTimeout(call.timer);
      call.reject(new Error('豆包客户端连接已断开。'));
    }
    this.pending.clear();
    this.listeners.clear();
    if (
      this.socket.readyState === WebSocket.OPEN ||
      this.socket.readyState === WebSocket.CONNECTING
    )
      this.socket.close();
  }
}

export function doubaoEndpoint(
  value = process.env.OFFICEBRIDGE_DOUBAO_ENDPOINT ?? 'http://127.0.0.1:9225',
) {
  const url = new URL(value);
  if (
    url.protocol !== 'http:' ||
    url.hostname !== '127.0.0.1' ||
    url.username ||
    url.password ||
    url.search ||
    url.hash ||
    url.pathname !== '/' ||
    url.port === '0'
  )
    throw new Error('豆包接口地址只允许 http://127.0.0.1:端口。');
  return url.origin;
}

export function isDoubaoChat(url: string) {
  try {
    const parsed = new URL(url);
    return (
      ['doubao:', 'chrome:'].includes(parsed.protocol) &&
      parsed.hostname === 'doubao-chat' &&
      /^\/chat(?:\/|$)/.test(parsed.pathname)
    );
  } catch {
    return false;
  }
}

export class DoubaoDiscoveryError extends Error {
  override name = 'DoubaoDiscoveryError';
  constructor(
    readonly code:
      | 'endpoint_unavailable'
      | 'unexpected_endpoint'
      | 'no_chat'
      | 'ambiguous_chat',
    message: string,
  ) {
    super(message);
  }
}

export async function findDoubaoTarget(endpoint: string) {
  endpoint = doubaoEndpoint(endpoint);
  let response: Response;
  try {
    response = await fetch(endpoint + '/json/list', {
      signal: AbortSignal.timeout(3000),
      redirect: 'error',
    });
  } catch {
    throw new DoubaoDiscoveryError(
      'endpoint_unavailable',
      `豆包本机接口未开启或不可访问（${new URL(endpoint).host}）。请在侧栏重新连接豆包，按提示完成启动或重启。`,
    );
  }
  if (!response.ok)
    throw new DoubaoDiscoveryError(
      'unexpected_endpoint',
      '此端口未提供豆包所需接口，请检查本机连接端口。',
    );
  const bytes = await response.text();
  if (Buffer.byteLength(bytes) > 1_000_000)
    throw new Error('豆包接口响应过大。');
  const targets: unknown = JSON.parse(bytes);
  if (!Array.isArray(targets)) throw new Error('豆包接口响应无效。');
  const matching = targets.filter(
    (target) =>
      target?.type === 'page' &&
      typeof target.url === 'string' &&
      isDoubaoChat(target.url),
  );
  if (matching.length !== 1)
    throw new DoubaoDiscoveryError(
      matching.length === 0 ? 'no_chat' : 'ambiguous_chat',
      matching.length === 0
        ? '本机接口已连接，但未找到豆包主对话窗口。请打开已登录的豆包主窗口后重试。'
        : '检测到多个豆包主对话窗口，请只保留一个后重试。',
    );
  const socket = new URL(matching[0].webSocketDebuggerUrl);
  if (
    socket.protocol !== 'ws:' ||
    socket.host !== new URL(endpoint).host ||
    socket.username ||
    socket.password ||
    socket.hash ||
    socket.search
  )
    throw new Error('豆包调试连接身份不匹配。');
  return socket.href;
}
