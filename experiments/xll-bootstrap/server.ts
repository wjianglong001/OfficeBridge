import { randomBytes } from 'node:crypto';
import { createBridgeServer } from '../../apps/bridge/src/server.js';

declare const XLL_PROBE_ASSETS: [string, string, string][];
declare const XLL_PROBE_TOKEN: string;
const port = 43191;
const base = `http://127.0.0.1:${port}`;
const files = new Map(
  XLL_PROBE_ASSETS.map(([url, type, body]) => [url, { type, body }]),
);
const pageToken = randomBytes(32).toString('hex');
const state = {
  service: 'officebridge-xll-probe',
  version: 1,
  pid: process.pid,
  deployment: 'pending',
  invocation: 'pending',
  reads: 0,
  address: '',
  rows: 0,
  columns: 0,
};
const server = createBridgeServer({ adapters: [] });
const original = server.listeners('request')[0]! as Function;
server.removeAllListeners('request');
server.on('request', async (req, res) => {
  const pathname = new URL(req.url ?? '/', base).pathname;
  if (!pathname.startsWith('/probe/')) {
    original(req, res);
    return;
  }
  res.setHeader('Cache-Control', 'no-store');
  res.setHeader('X-Content-Type-Options', 'nosniff');
  res.setHeader(
    'Content-Security-Policy',
    "default-src 'self'; connect-src 'self'; frame-ancestors 'self'",
  );
  const reply = (code: number, data: unknown) => {
    res.writeHead(code, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify(data));
  };
  if (
    req.headers.host !== `127.0.0.1:${port}` ||
    (req.headers.origin && req.headers.origin !== base)
  ) {
    reply(403, { error: 'invalid_origin' });
    return;
  }
  if (pathname === '/probe/health' && req.method === 'GET') {
    if (req.headers['x-officebridge-xll-probe'] !== XLL_PROBE_TOKEN) {
      reply(403, { error: 'unauthorized' });
      return;
    }
    reply(200, state);
    return;
  }
  if (pathname === '/probe/session' && req.method === 'GET') {
    reply(200, { token: pageToken });
    return;
  }
  if (pathname === '/probe/read-result' && req.method === 'POST') {
    if (
      req.headers['x-officebridge-probe-token'] !== pageToken ||
      req.headers.origin !== base
    ) {
      reply(403, { error: 'unauthorized' });
      return;
    }
    try {
      const chunks: Buffer[] = [];
      let size = 0;
      for await (const chunk of req) {
        size += chunk.length;
        if (size > 1024) throw new Error('oversize');
        chunks.push(chunk);
      }
      const r = JSON.parse(Buffer.concat(chunks).toString());
      if (
        !r ||
        Object.keys(r).some(
          (k) => !['address', 'rows', 'columns'].includes(k),
        ) ||
        !/^[A-Z]{1,3}[1-9]\d*(?::[A-Z]{1,3}[1-9]\d*)?$/.test(r.address) ||
        !Number.isInteger(r.rows) ||
        !Number.isInteger(r.columns) ||
        r.rows < 1 ||
        r.rows > 20 ||
        r.columns < 1 ||
        r.columns > 10
      )
        throw new Error('invalid');
      state.reads++;
      state.address = r.address;
      state.rows = r.rows;
      state.columns = r.columns;
      reply(200, { ok: true });
    } catch {
      reply(400, { error: 'invalid_result' });
    }
    return;
  }
  const file = files.get(pathname);
  if (file && req.method === 'GET') {
    res.writeHead(200, { 'Content-Type': file.type });
    res.end(file.body);
    return;
  }
  reply(404, { error: 'not_found' });
});

async function wps(path: string, data: string) {
  const response = await fetch(`http://127.0.0.1:58890${path}`, {
    method: 'POST',
    body: data,
    signal: AbortSignal.timeout(15000),
    redirect: 'error',
  });
  if (!response.ok) throw new Error(`wps_http_${response.status}`);
  return response.text();
}
server.listen(port, '127.0.0.1', async () => {
  if (process.argv.includes('--verify-only')) {
    state.deployment = 'disabled_for_verification';
    state.invocation = 'disabled_for_verification';
    return;
  }
  try {
    // Same scoped enable payload and wire encoding as the pinned WPS publisher.
    const deploy = {
      cmd: 'enable',
      name: 'OfficeBridgeXllProbe',
      addonType: 'et',
      online: 'true',
      url: base + '/probe/',
      version: '1.0.0',
      customDomain: '',
    };
    const result = await wps(
      '/deployaddons/runParams',
      Buffer.from(JSON.stringify(deploy)).toString('base64'),
    );
    if (result !== 'OK' && result !== '') throw new Error('deployment_refused');
    state.deployment = 'enabled';
    const invoke = {
      name: 'OfficeBridgeXllProbe',
      function: 'OfficeBridgeProbeOpen',
      info: '{}',
      jsPluginsXml: '',
    };
    await wps(
      '/et/runParams',
      'ksowebstartupet://' +
        Buffer.from(JSON.stringify(invoke)).toString('base64'),
    );
    state.invocation = 'sent';
  } catch (error) {
    const safe =
      error instanceof Error &&
      /^wps_http_\d+$|^deployment_refused$/.test(error.message)
        ? error.message
        : 'wps_request_failed';
    if (state.deployment === 'pending') state.deployment = safe;
    else state.invocation = safe;
  }
});
server.on('error', () => {
  process.exitCode = 1;
});
