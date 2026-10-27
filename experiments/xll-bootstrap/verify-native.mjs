// Isolated native DLL checks. verify-only forbids WPS registration/activation.
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { spawn } from 'node:child_process';
import { once } from 'node:events';
import { request } from 'node:http';
import { createServer } from 'node:net';
const info = JSON.parse(
  await readFile(
    new URL('../../.officebridge/xll-native/build-info.json', import.meta.url),
    'utf8',
  ),
);
const reservation = createServer();
reservation.listen(0, '127.0.0.1');
await once(reservation, 'listening');
const port = reservation.address().port;
const base = 'http://127.0.0.1:' + port;
await new Promise((resolve) => reservation.close(resolve));
const child = spawn(info.testHost, [info.xll, String(port)], {
  windowsHide: true,
  stdio: 'ignore',
});
const exited = once(child, 'exit');
let ready = false;
try {
  let state;
  for (let n = 0; n < 40; n++) {
    if (child.exitCode !== null)
      throw new Error('Native process exited before readiness.');
    try {
      const response = await fetch(base + '/probe/health', {
        headers: { 'X-OfficeBridge-Xll-Probe': info.token },
        signal: AbortSignal.timeout(500),
      });
      if (response.ok) {
        state = await response.json();
        break;
      }
    } catch {}
    await new Promise((resolve) => setTimeout(resolve, 100));
  }
  assert.equal(state?.pid, child.pid);
  ready = true;
  assert.equal(state.service, 'officebridge-native-xll-probe');
  assert.equal(state.mode, 'in-process-thread');
  assert.equal(state.version, 6);
  assert.equal(state.deployment, 0);
  assert.equal(state.invocation, 0);
  assert.equal(state.reads, 0);
  assert.equal((await fetch(base + '/probe/health')).status, 403);
  assert.equal(
    (
      await fetch(base + '/probe/session', {
        headers: { Origin: 'https://example.invalid' },
      })
    ).status,
    403,
  );
  const wrongHost = await new Promise((resolve, reject) => {
    const req = request(
      base + '/probe/session',
      { headers: { Host: 'localhost:' + port } },
      (res) => {
        res.resume();
        resolve(res.statusCode);
      },
    );
    req.on('error', reject);
    req.end();
  });
  assert.equal(wrongHost, 403);
  for (const path of [
    '/probe/',
    '/probe/main.js',
    '/probe/ribbon.xml',
    '/probe/taskpane.html',
    '/probe/probe.js',
    '/probe/probe.css',
    '/probe/execution-worker.js',
    '/wps/host.js',
    '/wps/contract.js',
    '/wps/host-operations.js',
    '/wps/executor.js',
  ]) {
    const response = await fetch(base + path);
    assert.equal(response.status, 200, path);
    assert.ok((await response.text()).length > 30, path);
  }
  const session = await (await fetch(base + '/probe/session')).json();
  const report = (body, token = session.token) =>
    fetch(base + '/probe/read-result', {
      method: 'POST',
      headers: {
        Origin: base,
        'Content-Type': 'application/json',
        'X-OfficeBridge-Probe-Token': token,
      },
      body: JSON.stringify(body),
    });
  assert.equal(
    (await report({ address: 'B2:C3', rows: 2, columns: 2 }, 'wrong')).status,
    403,
  );
  assert.equal(
    (
      await report({
        address: 'B2:C3',
        rows: 2,
        columns: 2,
        values: [['private']],
      })
    ).status,
    400,
  );
  assert.equal(
    (await report({ address: 'B2:C3', rows: 21, columns: 2 })).status,
    400,
  );
  assert.equal(
    (await report({ address: 'B2:C3', rows: 2, columns: 2 })).status,
    200,
  );
  const after = await (
    await fetch(base + '/probe/health', {
      headers: { 'X-OfficeBridge-Xll-Probe': info.token },
    })
  ).json();
  assert.equal(after.reads, 1);
  assert.equal(after.address, 'B2:C3');
  console.info(
    'PASS: native DLL in-process thread, resources/Worker, host/origin/token guards, bounded metadata reporting; WPS registration disabled. WPS loading and real selection remain separate acceptance steps.',
  );
} finally {
  if (ready)
    await fetch(base + '/probe/stop', {
      method: 'POST',
      headers: { 'X-OfficeBridge-Xll-Probe': info.token },
      body: '',
    }).catch(() => {});
  const completed = await Promise.race([
    exited.then(() => true),
    new Promise((resolve) => setTimeout(() => resolve(false), 3000)),
  ]);
  if (!completed) {
    child.kill();
    await exited;
  }
}
