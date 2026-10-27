// Developer-only isolated backend checks. Does not deploy a WPS add-in, open a
// workbook, start an Agent, or warm the XLL startup cache.
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { spawn } from 'node:child_process';
import { once } from 'node:events';
import { request } from 'node:http';
const info = JSON.parse(
  await readFile(
    new URL('../../.officebridge/xll-probe/build-info.json', import.meta.url),
    'utf8',
  ),
);
const base = 'http://127.0.0.1:43191';
try {
  await fetch(base + '/probe/health', { signal: AbortSignal.timeout(500) });
  throw new Error(
    'Port 43191 is occupied; verification never reuses a running process.',
  );
} catch (error) {
  if (error.message?.includes('occupied')) throw error;
}
const child = spawn(info.backend, ['--verify-only'], {
  windowsHide: true,
  stdio: ['ignore', 'pipe', 'pipe'],
});
const exited = once(child, 'exit');
try {
  let health;
  for (let attempt = 0; attempt < 40; attempt++) {
    if (child.exitCode !== null)
      throw new Error('Probe backend exited before readiness.');
    try {
      const response = await fetch(base + '/probe/health', {
        headers: { 'X-OfficeBridge-Xll-Probe': info.token },
        signal: AbortSignal.timeout(500),
      });
      if (response.ok) {
        health = await response.json();
        break;
      }
    } catch {}
    await new Promise((resolve) => setTimeout(resolve, 100));
  }
  assert.equal(health?.pid, child.pid);
  assert.equal(health.deployment, 'disabled_for_verification');
  assert.equal(health.reads, 0);
  assert.equal((await fetch(base + '/probe/health')).status, 403);
  assert.equal(
    (
      await fetch(base + '/probe/session', {
        headers: { Origin: 'https://example.invalid' },
      })
    ).status,
    403,
  );
  const wrongHostStatus = await new Promise((resolve, reject) => {
    const req = request(
      base + '/probe/session',
      { headers: { Host: 'localhost:43191' } },
      (res) => {
        res.resume();
        resolve(res.statusCode);
      },
    );
    req.on('error', reject);
    req.end();
  });
  assert.equal(wrongHostStatus, 403);
  for (const file of [
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
    const response = await fetch(base + file);
    assert.equal(response.status, 200, file);
    assert.ok((await response.text()).length > 30, file);
  }
  const session = await (await fetch(base + '/probe/session')).json();
  const report = (data, token = session.token) =>
    fetch(base + '/probe/read-result', {
      method: 'POST',
      headers: {
        Origin: base,
        'Content-Type': 'application/json',
        'X-OfficeBridge-Probe-Token': token,
      },
      body: JSON.stringify(data),
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
    (await report({ address: 'A1:XFD1048576', rows: 1048576, columns: 16384 }))
      .status,
    400,
  );
  assert.equal(
    (await report({ address: 'B2:C3', rows: 2, columns: 2 })).status,
    200,
  );
  console.info(
    'Standalone SEA backend: resources, worker route, origin/token guards, bounded metadata report passed. No WPS or Agent integration claimed.',
  );
} finally {
  child.kill();
  await exited;
}
