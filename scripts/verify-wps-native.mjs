// Isolated full XLL bootstrap checks. Never registers with WPS or starts an Agent.
import assert from 'node:assert/strict';
import { execFile, spawn } from 'node:child_process';
import { once } from 'node:events';
import {
  copyFile,
  lstat,
  mkdtemp,
  open,
  readFile,
  realpath,
  rm,
  writeFile,
} from 'node:fs/promises';
import { createServer, request as httpRequest } from 'node:http';
import { createServer as reserveServer } from 'node:net';
import { tmpdir } from 'node:os';
import { basename, join, resolve, sep } from 'node:path';
import { promisify } from 'node:util';

const info = JSON.parse(
  await readFile(
    new URL('../.officebridge/wps-native/build-info.json', import.meta.url),
    'utf8',
  ),
);
const temporary = await mkdtemp(join(tmpdir(), 'OfficeBridge 完整引导测试 '));
const xll = join(temporary, 'OfficeBridge.xll');
const runtime = join(temporary, 'OfficeBridgeService.exe');
const protectedPorts = new Set([4318, 43191, 43192, 58890, 58891]);
const protectedPids = new Set([0, 4, 30900, process.pid]);
const ownedServices = new Map();
const delay = (ms) => new Promise((done) => setTimeout(done, ms));
const execFileAsync = promisify(execFile);
let child,
  nativeDone,
  nativeExited = false,
  nativeVerified = false;
let foreignClosed = true,
  base,
  backendPort,
  session;
let startRequested = false,
  noChildConfirmed = false;
let primaryError;
const auth = { 'X-OfficeBridge-Xll-Probe': info.token };
const foreign = createServer((_request, response) => {
  response.setHeader('Content-Type', 'application/json');
  response.end(
    JSON.stringify({
      service: 'foreign-service',
      status: 'ok',
      protocolVersion: info.protocolVersion,
    }),
  );
});

async function freePort() {
  for (;;) {
    const server = reserveServer();
    server.listen(0, '127.0.0.1');
    await once(server, 'listening');
    const port = server.address().port;
    await new Promise((done) => server.close(done));
    if (!protectedPorts.has(port)) return port;
  }
}
async function closeForeign() {
  if (foreignClosed) return;
  foreign.closeAllConnections();
  await new Promise((done, reject) =>
    foreign.close((error) => (error ? reject(error) : done())),
  );
  foreignClosed = true;
}
async function request(path, options = {}) {
  return fetch(base + path, { ...options, signal: AbortSignal.timeout(3500) });
}
async function post(body, token = session.token, origin = base) {
  return request('/native/start', {
    method: 'POST',
    headers: {
      Origin: origin,
      'Content-Type': 'application/json',
      'X-OfficeBridge-Probe-Token': token,
    },
    body: typeof body === 'string' ? body : JSON.stringify(body),
  });
}
async function rejection(body, error) {
  const response = await post(body);
  assert.equal(response.status, 409);
  assert.equal((await response.json()).error, error);
}

// A PID is eligible for cleanup only after this exact native host reported it.
// Keep a process handle open while checking identity and terminating, so a reused
// PID cannot redirect cleanup to a different process. No process-name/port kills.
const processScript = String.raw`
$ErrorActionPreference = 'Stop'
$targetId = [int]$env:OFFICEBRIDGE_VERIFY_PID
try { $targetProcess = [Diagnostics.Process]::GetProcessById($targetId) }
catch [ArgumentException] { @{ exists = $false } | ConvertTo-Json -Compress; exit 0 }
try {
  $heldHandle = $targetProcess.SafeHandle
  if ($targetProcess.HasExited) { @{ exists = $false } | ConvertTo-Json -Compress; exit 0 }
  $targetPath = $targetProcess.MainModule.FileName
  $created = $targetProcess.StartTime.ToUniversalTime().Ticks.ToString()
  $metadata = Get-CimInstance Win32_Process -Filter ('ProcessId=' + $targetId)
  if ($targetProcess.HasExited) { @{ exists = $false } | ConvertTo-Json -Compress; exit 0 }
  if (!$metadata) { throw 'Cannot establish process ownership.' }
  if (![String]::Equals($targetPath, $env:OFFICEBRIDGE_VERIFY_PATH, [StringComparison]::OrdinalIgnoreCase)) { throw 'Unexpected process executable; refusing cleanup.' }
  if ([int]$metadata.ParentProcessId -ne [int]$env:OFFICEBRIDGE_VERIFY_PARENT) { throw 'Unexpected process parent; refusing cleanup.' }
  if ($env:OFFICEBRIDGE_VERIFY_CREATED -and $created -ne $env:OFFICEBRIDGE_VERIFY_CREATED) { throw 'Process identity changed; refusing cleanup.' }
  if ($env:OFFICEBRIDGE_VERIFY_ACTION -eq 'terminate') {
    $targetProcess.Kill()
    if (!$targetProcess.WaitForExit(5000)) { throw 'Owned service did not exit.' }
    @{ exists = $false } | ConvertTo-Json -Compress
  } else {
    @{ exists = $true; created = $created } | ConvertTo-Json -Compress
  }
} finally { $targetProcess.Dispose() }
`;
async function serviceIdentity(pid, action = 'inspect', created = '') {
  assert.ok(
    nativeVerified,
    'Native instance must be verified before accepting its PID.',
  );
  assert.ok(
    Number.isInteger(pid) && pid > 4 && !protectedPids.has(pid),
    'Refusing a protected or invalid PID.',
  );
  assert.ok(
    !protectedPorts.has(backendPort),
    'Refusing a production backend port.',
  );
  const { stdout } = await execFileAsync(
    join(
      process.env.SystemRoot ?? 'C:\\Windows',
      'System32',
      'WindowsPowerShell',
      'v1.0',
      'powershell.exe',
    ),
    ['-NoProfile', '-NonInteractive', '-Command', processScript],
    {
      windowsHide: true,
      timeout: 10000,
      env: {
        ...process.env,
        OFFICEBRIDGE_VERIFY_PID: String(pid),
        OFFICEBRIDGE_VERIFY_PATH: runtime,
        OFFICEBRIDGE_VERIFY_PARENT: String(child.pid),
        OFFICEBRIDGE_VERIFY_CREATED: created,
        OFFICEBRIDGE_VERIFY_ACTION: action,
      },
    },
  );
  return JSON.parse(stdout.trim());
}
async function status() {
  const response = await request('/native/bridge-status');
  assert.equal(response.status, 200);
  const state = await response.json();
  if (state.pid !== undefined) {
    assert.ok(
      Number.isInteger(state.pid) &&
        state.pid > 4 &&
        !protectedPids.has(state.pid),
    );
    if (!ownedServices.has(state.pid)) {
      const identity = await serviceIdentity(state.pid);
      ownedServices.set(state.pid, identity);
    }
  } else if (state.state === 'failed') noChildConfirmed = true;
  return state;
}
async function stopNative() {
  if (!child) return;
  if (!nativeExited)
    await request('/probe/stop', {
      method: 'POST',
      headers: auth,
      body: '',
    }).catch(() => {});
  await Promise.race([nativeDone, delay(3000)]);
  if (!nativeExited) {
    child.kill(); // Only the test-host ChildProcess created by this script.
    await Promise.race([nativeDone, delay(3000)]);
  }
  assert.ok(nativeExited, 'Owned native test host did not exit.');
}

try {
  await copyFile(info.xll, xll);
  // Reserve the test backend port with a deliberately incompatible listener.
  for (;;) {
    foreign.listen(0, '127.0.0.1');
    await once(foreign, 'listening');
    foreignClosed = false;
    backendPort = foreign.address().port;
    if (!protectedPorts.has(backendPort)) break;
    await closeForeign();
  }
  const nativePort = await freePort();
  assert.notEqual(nativePort, backendPort);
  base = 'http://127.0.0.1:' + nativePort;
  child = spawn(
    info.testHost,
    [xll, '--full', String(nativePort), String(backendPort)],
    {
      windowsHide: true,
      stdio: 'ignore',
      // CreateProcessW inherits these values from its isolated native test host.
      env: {
        ...process.env,
        OFFICEBRIDGE_DATA_DIR: join(temporary, 'data'),
        OFFICEBRIDGE_CODEX_PATH: join(temporary, 'not-installed-codex.exe'),
        OFFICEBRIDGE_WORKBUDDY_PATH: join(
          temporary,
          'not-installed-workbuddy.exe',
        ),
        OFFICEBRIDGE_DOUBAO_ENDPOINT: 'http://127.0.0.1:1',
      },
    },
  );
  protectedPids.add(child.pid);
  nativeDone = once(child, 'exit').then(
    ([code]) => {
      nativeExited = true;
      return code;
    },
    (error) => {
      nativeExited = true;
      throw error;
    },
  );
  void nativeDone.catch(() => {});
  for (let n = 0; n < 60; n++) {
    if (nativeExited)
      throw new Error('Native test host exited before readiness.');
    try {
      const response = await request('/probe/health', { headers: auth });
      if (response.ok) {
        const health = await response.json();
        assert.equal(health.pid, child.pid);
        assert.equal(health.registration, 0);
        nativeVerified = true;
        break;
      }
    } catch (error) {
      if (error.code === 'ERR_ASSERTION') throw error;
    }
    await delay(50);
  }
  assert.ok(nativeVerified, 'Owned native test host did not become ready.');
  assert.equal(
    (await (await request('/native/health')).json()).service,
    'officebridge-wps-bootstrap',
  );
  // fetch may replace Host; use the native HTTP client to send this guard probe.
  const invalidHostStatus = await new Promise((done, reject) => {
    const probe = httpRequest(
      new URL('/native/health', base),
      {
        headers: { Host: 'localhost:' + nativePort },
        agent: false,
        signal: AbortSignal.timeout(3500),
      },
      (response) => {
        response.once('error', reject);
        response.once('end', () => done(response.statusCode));
        response.resume();
      },
    );
    probe.once('error', reject);
    probe.end();
  });
  assert.equal(invalidHostStatus, 403);
  for (const asset of [
    'index.html',
    'main.js',
    'ribbon.xml',
    'launcher.html',
    'launcher.js',
    'styles.css',
  ])
    assert.equal((await request('/bootstrap/' + asset)).status, 200, asset);
  assert.equal((await request('/bootstrap/../../package.json')).status, 404);
  assert.equal((await status()).state, 'conflict');
  session = await (await request('/probe/session')).json();
  assert.equal((await post({ retry: false }, 'wrong')).status, 403);
  assert.equal(
    (await post({ retry: false }, session.token, 'https://example.invalid'))
      .status,
    403,
  );
  assert.equal((await post({ retry: false }, session.token, '')).status, 403);
  assert.equal(
    (await post({ retry: false, executable: 'C:\\unexpected.exe' })).status,
    400,
  );
  assert.equal((await post('{"retry":false,"retry":true}')).status, 400);
  assert.equal((await post({ retry: 'false' })).status, 400);
  await rejection({ retry: false }, 'port_conflict');
  await rejection({ retry: true }, 'port_conflict');
  await closeForeign();
  assert.equal((await status()).state, 'offline');
  await rejection({ retry: false }, 'runtime_unavailable');
  await writeFile(runtime, 'synthetic invalid runtime');
  await rejection({ retry: false }, 'runtime_unavailable');
  await copyFile(info.backend, runtime);
  const altered = await open(runtime, 'r+');
  try {
    const last = (await altered.stat()).size - 1;
    const byte = Buffer.alloc(1);
    assert.equal((await altered.read(byte, 0, 1, last)).bytesRead, 1);
    byte[0] ^= 1;
    await altered.write(byte, 0, 1, last);
  } finally {
    await altered.close();
  }
  await rejection({ retry: true }, 'runtime_unavailable');
  assert.equal((await status()).state, 'offline');
  await copyFile(info.backend, runtime);

  // Never spawn the service here: only the native /start implementation may do it.
  startRequested = true;
  const queued = await post({ retry: false });
  assert.equal(
    queued.status,
    202,
    'Native start must acknowledge a queued thread.',
  );
  assert.deepEqual(await queued.json(), { state: 'starting' });
  for (const retry of [false, true]) {
    const duplicate = await post({ retry });
    const body = await duplicate.json();
    // The service can become healthy between requests. Neither branch may queue again.
    assert.ok(duplicate.status === 409 || duplicate.status === 200);
    if (duplicate.status === 409) assert.equal(body.error, 'already_requested');
    else assert.equal(body.state, 'ready');
  }
  let connected;
  for (let n = 0; n < 80; n++) {
    connected = await status();
    assert.notEqual(
      connected.state,
      'failed',
      'Native startup failed: ' + JSON.stringify(connected),
    );
    assert.notEqual(connected.state, 'conflict');
    assert.ok(
      ownedServices.size <= 1,
      'Duplicate native starts created another process.',
    );
    if (connected.state === 'ready') break;
    assert.equal(connected.state, 'starting');
    await delay(50);
  }
  assert.equal(connected.state, 'ready');
  assert.ok(
    Number.isInteger(connected.pid),
    'Ready native status must report its created child.',
  );
  assert.equal(ownedServices.size, 1);
  assert.equal(ownedServices.get(connected.pid)?.exists, true);
  const servicePid = connected.pid;
  const backendBase = 'http://127.0.0.1:' + backendPort;
  assert.equal(connected.entryUrl, backendBase + '/wps/taskpane.html');
  for (const retry of [false, true]) {
    const duplicate = await post({ retry });
    assert.equal(duplicate.status, 200);
    assert.equal((await duplicate.json()).state, 'ready');
    assert.equal((await status()).pid, servicePid);
  }
  await assert.rejects(
    open(runtime, 'r+'),
    (error) => ['EACCES', 'EPERM', 'EBUSY'].includes(error.code),
    'Verified runtime must remain write-locked.',
  );
  const pane = await (
    await fetch(connected.entryUrl, { signal: AbortSignal.timeout(3500) })
  ).text();
  assert.match(pane, /analysis-form/);
  assert.doesNotMatch(pane, /OfficeBridge-XLL-Probe/);
  for (const asset of [
    'taskpane.js',
    'host.js',
    'host-operations.js',
    'executor.js',
    'execution-worker.js',
    'contract.js',
  ])
    assert.equal(
      (
        await fetch(backendBase + '/wps/' + asset, {
          signal: AbortSignal.timeout(3500),
        })
      ).status,
      200,
    );
  assert.equal(
    (
      await fetch(backendBase + '/api/v1/run', {
        method: 'POST',
        body: '{}',
        signal: AbortSignal.timeout(3500),
      })
    ).status,
    401,
  );

  // Intentional lifecycle assertion: unloading the owned DLL leaves its shared
  // backend alive. Cleanup still kills only the verified child PID below.
  await stopNative();
  assert.equal(await nativeDone, 0);
  const surviving = await fetch(backendBase + '/api/v1/health', {
    signal: AbortSignal.timeout(3500),
  });
  assert.equal(surviving.status, 200);
  assert.equal(
    (
      await serviceIdentity(
        servicePid,
        'inspect',
        ownedServices.get(servicePid).created,
      )
    ).exists,
    true,
  );
} catch (error) {
  primaryError = error;
} finally {
  const cleanupErrors = [];
  // If a response/assertion failed after queuing, learn the PID from this exact
  // native instance before cleanup. Never search for/kill by filename or port.
  if (
    startRequested &&
    !ownedServices.size &&
    !noChildConfirmed &&
    !nativeExited
  ) {
    for (let n = 0; n < 12 && !ownedServices.size && !noChildConfirmed; n++) {
      try {
        await status();
      } catch (error) {
        cleanupErrors.push(error);
        break;
      }
      if (!ownedServices.size && !noChildConfirmed) await delay(50);
    }
  }
  if (startRequested && !ownedServices.size && !noChildConfirmed)
    cleanupErrors.push(
      new Error(
        'Could not prove whether the native startup created a service; preserving its temporary directory.',
      ),
    );
  for (const [pid, identity] of ownedServices) {
    if (!identity.exists) continue;
    try {
      const ended = await serviceIdentity(pid, 'terminate', identity.created);
      assert.equal(
        ended.exists,
        false,
        'Owned service exit was not confirmed.',
      );
      identity.exists = false;
    } catch (error) {
      cleanupErrors.push(error);
    }
  }
  // Outside the explicit unload-survival assertion above, the service is always
  // confirmed exited before its native test host is stopped.
  if (!cleanupErrors.length) {
    try {
      await stopNative();
    } catch (error) {
      cleanupErrors.push(error);
    }
  }
  try {
    await closeForeign();
  } catch (error) {
    cleanupErrors.push(error);
  }
  if (!cleanupErrors.length) {
    try {
      const root = await realpath(tmpdir());
      const target = await realpath(temporary);
      assert.ok(target.toLowerCase().startsWith((root + sep).toLowerCase()));
      assert.ok(basename(target).startsWith('OfficeBridge 完整引导测试 '));
      assert.equal((await lstat(temporary)).isSymbolicLink(), false);
      assert.equal(
        resolve(target).toLowerCase(),
        resolve(temporary).toLowerCase(),
      );
      await rm(target, { recursive: true, force: true });
    } catch (error) {
      cleanupErrors.push(error);
    }
  }
  if (cleanupErrors.length)
    primaryError = new AggregateError(
      [...(primaryError ? [primaryError] : []), ...cleanupErrors],
      'Native verification failed or cleanup could not be proved safe. Temporary files: ' +
        temporary,
    );
}
if (primaryError) throw primaryError;
console.info(
  'PASS: isolated native CreateProcessW startup, real child PID and health, authorization, foreign-port refusal, missing/tampered fixed sibling, duplicate/retry suppression, write lock, complete UI/API, backend survival after native unload, and verified owned-process cleanup. No WPS registration or Agent requests.',
);
