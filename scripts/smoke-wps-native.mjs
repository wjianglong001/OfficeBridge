// Packaged GUI-service checks only; no WPS interaction or model request.
import assert from 'node:assert/strict';
import { spawn, execFileSync } from 'node:child_process';
import { once } from 'node:events';
import { copyFile, mkdtemp, readFile, rm } from 'node:fs/promises';
import { createServer } from 'node:net';
import { tmpdir } from 'node:os';
import { basename, join, resolve, sep } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createHash } from 'node:crypto';

if (process.platform !== 'win32') throw new Error('Windows required.');
const info = JSON.parse(
  await readFile(
    new URL('../.officebridge/wps-native/build-info.json', import.meta.url),
    'utf8',
  ),
);
const image = await readFile(info.backend);
const pe = image.readUInt32LE(0x3c);
assert.equal(image.toString('ascii', 0, 2), 'MZ');
assert.equal(image.toString('ascii', pe, pe + 4), 'PE\0\0');
assert.equal(image.readUInt16LE(pe + 4), 0x8664);
assert.equal(image.readUInt16LE(pe + 24), 0x20b);
assert.equal(
  image.readUInt16LE(pe + 24 + 68),
  2,
  'Service must use the GUI subsystem.',
);
assert.equal(
  createHash('sha256').update(image).digest('hex'),
  info.backendSha256,
);
const xll = await readFile(info.xll);
assert.equal(xll.readUInt16LE(xll.readUInt32LE(0x3c) + 4), 0x14c);
assert.equal(createHash('sha256').update(xll).digest('hex'), info.sha256);
const zipListing = [
  '$ErrorActionPreference = "Stop"',
  '[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)',
  'Add-Type -AssemblyName System.IO.Compression',
  'Add-Type -AssemblyName System.IO.Compression.FileSystem',
  '$archive = [System.IO.Compression.ZipFile]::OpenRead($env:OFFICEBRIDGE_NATIVE_ZIP_CHECK)',
  'try { @($archive.Entries | ForEach-Object { $_.FullName }) | ConvertTo-Json -Compress } finally { $archive.Dispose() }',
].join('\n');
const entries = JSON.parse(
  execFileSync(
    join(
      process.env.SystemRoot ?? 'C:\\Windows',
      'System32/WindowsPowerShell/v1.0/powershell.exe',
    ),
    ['-NoProfile', '-NonInteractive', '-Command', zipListing],
    {
      windowsHide: true,
      encoding: 'utf8',
      env: { ...process.env, OFFICEBRIDGE_NATIVE_ZIP_CHECK: info.archive },
    },
  ),
);
assert.deepEqual(
  entries.sort(),
  [
    'OfficeBridge/OfficeBridge.xll',
    'OfficeBridge/OfficeBridgeService.exe',
    'OfficeBridge/THIRD-PARTY-NOTICES.txt',
    'OfficeBridge/使用说明.txt',
  ].sort(),
);

const temporary = await mkdtemp(join(tmpdir(), 'OfficeBridge 原生包烟测 '));
const executable = join(temporary, 'OfficeBridgeService.exe');
await copyFile(info.backend, executable);
const reserved = createServer();
reserved.listen(0, '127.0.0.1');
await once(reserved, 'listening');
const port = reserved.address().port;
await new Promise((resolve) => reserved.close(resolve));
const base = `http://127.0.0.1:${port}`;
const child = spawn(executable, ['--no-open', '--port', String(port)], {
  cwd: temporary,
  windowsHide: true,
  stdio: ['ignore', 'pipe', 'pipe'],
  env: {
    ...process.env,
    PATH: join(process.env.SystemRoot ?? 'C:\\Windows', 'System32'),
    OFFICEBRIDGE_DATA_DIR: join(temporary, 'data'),
    OFFICEBRIDGE_CODEX_PATH: join(temporary, 'codex-not-installed.exe'),
    OFFICEBRIDGE_WORKBUDDY_PATH: join(temporary, 'workbuddy-not-installed.exe'),
    OFFICEBRIDGE_DOUBAO_ENDPOINT: 'http://127.0.0.1:1',
    OFFICEBRIDGE_PORT: 'invalid-env-port',
  },
});
let output = '';
child.stdout.on('data', (data) => {
  output += data.toString();
});
child.stderr.on('data', (data) => {
  output += data.toString();
});
let exited = false;
const done = once(child, 'exit').then(([code]) => {
  exited = true;
  return code;
});
void done.catch(() => {});
try {
  let ready = false;
  for (let attempt = 0; attempt < 100; attempt++) {
    if (exited || child.exitCode !== null)
      throw new Error('GUI service exited before readiness: ' + output);
    try {
      const response = await fetch(base + '/api/v1/health', {
        signal: AbortSignal.timeout(300),
        redirect: 'error',
      });
      assert.equal(response.status, 200);
      const health = await response.json();
      assert.equal(health.service, 'office-bridge');
      assert.equal(health.status, 'ok');
      assert.equal(health.protocolVersion, info.protocolVersion);
      ready = true;
      break;
    } catch (error) {
      if (error.code === 'ERR_ASSERTION') throw error;
    }
    await new Promise((resolve) => setTimeout(resolve, 50));
  }
  assert.ok(ready, 'GUI service did not become ready: ' + output);
  for (const path of [
    'index.html',
    'taskpane.html',
    'taskpane.js',
    'executor.js',
    'execution-worker.js',
    'host.js',
    'host-operations.js',
    'contract.js',
    'styles.css',
    'main.js',
    'ribbon.xml',
    'install.html',
  ]) {
    const response = await fetch(base + '/wps/' + path, {
      signal: AbortSignal.timeout(1500),
    });
    assert.equal(response.status, 200, path);
    assert.ok((await response.text()).length > 20, path);
  }
  assert.equal((await fetch(base + '/wps/../../package.json')).status, 404);
  assert.equal(
    (
      await fetch(base + '/api/v1/conversations', {
        method: 'POST',
        body: '{}',
      })
    ).status,
    401,
  );
  assert.equal(
    (await fetch(base + '/api/v1/run', { method: 'POST', body: '{}' })).status,
    401,
  );
  assert.equal(child.exitCode, null, 'The test must own the running service.');
  console.info(
    'PASS: WPS package allowlist, x86 XLL/hash, GUI x64 service/hash, isolated Chinese path without Node, embedded full resources and API authentication. No WPS or model requests.',
  );
} finally {
  if (!exited && child.exitCode === null) child.kill();
  const finished = await Promise.race([
    done.then(() => true),
    new Promise((resolve) => setTimeout(() => resolve(false), 3000)),
  ]);
  assert.ok(
    finished,
    'Owned GUI service did not stop; preserving its temporary directory.',
  );
  assert.ok(resolve(temporary).startsWith(resolve(tmpdir()) + sep));
  assert.ok(basename(temporary).startsWith('OfficeBridge 原生包烟测 '));
  await rm(temporary, { recursive: true, force: true });
}
