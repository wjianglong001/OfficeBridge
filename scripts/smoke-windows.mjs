import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { once } from 'node:events';
import { copyFile, mkdtemp, rm, readFile } from 'node:fs/promises';
import { createServer } from 'node:http';
import { tmpdir } from 'node:os';
import { basename, join, resolve, sep } from 'node:path';
import { fileURLToPath } from 'node:url';

const source = fileURLToPath(
  new URL(
    process.argv.includes('--versioned')
      ? '../release/' +
          JSON.parse(
            await readFile(new URL('../package.json', import.meta.url), 'utf8'),
          ).version +
          '/OfficeBridge.exe'
      : '../release/OfficeBridge.exe',
    import.meta.url,
  ),
);
const npmEntry = fileURLToPath(
  new URL('../dist/apps/bridge/src/index.js', import.meta.url),
);
const temporary = await mkdtemp(join(tmpdir(), 'OfficeBridge 独立测试 '));
const executable = join(temporary, 'OfficeBridge.exe');
await copyFile(source, executable);
// No node/npm on PATH; cwd has no source tree, node_modules or WPS resources.
const env = {
  ...process.env,
  PATH: join(process.env.SystemRoot ?? 'C:\\Windows', 'System32'),
  OFFICEBRIDGE_DATA_DIR: join(temporary, 'data'),
  OFFICEBRIDGE_CODEX_PATH: join(temporary, 'not-installed.exe'),
  OFFICEBRIDGE_WORKBUDDY_PATH: join(temporary, 'workbuddy-not-installed.exe'),
  OFFICEBRIDGE_DOUBAO_ENDPOINT: 'http://127.0.0.1:1',
  OFFICEBRIDGE_PORT: 'invalid-env-port',
};
let running;
let occupied;
function launch(args, npmMode = false, overrides = {}) {
  const child = spawn(
    npmMode ? process.execPath : executable,
    npmMode ? [npmEntry, ...args] : args,
    {
      cwd: temporary,
      env: { ...env, ...overrides },
      stdio: ['ignore', 'pipe', 'pipe'],
      windowsHide: true,
    },
  );
  let output = '';
  child.stdout.on('data', (chunk) => {
    output += chunk.toString();
  });
  child.stderr.on('data', (chunk) => {
    output += chunk.toString();
  });
  const done = new Promise((resolve, reject) => {
    const timer = setTimeout(() => {
      child.kill();
      reject(new Error(`Executable did not exit: ${output}`));
    }, 45000);
    child.once('error', (error) => {
      clearTimeout(timer);
      reject(error);
    });
    child.once('exit', (code) => {
      clearTimeout(timer);
      resolve({ code, output });
    });
  });
  // The long-running service is awaited later after termination.
  void done.catch(() => {});
  return { child, done, output: () => output };
}
async function close(server) {
  await new Promise((resolve) => server.close(resolve));
  server.closeAllConnections();
}
try {
  const helper = await launch(['--start-doubao'], false, {
    OFFICEBRIDGE_DOUBAO_ENDPOINT: 'http://not-local.invalid',
  }).done;
  assert.equal(helper.code, 1, helper.output);
  assert.match(helper.output, /只允许.*127\.0\.0\.1/);
  occupied = createServer((_request, response) => {
    response.setHeader('Content-Type', 'application/json');
    response.end(JSON.stringify({ service: 'another-program', status: 'ok' }));
  });
  occupied.listen(0, '127.0.0.1');
  await once(occupied, 'listening');
  const port = occupied.address().port;
  const conflict = await launch(['--no-open', '--port', String(port)]).done;
  assert.equal(conflict.code, 1, conflict.output);
  assert.match(conflict.output, /已被其他程序/);
  await close(occupied);
  occupied = undefined;

  running = launch(['--no-open', '--port', String(port)]);
  const base = `http://127.0.0.1:${port}`;
  let ready = false;
  for (let attempt = 0; attempt < 80; attempt++) {
    try {
      const response = await fetch(`${base}/api/v1/health`, {
        signal: AbortSignal.timeout(300),
      });
      const health = await response.json();
      assert.equal(health.service, 'office-bridge');
      ready = true;
      break;
    } catch {
      if (running.child.exitCode !== null) break;
      await new Promise((resolve) => setTimeout(resolve, 50));
    }
  }
  assert.ok(ready, running.output());
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
    const response = await fetch(`${base}/wps/${path}`);
    assert.equal(response.status, 200, path);
    const body = await response.text();
    assert.ok(body.length > 0, path);
    if (path === 'taskpane.html') assert.match(body, /id="cancel-connection"/);
    if (path === 'taskpane.js') assert.match(body, /agents\/doubao\/connect/);
    if (path === 'install.html') {
      assert.ok(body.includes(`${base}/wps/`));
      assert.ok(!body.includes('__OFFICEBRIDGE_ADDON_BASE_URL__'));
    }
  }
  assert.equal((await fetch(`${base}/wps/../../package.json`)).status, 404);
  // No authorization: exercise the embedded route without launching a real app.
  assert.equal(
    (
      await fetch(`${base}/api/v1/agents/doubao/connect`, {
        method: 'POST',
        body: '{}',
      })
    ).status,
    401,
  );
  const agents = await (await fetch(`${base}/api/v1/agents`)).json();
  assert.equal(agents.agents[0].status, 'unavailable');
  assert.equal(
    agents.agents.find((agent) => agent.id === 'doubao').status,
    'unavailable',
  );
  assert.equal(
    agents.agents.find((agent) => agent.id === 'workbuddy').status,
    'unavailable',
  );
  const duplicate = await launch(['--no-open', '--port', String(port)]).done;
  assert.equal(duplicate.code, 0, duplicate.output);
  assert.match(duplicate.output, /已在运行/);
  const npmDuplicate = await launch(['--no-open', '--port', String(port)], true)
    .done;
  assert.equal(npmDuplicate.code, 0, npmDuplicate.output);
  assert.match(npmDuplicate.output, /已在运行/);
  assert.equal((await fetch(`${base}/api/v1/health`)).status, 200);
  const invalid = await launch(['--no-open', '--port', '0']).done;
  assert.equal(invalid.code, 1);
  assert.match(invalid.output, /端口必须/);
  running.child.kill();
  await running.done;
  running = launch(['--no-open', '--port', String(port)], true);
  for (let attempt = 0; attempt < 80; attempt++) {
    try {
      if ((await fetch(`${base}/api/v1/health`)).ok) break;
    } catch {
      /* wait for compiled npm entry */
    }
    await new Promise((resolve) => setTimeout(resolve, 50));
  }
  const npmPage = await fetch(`${base}/wps/taskpane.html`);
  assert.equal(npmPage.status, 200);
  assert.match(await npmPage.text(), /OfficeBridge/);
  const exeAgainstNpm = await launch(['--no-open', '--port', String(port)])
    .done;
  assert.equal(exeAgainstNpm.code, 0, exeAgainstNpm.output);
  assert.match(exeAgainstNpm.output, /已在运行/);
  console.info(
    'PASS: standalone exe without Node on PATH, Chinese path, all embedded assets, unavailable Agent, npm/exe mutual reuse, foreign port conflict, invalid port. No model requests.',
  );
} finally {
  if (running && running.child.exitCode === null) {
    running.child.kill();
    await running.done;
  }
  if (occupied) await close(occupied);
  // Verify the resolved target before a recursive Windows cleanup.
  assert.ok(resolve(temporary).startsWith(resolve(tmpdir()) + sep));
  assert.ok(basename(temporary).startsWith('OfficeBridge 独立测试 '));
  await rm(temporary, { recursive: true, force: true });
}
