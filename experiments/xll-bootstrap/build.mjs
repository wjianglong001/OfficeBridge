import { build } from 'esbuild';
import { inject } from 'postject';
import { execFileSync } from 'node:child_process';
import { createHash, randomBytes } from 'node:crypto';
import { copyFile, mkdir, readFile, writeFile, stat } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { join, resolve } from 'node:path';

if (
  process.platform !== 'win32' ||
  process.arch !== 'x64' ||
  !process.versions.node.startsWith('24.')
)
  throw new Error('Build requires Windows x64 and Node.js 24.');
const root = fileURLToPath(new URL('../../', import.meta.url));
const source = join(root, 'experiments', 'xll-bootstrap');
const staging = join(root, '.officebridge', 'xll-probe');
const release = join(root, 'release', 'xll-probe');
const zigArgument = process.argv.indexOf('--zig');
if (zigArgument < 0 || !process.argv[zigArgument + 1])
  throw new Error('Provide --zig <zig.exe>.');
const zig = resolve(process.argv[zigArgument + 1]);
for (const dir of [staging, release]) await mkdir(dir, { recursive: true });
const run = (exe, args) =>
  execFileSync(exe, args, { cwd: root, windowsHide: true, stdio: 'inherit' });
run(process.execPath, ['scripts/prepare-wps.mjs']);
const token = randomBytes(32).toString('hex');
const assets = [];
for (const [name, type] of [
  ['main.js', 'text/javascript'],
  ['ribbon.xml', 'application/xml'],
  ['taskpane.html', 'text/html; charset=utf-8'],
  ['probe.js', 'text/javascript'],
  ['probe.css', 'text/css'],
])
  assets.push([
    '/probe/' + name,
    type,
    await readFile(join(source, name), 'utf8'),
  ]);
assets.push([
  '/probe/',
  'text/html; charset=utf-8',
  '<!doctype html><html><head><meta charset="utf-8"><script src="./main.js"></script></head><body>OfficeBridge XLL Probe</body></html>',
]);
assets.push([
  '/probe/index.html',
  'text/html; charset=utf-8',
  assets.at(-1)[2],
]);
assets.push([
  '/probe/execution-worker.js',
  'text/javascript',
  await readFile(
    join(root, 'apps/wps-addin/.generated/execution-worker.js'),
    'utf8',
  ),
]);
const bundle = join(staging, 'probe.cjs');
await build({
  absWorkingDir: root,
  entryPoints: ['experiments/xll-bootstrap/server.ts'],
  outfile: bundle,
  bundle: true,
  platform: 'node',
  target: 'node24',
  format: 'cjs',
  define: {
    'import.meta.url': '__officebridgeImportMetaUrl',
    XLL_PROBE_ASSETS: JSON.stringify(assets),
    XLL_PROBE_TOKEN: JSON.stringify(token),
  },
  banner: {
    js: "const __officebridgeImportMetaUrl = require('node:url').pathToFileURL(__filename).href;",
  },
});
const resources = [
  '.generated/index.html',
  '.generated/publish.template.html',
  '.generated/contract.js',
  'taskpane.html',
  'taskpane.js',
  'executor.js',
  '.generated/execution-worker.js',
  'host.js',
  'host-operations.js',
  'styles.css',
  'main.js',
  'ribbon.xml',
];
const blob = join(staging, 'probe.blob');
const config = join(staging, 'sea-config.json');
await writeFile(
  config,
  JSON.stringify({
    main: bundle,
    output: blob,
    disableExperimentalSEAWarning: true,
    useSnapshot: false,
    useCodeCache: false,
    execArgvExtension: 'none',
    assets: {
      ...Object.fromEntries(
        resources.map((name) => [
          'wps/' + name,
          join(root, 'apps/wps-addin', name),
        ]),
      ),
      'licenses/doubao-cli': join(
        root,
        'packages/agents/vendor/doubao-cli/LICENSE',
      ),
    },
  }),
);
run(process.execPath, ['--experimental-sea-config', config]);
const backend = join(staging, 'OfficeBridgeProbe.exe');
await copyFile(process.execPath, backend);
await inject(backend, 'NODE_SEA_BLOB', await readFile(blob), {
  sentinelFuse: 'NODE_SEA_FUSE_fce680ab2cc467b6e072b8b5df1996b2',
});
const sha = createHash('sha256')
  .update(await readFile(backend))
  .digest('hex');
const buildId = sha.slice(0, 16);
await writeFile(
  join(staging, 'build-config.h'),
  `#define XLL_TOKEN L"${token}"\n#define BUILD_ID L"${buildId}"\nstatic const BYTE BACKEND_SHA256[32] = {${sha
    .match(/../g)
    .map((b) => '0x' + b)
    .join(',')}};\n`,
);
const notices = join(root, 'release/THIRD-PARTY-NOTICES.txt');
await stat(notices).catch(() => {
  throw new Error(
    'Run npm run package:win first to prepare the exact runtime notices.',
  );
});
const rc = join(staging, 'payload.rc');
const res = join(staging, 'payload.res');
await writeFile(
  rc,
  `101 RCDATA "${backend.replaceAll('\\', '\\\\')}"\n102 RCDATA "${notices.replaceAll('\\', '\\\\')}"\n`,
);
run(zig, ['rc', '/fo', res, rc]);
const xll = join(release, 'OfficeBridgeProbe.xll');
run(zig, [
  'cc',
  '-target',
  'x86-windows-gnu',
  '-shared',
  '-O2',
  '-I' + staging,
  join(source, 'bootstrap.c'),
  res,
  '-lwinhttp',
  '-ladvapi32',
  '-luser32',
  '-o',
  xll,
]);
await copyFile(notices, join(release, 'THIRD-PARTY-NOTICES.txt'));
const xllSha = createHash('sha256')
  .update(await readFile(xll))
  .digest('hex');
await writeFile(
  join(release, 'SHA256SUMS.txt'),
  xllSha + '  OfficeBridgeProbe.xll\n',
);
await writeFile(
  join(staging, 'build-info.json'),
  JSON.stringify(
    { token, backend, xll, sha256: sha, xllSha256: xllSha, buildId },
    null,
    2,
  ),
);
console.info(
  JSON.stringify({
    xll,
    bytes: (await stat(xll)).size,
    architecture: 'x86',
    sha256: xllSha,
  }),
);
