import { execFileSync } from 'node:child_process';
import { createHash, randomBytes } from 'node:crypto';
import { copyFile, mkdir, readFile, stat, writeFile } from 'node:fs/promises';
import { join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
if (process.platform !== 'win32') throw new Error('Windows required.');
const root = fileURLToPath(new URL('../../', import.meta.url));
const source = join(root, 'experiments/xll-bootstrap');
const staging = join(root, '.officebridge/xll-native');
const release = join(root, 'release/xll-native');
const zig = resolve(
  process.argv[process.argv.indexOf('--zig') + 1] || 'zig.exe',
);
await mkdir(staging, { recursive: true });
await mkdir(release, { recursive: true });
execFileSync(process.execPath, ['scripts/prepare-wps.mjs'], {
  cwd: root,
  stdio: 'inherit',
  windowsHide: true,
});
const index = join(staging, 'index.html');
await writeFile(
  index,
  '<!doctype html><html><head><meta charset="utf-8"><script src="./main.js"></script></head><body>OfficeBridge Native Probe</body></html>',
);
const assets = [
  ['/probe/', 'text/html; charset=utf-8', index],
  ['/probe/index.html', 'text/html; charset=utf-8', index],
  ...[
    ['main.js', 'text/javascript'],
    ['ribbon.xml', 'application/xml'],
    ['taskpane.html', 'text/html; charset=utf-8'],
    ['probe.js', 'text/javascript'],
    ['probe.css', 'text/css'],
  ].map(([name, type]) => ['/probe/' + name, type, join(source, name)]),
  ...[
    ['host.js', 'host.js'],
    ['contract.js', '.generated/contract.js'],
    ['host-operations.js', 'host-operations.js'],
    ['executor.js', 'executor.js'],
  ].map(([url, name]) => [
    '/wps/' + url,
    'text/javascript',
    join(root, 'apps/wps-addin', name),
  ]),
  [
    '/probe/execution-worker.js',
    'text/javascript',
    join(root, 'apps/wps-addin/.generated/execution-worker.js'),
  ],
  [
    '/probe/licenses',
    'text/plain; charset=utf-8',
    join(root, 'node_modules/wpsjs/LICENSE'),
  ],
];
const notices = join(staging, 'THIRD-PARTY-NOTICES.txt');
await writeFile(
  notices,
  (await readFile(join(root, 'node_modules/wpsjs/LICENSE'), 'utf8')) +
    '\n\n--- cJSON v1.7.19 ---\n' +
    (await readFile(join(source, 'vendor/cJSON/LICENSE'), 'utf8')),
);
assets[assets.length - 1][2] = notices;
const token = randomBytes(32).toString('hex');
await writeFile(
  join(staging, 'native-config.h'),
  `#define NATIVE_TOKEN "${token}"\n#define NATIVE_TOKEN_W L"${token}"\nstatic const struct {const char *path;const char *type;int id;} native_assets[] = {${assets.map(([url, type], i) => `{${JSON.stringify(url)},${JSON.stringify(type)},${100 + i}}`).join(',')}};\n`,
);
const rc = join(staging, 'assets.rc');
const res = join(staging, 'assets.res');
await writeFile(
  rc,
  assets
    .map(
      ([, , path], i) => `${100 + i} RCDATA "${path.replaceAll('\\', '\\\\')}"`,
    )
    .join('\n'),
);
const run = (args) =>
  execFileSync(zig, args, { cwd: root, windowsHide: true, stdio: 'inherit' });
run(['rc', '/fo', res, rc]);
const filename = 'OfficeBridge-v6.xll';
const xll = join(release, filename);
run([
  'cc',
  '-target',
  'x86-windows-gnu',
  '-shared',
  '-DCJSON_HIDE_SYMBOLS',
  '-DCJSON_NESTING_LIMIT=32',
  '-O2',
  '-I' + staging,
  join(source, 'native-server.c'),
  join(source, 'vendor/cJSON/cJSON.c'),
  res,
  '-lwinhttp',
  '-lole32',
  '-loleaut32',
  '-luuid',
  '-lcrypt32',
  '-ladvapi32',
  '-luser32',
  '-lws2_32',
  '-o',
  xll,
]);
const testHost = join(staging, 'native-test-host.exe');
run([
  'cc',
  '-target',
  'x86-windows-gnu',
  '-O2',
  '-municode',
  join(source, 'native-test-host.c'),
  '-o',
  testHost,
]);
const sha256 = createHash('sha256')
  .update(await readFile(xll))
  .digest('hex');
await writeFile(
  join(staging, 'build-info.json'),
  JSON.stringify(
    { token, xll, sha256, testHost, mode: 'in-process-thread' },
    null,
    2,
  ),
);
await writeFile(join(release, 'SHA256SUMS.txt'), `${sha256}  ${filename}\n`);
await copyFile(notices, join(release, 'THIRD-PARTY-NOTICES.txt'));
console.info(JSON.stringify({ xll, bytes: (await stat(xll)).size, sha256 }));
