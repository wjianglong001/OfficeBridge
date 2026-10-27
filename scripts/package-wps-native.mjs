// Build a WPS-native portable folder. This script never starts WPS or an Agent.
import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { createHash, randomBytes, randomUUID } from 'node:crypto';
import {
  lstat,
  mkdir,
  readFile,
  rename,
  stat,
  unlink,
  writeFile,
} from 'node:fs/promises';
import { isAbsolute, join, relative, resolve, sep } from 'node:path';
import { fileURLToPath } from 'node:url';

if (process.platform !== 'win32' || process.arch !== 'x64')
  throw new Error('WPS native packaging requires Windows x64.');
const root = fileURLToPath(new URL('../', import.meta.url));
const source = join(root, 'experiments/xll-bootstrap');
const staging = join(root, '.officebridge/wps-native');
const distribution = join(root, 'release/wps-native');
const args = process.argv.slice(2);
let zig = join(
  root,
  '.officebridge/toolchains/zig-windows-x86_64-0.13.0/zig.exe',
);
for (let i = 0; i < args.length; i++) {
  if (args[i] !== '--zig' || !args[i + 1])
    throw new Error(
      'Usage: node scripts/package-wps-native.mjs [--zig <zig.exe>]',
    );
  zig = resolve(args[++i]);
}
const version = JSON.parse(
  await readFile(join(root, 'package.json'), 'utf8'),
).version;
if (
  typeof version !== 'string' ||
  !/^[0-9A-Za-z][0-9A-Za-z.+-]*$/.test(version)
)
  throw new Error('Invalid package version for the archive filename.');
// Versioned output never replaces a loaded XLL or running portable service.
const release = join(distribution, version, 'OfficeBridge');
const protocolSource = await readFile(
  join(root, 'packages/protocol/src/index.ts'),
  'utf8',
);
const protocolMatch = protocolSource.match(
  /export\s+const\s+PROTOCOL_VERSION\s*=\s*(['"])([^'"\r\n]+)\1\s*;/,
);
if (!protocolMatch) throw new Error('Cannot read the bridge protocol version.');
const protocolVersion = protocolMatch[2];
const inputBackend = join(root, 'release/OfficeBridge.exe');
const inputNotices = join(root, 'release/THIRD-PARTY-NOTICES.txt');

async function regular(path) {
  const entry = await lstat(path);
  if (entry.isSymbolicLink() || !entry.isFile())
    throw new Error('Expected a regular build artifact: ' + path);
}
async function ownDirectory(path) {
  const local = relative(root, path);
  assert.ok(
    local &&
      !isAbsolute(local) &&
      local !== '..' &&
      !local.startsWith('..' + sep),
  );
  let parent = root;
  for (const segment of local.split(sep)) {
    parent = join(parent, segment);
    let entry;
    try {
      entry = await lstat(parent);
    } catch (error) {
      if (error.code !== 'ENOENT') throw error;
      await mkdir(parent);
      entry = await lstat(parent);
    }
    if (entry.isSymbolicLink() || !entry.isDirectory())
      throw new Error(
        'Build directory cannot be a link or non-directory: ' + parent,
      );
  }
}
async function replaceOwnFile(path, body) {
  try {
    await regular(path);
  } catch (error) {
    if (error.code !== 'ENOENT') throw error;
  }
  await writeFile(path, body);
}
for (const path of [inputBackend, inputNotices]) {
  try {
    await regular(path);
  } catch {
    throw new Error(
      'Missing or invalid Windows package. Run npm run package:win first: ' +
        path,
    );
  }
}
await regular(zig).catch(() => {
  throw new Error('Zig compiler not found. Pass --zig <zig.exe>: ' + zig);
});
await ownDirectory(staging);
await ownDirectory(release);

// Node SEA is an AMD64 PE32+ image. Change only its 2-byte subsystem field;
// the native bootstrap pins the SHA-256 of this final GUI image.
const original = await readFile(inputBackend);
if (original.length < 64 || original.toString('ascii', 0, 2) !== 'MZ')
  throw new Error('OfficeBridge.exe is not a PE executable.');
const pe = original.readUInt32LE(0x3c);
if (
  pe > original.length - 24 ||
  original.toString('ascii', pe, pe + 4) !== 'PE\0\0'
)
  throw new Error('Invalid PE signature.');
const optional = pe + 24;
const optionalSize = original.readUInt16LE(pe + 20);
if (
  original.readUInt16LE(pe + 4) !== 0x8664 ||
  optionalSize < 70 ||
  optional + optionalSize > original.length ||
  original.readUInt16LE(optional) !== 0x20b
)
  throw new Error(
    'OfficeBridgeService requires an AMD64 PE32+ Node executable.',
  );
const subsystemOffset = optional + 68;
const oldSubsystem = original.readUInt16LE(subsystemOffset);
if (oldSubsystem !== 2 && oldSubsystem !== 3)
  throw new Error('Unexpected PE subsystem: ' + oldSubsystem);
const gui = Buffer.from(original);
gui.writeUInt16LE(2, subsystemOffset);
assert.deepEqual(
  gui.subarray(0, subsystemOffset),
  original.subarray(0, subsystemOffset),
);
assert.deepEqual(
  gui.subarray(subsystemOffset + 2),
  original.subarray(subsystemOffset + 2),
);
const backend = join(release, 'OfficeBridgeService.exe');
await replaceOwnFile(backend, gui);
const backendSha256 = createHash('sha256').update(gui).digest('hex');
const notices = join(release, 'THIRD-PARTY-NOTICES.txt');
await replaceOwnFile(
  notices,
  (await readFile(inputNotices, 'utf8')) +
    '\n\n----- cJSON v1.7.19 -----\n' +
    (await readFile(join(source, 'vendor/cJSON/LICENSE'), 'utf8')),
);

const bootstrap = join(root, 'apps/wps-bootstrap');
const assets = [
  ['/bootstrap/', 'text/html; charset=utf-8', join(bootstrap, 'index.html')],
  ...[
    ['index.html', 'text/html; charset=utf-8'],
    ['main.js', 'text/javascript'],
    ['ribbon.xml', 'application/xml'],
    ['launcher.html', 'text/html; charset=utf-8'],
    ['launcher.js', 'text/javascript'],
    ['styles.css', 'text/css'],
  ].map(([name, type]) => ['/bootstrap/' + name, type, join(bootstrap, name)]),
  ['/bootstrap/licenses', 'text/plain; charset=utf-8', notices],
];
for (const [, , path] of assets) await regular(path);
const token = randomBytes(32).toString('hex');
const deploy = {
  cmd: 'enable',
  name: 'OfficeBridgeNative',
  addonType: 'et',
  online: 'true',
  url: 'http://127.0.0.1:43192/bootstrap/',
  version: '1.0.0',
  customDomain: '',
};
const invoke = {
  name: 'OfficeBridgeNative',
  function: 'OfficeBridgeNativeOpen',
  info: '{}',
  jsPluginsXml: '',
};
await replaceOwnFile(
  join(staging, 'native-config.h'),
  [
    '#define NATIVE_TOKEN ' + JSON.stringify(token),
    '#define NATIVE_TOKEN_W L' + JSON.stringify(token),
    '#define NATIVE_LISTEN_PORT 43192',
    '#define NATIVE_ADDIN_NAME ' + JSON.stringify(deploy.name),
    '#define NATIVE_ADDIN_URL ' + JSON.stringify(deploy.url),
    '#define NATIVE_ADDIN_VERSION ' + JSON.stringify(deploy.version),
    '#define NATIVE_DEPLOY_JSON ' + JSON.stringify(JSON.stringify(deploy)),
    '#define NATIVE_INVOKE_JSON ' + JSON.stringify(JSON.stringify(invoke)),
    '#define NATIVE_BACKEND_PORT 4318',
    '#define NATIVE_BRIDGE_PROTOCOL_VERSION ' + JSON.stringify(protocolVersion),
    'static const BYTE NATIVE_BACKEND_SHA256[32] = {' +
      backendSha256
        .match(/../g)
        .map((value) => '0x' + value)
        .join(',') +
      '};',
    'static const struct {const char *path;const char *type;int id;} native_assets[] = {' +
      assets
        .map(
          ([url, type], i) =>
            '{' +
            JSON.stringify(url) +
            ',' +
            JSON.stringify(type) +
            ',' +
            (100 + i) +
            '}',
        )
        .join(',') +
      '};',
    '',
  ].join('\n'),
);
const rc = join(staging, 'assets.rc');
const res = join(staging, 'assets.res');
await replaceOwnFile(
  rc,
  '#pragma code_page(65001)\n' +
    assets
      .map(
        ([, , path], i) =>
          `${100 + i} RCDATA "${path.replaceAll('\\', '\\\\')}"`,
      )
      .join('\n') +
    '\n',
);
const run = (parameters) =>
  execFileSync(zig, parameters, {
    cwd: root,
    windowsHide: true,
    stdio: 'inherit',
  });
const xll = join(release, 'OfficeBridge.xll');
try {
  await regular(xll);
} catch (error) {
  if (error.code !== 'ENOENT') throw error;
}
run(['rc', '/fo', res, rc]);
run([
  'cc',
  '-target',
  'x86-windows-gnu',
  '-shared',
  '-O2',
  '-DOFFICEBRIDGE_FULL_BOOTSTRAP=1',
  '-DCJSON_HIDE_SYMBOLS',
  '-DCJSON_NESTING_LIMIT=32',
  '-I' + staging,
  join(source, 'native-server.c'),
  join(source, 'vendor/cJSON/cJSON.c'),
  res,
  '-lwinhttp',
  '-liphlpapi',
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
try {
  await regular(testHost);
} catch (error) {
  if (error.code !== 'ENOENT') throw error;
}
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
await replaceOwnFile(
  join(release, '使用说明.txt'),
  '\ufeff' +
    [
      '办公桥 · WPS 原生加载项（Windows x64 / WPS 表格32位）',
      '',
      '1. 先完整解压压缩包，保持 OfficeBridge 文件夹内文件在一起，不要只复制 XLL。',
      '2. WPS 表格 → 工具 → 加载项 → 浏览，选择本文件夹的 OfficeBridge.xll，勾选并确定。',
      '3. 如出现 WPS 的“信任并安装”，由你确认；取消后不会自动重试。',
      '4. 办公桥将启动本文件夹的 OfficeBridgeService.exe，并在 WPS 内打开完整侧栏。无需单独双击后台或安装 Node.js/npm。',
      '5. 直接打开自己的表格并选中需要的数据，在侧栏对话；没有选区时由 Agent 按需定位读取。',
      '6. 在侧栏选择已安装、已登录的 Codex、WorkBuddy 或豆包（实验）；启动连接检查不会发送模型请求。',
      '7. 分析直接返回；修改先预览，只有点击“应用”后才写入，可按现有能力恢复最近一次修改。',
      '8. 不要删除或移动正在使用的文件夹。WPS 退出不保证独立后台同步退出。',
      '',
      '默认后台仅监听 http://127.0.0.1:4318；引导仅监听 http://127.0.0.1:43192。',
      '端口被其他程序占用、后台身份不匹配，或系统拒绝启动时明确显示原因及错误码；不关闭其他程序，不自动换端口，不自动重试。',
      '此包尚未配置代码签名；WPS 实机安装、后台启动和真实 Agent 操作需单独验收。',
      '发布时保留 THIRD-PARTY-NOTICES.txt。开发者仍可按项目原有 npm 方式运行。',
      '',
    ].join('\r\n'),
);

// Archive only the fixed artifact allowlist; never sweep unrelated old files
// or user workbooks from an existing output folder into the distribution.
const archive = join(distribution, `OfficeBridge-WPS-${version}.zip`);
try {
  await regular(archive);
} catch (error) {
  if (error.code !== 'ENOENT') throw error;
}
const temporaryArchive = join(
  staging,
  'OfficeBridge-WPS-' + randomUUID() + '.zip',
);
const compression = [
  '$ErrorActionPreference = "Stop"',
  'Add-Type -AssemblyName System.IO.Compression',
  'Add-Type -AssemblyName System.IO.Compression.FileSystem',
  '$archive = [System.IO.Compression.ZipFile]::Open($env:OFFICEBRIDGE_NATIVE_ZIP_TARGET, [System.IO.Compression.ZipArchiveMode]::Create)',
  'try {',
  '  foreach ($name in @("OfficeBridge.xll", "OfficeBridgeService.exe", "THIRD-PARTY-NOTICES.txt", "使用说明.txt")) {',
  '    $source = [System.IO.Path]::Combine($env:OFFICEBRIDGE_NATIVE_ZIP_SOURCE, $name)',
  '    [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($archive, $source, ("OfficeBridge/" + $name), [System.IO.Compression.CompressionLevel]::Optimal) | Out-Null',
  '  }',
  '} finally { $archive.Dispose() }',
].join('\n');
try {
  execFileSync(
    join(
      process.env.SystemRoot ?? 'C:\\Windows',
      'System32/WindowsPowerShell/v1.0/powershell.exe',
    ),
    ['-NoProfile', '-NonInteractive', '-Command', compression],
    {
      cwd: root,
      windowsHide: true,
      stdio: 'inherit',
      env: {
        ...process.env,
        OFFICEBRIDGE_NATIVE_ZIP_SOURCE: release,
        OFFICEBRIDGE_NATIVE_ZIP_TARGET: temporaryArchive,
      },
    },
  );
  await rename(temporaryArchive, archive);
} finally {
  await unlink(temporaryArchive).catch((error) => {
    if (error.code !== 'ENOENT') throw error;
  });
}
await replaceOwnFile(
  join(distribution, 'SHA256SUMS.txt'),
  [
    `${sha256}  ${version}/OfficeBridge/OfficeBridge.xll`,
    `${backendSha256}  ${version}/OfficeBridge/OfficeBridgeService.exe`,
    `${createHash('sha256')
      .update(await readFile(archive))
      .digest('hex')}  ${archive.slice(distribution.length + 1)}`,
    '',
  ].join('\n'),
);
await replaceOwnFile(
  join(staging, 'build-info.json'),
  JSON.stringify(
    {
      token,
      xll,
      testHost,
      backend,
      sha256,
      backendSha256,
      archive,
      version,
      protocolVersion,
      mode: 'full-bootstrap',
    },
    null,
    2,
  ),
);
console.info(
  JSON.stringify({
    xll,
    backend,
    archive,
    bytes: (await stat(xll)).size,
    sha256,
    backendSha256,
  }),
);
