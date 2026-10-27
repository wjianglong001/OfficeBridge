import { build } from 'esbuild';
import { inject } from 'postject';
import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import {
  copyFile,
  mkdir,
  readFile,
  stat,
  writeFile,
  unlink,
} from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { join } from 'node:path';

if (process.platform !== 'win32' || process.arch !== 'x64') {
  throw new Error(
    'Windows 打包需要 Windows x64 主机，请使用 Windows CI 或本机打包。',
  );
}
if (Number(process.versions.node.split('.')[0]) !== 24) {
  throw new Error(
    'Windows 打包请使用 Node.js 24 LTS；npm 源码运行仍支持 package.json 中的版本范围。',
  );
}
const root = fileURLToPath(new URL('../', import.meta.url));
const staging = join(root, '.officebridge', 'packaging');
const version = JSON.parse(
  await readFile(join(root, 'package.json'), 'utf8'),
).version;
const release = process.argv.includes('--versioned')
  ? join(root, 'release', version)
  : join(root, 'release');
await mkdir(staging, { recursive: true });
await mkdir(release, { recursive: true });
// Remove only the old generated helper when rebuilding an existing output folder.
await unlink(join(release, '启用豆包连接.cmd')).catch((error) => {
  if (error.code !== 'ENOENT') throw error;
});
const bundle = join(staging, 'bridge.cjs');
const blob = join(staging, 'bridge.blob');
const executable = join(release, 'OfficeBridge.exe');

await build({
  absWorkingDir: root,
  entryPoints: ['apps/bridge/src/index.ts'],
  outfile: bundle,
  bundle: true,
  platform: 'node',
  target: 'node24',
  format: 'cjs',
  // This URL is only used by the non-SEA file reader; SEA uses embedded assets.
  define: { 'import.meta.url': '__officebridgeImportMetaUrl' },
  banner: {
    js: "const __officebridgeImportMetaUrl = require('node:url').pathToFileURL(__filename).href;",
  },
  metafile: true,
}).then((result) => {
  const unexpected = Object.keys(result.metafile.inputs).filter((name) =>
    name.includes('node_modules/'),
  );
  if (unexpected.length)
    throw new Error(
      `Unexpected bundled dependencies: ${unexpected.join(', ')}`,
    );
});

const resourceNames = [
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
const config = join(staging, 'sea-config.json');
await writeFile(
  config,
  JSON.stringify(
    {
      main: bundle,
      output: blob,
      disableExperimentalSEAWarning: true,
      useSnapshot: false,
      useCodeCache: false,
      execArgvExtension: 'none',
      assets: {
        ...Object.fromEntries(
          resourceNames.map((name) => [
            `wps/${name}`,
            join(root, 'apps', 'wps-addin', name),
          ]),
        ),
        'licenses/doubao-cli': join(
          root,
          'packages',
          'agents',
          'vendor',
          'doubao-cli',
          'LICENSE',
        ),
      },
    },
    null,
    2,
  ),
);

// Include the exact runtime's complete third-party notices in the release.
const nodeLicenseUrl = `https://raw.githubusercontent.com/nodejs/node/v${process.versions.node}/LICENSE`;
const licenseCache = join(staging, `node-${process.versions.node}-LICENSE`);
let nodeLicense;
try {
  nodeLicense = await readFile(licenseCache, 'utf8');
} catch {
  try {
    const response = await fetch(nodeLicenseUrl, {
      signal: AbortSignal.timeout(10000),
    });
    if (!response.ok)
      throw new Error(`Cannot download Node.js license: ${response.status}`);
    nodeLicense = await response.text();
  } catch {
    // Windows curl also works in environments whose proxy is not used by fetch.
    nodeLicense = execFileSync(
      join(process.env.SystemRoot ?? 'C:\\Windows', 'System32', 'curl.exe'),
      [
        '--silent',
        '--show-error',
        '--fail',
        '--location',
        '--retry',
        '1',
        '--max-time',
        '25',
        nodeLicenseUrl,
      ],
      { encoding: 'utf8', maxBuffer: 2 * 1024 * 1024, windowsHide: true },
    );
  }
  await writeFile(licenseCache, nodeLicense);
}
if (!nodeLicense.includes('Node.js') || !nodeLicense.includes('Permission'))
  throw new Error('Invalid Node.js license response');
const notices = [
  `Node.js ${process.versions.node}\nSource: ${nodeLicenseUrl}\n${nodeLicense}`,
];
for (const [name, license] of [
  ['wpsjs', 'LICENSE'],
  ['esbuild', 'LICENSE.md'],
  ['postject', 'LICENSE'],
]) {
  notices.push(
    `${name}\n${await readFile(join(root, 'node_modules', name, license), 'utf8')}`,
  );
}
notices.push(
  `doubao-cli native protocol subset (MIT, Fullstop000)\nSource revision: b790ea2d127922902f7dc461b4169c28a2438f26\n${await readFile(join(root, 'packages', 'agents', 'vendor', 'doubao-cli', 'LICENSE'), 'utf8')}`,
);
await writeFile(
  join(release, 'THIRD-PARTY-NOTICES.txt'),
  notices.join('\n\n-----\n\n'),
);

execFileSync(process.execPath, ['--experimental-sea-config', config], {
  cwd: root,
  stdio: 'inherit',
});
await copyFile(process.execPath, executable);
await inject(executable, 'NODE_SEA_BLOB', await readFile(blob), {
  sentinelFuse: 'NODE_SEA_FUSE_fce680ab2cc467b6e072b8b5df1996b2',
});
const digest = createHash('sha256')
  .update(await readFile(executable))
  .digest('hex');
await writeFile(
  join(release, 'SHA256SUMS.txt'),
  `${digest}  OfficeBridge.exe\n`,
);
await writeFile(
  join(release, '使用说明.txt'),
  '\ufeff' +
    [
      'OfficeBridge · 办公桥（Windows x64 便携版）',
      '',
      '1. 双击 OfficeBridge.exe，无需安装 Node.js 或 npm。',
      '2. 浏览器自动打开 WPS 安装页，在 OfficeBridge 行点击安装，确认 WPS 信任提示。',
      '3. 保存文件后重开 WPS 表格，从 OfficeBridge 选项卡打开办公桥。',
      '4. 保持 OfficeBridge 窗口运行；关闭窗口或按 Ctrl+C 停止服务。',
      '5. 侧栏右上角可选择 Codex / WorkBuddy / 豆包（实验）；请先在所选 Agent 中登录。切换 Agent 会开启新对话。',
      '',
      '已安装过加载项时，启动程序后即可直接在 WPS 中使用。',
      '默认地址：http://127.0.0.1:4318/wps/install.html',
      '分析功能另需已安装、已登录的 Agent；本程序不会安装 Agent 或代为登录。',
      '在侧栏选择豆包即可连接；未运行时自动启动，需要重启时在侧栏点击“重启并连接”。取消后保持豆包运行，不影响原 Agent 对话。',
      '豆包只开启本次进程的本机接口，默认 127.0.0.1:9225；不改登录信息、快捷方式或全局配置，连接检查不发送模型请求。',
      '支持选择、行列、工作表及工作簿范围；分析直接返回，修改先预览再应用。',
      '发送时附带选区值，Agent 判断是否足够，再按需补充取数；统计在本地执行，修改时才准备恢复信息。',
      '左下角直接显示当前选区；完整选区值随消息发送，AI 判断是否足够，必要时补充读取同文件数据。没有选区也可以直接提问。',
      '支持值、常见公式、数字格式/加粗/填色及新增结果表，可恢复最近一次应用。',
      '不设固定单元格数量上限；单次数据及恢复基线仍受 512 KB 体积限制。新增操作仍需 WPS 实机验收。',
      '端口被其他程序占用会报错；不会终止其他程序或自动更换端口。',
      '可选参数：OfficeBridge.exe --no-open，OfficeBridge.exe --port 4319。',
      '更改端口后需在新地址重新安装/更新加载项。',
      '开发者仍可从源码 npm install、npm run dev；或 npm run build、npm start。',
      '',
      '本构建尚未配置 Windows 代码签名。分发时请一并保留第三方许可与校验文件。',
      `运行时：Node.js ${process.versions.node} / Windows x64`,
      '',
    ].join('\r\n'),
);
console.info(
  `Windows executable: ${executable} (${((await stat(executable)).size / 1024 / 1024).toFixed(1)} MB)`,
);
