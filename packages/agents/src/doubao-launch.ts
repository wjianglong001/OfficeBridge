import { execFile, spawn } from 'node:child_process';
import { promisify } from 'node:util';
import { stat } from 'node:fs/promises';
import { basename, join } from 'node:path';
import {
  doubaoEndpoint,
  findDoubaoTarget,
  DoubaoDiscoveryError,
} from './doubao-cdp.js';

const execute = promisify(execFile);
export interface DoubaoProcess {
  pid: number;
  executable: string;
  created: string;
  main: boolean;
}
export interface DoubaoInstallation {
  executable: string;
  running: boolean;
  processes?: DoubaoProcess[];
}
export async function discoverDoubaoLaunch() {
  if (process.platform !== 'win32')
    throw new Error('此豆包连接助手仅支持 Windows。');
  // Read only: retains a main/child flag, never command lines or account data.
  const { stdout } = await execute(
    'powershell.exe',
    [
      '-NoProfile',
      '-NonInteractive',
      '-Command',
      '$ErrorActionPreference="Stop"; ConvertTo-Json -Compress -InputObject @(Get-CimInstance Win32_Process -Filter "Name=\'Doubao.exe\'" | ForEach-Object { $p = Get-Process -Id $_.ProcessId -ErrorAction SilentlyContinue; if ($p) { @{pid=$p.Id; executable=$p.Path; created=$p.StartTime.ToUniversalTime().ToString("o"); main=($_.CommandLine -notmatch "--type=")} } })',
    ],
    { windowsHide: true, timeout: 5000, maxBuffer: 65536 },
  );
  const processes: DoubaoProcess[] = JSON.parse(stdout);
  if (
    !Array.isArray(processes) ||
    processes.some(
      (p) =>
        !p ||
        !Number.isSafeInteger(p.pid) ||
        p.pid <= 0 ||
        typeof p.executable !== 'string' ||
        !p.executable ||
        typeof p.created !== 'string' ||
        !p.created,
    )
  )
    throw new Error('无法确认豆包进程状态，未启动客户端。');
  const candidates: string[] = [];
  if (process.env.OFFICEBRIDGE_DOUBAO_PATH)
    candidates.push(process.env.OFFICEBRIDGE_DOUBAO_PATH);
  candidates.push(...processes.map((p) => p.executable));
  for (const hive of ['HKCU', 'HKLM', 'HKLM\\Software\\WOW6432Node']) {
    const key = hive.includes('WOW6432Node')
      ? hive + '\\Microsoft\\Windows\\CurrentVersion\\App Paths\\Doubao.exe'
      : hive +
        '\\Software\\Microsoft\\Windows\\CurrentVersion\\App Paths\\Doubao.exe';
    try {
      const { stdout } = await execute('reg.exe', ['query', key, '/ve'], {
        windowsHide: true,
        timeout: 3000,
        maxBuffer: 65536,
      });
      const match = /REG_SZ\s+(.+)/.exec(stdout);
      if (match) candidates.push(match[1]!.trim().replace(/^"|"$/g, ''));
    } catch {}
  }
  for (const base of [
    process.env.LOCALAPPDATA,
    process.env.ProgramFiles,
  ].filter((v): v is string => !!v))
    candidates.push(
      join(base, 'Doubao', 'Application', 'Doubao.exe'),
      join(base, 'Doubao', 'app', 'Doubao.exe'),
      join(base, 'Doubao', 'Doubao.exe'),
      join(base, 'Programs', 'Doubao', 'Doubao.exe'),
    );
  for (const path of new Set(candidates)) {
    try {
      if (
        basename(path).toLowerCase() === 'doubao.exe' &&
        (await stat(path)).isFile()
      )
        return { executable: path, running: processes.length > 0, processes };
    } catch {}
  }
  if (processes.length)
    throw new Error('豆包仍在运行，未改动客户端；请先完全退出后再重试。');
  throw new Error(
    '未找到豆包安装位置，请通过 OFFICEBRIDGE_DOUBAO_PATH 指定 Doubao.exe。',
  );
}

/** Only called after a short-lived confirmation for these exact processes. */
export async function stopConfirmedDoubao(installation: DoubaoInstallation) {
  if (process.platform !== 'win32' || !installation.processes?.length)
    throw new Error('无法确认豆包进程，未重启客户端。请重新连接。');
  const encoded = Buffer.from(JSON.stringify(installation.processes)).toString(
    'base64',
  );
  // No user-supplied commands or paths are executed. Validate every process
  // identity before stopping any; a replaced PID is never intentionally killed.
  const script = `$ErrorActionPreference='Stop'
$expected = @([Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('${encoded}')) | ConvertFrom-Json)
$targets = @()
foreach ($item in $expected) {
  $process = Get-Process -Id $item.pid -ErrorAction SilentlyContinue
  if (!$process) { continue }
  if ($process.ProcessName -ne 'Doubao' -or $process.Path -ine $item.executable -or $process.StartTime.ToUniversalTime().ToString('o') -ne $item.created) { throw 'identity_changed' }
  $targets += $process
}
foreach ($target in $targets) { if (!$target.HasExited) { try { $target.Kill() } catch { if (!$target.HasExited) { throw } } } }
foreach ($target in $targets) { if (!$target.WaitForExit(5000)) { throw 'exit_timeout' } }
`;
  try {
    await execute(
      'powershell.exe',
      [
        '-NoProfile',
        '-NonInteractive',
        '-EncodedCommand',
        Buffer.from(script, 'utf16le').toString('base64'),
      ],
      { windowsHide: true, timeout: 15000, maxBuffer: 65536 },
    );
  } catch {
    throw new Error(
      '豆包未能完成退出，未再次启动。请完全退出豆包后在侧栏重试。',
    );
  }
}

interface Options {
  endpoint?: string;
  discover?: () => Promise<DoubaoInstallation>;
  probe?: (endpoint: string) => Promise<unknown>;
  launch?: (executable: string, args: string[]) => Promise<void>;
  wait?: () => Promise<void>;
}
/** Explicit CLI action only; normal startup/probe never calls this. */
export async function startDoubaoConnection(options: Options = {}) {
  const endpoint = doubaoEndpoint(options.endpoint),
    probe = options.probe ?? findDoubaoTarget;
  try {
    await probe(endpoint);
    return {
      started: false,
      message: '豆包本机接口已就绪，请在 WPS 侧栏重新连接。',
    };
  } catch (error) {
    if (
      !(error instanceof DoubaoDiscoveryError) ||
      error.code !== 'endpoint_unavailable'
    )
      throw error;
  }
  const launch = await (options.discover ?? discoverDoubaoLaunch)();
  if (launch.running)
    throw new Error(
      '豆包正在运行，但本机接口未开启。请在 WPS 侧栏选择豆包，按提示确认重启。本命令不会终止客户端。',
    );
  const args = [
    '--remote-debugging-address=127.0.0.1',
    '--remote-debugging-port=' + (new URL(endpoint).port || '80'),
  ];
  await (
    options.launch ??
    (async (executable, args) => {
      const child = spawn(executable, args, {
        stdio: 'ignore',
        windowsHide: true,
        detached: true,
      });
      await new Promise<void>((resolve, reject) => {
        child.once('spawn', resolve);
        child.once('error', () =>
          reject(new Error('无法启动豆包客户端，请检查安装位置。')),
        );
      });
      child.unref();
    })
  )(launch.executable, args);
  for (let attempt = 0; attempt < 20; attempt++) {
    try {
      await probe(endpoint);
      return {
        started: true,
        message:
          '豆包本机接口已启动。请打开已登录的主对话窗口，并在 WPS 侧栏重新连接。',
      };
    } catch (error) {
      if (
        error instanceof DoubaoDiscoveryError &&
        !['endpoint_unavailable', 'no_chat'].includes(error.code)
      )
        throw error;
    }
    await (
      options.wait ??
      (() => new Promise<void>((resolve) => setTimeout(resolve, 500)))
    )();
  }
  throw new Error(
    '已启动豆包，但本机接口或主对话窗口尚未就绪。请等待客户端载入并登录后重新连接；未再次启动客户端。',
  );
}
