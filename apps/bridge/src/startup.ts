import { spawn } from 'node:child_process';
import { once } from 'node:events';
import { homedir } from 'node:os';
import { join } from 'node:path';
import { isSea } from 'node:sea';
import { createInterface } from 'node:readline/promises';
import {
  DEFAULT_PORT,
  PROTOCOL_VERSION,
} from '../../../packages/protocol/src/index.js';
import { createBridgeServer } from './server.js';
import { startDoubaoConnection } from '../../../packages/agents/src/doubao-launch.js';

export function parseStartupOptions(args: string[]) {
  let portText = process.env.OFFICEBRIDGE_PORT ?? String(DEFAULT_PORT);
  let open = isSea();
  let help = false;
  for (let index = 0; index < args.length; index++) {
    const arg = args[index];
    if (arg === '--port') {
      const value = args[++index];
      if (!value) throw new Error('--port 后需要指定端口。');
      portText = value;
    } else if (arg === '--open') open = true;
    else if (arg === '--no-open') open = false;
    else if (arg === '--help' || arg === '-h') help = true;
    else throw new Error(`未知启动参数：${arg}`);
  }
  const port = Number(portText);
  if (
    !/^\d+$/.test(portText) ||
    !Number.isInteger(port) ||
    port < 1 ||
    port > 65535
  )
    throw new Error('端口必须是 1 到 65535 之间的整数。');
  return { port, open, help };
}

function openBrowser(url: string) {
  const command =
    process.platform === 'win32'
      ? join(
          process.env.SystemRoot ?? 'C:\\Windows',
          'System32',
          'rundll32.exe',
        )
      : process.platform === 'darwin'
        ? 'open'
        : 'xdg-open';
  const args =
    process.platform === 'win32' ? ['url.dll,FileProtocolHandler', url] : [url];
  const child = spawn(command, args, {
    stdio: 'ignore',
    windowsHide: true,
    detached: true,
  });
  const reportFailure = () =>
    console.error(`无法自动打开浏览器，请手动打开：${url}`);
  child.once('error', reportFailure);
  child.once('exit', (code) => {
    if (code !== null && code !== 0) reportFailure();
  });
  child.unref();
}

async function existingBridge(base: string): Promise<boolean> {
  try {
    const response = await fetch(`${base}/api/v1/health`, {
      signal: AbortSignal.timeout(1500),
      redirect: 'error',
    });
    if (!response.ok) return false;
    const body = (await response.json()) as Record<string, unknown>;
    return (
      body.service === 'office-bridge' &&
      body.status === 'ok' &&
      body.protocolVersion === PROTOCOL_VERSION
    );
  } catch {
    return false;
  }
}

export async function startBridge(args = process.argv.slice(2)) {
  if (args.length === 1 && args[0] === '--start-doubao') {
    console.info((await startDoubaoConnection()).message);
    return;
  }
  const options = parseStartupOptions(args);
  if (options.help) {
    console.info(
      'OfficeBridge · 办公桥\n用法：OfficeBridge.exe [--port 4318] [--no-open | --open]\nnpm start -- [--port 4318] [--open]\nOfficeBridge.exe --start-doubao：显式启动豆包本机连接；运行中的豆包不会被终止。\n默认端口 4318；exe 自动打开 WPS 安装页，npm 默认不打开浏览器。',
    );
    return;
  }
  if (isSea() && !process.env.OFFICEBRIDGE_DATA_DIR) {
    process.env.OFFICEBRIDGE_DATA_DIR = join(
      process.env.LOCALAPPDATA ?? join(homedir(), 'AppData', 'Local'),
      'OfficeBridge',
    );
  }
  const base = `http://127.0.0.1:${options.port}`;
  const installUrl = `${base}/wps/install.html`;
  const server = createBridgeServer();
  try {
    await new Promise<void>((resolve, reject) => {
      server.once('error', reject);
      server.listen(options.port, '127.0.0.1', () => {
        server.removeListener('error', reject);
        resolve();
      });
    });
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code !== 'EADDRINUSE') throw error;
    if (!(await existingBridge(base))) {
      throw new Error(
        `端口 ${options.port} 已被其他程序或不兼容的服务占用。请先检查占用情况，或使用 --port 指定其他端口；更换端口后需重新安装加载项。`,
      );
    }
    console.info(`OfficeBridge 已在运行，本次复用现有服务：${base}`);
    console.info(`WPS 安装页面：${installUrl}`);
    if (options.open) openBrowser(installUrl);
    return;
  }

  console.info('OfficeBridge · 办公桥 已启动');
  console.info(`服务地址：${base}`);
  console.info(`WPS 安装页面：${installUrl}`);
  console.info('首次使用：在安装页面点击安装，然后重新打开 WPS 表格。');
  console.info('保持此窗口运行；关闭窗口或按 Ctrl+C 停止服务。');
  console.info(
    '表格处理需要已安装并登录 Codex 或 WorkBuddy，可在侧栏右上角选择；修改需预览后应用。',
  );
  if (options.open) openBrowser(installUrl);

  let stopping = false;
  const stop = () => {
    if (stopping) return;
    stopping = true;
    server.close();
    server.closeAllConnections();
  };
  const signals = ['SIGINT', 'SIGTERM'] as const;
  for (const signal of signals) process.once(signal, stop);
  server.on('error', (error) => {
    console.error(`OfficeBridge 服务异常：${error.message}`);
    process.exitCode = 1;
    stop();
  });
  await once(server, 'close');
  for (const signal of signals) process.removeListener(signal, stop);
}

export async function reportStartupError(error: unknown) {
  console.error(
    `OfficeBridge 启动失败：${error instanceof Error ? error.message : String(error)}`,
  );
  process.exitCode = 1;
  // A double-clicked executable must leave a startup error readable.
  if (isSea() && process.stdin.isTTY) {
    const terminal = createInterface({
      input: process.stdin,
      output: process.stdout,
    });
    try {
      await terminal.question('按回车键退出…');
    } finally {
      terminal.close();
    }
  }
}
