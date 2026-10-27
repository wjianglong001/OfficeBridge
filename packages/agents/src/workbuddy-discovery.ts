import { access, readFile, stat } from 'node:fs/promises';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { dirname, join, resolve } from 'node:path';
import { homedir } from 'node:os';

const execute = promisify(execFile);
export interface WorkBuddyLaunch {
  command: string;
  cli: string;
  env: NodeJS.ProcessEnv;
}

export async function findWorkBuddy(): Promise<WorkBuddyLaunch> {
  const explicit = process.env.OFFICEBRIDGE_WORKBUDDY_PATH;
  const candidates: string[] = [];
  if (explicit) candidates.push(resolve(explicit));
  else {
    for (const base of [
      process.env.LOCALAPPDATA && join(process.env.LOCALAPPDATA, 'Programs'),
      process.env.ProgramFiles,
      process.env['ProgramFiles(x86)'],
    ].filter((x): x is string => !!x)) {
      candidates.push(
        join(base, 'WorkBuddyAI', 'WorkBuddyAI.exe'),
        join(base, 'WorkBuddy', 'WorkBuddy.exe'),
      );
    }
    if (process.platform === 'win32') {
      for (const key of [
        'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall',
        'HKLM\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall',
        'HKLM\\Software\\WOW6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall',
      ]) {
        try {
          const { stdout } = await execute(
            'reg.exe',
            ['query', key, '/s', '/v', 'DisplayIcon'],
            { windowsHide: true, timeout: 5000, maxBuffer: 2_000_000 },
          );
          for (const match of stdout.matchAll(
            /([A-Z]:\\[^\r\n"]*?\\WorkBuddy(?:AI)?\.exe)/gi,
          ))
            candidates.push(match[1]!);
        } catch {
          /* Absent registry hive or no matches. */
        }
      }
    }
  }
  for (let candidate of new Set(candidates)) {
    try {
      if ((await stat(candidate)).isDirectory())
        candidate = join(candidate, 'WorkBuddyAI.exe');
      const root = dirname(candidate);
      const cliRoot = join(root, 'resources', 'app.asar.unpacked', 'cli');
      const cli = join(cliRoot, 'dist', 'codebuddy.js');
      await access(cli);
      const product = JSON.parse(
        await readFile(join(cliRoot, 'product.json'), 'utf8'),
      );
      if (
        typeof product.productName !== 'string' ||
        !/workbuddy/i.test(product.productName)
      )
        continue;
      const env: NodeJS.ProcessEnv = {
        ...process.env,
        ELECTRON_RUN_AS_NODE: '1',
        DISABLE_TELEMETRY: '1',
        CODEBUDDY_CODE_DISABLE_BACKGROUND_TASKS: '1',
      };
      const custom = product.config?.customUserDataDir;
      if (typeof custom === 'string' && /^\.[\w.-]+$/.test(custom)) {
        env.WORKBUDDY_CONFIG_DIR =
          process.env.WORKBUDDY_CONFIG_DIR ?? join(homedir(), custom);
        env.CODEBUDDY_CONFIG_DIR =
          process.env.CODEBUDDY_CONFIG_DIR ?? env.WORKBUDDY_CONFIG_DIR;
      }
      // Verify the package runtime can run CLI code, without opening a desktop window.
      const { stdout } = await execute(candidate, ['--version'], {
        env,
        windowsHide: true,
        timeout: 5000,
      });
      if (!/^v\d+\.[\d.]+\s*$/.test(stdout)) continue;
      return { command: candidate, cli, env };
    } catch {
      /* Try the next genuine desktop installation. */
    }
  }
  throw new Error(
    '未找到可用的 WorkBuddy 桌面运行时。请安装并登录 WorkBuddy，或用 OFFICEBRIDGE_WORKBUDDY_PATH 指定其 exe 路径。',
  );
}
