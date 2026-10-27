import { randomBytes } from 'node:crypto';
import {
  discoverDoubaoLaunch,
  startDoubaoConnection,
  stopConfirmedDoubao,
  type DoubaoInstallation,
} from './doubao-launch.js';
import {
  doubaoEndpoint,
  findDoubaoTarget,
  DoubaoDiscoveryError,
} from './doubao-cdp.js';

interface Options {
  discover?: typeof discoverDoubaoLaunch;
  probe?: typeof findDoubaoTarget;
  start?: typeof startDoubaoConnection;
  stop?: typeof stopConfirmedDoubao;
  now?: () => number;
}
export class DoubaoConnectionChangedError extends Error {
  override name = 'DoubaoConnectionChangedError';
}
function identity(installation: DoubaoInstallation) {
  const processes = installation.processes;
  if (
    !processes?.length ||
    processes.some(
      (p) =>
        p.executable.toLowerCase() !== installation.executable.toLowerCase(),
    )
  )
    throw new Error(
      '无法唯一确认豆包安装位置，未重启客户端。请完全退出豆包后在侧栏重试。',
    );
  const main = processes.filter((p) => p.main);
  if (!main.length) throw new Error('豆包正在退出，请稍后重新连接。');
  return JSON.stringify([
    installation.executable.toLowerCase(),
    main.sort((a, b) => a.pid - b.pid),
  ]);
}

/** Connection setup is separate from probes: GET/status never starts or stops apps. */
export function createDoubaoConnection(options: Options = {}) {
  const confirmations = new Map<
    string,
    { identity: string; expires: number }
  >();
  const now = options.now ?? Date.now;
  let working = false;
  return {
    async prepare(restartToken?: string) {
      if (working) throw new Error('正在连接豆包，请稍候。');
      working = true;
      let changed = false;
      try {
        let confirmation;
        if (restartToken !== undefined) {
          confirmation = confirmations.get(restartToken);
          confirmations.delete(restartToken);
          if (!confirmation || confirmation.expires <= now())
            throw new Error('重启确认已过期，请重新连接豆包。');
        }
        const endpoint = doubaoEndpoint();
        try {
          await (options.probe ?? findDoubaoTarget)(endpoint);
          return { status: 'ready' as const, changed: false };
        } catch (error) {
          if (
            !(error instanceof DoubaoDiscoveryError) ||
            error.code !== 'endpoint_unavailable'
          )
            throw error;
        }
        const installation = await (options.discover ?? discoverDoubaoLaunch)();
        if (installation.running) {
          const currentIdentity = identity(installation);
          if (confirmation?.identity !== currentIdentity) {
            for (const [key, value] of confirmations)
              if (value.expires <= now()) confirmations.delete(key);
            if (confirmations.size >= 8)
              confirmations.delete(confirmations.keys().next().value!);
            const nextToken = randomBytes(24).toString('hex');
            confirmations.set(nextToken, {
              identity: currentIdentity,
              expires: now() + 120000,
            });
            return {
              status: 'confirmation_required' as const,
              restartToken: nextToken,
              message: '需要重启豆包才能连接',
              hint: '重启会中断豆包当前任务。确认后自动重启并连接。',
            };
          }
          changed = true;
          await (options.stop ?? stopConfirmedDoubao)(installation);
        }
        changed = true;
        await (options.start ?? startDoubaoConnection)();
        return { status: 'ready' as const, changed: true };
      } catch (error) {
        if (changed)
          throw new DoubaoConnectionChangedError(
            error instanceof Error
              ? error.message
              : '豆包启动失败，请重新连接。',
          );
        throw error;
      } finally {
        working = false;
      }
    },
  };
}
