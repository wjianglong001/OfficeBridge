// OfficeBridge replacement for upstream application discovery.
// Explicit, in-memory context only; never reads a profile or credentials.
import { AsyncLocalStorage } from 'node:async_hooks';
const context = new AsyncLocalStorage();
export function withApp(app, callback) { return context.run(app, callback); }
export function currentApp() {
  const app = context.getStore();
  if (!app) throw new Error('Doubao app context is required');
  return app;
}
export function activeProfile() { return 'Default'; }
export function agentWorkspace() { return ''; }
