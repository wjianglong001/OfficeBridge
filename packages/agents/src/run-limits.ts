// Wall time is bounded even when a provider continuously reports real progress.
export const AGENT_RUN_LIMITS = Object.freeze({
  idleMs: 120000,
  maxMs: 600000,
});
