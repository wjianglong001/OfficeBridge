import { cp, mkdir } from 'node:fs/promises';

const source = new URL('../apps/wps-addin/', import.meta.url);
const destination = new URL('../dist/apps/wps-addin/', import.meta.url);
await mkdir(destination, { recursive: true });
await cp(source, destination, { recursive: true });
console.info('WPS source assets copied to dist/apps/wps-addin.');
await cp(
  new URL('../packages/agents/vendor/', import.meta.url),
  new URL('../dist/packages/agents/vendor/', import.meta.url),
  { recursive: true },
);
