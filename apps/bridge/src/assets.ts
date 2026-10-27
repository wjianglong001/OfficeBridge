import { readFile } from 'node:fs/promises';
import { getAsset, isSea } from 'node:sea';

// SEA reads the same resources from the executable, without extracting files.
export async function readWpsAsset(name: string): Promise<Buffer> {
  if (isSea()) return Buffer.from(getAsset(`wps/${name}`));
  return readFile(new URL(`../../wps-addin/${name}`, import.meta.url));
}
