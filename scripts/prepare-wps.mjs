import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { build } from 'esbuild';
import { fileURLToPath } from 'node:url';

// Reuse the publisher template shipped with the pinned official wpsjs package.
// Generating these files does not install anything or modify WPS configuration.
const templateUrl = new URL(
  '../node_modules/wpsjs/src/lib/res/publish.html',
  import.meta.url,
);
const destination = new URL('../apps/wps-addin/.generated/', import.meta.url);
let template = await readFile(templateUrl, 'utf8');
for (const placeholder of [
  'PUBLISH_REPLACE_STRING',
  'SERVERID_REPLEASE_STRING',
]) {
  if (!template.includes(placeholder))
    throw new Error(`wpsjs template changed: ${placeholder} is missing`);
}
const addons = [
  {
    name: 'OfficeBridge',
    addonType: 'et',
    online: 'true',
    multiUser: 'false',
    customDomain: '',
    url: '__OFFICEBRIDGE_ADDON_BASE_URL__',
  },
];
template = template
  .replace('PUBLISH_REPLACE_STRING', JSON.stringify(addons))
  .replace('SERVERID_REPLEASE_STRING', 'undefined')
  .replace(/WPS加载项配置/g, 'OfficeBridge 安装与卸载')
  // Do not expose management actions for unrelated user add-ins.
  .replace(
    /curList\.forEach\(function \(element\) \{/g,
    'curList.forEach(function (element) { if (element.name !== "OfficeBridge" || element.addonType !== "et") return;',
  )
  .replace(/<div class="ClearAll"[^>]*>.*?<\/div>/, '')
  .replace(
    'ClearAll.style.maxWidth =',
    'if (ClearAll) ClearAll.style.maxWidth =',
  );
await mkdir(destination, { recursive: true });
await build({
  entryPoints: [
    fileURLToPath(
      new URL('../apps/wps-addin/contract-entry.ts', import.meta.url),
    ),
  ],
  outfile: fileURLToPath(new URL('contract.js', destination)),
  bundle: true,
  format: 'iife',
  platform: 'browser',
  globalName: 'OfficeBridgeContract',
  target: 'es2020',
});
await writeFile(new URL('publish.template.html', destination), template);
await build({
  entryPoints: [
    fileURLToPath(
      new URL('../apps/wps-addin/execution-worker.ts', import.meta.url),
    ),
  ],
  outfile: fileURLToPath(new URL('execution-worker.js', destination)),
  bundle: true,
  format: 'iife',
  platform: 'browser',
  target: 'es2020',
});
// WPS publisher validation requests both ribbon.xml and index.html.
await writeFile(
  new URL('index.html', destination),
  '<!doctype html>\n<html><head><meta charset="UTF-8"><script src="./main.js"></script></head><body></body></html>\n',
);
console.info(
  'WPS installation resources prepared. Run npm run dev, then open /wps/install.html.',
);
