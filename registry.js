// registry.js
// Generator registry: scans generators/**/*.js, imports each module's `metadata`
// and exposes a machine-readable catalog. The orchestrator queries this to map
// user intent → the right deterministic generator script.
//
//   node registry.js            → human-readable list
//   node registry.js --json     → JSON catalog (for the orchestrator)

import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const ROOT = path.dirname(fileURLToPath(import.meta.url));
const GENERATORS_DIR = path.join(ROOT, 'generators');

export async function loadRegistry() {
  const entries = [];
  const verticals = fs.readdirSync(GENERATORS_DIR, { withFileTypes: true })
    .filter((d) => d.isDirectory())
    .map((d) => d.name)
    .sort();

  for (const vertical of verticals) {
    const dir = path.join(GENERATORS_DIR, vertical);
    const files = fs.readdirSync(dir).filter((f) => f.endsWith('.js')).sort();
    for (const file of files) {
      const fullPath = path.join(dir, file);
      const mod = await import(pathToFileURL(fullPath).href);
      if (!mod.metadata || typeof mod.generate !== 'function') {
        throw new Error(`${vertical}/${file} must export { metadata, generate }`);
      }
      entries.push({
        ...mod.metadata,
        script: `generators/${vertical}/${file}`,
      });
    }
  }
  return entries;
}

const isMain = process.argv[1]
  && path.resolve(process.argv[1]).toLowerCase() === fileURLToPath(import.meta.url).toLowerCase();

if (isMain) {
  const registry = await loadRegistry();
  if (process.argv.includes('--json')) {
    console.log(JSON.stringify(registry, null, 2));
  } else {
    let currentVertical = '';
    for (const g of registry) {
      if (g.vertical !== currentVertical) {
        currentVertical = g.vertical;
        console.log(`\n=== ${currentVertical.toUpperCase()} ===`);
      }
      console.log(`  ${g.id.padEnd(28)} ${String(g.outputType).padEnd(12)} ${g.name}`);
    }
    console.log(`\nTotal: ${registry.length} generators`);
  }
}
